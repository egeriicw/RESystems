"""The diagram engine against the reference models in simulation.py, plus validation."""
import pytest

from app.diagram import SEP, Diagram, DiagramError, simulate_diagram, template, validate
from app.schemas import CollectorParams, ModelType, SimulationRequest, SystemParams
from app.simulation import simulate
from app.weather import synthetic_year

WEATHER = synthetic_year()

# Reference series -> (diagram node, output) for each original model.
MAPPING = {
    ModelType.q_useful: {"It": "collector:It", "Qu": "collector:Qu"},
    ModelType.house_heating: {"It": "collector:It", "Qu": "collector:Qu", "Qts": "collector:Qts",
                              "Qfs": "emitter:Q", "Qload": "house:Qload", "Qaux": "house:Qaux", "Ts": "tank:Ts"},
    ModelType.dhw: {"It": "collector:It", "Qu": "collector:Qu", "Qts": "collector:Qts", "Qfsw": "dhw:Qsolar",
                    "Qloadw": "dhw:Qload", "Qauxw": "dhw:Qaux", "Ts": "tank:Ts"},
    ModelType.house_fan_coil: {"It": "collector:It", "Qu": "collector:Qu", "Qts": "collector:Qts",
                               "Qfsh": "fan_coil:Q", "Qloadh": "house:Qload", "Qauxh": "house:Qaux", "Ts": "tank:Ts"},
    ModelType.house_fan_coil_dhw: {"It": "collector:It", "Qu": "collector:Qu", "Qts": "collector:Qts",
                                   "Qfsh": "fan_coil:Q", "Qloadh": "house:Qload", "Qauxh": "house:Qaux",
                                   "Qfsw": "dhw:Qsolar", "Qloadw": "dhw:Qload", "Qauxw": "dhw:Qaux", "Ts": "tank:Ts"},
}

CASES = [
    (CollectorParams(), SystemParams()),
    (CollectorParams(num_collectors=8), SystemParams()),
    (CollectorParams(num_collectors=5, tilt=55, azimuth=-20, num_glazings=2, fr_ul=12),
     SystemParams(storage_mass=800, storage_initial_temp=40, collector_hx_effectiveness=0.6,
                  flow_per_collector_gpm=0.4, ua_house=900, fan_coil_effectiveness=0.5, heating_months=[11, 12, 1, 2])),
]


@pytest.mark.parametrize("model", list(ModelType))
@pytest.mark.parametrize("case", range(len(CASES)))
def test_templates_reproduce_reference_models_hourly(model, case):
    col, sys = CASES[case]
    ref = simulate(SimulationRequest(model=model, collector=col, system=sys), lambda _: WEATHER)
    got = simulate_diagram(template(model, col, sys), WEATHER)
    for ref_key, key in MAPPING[model].items():
        assert got.hourly[key] == pytest.approx(ref.hourly[ref_key], rel=1e-9, abs=1e-9), (ref_key, key)
        assert got.monthly[0][key] == pytest.approx(ref.monthly[0][ref_key], rel=1e-9, abs=1e-6)
    for name, value in ref.metrics.items():
        assert got.metrics[name] == pytest.approx(value, abs=1e-4), name
    # Every series covers the year once, and the system totals match the single-collector values.
    assert {len(v) for v in got.hourly.values()} == {8760}
    assert got.hourly["Ta"] == ref.hourly["Ta"]
    assert got.hourly["It"] == pytest.approx(ref.hourly["It"], rel=1e-9, abs=1e-9)
    assert got.annual["Qu"] == pytest.approx(ref.annual["Qu"], rel=1e-9)


def test_two_arrays_two_tanks_are_independent():
    """Two copies of the DHW system side by side give the same results as one alone."""
    single = template(ModelType.dhw)
    double = Diagram(nodes=list(single.nodes), edges=list(single.edges))
    for n in single.nodes:
        double.nodes.append(n.model_copy(update={"id": f"{n.id}-b"}))
    for e in single.edges:
        double.edges.append(e.model_copy(update={"id": f"{e.id}-b", "source": f"{e.source}-b", "target": f"{e.target}-b"}))
    one = simulate_diagram(single, WEATHER)
    two = simulate_diagram(double, WEATHER)
    assert two.hourly[f"tank-b{SEP}Ts"] == one.hourly[f"tank{SEP}Ts"]
    assert two.annual["Qsolar"] == pytest.approx(2 * one.annual["Qsolar"], rel=1e-6)
    assert two.metrics["solar_fraction"] == pytest.approx(one.metrics["solar_fraction"], abs=1e-4)
    # Drawing order doesn't matter.
    rev = Diagram(nodes=list(reversed(double.nodes)), edges=list(reversed(double.edges)))
    assert simulate_diagram(rev, WEATHER).annual["Qsolar"] == pytest.approx(two.annual["Qsolar"], rel=1e-9)


def test_two_arrays_on_one_tank_add_up():
    d = template(ModelType.house_heating, CollectorParams(num_collectors=4))
    d.nodes.append(d.nodes[0].model_copy(update={"id": "collector2"}))
    d.edges.append(d.edges[0].model_copy(update={"id": "w2", "source": "collector2"}))
    eight = simulate_diagram(template(ModelType.house_heating, CollectorParams(num_collectors=8),
                                      SystemParams(flow_per_collector_gpm=0.32)), WEATHER)
    four_plus_four = simulate_diagram(d, WEATHER)
    assert four_plus_four.hourly[f"tank{SEP}Ts"] == pytest.approx(eight.hourly[f"tank{SEP}Ts"], rel=1e-9)
    assert set(four_plus_four.diagram["edges"]) == {e.id for e in d.edges}  # a summary for each wire


def issues(d: dict):
    return [(i.node_id, i.edge_id, i.message) for i in validate(Diagram.model_validate(d))]


def test_validation_messages():
    assert issues({"nodes": [], "edges": []}) == [(None, None, "Add at least one solar collector array")]
    base = template(ModelType.house_fan_coil_dhw).model_dump()

    bad_param = {**base, "nodes": [{**n, "params": {**n["params"], "mass": -5}} if n["id"] == "tank" else n for n in base["nodes"]]}
    assert ("tank", None, "Water mass must be at least 1") in issues(bad_param)

    wrong_kind = {**base, "edges": [*base["edges"], {"id": "x", "source": "tank", "source_handle": "hot_water",
                                                     "target": "fan_coil", "target_handle": "in"}]}
    msgs = issues(wrong_kind)
    assert any(e == "x" and "Can't connect hot water" in m for _, e, m in msgs)

    no_house_wire = {**base, "edges": [e for e in base["edges"] if e["target"] != "house"]}
    assert ("house", None, "House heating load must be wired to a heat emitter") in issues(no_house_wire)
    assert ("fan_coil", None, "Fan coil must be wired to a house") in issues(no_house_wire)

    two_into_dhw = {**base, "nodes": [*base["nodes"], {"id": "t2", "type": "tank"}],
                    "edges": [*base["edges"], {"id": "y", "source": "t2", "source_handle": "hot_water", "target": "dhw", "target_handle": "in"}]}
    assert any(e == "y" and "takes only 1 wire" in m for _, e, m in issues(two_into_dhw))

    chained = {"nodes": [{"id": "c", "type": "collector"}, {"id": "h1", "type": "heat_exchanger"},
                         {"id": "h2", "type": "heat_exchanger"}, {"id": "t", "type": "tank"}],
               "edges": [{"id": "a", "source": "c", "source_handle": "out", "target": "h1", "target_handle": "in"},
                         {"id": "b", "source": "h1", "source_handle": "out", "target": "h2", "target_handle": "in"},
                         {"id": "c2", "source": "h2", "source_handle": "out", "target": "t", "target_handle": "in"}]}
    assert ("b" in [e for _, e, _ in issues(chained)])

    with pytest.raises(DiagramError):
        simulate_diagram(Diagram.model_validate(no_house_wire), WEATHER)

    # An exchanger sized by the hot water draw needs a hot water load on its tank.
    dhw = template(ModelType.dhw).model_dump()
    no_load = {**dhw, "nodes": [n for n in dhw["nodes"] if n["id"] != "dhw"], "edges": [e for e in dhw["edges"] if e["target"] != "dhw"]}
    assert ("hx", None, "Tank-side flow is set to the hot water draw, but its tank has no hot water load") in issues(no_load)


def test_wire_params_are_validated_and_used():
    d = template(ModelType.house_fan_coil).model_dump()
    for e in d["edges"]:
        if e["source"] == "hx":
            e["params"]["flow"] = -1
    assert any(i.edge_id == "w-hx-tank" and i.param == "flow" for i in validate(Diagram.model_validate(d)))
    slow, fast = (template(ModelType.house_heating), template(ModelType.house_heating))
    for dd, f in ((slow, 50.0), (fast, 2000.0)):
        next(e for e in dd.edges if e.source == "tank").params["flow"] = f
    assert simulate_diagram(fast, WEATHER).annual["Qsolar"] > simulate_diagram(slow, WEATHER).annual["Qsolar"]


# ---------------------------------------------------------------- API

def test_components_and_templates_api(client):
    lib = client.get("/api/components").json()
    types = {c["type"]: c for c in lib["components"]}
    assert {"collector", "heat_exchanger", "tank", "fan_coil", "direct_emitter", "house", "dhw_load"} <= set(types)
    assert "electric" in lib["port_kinds"]
    tank_mass = next(p for p in types["tank"]["params"] if p["key"] == "mass")
    assert tank_mass["unit"] == "kg" and tank_mass["default"] == 1500
    tpls = client.get("/api/diagram/templates").json()
    assert [t["id"] for t in tpls] == [m.value for m in ModelType]
    for t in tpls:
        assert client.post("/api/diagram/validate", json={"diagram": t["diagram"]}).json() == {"issues": []}


def test_simulate_api_and_errors(client):
    d = template(ModelType.dhw).model_dump()
    r = client.post("/api/diagram/simulate", json={"diagram": d, "include_hourly": False})
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["model"] == "diagram" and body["hourly"] is None
    assert body["diagram"]["nodes"]["tank"]["Ts"] > 0
    assert body["diagram"]["edges"]["w-tank-dhw"]["value"] == pytest.approx(body["annual"]["dhw:Qsolar"], abs=0.01)
    bad = client.post("/api/diagram/simulate", json={"diagram": {"nodes": [], "edges": []}})
    assert bad.status_code == 422
    assert bad.json()["detail"]["issues"][0]["message"] == "Add at least one solar collector array"
    assert client.post("/api/diagram/simulate", json={"diagram": d, "weather_id": "nope"}).status_code == 404


def test_diagram_scenarios_are_stored_run_and_compared(client):
    d = template(ModelType.house_fan_coil_dhw, CollectorParams(num_collectors=8)).model_dump()
    sc = client.post("/api/scenarios", json={"name": "diagram A", "model": "diagram", "diagram": d})
    assert sc.status_code == 201, sc.text
    sc = sc.json()
    assert sc["diagram"]["nodes"][0]["id"] == "collector"
    run = client.post(f"/api/scenarios/{sc['id']}/run").json()
    assert run["result"]["model"] == "diagram"
    stored = client.get(f"/api/runs/{run['id']}").json()
    assert stored["result"]["diagram"]["edges"] == run["result"]["diagram"]["edges"]
    assert len(stored["result"]["hourly"]["tank:Ts"]) == 8760
    # The same system as a classic scenario gives the same metrics.
    classic = client.post("/api/scenarios", json={"name": "classic", "model": "house_fan_coil_dhw",
                                                  "collector": {"num_collectors": 8}}).json()
    crun = client.post(f"/api/scenarios/{classic['id']}/run").json()
    assert run["metrics"]["solar_fraction"] == pytest.approx(crun["metrics"]["solar_fraction"], abs=1e-4)

    # Drafts can be saved; running them reports the problems.
    draft = client.put(f"/api/scenarios/{sc['id']}", json={"name": "draft", "model": "diagram",
                                                           "diagram": {"nodes": [], "edges": []}})
    assert draft.status_code == 200
    r = client.post(f"/api/scenarios/{sc['id']}/run")
    assert r.status_code == 422 and r.json()["detail"]["issues"]
    assert client.post("/api/scenarios", json={"name": "x", "model": "diagram"}).status_code == 422
    assert client.post("/api/scenarios", json={"name": "x", "model": "dhw", "diagram": d}).status_code == 422
    for s in (sc, classic):
        client.delete(f"/api/scenarios/{s['id']}")
