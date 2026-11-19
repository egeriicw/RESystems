"""One-line system diagrams: schema, validation and hourly simulation.

A diagram is a set of components (see components.py) joined by wires. The
engine supports the topologies the original program's physics covers:

  * a collector array feeds at most one tank, directly or through one heat
    exchanger; an unconnected collector runs at a fixed inlet temperature;
  * each heat emitter draws from one tank and serves one house;
  * each hot water load draws from one tank;
  * any number of arrays, tanks and loads.

Each hour every component reads the start-of-hour tank temperatures, then
every tank updates once, so results don't depend on drawing order. The five
templates from `template()` reproduce simulation.py's models hour by hour
(tests/test_diagram.py checks this).

Differences from simulation.py, which is kept unchanged as the reference:
  * a fan coil never delivers negative heat. simulation.py lets it go
    negative when its minimum supply temperature is set below the room
    temperature; with the defaults (48.9 C vs 22.22 C) the two agree.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

from pydantic import BaseModel, Field

from .components import COMPONENTS, GPM_TO_KG_PER_HR, ComponentType, Param, Port, wire_params
from .constants import CP_WATER
from .schemas import CollectorParams, ModelType, SeriesInfo, SimulationResult, SystemParams
from .simulation import incidence_modifier
from .solar import tilted_radiation
from .weather import WeatherData

ID_PATTERN = r"^[A-Za-z0-9_-]{1,40}$"
SEP = ":"  # series key separator, e.g. "tank-1:Ts" ("." would clash with chart libraries' nested paths)


class Position(BaseModel):
    x: float = 0
    y: float = 0


class DiagramNode(BaseModel):
    id: str = Field(pattern=ID_PATTERN)
    type: str
    label: str = Field("", max_length=60)
    position: Position = Position()
    params: dict = {}


class DiagramEdge(BaseModel):
    id: str = Field(pattern=ID_PATTERN)
    source: str
    source_handle: str
    target: str
    target_handle: str
    params: dict = {}


class Diagram(BaseModel):
    nodes: list[DiagramNode] = Field(default_factory=list, max_length=80)
    edges: list[DiagramEdge] = Field(default_factory=list, max_length=200)


class Issue(BaseModel):
    message: str
    node_id: str | None = None
    edge_id: str | None = None
    param: str | None = None


class DiagramError(ValueError):
    def __init__(self, issues: list[Issue]):
        super().__init__("; ".join(i.message for i in issues))
        self.issues = issues


# ---------------------------------------------------------------- validation

def _check_param(p: Param, value) -> str | None:
    if p.type in ("number", "integer"):
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
            return "must be a number"
        if p.type == "integer" and value != int(value):
            return "must be a whole number"
        if p.min is not None and value < p.min:
            return f"must be at least {p.min:g}"
        if p.max is not None and value > p.max:
            return f"must be at most {p.max:g}"
    elif p.type == "select":
        if value not in [o["value"] for o in p.options or []]:
            return "is not one of the allowed options"
    elif p.type == "profile":
        if (not isinstance(value, list) or len(value) != 24
                or not all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) and v >= 0 for v in value)):
            return "must be 24 non-negative numbers"
    elif p.type == "months":
        if not isinstance(value, list) or not all(isinstance(v, int) and 1 <= v <= 12 for v in value) or len(set(value)) != len(value):
            return "must be a list of months 1–12"
    return None


def _resolve_params(schema: list[Param], given: dict, issues: list[Issue], **where) -> dict:
    known = {p.key for p in schema}
    for key in given:
        if key not in known:
            issues.append(Issue(message=f"Unknown parameter “{key}”", param=key, **where))
    out = {}
    for p in schema:
        value = given.get(p.key, p.default)
        if (err := _check_param(p, value)) is not None:
            issues.append(Issue(message=f"{p.label} {err}", param=p.key, **where))
        if p.type == "integer" and isinstance(value, float) and value == int(value):
            value = int(value)
        out[p.key] = value
    return out


@dataclass
class Node:
    id: str
    type: str
    label: str
    spec: ComponentType
    params: dict


@dataclass
class Wire:
    id: str
    source: Node
    source_port: Port
    target: Node
    target_port: Port
    params: dict


@dataclass
class Compiled:
    nodes: dict[str, Node]
    wires: list[Wire]
    inputs: dict[tuple[str, str], list[Wire]] = field(default_factory=dict)   # (node, port) -> wires in
    outputs: dict[tuple[str, str], list[Wire]] = field(default_factory=dict)  # (node, port) -> wires out

    def wire_in(self, node: str, port: str) -> Wire | None:
        w = self.inputs.get((node, port), [])
        return w[0] if w else None

    def wire_out(self, node: str, port: str) -> Wire | None:
        w = self.outputs.get((node, port), [])
        return w[0] if w else None


def node_label(n: DiagramNode | Node) -> str:
    return n.label or (COMPONENTS[n.type].title if n.type in COMPONENTS else n.type)


def compile_diagram(d: Diagram) -> Compiled:
    """Validate a diagram; raise DiagramError listing every problem found."""
    issues: list[Issue] = []
    nodes: dict[str, Node] = {}
    for n in d.nodes:
        if n.id in nodes:
            issues.append(Issue(message=f"Duplicate component id “{n.id}”", node_id=n.id))
            continue
        spec = COMPONENTS.get(n.type)
        if spec is None:
            issues.append(Issue(message=f"Unknown component type “{n.type}”", node_id=n.id))
            continue
        nodes[n.id] = Node(n.id, n.type, node_label(n), spec, _resolve_params(spec.params, n.params, issues, node_id=n.id))

    c = Compiled(nodes=nodes, wires=[])
    seen_edges, seen_pairs = set(), set()
    for e in d.edges:
        where = {"edge_id": e.id}
        if e.id in seen_edges:
            issues.append(Issue(message=f"Duplicate wire id “{e.id}”", **where))
            continue
        seen_edges.add(e.id)
        src, tgt = nodes.get(e.source), nodes.get(e.target)
        if src is None or tgt is None:
            if not any(i.node_id in (e.source, e.target) for i in issues):
                issues.append(Issue(message="Wire is connected to a component that doesn't exist", **where))
            continue
        sp, tp = src.spec.port(e.source_handle), tgt.spec.port(e.target_handle)
        if sp is None or tp is None or sp.direction != "out" or tp.direction != "in":
            issues.append(Issue(message="Wire must run from an output to an input", **where))
            continue
        if src.id == tgt.id:
            issues.append(Issue(message="A component can't be wired to itself", **where))
            continue
        if sp.kind != tp.kind:
            issues.append(Issue(message=f"Can't connect {sp.label.lower()} ({sp.kind.replace('_', ' ')}) "
                                        f"to {tp.label.lower()} ({tp.kind.replace('_', ' ')})", **where))
            continue
        pair = (src.id, sp.id, tgt.id, tp.id)
        if pair in seen_pairs:
            issues.append(Issue(message="These ports are already wired together", **where))
            continue
        seen_pairs.add(pair)
        schema = wire_params(src.spec, sp, tgt.spec, tp)
        w = Wire(e.id, src, sp, tgt, tp, _resolve_params(schema, e.params, issues, **where))
        c.wires.append(w)
        c.outputs.setdefault((src.id, sp.id), []).append(w)
        c.inputs.setdefault((tgt.id, tp.id), []).append(w)

    for (nid, pid), ws in [*c.outputs.items(), *c.inputs.items()]:
        port = nodes[nid].spec.port(pid)
        if port.max_connections is not None and len(ws) > port.max_connections:
            for w in ws[port.max_connections:]:
                issues.append(Issue(message=f"{nodes[nid].label}: “{port.label}” takes only "
                                            f"{port.max_connections} wire", edge_id=w.id))

    for n in nodes.values():
        need = {
            "heat_exchanger": [("in", "a collector"), ("out", "a tank")],
            "fan_coil": [("in", "a tank"), ("out", "a house")],
            "direct_emitter": [("in", "a tank"), ("out", "a house")],
            "house": [("in", "a heat emitter")],
            "dhw_load": [("in", "a tank")],
        }.get(n.type, [])
        for pid, what in need:
            connected = c.wire_in(n.id, pid) if n.spec.port(pid).direction == "in" else c.wire_out(n.id, pid)
            if connected is None:
                issues.append(Issue(message=f"{n.label} must be wired to {what}", node_id=n.id))
        if n.type == "heat_exchanger" and (w := c.wire_in(n.id, "in")) and w.source.type != "collector":
            issues.append(Issue(message="A heat exchanger must be fed directly by a collector", edge_id=w.id))
        if n.type == "heat_exchanger" and n.params["tank_side_flow"] == "hot_water_draw" and (out := c.wire_out(n.id, "out")):
            if not any(w.target.type == "dhw_load" for w in c.outputs.get((out.target.id, "hot_water"), [])):
                issues.append(Issue(message="Tank-side flow is set to the hot water draw, but its tank has no hot water load",
                                    node_id=n.id, param="tank_side_flow"))

    if not any(n.type == "collector" for n in nodes.values()):
        issues.append(Issue(message="Add at least one solar collector array"))
    if issues:
        raise DiagramError(issues)
    return c


def validate(d: Diagram) -> list[Issue]:
    try:
        compile_diagram(d)
        return []
    except DiagramError as exc:
        return exc.issues


# ---------------------------------------------------------------- simulation

def _aggregate(hourly: dict[str, list[float]], kinds: dict[str, str], w: WeatherData):
    """Daily totals (means for temperatures), monthly averages of daily values, annual average."""
    n_days = len(w.ghi) // 24
    daily = {}
    for key, values in hourly.items():
        is_temp = kinds[key] == "temperature"
        daily[key] = [(sum(values[d * 24:(d + 1) * 24]) / 24 if is_temp else sum(values[d * 24:(d + 1) * 24]))
                      for d in range(n_days)]
    monthly = []
    for m in range(1, 13):
        idx = [d for d in range(n_days) if w.month[d * 24] == m]
        row = {"month": m}
        for key, days in daily.items():
            row[key] = sum(days[d] for d in idx) / len(idx) if idx else 0.0
        monthly.append(row)
    annual = {key: sum(days) / len(days) for key, days in daily.items()}
    return daily, monthly, annual


def simulate_diagram(d: Diagram, w: WeatherData, include_hourly: bool = True) -> SimulationResult:
    c = compile_diagram(d)
    cp = CP_WATER
    by_type = lambda t: [n for n in c.nodes.values() if n.type == t]  # noqa: E731

    tanks = {n.id: {"node": n, "ts": n.params["initial_temp"]} for n in by_type("tank")}
    collectors = []
    for n in by_type("collector"):
        p = n.params
        col = {"node": n, "area": p["num_collectors"] * p["area_per_collector"], "mode": "fixed", "tank": None, "hx": None}
        if (wire := c.wire_out(n.id, "out")) is not None:
            if wire.target.type == "tank":
                col |= {"mode": "direct", "tank": wire.target.id}
            else:  # heat exchanger
                hx = wire.target
                hx_out = c.wire_out(hx.id, "out")
                col |= {"mode": "hx", "tank": hx_out.target.id, "hx": hx, "loop_flow": hx_out.params["flow"]}
                col["tc1"] = tanks[col["tank"]]["node"].params["initial_temp"]
        solar = tilted_radiation(w, p["tilt"], p["azimuth"], p["ground_reflectance"])
        col["solar"] = solar
        collectors.append(col)

    houses = []
    for n in by_type("house"):
        emitter = c.wire_in(n.id, "in").source
        supply = c.wire_in(emitter.id, "in")
        houses.append({"node": n, "emitter": emitter, "tank": supply.source.id,
                       "flow": supply.params.get("flow")})
    dhws = [{"node": n, "tank": c.wire_in(n.id, "in").source.id} for n in by_type("dhw_load")]

    # Output series: per component, then system totals.
    series: list[SeriesInfo] = []
    hourly: dict[str, list[float]] = {}

    def add(key: str, label: str, kind: str, unit: str) -> list[float]:
        series.append(SeriesInfo(key=key, label=label, kind=kind, unit=unit))
        hourly[key] = []
        return hourly[key]

    total_keys = [("Ta", "Ambient temperature", "temperature", "C"),
                  ("It", "Radiation on collectors (area-weighted)", "irradiance", "kJ/day-m2"),
                  ("Qu", "Collector useful gain", "energy", "kJ/day")]
    if tanks:
        total_keys.append(("Qts", "Solar to storage", "energy", "kJ/day"))
    if houses or dhws:
        total_keys += [("Qsolar", "Solar delivered to loads", "energy", "kJ/day"),
                       ("Qload", "Total load", "energy", "kJ/day"),
                       ("Qaux", "Auxiliary heat", "energy", "kJ/day")]
    totals = {k: add(k, label, kind, unit) for k, label, kind, unit in total_keys}
    out: dict[str, dict[str, list[float]]] = {}
    for n in c.nodes.values():
        out[n.id] = {o.key: add(f"{n.id}{SEP}{o.key}", f"{n.label} – {o.label}", o.kind, o.unit) for o in n.spec.outputs}

    total_area = sum(col["area"] for col in collectors)
    for j in range(len(w.ghi)):
        ta, mo = w.temp[j], w.month[j]
        mwater_hour = (w.hour[j] - 1) % 24
        ts0 = {tid: t["ts"] for tid, t in tanks.items()}  # start-of-hour state
        tank_in = {tid: 0.0 for tid in tanks}
        tank_out = {tid: 0.0 for tid in tanks}
        draws = {tid: 0.0 for tid in tanks}
        for dh in dhws:
            draws[dh["tank"]] += dh["node"].params["profile"][mwater_hour]
        sums = {k: 0.0 for k in totals if k not in ("Ta", "It")}
        it_weighted = 0.0

        for col in collectors:
            n, p = col["node"], col["node"].params
            it = col["solar"].it[j]
            kta = incidence_modifier(col["solar"].cos_theta[j], p["num_glazings"])
            absorbed = p["fr_ta"] * kta * it
            area, frul = col["area"], p["fr_ul"]
            qts = 0.0
            if col["mode"] == "fixed":
                qu = max(area * (absorbed - frul * (p["inlet_temp"] - ta)), 0.0)
            elif col["mode"] == "direct":
                ts = ts0[col["tank"]]
                raw = area * (absorbed - frul * (ts - ta))
                qts = raw if raw > 0 and ts <= tanks[col["tank"]]["node"].params["max_temp"] else 0.0
                qu = max(raw, 0.0)
            else:
                ts = ts0[col["tank"]]
                hx = col["hx"].params
                qu = max(area * (absorbed - frul * (col["tc1"] - ta)), 0.0)
                flow = col["loop_flow"] if hx["tank_side_flow"] == "loop" else draws[col["tank"]]
                e = hx["effectiveness"]
                denom = e * flow * cp + (1 - e) * area * frul
                if denom > 0:
                    col["tc1"] = ((1 - e) * area * (absorbed + frul * ta) + e * flow * cp * ts) / denom
                qts = qu if qu > 0 and ts <= tanks[col["tank"]]["node"].params["max_temp"] else 0.0
                out[col["hx"].id]["Q"].append(qts)
            if col["tank"] is not None:
                tank_in[col["tank"]] += qts
            out[n.id]["It"].append(it)
            out[n.id]["Qu"].append(qu)
            out[n.id]["Qts"].append(qts)
            it_weighted += area * it
            sums["Qu"] += qu
            if "Qts" in sums:
                sums["Qts"] += qts

        for h in houses:
            hp = h["node"].params
            ts = ts0[h["tank"]]
            qload = hp["ua"] * (hp["room_temp"] - ta) - hp["internal_gains"]
            qload = qload if qload > 0 and mo in hp["heating_months"] else 0.0
            em = h["emitter"]
            q = 0.0
            if em.type == "fan_coil":
                if qload > 0 and ts > em.params["min_supply_temp"]:
                    q = max(min(em.params["effectiveness"] * em.params["capacitance"] * (ts - hp["room_temp"]), qload), 0.0)
            elif ts > hp["room_temp"] and qload > 0:
                q = min(h["flow"] * cp * (ts - hp["room_temp"]), qload)
            tank_out[h["tank"]] += q
            out[em.id]["Q"].append(q)
            out[h["node"].id]["Qload"].append(qload)
            out[h["node"].id]["Qsolar"].append(q)
            out[h["node"].id]["Qaux"].append(max(qload - q, 0.0))
            sums["Qsolar"] += q
            sums["Qload"] += qload
            sums["Qaux"] += max(qload - q, 0.0)

        for dh in dhws:
            p = dh["node"].params
            m = p["profile"][mwater_hour]
            ts = ts0[dh["tank"]]
            qload = max(m * cp * (p["supply_temp"] - p["ground_water_temp"]), 0.0)
            qsolar = min(max(m * cp * (ts - p["ground_water_temp"]), 0.0), qload)
            tank_out[dh["tank"]] += qsolar
            o = out[dh["node"].id]
            o["Qload"].append(qload)
            o["Qsolar"].append(qsolar)
            o["Qaux"].append(qload - qsolar)
            sums["Qsolar"] += qsolar
            sums["Qload"] += qload
            sums["Qaux"] += qload - qsolar

        for tid, t in tanks.items():
            t["ts"] += (tank_in[tid] - tank_out[tid]) / (t["node"].params["mass"] * cp)
            out[tid]["Ts"].append(t["ts"])
            out[tid]["Qin"].append(tank_in[tid])
            out[tid]["Qout"].append(tank_out[tid])

        totals["Ta"].append(ta)
        totals["It"].append(it_weighted / total_area if total_area > 0 else 0.0)
        for k, v in sums.items():
            totals[k].append(v)

    kinds = {s.key: s.kind for s in series}
    daily, monthly, annual = _aggregate(hourly, kinds, w)

    incident = sum(col["area"] * annual[f"{col['node'].id}{SEP}It"] for col in collectors)
    ratio = lambda a, b: a / b if b > 0 else None  # noqa: E731
    if tanks:
        metrics = {"collector_efficiency": ratio(annual["Qts"], incident)}
    else:
        metrics = {"collector_efficiency": ratio(annual["Qu"], incident)}
    if houses or dhws:
        metrics |= {"system_efficiency": ratio(annual["Qsolar"], incident),
                    "solar_fraction": ratio(annual["Qsolar"], annual["Qload"])}

    nodes_summary = {nid: {k: round(annual[f"{nid}{SEP}{k}"], 2) for k in o} for nid, o in out.items()}
    carried = {
        "collector": lambda wr: ("Delivered to storage", annual[f"{wr.source.id}{SEP}Qts"], "kJ/day"),
        "heat_exchanger": lambda wr: ("Heat transferred", annual[f"{wr.source.id}{SEP}Q"], "kJ/day"),
        "fan_coil": lambda wr: ("Heat delivered", annual[f"{wr.source.id}{SEP}Q"], "kJ/day"),
        "direct_emitter": lambda wr: ("Heat delivered", annual[f"{wr.source.id}{SEP}Q"], "kJ/day"),
    }
    edges_summary = {}
    for wr in c.wires:
        if wr.source.type == "tank":
            key = f"{wr.target.id}{SEP}" + ("Qsolar" if wr.target.type == "dhw_load" else "Q")
            label, value, unit = "Delivered from tank", annual[key], "kJ/day"
        else:
            label, value, unit = carried[wr.source.type](wr)
        edges_summary[wr.id] = {"label": label, "value": round(value, 2), "unit": unit}

    rnd = lambda vals, digits: [round(v, digits) for v in vals]  # noqa: E731
    return SimulationResult(
        model="diagram",
        title="One-line diagram",
        weather={"name": w.name, "source": w.source, "latitude": w.latitude,
                 "longitude": w.longitude, "time_zone": w.time_zone},
        inputs={"diagram": d.model_dump(), "collector_area_m2": total_area, "house_ua": None},
        series=series,
        monthly=[{k: (round(v, 2) if isinstance(v, float) else v) for k, v in r.items()} for r in monthly],
        annual={k: round(v, 2) for k, v in annual.items()},
        metrics={k: (round(v, 4) if v is not None else None) for k, v in metrics.items()},
        daily={k: rnd(v, 2) for k, v in daily.items()},
        hourly={k: rnd(v, 1) for k, v in hourly.items()} if include_hourly else None,
        notes=["Weather is a synthetic typical year; upload a TMY2 file for site-specific results."] if w.source == "synthetic" else [],
        diagram={"nodes": nodes_summary, "edges": edges_summary},
    )


# ---------------------------------------------------------------- templates

def template(model: ModelType, collector: CollectorParams | None = None, system: SystemParams | None = None) -> Diagram:
    """The diagram equivalent of one of the original program's five models."""
    col, sys = collector or CollectorParams(), system or SystemParams()
    loop_flow = col.num_collectors * sys.flow_per_collector_gpm * GPM_TO_KG_PER_HR
    nodes: list[DiagramNode] = []
    edges: list[DiagramEdge] = []

    def node(nid: str, ntype: str, x: float, y: float, **params) -> None:
        nodes.append(DiagramNode(id=nid, type=ntype, position=Position(x=x, y=y), params=params))

    def wire(src: str, sh: str, tgt: str, th: str, **params) -> None:
        edges.append(DiagramEdge(id=f"w-{src}-{tgt}", source=src, source_handle=sh, target=tgt, target_handle=th, params=params))

    node("collector", "collector", 0, 120, num_collectors=col.num_collectors, area_per_collector=col.area_per_collector,
         tilt=col.tilt, azimuth=col.azimuth, ground_reflectance=col.ground_reflectance, fr_ta=col.fr_ta,
         fr_ul=col.fr_ul, num_glazings=col.num_glazings, inlet_temp=sys.collector_inlet_temp)
    if model == ModelType.q_useful:
        return Diagram(nodes=nodes, edges=edges)

    tank = dict(mass=sys.storage_mass, initial_temp=sys.storage_initial_temp, max_temp=sys.storage_max_temp)
    house = dict(ua=sys.ua, internal_gains=sys.internal_gains, room_temp=sys.room_temp,
                 heating_months=list(sys.heating_months))
    dhw = dict(ground_water_temp=sys.ground_water_temp, supply_temp=sys.hot_water_supply_temp,
               profile=list(sys.hot_water_profile))

    if model == ModelType.house_heating:
        node("tank", "tank", 320, 120, **tank)
        node("emitter", "direct_emitter", 600, 40)
        node("house", "house", 880, 40, **house)
        wire("collector", "out", "tank", "in")
        wire("tank", "heating", "emitter", "in", flow=loop_flow)
        wire("emitter", "out", "house", "in")
        return Diagram(nodes=nodes, edges=edges)

    node("hx", "heat_exchanger", 280, 120, effectiveness=sys.collector_hx_effectiveness,
         tank_side_flow="hot_water_draw" if model == ModelType.dhw else "loop")
    node("tank", "tank", 560, 120, **tank)
    wire("collector", "out", "hx", "in")
    wire("hx", "out", "tank", "in", flow=loop_flow)
    if model in (ModelType.house_fan_coil, ModelType.house_fan_coil_dhw):
        node("fan_coil", "fan_coil", 840, 20, effectiveness=sys.fan_coil_effectiveness,
             capacitance=sys.fan_coil_capacitance, min_supply_temp=sys.supply_air_temp)
        node("house", "house", 1120, 20, **house)
        wire("tank", "heating", "fan_coil", "in")
        wire("fan_coil", "out", "house", "in")
    if model in (ModelType.dhw, ModelType.house_fan_coil_dhw):
        node("dhw", "dhw_load", 840, 240, **dhw)
        wire("tank", "hot_water", "dhw", "in")
    return Diagram(nodes=nodes, edges=edges)


TEMPLATE_TITLES = {
    ModelType.q_useful: "Collector only",
    ModelType.house_heating: "House heating, direct supply",
    ModelType.dhw: "Domestic hot water",
    ModelType.house_fan_coil: "House heating with fan coil",
    ModelType.house_fan_coil_dhw: "Fan coil and hot water",
}


def templates() -> list[dict]:
    """Starting diagrams. House models use 8 collectors, as the original program's examples did."""
    out = []
    for m in ModelType:
        n = 8 if m.value.startswith("house") else 2
        out.append({"id": m.value, "title": TEMPLATE_TITLES[m],
                    "diagram": template(m, CollectorParams(num_collectors=n)).model_dump()})
    return out
