import pytest
from app.schemas import CollectorParams, ModelType, SimulationRequest, SystemParams
from app.simulation import incidence_modifier, simulate
from app.solar import erbs_diffuse_fraction, tilted_radiation
from app.weather import parse_tmy2, synthetic_year, to_tmy2

WEATHER = synthetic_year()


def lookup(_):
    return WEATHER


def test_synthetic_year_shape():
    assert len(WEATHER.ghi) == 8760
    assert WEATHER.month[0] == 1 and WEATHER.month[-1] == 12
    assert min(WEATHER.ghi) == 0  # nights
    july = [t for t, m in zip(WEATHER.temp, WEATHER.month) if m == 7]
    jan = [t for t, m in zip(WEATHER.temp, WEATHER.month) if m == 1]
    assert sum(july) / len(july) > sum(jan) / len(jan) + 15


def test_tmy2_roundtrip():
    parsed = parse_tmy2(to_tmy2(WEATHER))
    assert parsed.latitude == pytest.approx(39.9, abs=0.02)
    assert parsed.longitude == pytest.approx(-84.2, abs=0.02)
    assert parsed.time_zone == -5
    assert parsed.hour[:3] == [1, 2, 3]
    assert parsed.ghi[12] == pytest.approx(WEATHER.ghi[12], abs=2)
    assert parsed.temp[100] == pytest.approx(WEATHER.temp[100], abs=0.05)


def test_erbs_and_iam():
    assert erbs_diffuse_fraction(0.1) == pytest.approx(0.991)
    assert erbs_diffuse_fraction(0.9) == 0.165
    assert incidence_modifier(1.0, 1) == 1.0
    assert incidence_modifier(0.0, 1) == 0.0
    assert incidence_modifier(0.5, 2) == pytest.approx(0.83)


def test_tilt_increases_winter_radiation():
    flat = tilted_radiation(WEATHER, 0, 0, 0.2).it
    tilted = tilted_radiation(WEATHER, 60, 0, 0.2).it
    jan = range(0, 31 * 24)
    assert sum(tilted[i] for i in jan) > 1.3 * sum(flat[i] for i in jan)
    # Horizontal surface: tilted radiation equals global horizontal during the day.
    assert sum(flat) == pytest.approx(sum(WEATHER.ghi), rel=0.02)


@pytest.mark.parametrize("model", list(ModelType))
def test_every_model_runs(model):
    res = simulate(SimulationRequest(model=model), lookup)
    assert len(res.monthly) == 12
    assert len(res.daily["It"]) == 365
    assert len(res.hourly["It"]) == 8760
    eff = res.metrics["collector_efficiency"]
    assert eff is not None and 0 < eff < 1
    sf = res.metrics.get("solar_fraction")
    if sf is not None:
        assert 0 <= sf <= 1
    if "Ts" in res.annual:
        assert -20 < min(res.hourly["Ts"]) and max(res.hourly["Ts"]) < 110


def test_more_collectors_more_solar_dhw():
    small = simulate(SimulationRequest(model="dhw", collector=CollectorParams(num_collectors=1)), lookup)
    big = simulate(SimulationRequest(model="dhw", collector=CollectorParams(num_collectors=4)), lookup)
    assert big.metrics["solar_fraction"] > small.metrics["solar_fraction"]


def test_heating_only_in_heating_months():
    res = simulate(SimulationRequest(model="house_heating"), lookup)
    july = res.monthly[6]
    assert july["Qload"] == 0
    assert res.monthly[0]["Qload"] > 0


def test_ua_components_override():
    sys = SystemParams(ua_components={"walls": 500, "windows": 300, "ceiling": 200, "infiltration": 100})
    assert sys.ua == 1100


def test_api_endpoints(client):
    assert client.get("/api/health").json() == {"status": "ok"}
    assert len(client.get("/api/models").json()) == 5
    weather = client.get("/api/weather").json()
    assert weather[0]["id"] == "sample"
    r = client.post("/api/simulate", json={"model": "house_fan_coil_dhw", "include_hourly": False})
    assert r.status_code == 200
    body = r.json()
    assert body["hourly"] is None
    assert set(body["annual"]) >= {"Qfsh", "Qfsw", "Ts"}
    assert client.post("/api/simulate", json={"weather_id": "nope"}).status_code == 404


def test_upload_and_synthetic(client):
    r = client.post("/api/weather/upload", files={"file": ("x.tm2", to_tmy2(WEATHER).encode(), "text/plain")})
    assert r.status_code == 201, r.text
    wid = r.json()["id"]
    assert client.post("/api/simulate", json={"weather_id": wid, "model": "q_useful"}).status_code == 200
    bad = client.post("/api/weather/upload", files={"file": ("x.tm2", b"garbage", "text/plain")})
    assert bad.status_code == 422
    s = client.post("/api/weather/synthetic", json={"name": "Phoenix", "latitude": 33.4, "longitude": -112, "time_zone": -7})
    assert s.status_code == 201
    assert client.delete(f"/api/weather/{s.json()['id']}").status_code == 204
    assert client.delete(f"/api/weather/{wid}").status_code == 204
    assert client.get(f"/api/weather/{wid}").status_code == 404
    assert client.delete("/api/weather/sample").status_code == 400
