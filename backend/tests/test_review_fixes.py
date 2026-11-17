"""Regression tests for the code-review findings (one block per finding)."""
import math

import pytest
from sqlalchemy import text

from app.db import get_engine
from app.schemas import CollectorParams, ModelType, SimulationRequest, SystemParams
from app.simulation import simulate
from app.solar import tilted_radiation
from app.weather import WeatherData, synthetic_year

WEATHER = synthetic_year()


def lookup(_):
    return WEATHER


# --- #2: solar to hot water can't exceed the hot water load -------------------

def test_hot_tank_does_not_overdeliver_hot_water():
    sys = SystemParams(storage_initial_temp=80)
    for model in ("dhw", "house_fan_coil_dhw"):
        res = simulate(SimulationRequest(model=model, collector=CollectorParams(num_collectors=8), system=sys), lookup)
        h = res.hourly
        assert all(f <= l + 1e-6 for f, l in zip(h["Qfsw"], h["Qloadw"]))
        assert all(f >= 0 for f in h["Qfsw"])
        assert res.metrics["solar_fraction"] <= 1


# --- #7: hours with GHI >= extraterrestrial, or sun below horizon at midpoint ----

def _one_day(ghi_by_hour: dict[int, float]) -> WeatherData:
    w = WeatherData(name="t", source="test", latitude=39.9, longitude=-84.2, time_zone=-5)
    for hr in range(1, 25):
        w.month.append(1); w.day.append(1); w.hour.append(hr)
        w.ghi.append(ghi_by_hour.get(hr, 0.0)); w.temp.append(0.0); w.dew_point.append(0.0)
    return w


def test_radiation_kept_when_kt_at_least_one_or_sun_below_midpoint():
    # Hour 8 midpoint is before sunrise in January at Dayton; hour 13 GHI far above I0.
    w = _one_day({8: 50.0, 12: 1500.0, 13: 100000.0})
    flat = tilted_radiation(w, 0, 0, 0.2).it
    assert flat == pytest.approx(w.ghi, abs=1e-9)  # horizontal surface: It == Ih every hour
    tilted = tilted_radiation(w, 60, 0, 0.2).it
    assert tilted[7] > 0 and tilted[12] > 0
    assert all(math.isfinite(v) and v >= 0 for v in tilted)


def test_synthetic_horizontal_matches_ghi_exactly():
    flat = tilted_radiation(WEATHER, 0, 0, 0.2).it
    assert flat == pytest.approx(WEATHER.ghi, abs=1e-6)


# --- #9: collector efficiency means the same thing for every storage model ----

@pytest.mark.parametrize("model", [m for m in ModelType if m != ModelType.q_useful])
def test_collector_efficiency_is_collected_over_incident(model):
    col = CollectorParams(num_collectors=8)
    res = simulate(SimulationRequest(model=model, collector=col, include_hourly=False), lookup)
    expected = res.annual["Qts"] / (col.area * res.annual["It"])
    assert res.metrics["collector_efficiency"] == pytest.approx(expected, abs=1e-4)
    assert 0 < res.metrics["system_efficiency"] <= res.metrics["collector_efficiency"] + 1e-4


# --- API findings -------------------------------------------------------------

@pytest.fixture
def phoenix(client):
    w = client.post("/api/weather/synthetic", json={"name": "Phoenix", "latitude": 33.4, "longitude": -112.0, "time_zone": -7}).json()
    yield w
    client.delete(f"/api/weather/{w['id']}")


def test_moving_a_site_reassigns_weather_and_time_zone(client, phoenix):  # #4
    site = client.post("/api/sites", json={"name": "Mover", "latitude": 39.76, "longitude": -84.19}).json()
    assert site["weather_id"] == "sample"
    moved = client.patch(f"/api/sites/{site['id']}", json={"latitude": 33.5, "longitude": -112.1}).json()
    assert moved["weather_id"] == phoenix["id"]
    assert moved["time_zone"] == -7
    # Explicit values in the same request win over the automatic ones.
    back = client.patch(f"/api/sites/{site['id']}", json={"latitude": 39.7, "longitude": -84.2, "weather_id": phoenix["id"], "time_zone": -4}).json()
    assert back["weather_id"] == phoenix["id"] and back["time_zone"] == -4
    # A non-location edit leaves both alone.
    same = client.patch(f"/api/sites/{site['id']}", json={"occupants": 3}).json()
    assert same["weather_id"] == phoenix["id"] and same["time_zone"] == -4
    client.delete(f"/api/sites/{site['id']}")


def test_empty_weather_id_is_rejected(client):  # #6
    site = client.post("/api/sites", json={"name": "x", "latitude": 39.7, "longitude": -84.2}).json()
    assert client.patch(f"/api/sites/{site['id']}", json={"weather_id": ""}).status_code == 404
    assert client.post("/api/sites", json={"name": "y", "latitude": 1, "longitude": 1, "weather_id": ""}).status_code == 404
    assert client.patch(f"/api/sites/{site['id']}", json={"weather_id": None}).json()["weather_id"] is None
    client.delete(f"/api/sites/{site['id']}")


def test_deleted_dataset_is_not_served_from_cache(client):  # #5
    w = client.post("/api/weather/synthetic", json={"name": "Cached"}).json()
    assert client.post("/api/simulate", json={"weather_id": w["id"], "include_hourly": False}).status_code == 200
    # Delete behind this process's back, as another worker would.
    with get_engine().begin() as conn:
        conn.execute(text("DELETE FROM weather_datasets WHERE id = :id"), {"id": w["id"]})
    assert client.post("/api/simulate", json={"weather_id": w["id"], "include_hourly": False}).status_code == 404


def test_weather_cache_is_bounded(client):  # #5
    from app import store
    ids = [client.post("/api/weather/synthetic", json={"name": f"c{i}", "seed": i}).json()["id"] for i in range(store.CACHE_SIZE + 3)]
    for wid in ids:
        client.post("/api/simulate", json={"weather_id": wid, "include_hourly": False})
    assert len(store._weather_cache) <= store.CACHE_SIZE
    for wid in ids:
        client.delete(f"/api/weather/{wid}")


def test_upload_size_limit(client):  # #10
    big = b"x" * (5_000_001)
    assert client.post("/api/weather/upload", files={"file": ("big.tm2", big, "text/plain")}).status_code == 413


def test_scenario_reports_effective_weather(client, phoenix):  # #3
    site = client.post("/api/sites", json={"name": "PHX", "latitude": 33.5, "longitude": -112.1}).json()
    sc = client.post("/api/scenarios", json={"name": "inherits", "site_id": site["id"], "model": "dhw"}).json()
    assert sc["weather_id"] is None
    assert sc["effective_weather_id"] == phoenix["id"]
    standalone = client.post("/api/scenarios", json={"name": "none", "model": "dhw"}).json()
    assert standalone["effective_weather_id"] == "sample"
    client.delete(f"/api/scenarios/{standalone['id']}")
    client.delete(f"/api/sites/{site['id']}")


def test_scenario_listing_shows_latest_run_only(client):  # #8
    sc = client.post("/api/scenarios", json={"name": "many runs", "model": "q_useful"}).json()
    run_ids = [client.post(f"/api/scenarios/{sc['id']}/run").json()["id"] for _ in range(3)]
    listed = next(s for s in client.get("/api/scenarios").json() if s["id"] == sc["id"])
    assert listed["latest_run"]["id"] == run_ids[-1]
    assert listed["latest_run"]["metrics"]["collector_efficiency"] > 0
    assert listed["latest_run"]["weather_name"].startswith("Dayton")
    assert client.get(f"/api/scenarios/{sc['id']}").json()["latest_run"]["id"] == run_ids[-1]
    assert [r["id"] for r in client.get(f"/api/scenarios/{sc['id']}/runs").json()] == run_ids[::-1]
    client.delete(f"/api/scenarios/{sc['id']}")
