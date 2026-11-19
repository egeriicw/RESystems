import pytest

DAYTON = {"name": "Dayton house", "latitude": 39.76, "longitude": -84.19, "floor_area_m2": 180, "ua_house": 1185}


@pytest.fixture
def site(client):
    r = client.post("/api/sites", json=DAYTON)
    assert r.status_code == 201, r.text
    yield r.json()
    client.delete(f"/api/sites/{r.json()['id']}")


def test_sample_weather_survives_roundtrip(client):
    detail = client.get("/api/weather/sample?hourly=true").json()
    assert detail["hours"] == 8760 and len(detail["hourly"]["Ih"]) == 8760
    assert len(detail["monthly"]) == 12
    # Stored series gives the same answer as before it was persisted.
    r = client.post("/api/simulate", json={"model": "dhw", "include_hourly": False}).json()
    assert r["metrics"]["solar_fraction"] == pytest.approx(0.477, abs=0.005)


def test_site_defaults_and_nearest_weather(client, site):
    assert site["time_zone"] == -6 or site["time_zone"] == -5  # estimated from longitude
    assert site["weather_id"] == "sample"
    assert site["weather_distance_km"] == pytest.approx(15.6, abs=2)

    far = client.post("/api/weather/synthetic", json={"name": "Phoenix", "latitude": 33.4, "longitude": -112.0, "time_zone": -7}).json()
    phx = client.post("/api/sites", json={"name": "Phoenix office", "latitude": 33.5, "longitude": -112.1}).json()
    assert phx["weather_id"] == far["id"] and phx["weather_name"] == "Phoenix"

    # Deleting the weather dataset detaches the site rather than deleting it.
    client.delete(f"/api/weather/{far['id']}")
    assert client.get(f"/api/sites/{phx['id']}").json()["weather_id"] is None
    client.delete(f"/api/sites/{phx['id']}")


def test_site_crud_and_radius_search(client, site):
    sid = site["id"]
    r = client.patch(f"/api/sites/{sid}", json={"occupants": 4, "building_type": "house"})
    assert r.json()["occupants"] == 4 and r.json()["name"] == "Dayton house"
    assert client.patch(f"/api/sites/{sid}", json={"name": None}).status_code == 422
    assert client.post("/api/sites", json={"name": "x", "latitude": 95, "longitude": 0}).status_code == 422

    near = client.get("/api/sites", params={"near_lat": 39.9, "near_lon": -84.2, "radius_km": 30}).json()
    assert [s["id"] for s in near] == [sid]
    assert client.get("/api/sites", params={"near_lat": 33.4, "near_lon": -112, "radius_km": 30}).json() == []
    assert client.get("/api/sites/99999").status_code == 404


def test_scenario_run_is_stored_with_hourly(client, site):
    body = {"name": "8 collectors", "site_id": site["id"], "model": "house_fan_coil_dhw",
            "collector": {"num_collectors": 8}}
    s = client.post("/api/scenarios", json=body)
    assert s.status_code == 201, s.text
    sc = s.json()
    assert sc["collector"]["tilt"] == 40 and sc["site_name"] == "Dayton house" and sc["latest_run"] is None

    run = client.post(f"/api/scenarios/{sc['id']}/run")
    assert run.status_code == 201, run.text
    run = run.json()
    assert run["weather_name"].startswith("Dayton")  # weather came from the site
    assert run["metrics"]["solar_fraction"] > 0.1

    stored = client.get(f"/api/runs/{run['id']}").json()
    assert stored["result"]["monthly"] == run["result"]["monthly"]
    assert stored["result"]["hourly"]["Ts"] == run["result"]["hourly"]["Ts"]
    assert client.get(f"/api/runs/{run['id']}?hourly=false").json()["result"]["hourly"] is None

    # Edit the scenario and run again: both runs are kept, newest first.
    body["collector"]["num_collectors"] = 2
    assert client.put(f"/api/scenarios/{sc['id']}", json=body).status_code == 200
    second = client.post(f"/api/scenarios/{sc['id']}/run").json()
    runs = client.get(f"/api/scenarios/{sc['id']}/runs").json()
    assert [r["id"] for r in runs] == [second["id"], run["id"]]
    assert second["metrics"]["solar_fraction"] < run["metrics"]["solar_fraction"]
    assert client.get(f"/api/scenarios?site_id={site['id']}").json()[0]["latest_run"]["id"] == second["id"]
    assert client.get(f"/api/sites/{site['id']}").json()["scenario_count"] == 1

    assert client.delete(f"/api/runs/{run['id']}").status_code == 204
    assert client.get(f"/api/runs/{run['id']}").status_code == 404


def test_deleting_site_cascades(client):
    site = client.post("/api/sites", json=DAYTON).json()
    sc = client.post("/api/scenarios", json={"name": "a", "site_id": site["id"], "model": "dhw"}).json()
    run = client.post(f"/api/scenarios/{sc['id']}/run").json()
    assert client.delete(f"/api/sites/{site['id']}").status_code == 204
    assert client.get(f"/api/scenarios/{sc['id']}").status_code == 404
    assert client.get(f"/api/runs/{run['id']}").status_code == 404


def test_scenario_validation(client):
    assert client.post("/api/scenarios", json={"name": "a", "model": "dhw", "site_id": 99999}).status_code == 404
    assert client.post("/api/scenarios", json={"name": "a", "model": "dhw", "weather_id": "nope"}).status_code == 404
    assert client.post("/api/scenarios", json={"name": "a", "model": "bogus"}).status_code == 422
    # A scenario without a site falls back to the sample weather.
    sc = client.post("/api/scenarios", json={"name": "standalone", "model": "q_useful"}).json()
    assert client.post(f"/api/scenarios/{sc['id']}/run").json()["weather_name"].startswith("Dayton")
    client.delete(f"/api/scenarios/{sc['id']}")
