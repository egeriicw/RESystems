# Renewable Energy Systems

A web port of the Visual Basic 6 **Renewable Energy Systems** solar-thermal simulator
(original source in [`Resources/`](Resources)). It has a FastAPI backend that runs the
hourly simulations and a React frontend for inputs and results.

```
backend/             FastAPI, simulation engine (pure Python), SQLAlchemy + Alembic
frontend/            React + TypeScript (Vite), Recharts, Leaflet
docker-compose.yml   Postgres 16 with TimescaleDB and PostGIS
Resources/           Original VB6 forms and modules
```

## System diagram

The **System Diagram** page is a one-line diagram editor. Pick components from the library,
wire them together, and the system is simulated once everything is connected. Click any
component or wire to open its box, edit its parameters, and save to rerun. Wires show the
energy they carried in the last run; components show a key result.

| Component | Ports | Parameters |
|---|---|---|
| Solar collector array | collector loop out | number and area of collectors, tilt, azimuth, ground reflectance, FR(τα), FR·UL, glazings |
| Heat exchanger | collector loop in → out | effectiveness, tank-side flow (loop or hot-water draw) |
| Storage tank | collector loop in; heating supply and hot water out | water mass, initial and maximum temperature |
| Fan coil | heating water in → heat out | effectiveness, capacitance rate, minimum tank temperature |
| Direct supply emitter | heating water in → heat out | (supply flow is set on its wire) |
| House heating load | heat in | UA, internal gains, set point, heating months |
| Hot water load | hot water in | mains and delivery temperature, hourly draw profile |

Wiring rules: ports join only to ports of the same kind (colour). A collector feeds at most
one tank, directly or through one heat exchanger; an unconnected collector runs at a fixed
inlet temperature. Each emitter serves one house from one tank; each hot water load draws
from one tank. Any number of arrays, tanks and loads can be combined. Problems are listed
by component or wire, and the diagram runs only when there are none.

Wires carry parameters only where the physics uses them: the flow rate on the wire from a
heat exchanger to its tank, and on the wire from a tank to a direct supply emitter.

The five templates rebuild the original program's models, and `tests/test_diagram.py`
checks that they reproduce `simulation.py` hour by hour. Diagrams are saved as scenarios
(model `diagram`), so they get run history and Compare like any other scenario.

Only thermal components exist so far, matching what the VB program modelled. Ports have a
kind, and `electric` is reserved for PV, battery and EV components.

## Quick start

Database (Docker):

```bash
docker compose up -d                 # Postgres + TimescaleDB + PostGIS on localhost:5432
```

Backend (Python 3.11+):

```bash
cd backend
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/uvicorn app.main:app --reload --port 8000
.venv/bin/python -m pytest          # tests (uses a separate resystems_test database)
```

On startup the backend applies migrations and seeds the sample weather year. Set
`AUTO_MIGRATE=0` to run `alembic upgrade head` yourself. `DATABASE_URL` overrides the
default `postgresql+psycopg://resystems:resystems@localhost:5432/resystems`.

Interactive API docs: http://localhost:8000/docs

Frontend (Node 18+):

```bash
cd frontend
npm install
npm run dev                          # http://localhost:5173, proxies /api to :8000
```

## What was ported

| VB6 | Python |
|---|---|
| `REMain.LoadTMY2_Click`: TMY2 reader | `weather.parse_tmy2` (same fixed columns) |
| `REMainCommon.SolDat`: solar geometry, Erbs split, HDKR tilted radiation | `solar.tilted_radiation` |
| `REMain.CalcQUseful` | model `q_useful` |
| `REMain.CalcSimHouse` | model `house_heating` |
| `REMain.CalcDHW` | model `dhw` |
| `REMain.CalcHouseWFan` | model `house_fan_coil` |
| `REMain.CalcHouseCombo` | model `house_fan_coil_dhw` |
| `System` form, Collectors dialog | `SystemParams`, `CollectorParams` (defaults from the commented values in the VB code) |
| `REMainData` monthly printout | `/api/simulate` monthly table plus daily and hourly series |

### Differences from the VB code

These are latent bugs in the original, fixed in the port (see the `app/simulation.py` docstring):

- Daily arrays were never indexed by day (`n` was never set). They are now computed per day.
- `If Quaxw < 0` / `If QuaxH < 0` typos meant auxiliary energy was never floored at zero.
- The Erbs correlation's `kT > 0.8` branch was unreachable because of an `Or` in the condition.
- Q-useful collector efficiency didn't divide by collector area.
- The incidence-angle modifier divided by zero when the sun was behind the collector.
- `b0` now depends on the glazing count in every model. Two models had hard-coded it to −0.1.
- Hours where measured radiation was at or above the extraterrestrial value were dropped (counted
  as zero). The clearness index is now capped at 1 instead. Sunrise and sunset hours with measured
  radiation but the sun below the horizon at mid-hour are treated as diffuse.
- In diagrams, a fan coil never delivers negative heat. `simulation.py` lets it go negative if its
  minimum supply temperature is set below the room temperature; with the defaults they agree.
- Solar heat to hot water is limited to the hot-water load (a tempering valve mixes in cold water).
  The VB code let a tank hotter than the supply temperature deliver more than the load, which
  drained the tank too fast and could push the solar fraction above 1.

### Results reported

- **Collector efficiency**: energy collected into storage ÷ radiation on the collector, for every
  model. The VB fan-coil models used energy delivered to the load under this name.
- **System efficiency**: solar energy delivered to the loads ÷ radiation on the collector (the VB
  fan-coil figure). Not shown for runs stored before it was added.
- **Solar fraction**: solar energy delivered ÷ load.

### Weather data

The original reads `c:/Engineering Software/TMY2/Dayton.tm2`, which is not in the repo.
The app starts with a deterministic **synthetic** Dayton-like year so it works out of the box.
For real results, upload an NREL TMY2 file on the Weather page or `POST /api/weather/upload`.
You can also generate a synthetic year for any latitude and longitude.

## API

| Method | Path | Purpose |
|---|---|---|
| GET | `/api/models` | The five models and their output series |
| GET | `/api/defaults` | Default collector and system parameters |
| GET | `/api/weather` | Weather datasets with monthly summaries |
| POST | `/api/weather/upload` | Upload a `.tm2` file (multipart `file`) |
| POST | `/api/weather/synthetic` | Generate a synthetic year |
| DELETE | `/api/weather/{id}` | Remove a dataset |
| POST | `/api/simulate` | `{weather_id, model, collector, system, include_hourly}` → monthly, annual, daily, hourly, metrics. Nothing is stored. |
| GET/POST | `/api/sites` | List or create sites. `?near_lat=&near_lon=&radius_km=` searches by distance. |
| GET/PATCH/DELETE | `/api/sites/{id}` | Read, edit or delete a site (deleting removes its scenarios and runs) |
| GET/POST | `/api/scenarios` | List (optionally `?site_id=`) or create saved scenarios. Each includes `effective_weather_id`: its own weather, else its site's. |
| GET/PUT/DELETE | `/api/scenarios/{id}` | Read, replace or delete a scenario |
| POST | `/api/scenarios/{id}/run` | Simulate and store the run, including hourly series |
| GET | `/api/scenarios/{id}/runs` | Run history, newest first |
| GET/DELETE | `/api/runs/{id}` | A stored run (`?hourly=false` to skip the hourly series) |
| GET | `/api/components` | Component library: types, ports, parameter schemas with units and defaults, outputs |
| GET | `/api/diagram/templates` | The five original models as diagrams |
| POST | `/api/diagram/validate` | `{diagram}` → `{issues: [{message, node_id, edge_id, param}]}` |
| POST | `/api/diagram/simulate` | `{weather_id, diagram, include_hourly}` → result; 422 with `issues` if incomplete. Nothing is stored. |

Units follow the original: energy in kJ (hourly kJ/hr, daily and monthly values as average kJ/day),
radiation per m² of collector, temperatures in °C.

## Data model

| Table | Contents |
|---|---|
| `weather_datasets` | Name, source, location (PostGIS point), monthly summary |
| `weather_hourly` | 8,760 hourly records per dataset (Timescale hypertable) |
| `sites` | Name, location (PostGIS point), time zone, building features, assigned weather dataset |
| `scenarios` | Model type plus collector and system inputs (or a diagram), optionally tied to a site |
| `runs` | Stored result of running a scenario (monthly, daily, annual, metrics) |
| `run_hourly` | Hourly series of each run (Timescale hypertable) |

A new site is assigned the nearest weather dataset automatically (PostGIS nearest-neighbour search).
Hourly series are laid out on the reference year 2001, since typical-year data has no real calendar year.

## Not built yet (from REQUIREMENTS.md)

Automatic weather lookup by location, upload of measured interval data with gap filling,
EV, PV, battery and demand-response models, rate schedules, greenhouse gas factors,
multi-year scenario analysis, and login.
