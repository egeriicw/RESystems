# Renewable Energy Systems

A web port of the Visual Basic 6 **Renewable Energy Systems** solar-thermal simulator
(original source in [`Resources/`](Resources)). It has a FastAPI backend that runs the
hourly simulations and a React frontend for inputs and results.

```
backend/    FastAPI + simulation engine (pure Python)
frontend/   React + TypeScript (Vite), Recharts
Resources/  Original VB6 forms and modules
```

## Quick start

Backend (Python 3.11+):

```bash
cd backend
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/uvicorn app.main:app --reload --port 8000
.venv/bin/python -m pytest          # tests
```

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
| POST | `/api/simulate` | `{weather_id, model, collector, system, include_hourly}` → monthly, annual, daily, hourly, metrics |

Units follow the original: energy in kJ (hourly kJ/hr, daily and monthly values as average kJ/day),
radiation per m² of collector, temperatures in °C.

## Not yet in the MVP (from REQUIREMENTS.md)

The following are not built yet: Postgres/PostGIS/Timescale storage (weather is held in memory),
maps, EV, PV, battery and demand-response models, rate schedules, scenario analysis, login,
and gap-filling for uploaded data.
