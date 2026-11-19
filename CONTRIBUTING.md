# Contributing

How the team handles revisions to this codebase.

## Ownership

| Area | Covers | Reviewer for |
|---|---|---|
| Simulation engine | `backend/app/solar.py`, `simulation.py`, `weather.py`, `components.py`, `diagram.py` | Any change to physics or numbers |
| API and data | `backend/app/routers`, `store.py`, `orm.py`, Alembic migrations | Schema and endpoint changes |
| Frontend | `frontend/src`, including the diagram editor in `components/diagram` | UI changes |
| Infrastructure | `docker-compose.yml`, CI, deployment | Build and environment changes |

One person can hold two areas, but every change gets a reviewer who isn't the author.

## How a revision moves

1. **Issue first.** Each change starts as a GitHub issue saying what and why, labelled by area.
2. **Short-lived branch** off `main`, one issue per branch.
3. **Pull request** with CI green and one approval from the area owner. `main` is protected: no direct pushes.
4. **Merge and tag.** Squash-merge, and tag a release when a phase finishes.

## Rules specific to this code

- **Physics changes need a reference test.** Any edit to the engine adds or updates a test with an
  expected number. Results that shift must be called out in the pull request.
- **Deviations from the VB original are recorded.** Keep the list of fixed bugs in `README.md` and
  the `backend/app/simulation.py` docstring current.
- **Schema changes are new migrations.** Never edit `0001_initial.py`; add `0002_…` and test
  upgrading a database that has data in it.
- **API changes stay backward-compatible**, or update `frontend/src/api.ts` in the same pull request.
- **Units are declared.** Every new field states its unit in the schema description.
- **`simulation.py` is the reference for the diagram engine.** Keep the template-equivalence tests in
  `tests/test_diagram.py` passing; if a deliberate change breaks them, change both and say why.
- **Adding a component** means: define its ports, parameters and outputs in `components.py`, handle it in
  `diagram.py` (validation and the hourly loop), add its icon in `frontend/src/components/diagram/Glyph.tsx`,
  and add a test. The palette and edit box are generated from `components.py`.

## Before opening a pull request

```bash
docker compose up -d
cd backend && .venv/bin/python -m pytest
cd frontend && npm run build
```

## First revisions to queue

1. Confirm the CI workflow runs, then turn on branch protection for `main`.
2. Validate the engine against a real TMY2 file and the original program's output, if available.
3. Add frontend tests, starting with the flow: create a site, save and run two scenarios, compare them.
4. Switch displayed units from kJ to kWh.
5. Add weather lookup by location.
6. Then the roadmap phases: interval data upload with gap filling, electric models (PV, battery,
   EV, demand response), rates and emissions, scenario analysis, login.

## Cadence

- Weekly 30-minute triage: new issues get an owner and a phase.
- A phase is done when its tests pass in CI, the README is updated, and someone other than the
  author has clicked through it.
