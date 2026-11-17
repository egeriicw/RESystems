from __future__ import annotations

from datetime import datetime

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, Field
from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from .. import store
from ..db import get_db
from ..orm import Run, Scenario
from ..schemas import CollectorParams, ModelType, SimulationRequest, SimulationResult, SystemParams
from ..simulation import simulate
from .sites import get_site

router = APIRouter(prefix="/api", tags=["scenarios"])


class ScenarioIn(BaseModel):
    name: str = Field(min_length=1, max_length=120)
    site_id: int | None = None
    weather_id: str | None = Field(None, description="Defaults to the site's weather, then the sample dataset")
    model: ModelType
    collector: CollectorParams = CollectorParams()
    system: SystemParams = SystemParams()


class RunSummary(BaseModel):
    id: int
    scenario_id: int
    created_at: datetime
    title: str
    weather_name: str
    metrics: dict
    annual: dict


class ScenarioOut(ScenarioIn):
    id: int
    site_name: str | None = None
    created_at: datetime
    updated_at: datetime
    latest_run: RunSummary | None = None


class RunOut(RunSummary):
    scenario_name: str
    result: SimulationResult


def _summary(run: Run) -> RunSummary:
    r = run.result
    return RunSummary(id=run.id, scenario_id=run.scenario_id, created_at=run.created_at, title=r["title"],
                      weather_name=r["weather"]["name"], metrics=r["metrics"], annual=r["annual"])


def _out(s: Scenario) -> ScenarioOut:
    return ScenarioOut(id=s.id, name=s.name, site_id=s.site_id, weather_id=s.weather_id, model=s.model,
                       collector=s.collector, system=s.system, site_name=s.site.name if s.site else None,
                       created_at=s.created_at, updated_at=s.updated_at,
                       latest_run=_summary(s.runs[0]) if s.runs else None)


def _get(db: Session, scenario_id: int) -> Scenario:
    s = db.get(Scenario, scenario_id)
    if s is None:
        raise HTTPException(404, f"Unknown scenario {scenario_id}")
    return s


def _get_run(db: Session, run_id: int) -> Run:
    run = db.get(Run, run_id)
    if run is None:
        raise HTTPException(404, f"Unknown run {run_id}")
    return run


def _validate_refs(db: Session, body: ScenarioIn) -> None:
    if body.site_id is not None:
        get_site(db, body.site_id)
    if body.weather_id is not None:
        store.get_dataset(db, body.weather_id)


@router.get("/scenarios", response_model=list[ScenarioOut])
def list_scenarios(site_id: int | None = None, db: Session = Depends(get_db)) -> list[ScenarioOut]:
    q = select(Scenario).options(selectinload(Scenario.runs), selectinload(Scenario.site)).order_by(Scenario.updated_at.desc())
    if site_id is not None:
        q = q.where(Scenario.site_id == site_id)
    return [_out(s) for s in db.scalars(q)]


@router.post("/scenarios", response_model=ScenarioOut, status_code=201)
def create_scenario(body: ScenarioIn, db: Session = Depends(get_db)) -> ScenarioOut:
    _validate_refs(db, body)
    s = Scenario(**body.model_dump(mode="json"))
    db.add(s)
    db.commit()
    return _out(s)


@router.get("/scenarios/{scenario_id}", response_model=ScenarioOut)
def read_scenario(scenario_id: int, db: Session = Depends(get_db)) -> ScenarioOut:
    return _out(_get(db, scenario_id))


@router.put("/scenarios/{scenario_id}", response_model=ScenarioOut)
def update_scenario(scenario_id: int, body: ScenarioIn, db: Session = Depends(get_db)) -> ScenarioOut:
    s = _get(db, scenario_id)
    _validate_refs(db, body)
    for key, value in body.model_dump(mode="json").items():
        setattr(s, key, value)
    db.commit()
    return _out(s)


@router.delete("/scenarios/{scenario_id}", status_code=204)
def delete_scenario(scenario_id: int, db: Session = Depends(get_db)) -> None:
    db.delete(_get(db, scenario_id))
    db.commit()


@router.post("/scenarios/{scenario_id}/run", response_model=RunOut, status_code=201)
def run_scenario(scenario_id: int, db: Session = Depends(get_db)) -> RunOut:
    """Simulate the scenario and store the result, including the hourly series."""
    s = _get(db, scenario_id)
    weather_id = store.sample_or(db, s.weather_id, s.site.weather_id if s.site else None)
    req = SimulationRequest(weather_id=weather_id, model=s.model, collector=s.collector, system=s.system)
    result = simulate(req, lambda wid: store.load_weather(db, wid))
    run = store.save_run(db, s.id, result)
    return RunOut(**_summary(run).model_dump(), scenario_name=s.name, result=result)


@router.get("/scenarios/{scenario_id}/runs", response_model=list[RunSummary])
def list_runs(scenario_id: int, db: Session = Depends(get_db)) -> list[RunSummary]:
    return [_summary(r) for r in _get(db, scenario_id).runs]


@router.get("/runs/{run_id}", response_model=RunOut)
def read_run(run_id: int, hourly: bool = True, db: Session = Depends(get_db)) -> RunOut:
    run = _get_run(db, run_id)
    result = SimulationResult(**run.result, hourly=store.load_run_hourly(db, run.id) if hourly else None)
    return RunOut(**_summary(run).model_dump(), scenario_name=run.scenario.name, result=result)


@router.delete("/runs/{run_id}", status_code=204)
def delete_run(run_id: int, db: Session = Depends(get_db)) -> None:
    db.delete(_get_run(db, run_id))
    db.commit()
