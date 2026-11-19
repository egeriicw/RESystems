from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from sqlalchemy.orm import Session

from .. import store
from ..components import library
from ..db import get_db
from ..diagram import Diagram, DiagramError, Issue, simulate_diagram, templates, validate
from ..schemas import SimulationResult

router = APIRouter(prefix="/api", tags=["diagram"])


@router.get("/components")
def components() -> dict:
    """Component library: types, ports, parameter schemas (units, limits, defaults) and outputs."""
    return library()


@router.get("/diagram/templates")
def diagram_templates() -> list[dict]:
    return templates()


class ValidateRequest(BaseModel):
    diagram: Diagram


@router.post("/diagram/validate")
def validate_diagram(req: ValidateRequest) -> dict[str, list[Issue]]:
    return {"issues": validate(req.diagram)}


class DiagramSimulateRequest(BaseModel):
    weather_id: str = "sample"
    diagram: Diagram
    include_hourly: bool = True


@router.post("/diagram/simulate", response_model=SimulationResult)
def simulate(req: DiagramSimulateRequest, db: Session = Depends(get_db)) -> SimulationResult:
    """Ad hoc run of a diagram; nothing is stored. 422 with `issues` if the diagram is incomplete."""
    weather = store.load_weather(db, req.weather_id)
    try:
        return simulate_diagram(req.diagram, weather, req.include_hourly)
    except DiagramError as exc:
        raise HTTPException(422, {"message": "The diagram has problems", "issues": [i.model_dump() for i in exc.issues]}) from exc
