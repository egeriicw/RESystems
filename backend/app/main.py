"""FastAPI service for the Renewable Energy Systems simulator."""
from __future__ import annotations

import os
from contextlib import asynccontextmanager

from fastapi import Depends, FastAPI
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session

from . import store
from .db import get_db, get_engine, migrate
from .routers import diagram, scenarios, sites, weather
from .schemas import CollectorParams, ModelType, SimulationRequest, SimulationResult, SystemParams
from .simulation import MODEL_SERIES, SERIES, TITLES, simulate


@asynccontextmanager
async def lifespan(_: FastAPI):
    if os.environ.get("AUTO_MIGRATE", "1") == "1":
        migrate()
    with Session(get_engine()) as db:
        store.seed_sample(db)
    yield


app = FastAPI(title="Renewable Energy Systems API", version="0.2.0", lifespan=lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=os.environ.get("CORS_ORIGINS", "http://localhost:5173").split(","),
    allow_methods=["*"],
    allow_headers=["*"],
)
app.include_router(weather.router)
app.include_router(sites.router)
app.include_router(scenarios.router)
app.include_router(diagram.router)


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok"}


@app.get("/api/models")
def list_models() -> list[dict]:
    return [{"id": m.value, "title": TITLES[m], "series": [SERIES[k] for k in MODEL_SERIES[m]]} for m in ModelType]


@app.get("/api/defaults")
def defaults() -> dict:
    return {"collector": CollectorParams().model_dump(), "system": SystemParams().model_dump()}


@app.post("/api/simulate", response_model=SimulationResult)
def run_simulation(req: SimulationRequest, db: Session = Depends(get_db)) -> SimulationResult:
    """Ad hoc simulation; nothing is stored. Use /api/scenarios/{id}/run to keep results."""
    return simulate(req, lambda wid: store.load_weather(db, wid))
