"""FastAPI service for the Renewable Energy Systems simulator."""
from __future__ import annotations

import os
import uuid

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

from .schemas import CollectorParams, ModelType, SimulationRequest, SimulationResult, SystemParams
from .simulation import MODEL_SERIES, SERIES, TITLES, simulate
from .weather import WeatherData, WeatherParseError, parse_tmy2, synthetic_year

app = FastAPI(title="Renewable Energy Systems API", version="0.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=os.environ.get("CORS_ORIGINS", "http://localhost:5173").split(","),
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory weather store. The requirements call for Postgres/Timescale later;
# this keeps the MVP dependency-free.
WEATHER: dict[str, WeatherData] = {"sample": synthetic_year()}


def get_weather(weather_id: str) -> WeatherData:
    try:
        return WEATHER[weather_id]
    except KeyError:
        raise HTTPException(404, f"Unknown weather dataset {weather_id!r}") from None


def weather_summary(weather_id: str, w: WeatherData) -> dict:
    monthly = []
    for m in range(1, 13):
        idx = [i for i, mo in enumerate(w.month) if mo == m]
        days = len(idx) / 24 or 1
        monthly.append({
            "month": m,
            "Ta": round(sum(w.temp[i] for i in idx) / max(len(idx), 1), 2),
            "Ih": round(sum(w.ghi[i] for i in idx) / days, 1),
        })
    return {"id": weather_id, "name": w.name, "source": w.source, "latitude": w.latitude,
            "longitude": w.longitude, "time_zone": w.time_zone, "hours": len(w.ghi), "monthly": monthly}


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok"}


@app.get("/api/models")
def list_models() -> list[dict]:
    return [{"id": m.value, "title": TITLES[m], "series": [SERIES[k] for k in MODEL_SERIES[m]]} for m in ModelType]


@app.get("/api/defaults")
def defaults() -> dict:
    return {"collector": CollectorParams().model_dump(), "system": SystemParams().model_dump()}


@app.get("/api/weather")
def list_weather() -> list[dict]:
    return [weather_summary(k, w) for k, w in WEATHER.items()]


@app.get("/api/weather/{weather_id}")
def weather_detail(weather_id: str, hourly: bool = False) -> dict:
    w = get_weather(weather_id)
    out = weather_summary(weather_id, w)
    if hourly:
        out["hourly"] = {"Ta": w.temp, "Ih": w.ghi}
    return out


@app.post("/api/weather/upload", status_code=201)
async def upload_weather(file: UploadFile = File(...)) -> dict:
    raw = await file.read()
    if len(raw) > 5_000_000:
        raise HTTPException(413, "File too large")
    try:
        w = parse_tmy2(raw.decode("latin-1"), name=file.filename or "upload")
    except WeatherParseError as exc:
        raise HTTPException(422, str(exc)) from exc
    weather_id = uuid.uuid4().hex[:8]
    WEATHER[weather_id] = w
    return weather_summary(weather_id, w)


class SyntheticRequest(BaseModel):
    name: str = Field("Synthetic site", max_length=60)
    latitude: float = Field(39.9, ge=-66, le=66)
    longitude: float = Field(-84.2, ge=-180, le=180)
    time_zone: float = Field(-5, ge=-12, le=14)
    mean_temp: float = Field(11.5, ge=-30, le=40)
    temp_amplitude: float = Field(13, ge=0, le=30)
    seed: int = 42


@app.post("/api/weather/synthetic", status_code=201)
def create_synthetic(req: SyntheticRequest) -> dict:
    weather_id = uuid.uuid4().hex[:8]
    WEATHER[weather_id] = synthetic_year(**req.model_dump())
    return weather_summary(weather_id, WEATHER[weather_id])


@app.delete("/api/weather/{weather_id}", status_code=204)
def delete_weather(weather_id: str) -> None:
    if weather_id == "sample":
        raise HTTPException(400, "The sample dataset cannot be deleted")
    get_weather(weather_id)
    del WEATHER[weather_id]


@app.post("/api/simulate", response_model=SimulationResult)
def run_simulation(req: SimulationRequest) -> SimulationResult:
    return simulate(req, get_weather)
