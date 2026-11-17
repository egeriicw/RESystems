from __future__ import annotations

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from pydantic import BaseModel, Field
from sqlalchemy import select
from sqlalchemy.orm import Session

from .. import store
from ..db import get_db
from ..orm import WeatherDataset
from ..weather import WeatherParseError, parse_tmy2, synthetic_year

router = APIRouter(prefix="/api/weather", tags=["weather"])


@router.get("")
def list_weather(db: Session = Depends(get_db)) -> list[dict]:
    return [store.weather_out(ds) for ds in db.scalars(select(WeatherDataset).order_by(WeatherDataset.created_at))]


@router.get("/{weather_id}")
def weather_detail(weather_id: str, hourly: bool = False, db: Session = Depends(get_db)) -> dict:
    out = store.weather_out(store.get_dataset(db, weather_id))
    if hourly:
        w = store.load_weather(db, weather_id)
        out["hourly"] = {"Ta": w.temp, "Ih": w.ghi}
    return out


@router.post("/upload", status_code=201)
async def upload_weather(file: UploadFile = File(...), db: Session = Depends(get_db)) -> dict:
    raw = await file.read()
    if len(raw) > 5_000_000:
        raise HTTPException(413, "File too large")
    try:
        w = parse_tmy2(raw.decode("latin-1"), name=file.filename or "upload")
    except WeatherParseError as exc:
        raise HTTPException(422, str(exc)) from exc
    return store.weather_out(store.save_weather(db, w))


class SyntheticRequest(BaseModel):
    name: str = Field("Synthetic site", max_length=60)
    latitude: float = Field(39.9, ge=-66, le=66)
    longitude: float = Field(-84.2, ge=-180, le=180)
    time_zone: float = Field(-5, ge=-12, le=14)
    mean_temp: float = Field(11.5, ge=-30, le=40)
    temp_amplitude: float = Field(13, ge=0, le=30)
    seed: int = 42


@router.post("/synthetic", status_code=201)
def create_synthetic(req: SyntheticRequest, db: Session = Depends(get_db)) -> dict:
    return store.weather_out(store.save_weather(db, synthetic_year(**req.model_dump())))


@router.delete("/{weather_id}", status_code=204)
def delete_weather(weather_id: str, db: Session = Depends(get_db)) -> None:
    if weather_id == store.SAMPLE_ID:
        raise HTTPException(400, "The sample dataset cannot be deleted")
    store.get_dataset(db, weather_id)
    store.delete_weather(db, weather_id)
