"""Persistence helpers for weather datasets and simulation runs."""
from __future__ import annotations

import uuid
from datetime import datetime, timedelta, timezone

from fastapi import HTTPException
from psycopg.types.json import Jsonb
from sqlalchemy import delete, select, text
from sqlalchemy.orm import Session

from .orm import Run, WeatherDataset
from .schemas import SimulationResult
from .weather import WeatherData, _month_day, synthetic_year

SAMPLE_ID = "sample"
REFERENCE_YEAR_START = datetime(2001, 1, 1, tzinfo=timezone.utc)
_weather_cache: dict[str, WeatherData] = {}


def hour_ts(i: int) -> datetime:
    return REFERENCE_YEAR_START + timedelta(hours=i)


def _copy(db: Session, table: str, columns: list[str], rows) -> None:
    """Bulk insert with COPY; much faster than INSERT for 8,760-row series."""
    cursor = db.connection().connection.driver_connection.cursor()
    with cursor.copy(f"COPY {table} ({', '.join(columns)}) FROM STDIN") as copy:
        for row in rows:
            copy.write_row(row)


def monthly_summary(w: WeatherData) -> list[dict]:
    out = []
    for m in range(1, 13):
        idx = [i for i, mo in enumerate(w.month) if mo == m]
        days = len(idx) / 24 or 1
        out.append({
            "month": m,
            "Ta": round(sum(w.temp[i] for i in idx) / max(len(idx), 1), 2),
            "Ih": round(sum(w.ghi[i] for i in idx) / days, 1),
        })
    return out


def weather_out(ds: WeatherDataset) -> dict:
    return {"id": ds.id, "name": ds.name, "source": ds.source, "latitude": ds.latitude,
            "longitude": ds.longitude, "time_zone": ds.time_zone, "hours": 8760, "monthly": ds.monthly}


def save_weather(db: Session, w: WeatherData, weather_id: str | None = None) -> WeatherDataset:
    ds = WeatherDataset(id=weather_id or uuid.uuid4().hex[:8], name=w.name, source=w.source, latitude=w.latitude,
                        longitude=w.longitude, time_zone=w.time_zone, monthly=monthly_summary(w))
    db.add(ds)
    db.flush()
    _copy(db, "weather_hourly", ["dataset_id", "ts", "ghi", "temp", "dew_point"],
          ((ds.id, hour_ts(i), w.ghi[i], w.temp[i], w.dew_point[i]) for i in range(len(w.ghi))))
    db.commit()
    return ds


def get_dataset(db: Session, weather_id: str) -> WeatherDataset:
    ds = db.get(WeatherDataset, weather_id)
    if ds is None:
        raise HTTPException(404, f"Unknown weather dataset {weather_id!r}")
    return ds


def load_weather(db: Session, weather_id: str) -> WeatherData:
    """Read a dataset's hourly series (cached in-process; datasets are immutable)."""
    if weather_id in _weather_cache:
        return _weather_cache[weather_id]
    ds = get_dataset(db, weather_id)
    rows = db.execute(
        text("SELECT ghi, temp, dew_point FROM weather_hourly WHERE dataset_id = :id ORDER BY ts"), {"id": weather_id}
    ).all()
    w = WeatherData(name=ds.name, source=ds.source, latitude=ds.latitude, longitude=ds.longitude, time_zone=ds.time_zone)
    for i, (ghi, temp, dew) in enumerate(rows):
        month, day = _month_day(i // 24 + 1)
        w.month.append(month)
        w.day.append(day)
        w.hour.append(i % 24 + 1)
        w.ghi.append(ghi)
        w.temp.append(temp)
        w.dew_point.append(dew)
    _weather_cache[weather_id] = w
    return w


def delete_weather(db: Session, weather_id: str) -> None:
    db.execute(delete(WeatherDataset).where(WeatherDataset.id == weather_id))
    db.commit()
    _weather_cache.pop(weather_id, None)


def seed_sample(db: Session) -> None:
    if db.get(WeatherDataset, SAMPLE_ID) is None:
        save_weather(db, synthetic_year(), SAMPLE_ID)


def nearest_weather(db: Session, latitude: float, longitude: float) -> tuple[str, float] | None:
    """Closest weather dataset to a point and its distance in km (PostGIS KNN)."""
    row = db.execute(text("""
        SELECT id, ST_Distance(geom, ST_SetSRID(ST_MakePoint(:lon, :lat), 4326)::geography) / 1000
        FROM weather_datasets
        ORDER BY geom <-> ST_SetSRID(ST_MakePoint(:lon, :lat), 4326)::geography
        LIMIT 1"""), {"lat": latitude, "lon": longitude}).first()
    return (row[0], row[1]) if row else None


def save_run(db: Session, scenario_id: int, result: SimulationResult) -> Run:
    hourly = result.hourly or {}
    run = Run(scenario_id=scenario_id, result=result.model_dump(mode="json", exclude={"hourly"}))
    db.add(run)
    db.flush()
    if hourly:
        keys = list(hourly)
        n = len(hourly[keys[0]])
        _copy(db, "run_hourly", ["run_id", "ts", "data"],
              ((run.id, hour_ts(i), Jsonb({k: hourly[k][i] for k in keys})) for i in range(n)))
    db.commit()
    return run


def load_run_hourly(db: Session, run_id: int) -> dict[str, list[float]] | None:
    rows = db.execute(text("SELECT data FROM run_hourly WHERE run_id = :id ORDER BY ts"), {"id": run_id}).scalars().all()
    if not rows:
        return None
    return {k: [r[k] for r in rows] for k in rows[0]}


def sample_or(db: Session, *candidates: str | None) -> str:
    """First weather id that is set, falling back to the sample dataset."""
    for c in candidates:
        if c:
            return c
    first = db.scalars(select(WeatherDataset.id).order_by(WeatherDataset.created_at).limit(1)).first()
    return first or SAMPLE_ID
