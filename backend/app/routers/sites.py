from __future__ import annotations

from datetime import datetime

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, Field
from sqlalchemy import func, select, text
from sqlalchemy.orm import Session

from .. import store
from ..db import get_db
from ..orm import Scenario, Site, WeatherDataset

router = APIRouter(prefix="/api/sites", tags=["sites"])


class SiteFields(BaseModel):
    name: str = Field(min_length=1, max_length=120)
    latitude: float = Field(ge=-90, le=90)
    longitude: float = Field(ge=-180, le=180)
    time_zone: float | None = Field(None, ge=-12, le=14, description="Hours from UTC; estimated from longitude if omitted")
    address: str | None = Field(None, max_length=300)
    building_type: str | None = Field(None, max_length=60)
    floor_area_m2: float | None = Field(None, ge=0)
    year_built: int | None = Field(None, ge=1600, le=2200)
    occupants: int | None = Field(None, ge=0)
    ua_house: float | None = Field(None, ge=0, description="Building heat loss coefficient (kJ/hr-C)")
    notes: str | None = Field(None, max_length=2000)
    weather_id: str | None = Field(None, description="Nearest dataset is assigned if omitted")


class SiteUpdate(SiteFields):
    name: str | None = Field(None, min_length=1, max_length=120)
    latitude: float | None = Field(None, ge=-90, le=90)
    longitude: float | None = Field(None, ge=-180, le=180)


class SiteOut(SiteFields):
    id: int
    time_zone: float
    created_at: datetime
    scenario_count: int = 0
    weather_name: str | None = None
    weather_distance_km: float | None = None


def get_site(db: Session, site_id: int) -> Site:
    site = db.get(Site, site_id)
    if site is None:
        raise HTTPException(404, f"Unknown site {site_id}")
    return site


def _distances(db: Session) -> dict[int, float]:
    rows = db.execute(text(
        "SELECT s.id, ST_Distance(s.geom, w.geom) / 1000 FROM sites s JOIN weather_datasets w ON w.id = s.weather_id"))
    return {r[0]: r[1] for r in rows}


def _out(site: Site, counts: dict[int, int], names: dict[str, str], dist: dict[int, float]) -> SiteOut:
    d = dist.get(site.id)
    return SiteOut.model_validate({
        **{c.name: getattr(site, c.name) for c in Site.__table__.columns},
        "scenario_count": counts.get(site.id, 0),
        "weather_name": names.get(site.weather_id) if site.weather_id else None,
        "weather_distance_km": round(d, 1) if d is not None else None,
    })


def _context(db: Session) -> tuple[dict[int, int], dict[str, str], dict[int, float]]:
    counts = dict(db.execute(select(Scenario.site_id, func.count()).group_by(Scenario.site_id)).all())
    names = dict(db.execute(select(WeatherDataset.id, WeatherDataset.name)).all())
    return counts, names, _distances(db)


@router.get("", response_model=list[SiteOut])
def list_sites(near_lat: float | None = None, near_lon: float | None = None, radius_km: float = 50,
               db: Session = Depends(get_db)) -> list[SiteOut]:
    """All sites, or those within radius_km of (near_lat, near_lon), nearest first."""
    if near_lat is not None and near_lon is not None:
        ids = db.execute(text("""
            SELECT id FROM sites
            WHERE ST_DWithin(geom, ST_SetSRID(ST_MakePoint(:lon, :lat), 4326)::geography, :m)
            ORDER BY geom <-> ST_SetSRID(ST_MakePoint(:lon, :lat), 4326)::geography"""),
            {"lat": near_lat, "lon": near_lon, "m": radius_km * 1000}).scalars().all()
        by_id = {s.id: s for s in db.scalars(select(Site).where(Site.id.in_(ids)))}
        sites = [by_id[i] for i in ids]
    else:
        sites = list(db.scalars(select(Site).order_by(Site.name)))
    ctx = _context(db)
    return [_out(s, *ctx) for s in sites]


@router.post("", response_model=SiteOut, status_code=201)
def create_site(body: SiteFields, db: Session = Depends(get_db)) -> SiteOut:
    data = body.model_dump()
    if data["time_zone"] is None:
        data["time_zone"] = round(body.longitude / 15)
    if "weather_id" in body.model_fields_set and data["weather_id"] is not None:
        store.get_dataset(db, data["weather_id"])
    else:
        nearest = store.nearest_weather(db, body.latitude, body.longitude)
        data["weather_id"] = nearest[0] if nearest else None
    site = Site(**data)
    db.add(site)
    db.commit()
    return _out(site, *_context(db))


@router.get("/{site_id}", response_model=SiteOut)
def read_site(site_id: int, db: Session = Depends(get_db)) -> SiteOut:
    return _out(get_site(db, site_id), *_context(db))


@router.patch("/{site_id}", response_model=SiteOut)
def update_site(site_id: int, body: SiteUpdate, db: Session = Depends(get_db)) -> SiteOut:
    site = get_site(db, site_id)
    changes = body.model_dump(exclude_unset=True)
    for required in ("name", "latitude", "longitude", "time_zone"):
        if required in changes and changes[required] is None:
            raise HTTPException(422, f"{required} cannot be null")
    if changes.get("weather_id") is not None:
        store.get_dataset(db, changes["weather_id"])
    moved = any(k in changes and changes[k] != getattr(site, k) for k in ("latitude", "longitude"))
    # Values that were set automatically for the old location; hand-picked ones are kept on a move.
    auto_tz = site.time_zone == round(site.longitude / 15)
    old_nearest = store.nearest_weather(db, site.latitude, site.longitude) if moved else None
    auto_weather = site.weather_id is None or (old_nearest is not None and site.weather_id == old_nearest[0])
    for key, value in changes.items():
        setattr(site, key, value)
    # A moved site gets new automatic values unless this request set them, or
    # they had been chosen by hand.
    if moved:
        if "time_zone" not in changes and auto_tz:
            site.time_zone = round(site.longitude / 15)
        if "weather_id" not in changes and auto_weather:
            nearest = store.nearest_weather(db, site.latitude, site.longitude)
            site.weather_id = nearest[0] if nearest else None
    db.commit()
    return _out(site, *_context(db))


@router.delete("/{site_id}", status_code=204)
def delete_site(site_id: int, db: Session = Depends(get_db)) -> None:
    db.delete(get_site(db, site_id))
    db.commit()
