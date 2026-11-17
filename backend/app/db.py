"""Database engine and session handling (Postgres with TimescaleDB + PostGIS)."""
from __future__ import annotations

import os
from collections.abc import Iterator
from functools import lru_cache
from pathlib import Path

from sqlalchemy import Engine, create_engine
from sqlalchemy.orm import Session

DEFAULT_URL = "postgresql+psycopg://resystems:resystems@localhost:5432/resystems"


def database_url() -> str:
    return os.environ.get("DATABASE_URL", DEFAULT_URL)


@lru_cache
def get_engine() -> Engine:
    return create_engine(database_url(), pool_pre_ping=True)


def get_db() -> Iterator[Session]:
    with Session(get_engine()) as session:
        yield session


def migrate() -> None:
    """Apply Alembic migrations up to head (same as `alembic upgrade head`)."""
    from alembic import command
    from alembic.config import Config

    cfg = Config(str(Path(__file__).resolve().parent.parent / "alembic.ini"))
    cfg.set_main_option("sqlalchemy.url", database_url().replace("%", "%%"))
    command.upgrade(cfg, "head")
