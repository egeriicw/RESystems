"""Tests run against a throwaway `resystems_test` database on the same Postgres
server as development (docker compose up -d). Override with TEST_DATABASE_URL."""
import os

import pytest
from sqlalchemy import create_engine, text
from sqlalchemy.engine import make_url

from app.db import DEFAULT_URL

TEST_URL = os.environ.get("TEST_DATABASE_URL", str(make_url(DEFAULT_URL).set(database="resystems_test").render_as_string(hide_password=False)))
os.environ["DATABASE_URL"] = TEST_URL


@pytest.fixture(scope="session")
def client():
    url = make_url(TEST_URL)
    admin = create_engine(url.set(database="postgres"), isolation_level="AUTOCOMMIT")
    with admin.connect() as conn:
        conn.execute(text(f'DROP DATABASE IF EXISTS "{url.database}" WITH (FORCE)'))
        conn.execute(text(f'CREATE DATABASE "{url.database}"'))
    admin.dispose()

    from fastapi.testclient import TestClient

    from app.main import app

    with TestClient(app) as c:  # lifespan migrates and seeds the sample weather
        yield c
