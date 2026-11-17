"""Sites, weather, scenarios and runs.

Revision ID: 0001
Revises:
"""
from alembic import op

revision = "0001"
down_revision = None
branch_labels = None
depends_on = None

POINT = "geography(Point, 4326) GENERATED ALWAYS AS (ST_SetSRID(ST_MakePoint(longitude, latitude), 4326)::geography) STORED"


def upgrade() -> None:
    op.execute("CREATE EXTENSION IF NOT EXISTS postgis")
    op.execute("CREATE EXTENSION IF NOT EXISTS timescaledb")

    op.execute(f"""
        CREATE TABLE weather_datasets (
            id text PRIMARY KEY,
            name text NOT NULL,
            source text NOT NULL,
            latitude double precision NOT NULL,
            longitude double precision NOT NULL,
            geom {POINT},
            time_zone double precision NOT NULL,
            monthly jsonb NOT NULL,
            created_at timestamptz NOT NULL DEFAULT now()
        )""")
    op.execute("CREATE INDEX weather_datasets_geom_idx ON weather_datasets USING gist (geom)")

    # Typical-year hourly records. ts is the start of each hour, laid out on the
    # non-leap reference year 2001 (site local standard time, stored as UTC).
    op.execute("""
        CREATE TABLE weather_hourly (
            dataset_id text NOT NULL REFERENCES weather_datasets(id) ON DELETE CASCADE,
            ts timestamptz NOT NULL,
            ghi real NOT NULL,
            temp real NOT NULL,
            dew_point real NOT NULL,
            PRIMARY KEY (dataset_id, ts)
        )""")
    op.execute("SELECT create_hypertable('weather_hourly', 'ts', chunk_time_interval => interval '3 months')")

    op.execute(f"""
        CREATE TABLE sites (
            id serial PRIMARY KEY,
            name text NOT NULL,
            latitude double precision NOT NULL,
            longitude double precision NOT NULL,
            geom {POINT},
            time_zone double precision NOT NULL DEFAULT 0,
            address text,
            building_type text,
            floor_area_m2 double precision,
            year_built integer,
            occupants integer,
            ua_house double precision,
            notes text,
            weather_id text REFERENCES weather_datasets(id) ON DELETE SET NULL,
            created_at timestamptz NOT NULL DEFAULT now()
        )""")
    op.execute("CREATE INDEX sites_geom_idx ON sites USING gist (geom)")

    op.execute("""
        CREATE TABLE scenarios (
            id serial PRIMARY KEY,
            site_id integer REFERENCES sites(id) ON DELETE CASCADE,
            weather_id text REFERENCES weather_datasets(id) ON DELETE SET NULL,
            name text NOT NULL,
            model text NOT NULL,
            collector jsonb NOT NULL,
            system jsonb NOT NULL,
            created_at timestamptz NOT NULL DEFAULT now(),
            updated_at timestamptz NOT NULL DEFAULT now()
        )""")
    op.execute("CREATE INDEX scenarios_site_idx ON scenarios (site_id)")

    op.execute("""
        CREATE TABLE runs (
            id serial PRIMARY KEY,
            scenario_id integer NOT NULL REFERENCES scenarios(id) ON DELETE CASCADE,
            created_at timestamptz NOT NULL DEFAULT now(),
            result jsonb NOT NULL
        )""")
    op.execute("CREATE INDEX runs_scenario_idx ON runs (scenario_id)")

    op.execute("""
        CREATE TABLE run_hourly (
            run_id integer NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
            ts timestamptz NOT NULL,
            data jsonb NOT NULL,
            PRIMARY KEY (run_id, ts)
        )""")
    op.execute("SELECT create_hypertable('run_hourly', 'ts', chunk_time_interval => interval '3 months')")


def downgrade() -> None:
    for table in ("run_hourly", "runs", "scenarios", "sites", "weather_hourly", "weather_datasets"):
        op.execute(f"DROP TABLE IF EXISTS {table} CASCADE")
