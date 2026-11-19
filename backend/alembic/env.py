from alembic import context
from sqlalchemy import create_engine

from app.db import database_url


def run() -> None:
    engine = create_engine(context.config.get_main_option("sqlalchemy.url") or database_url())
    with engine.connect() as connection:
        context.configure(connection=connection, target_metadata=None)
        with context.begin_transaction():
            context.run_migrations()


run()
