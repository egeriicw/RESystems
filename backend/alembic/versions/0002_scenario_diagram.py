"""Store one-line diagrams on scenarios.

A scenario whose model is "diagram" keeps its components and wires in
`diagram`; its collector and system columns hold defaults and are unused.

Revision ID: 0002
Revises: 0001
"""
from alembic import op

revision = "0002"
down_revision = "0001"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.execute("ALTER TABLE scenarios ADD COLUMN diagram jsonb")
    op.execute("""ALTER TABLE scenarios ADD CONSTRAINT scenarios_diagram_present
                  CHECK ((model = 'diagram') = (diagram IS NOT NULL))""")


def downgrade() -> None:
    op.execute("DELETE FROM scenarios WHERE model = 'diagram'")
    op.execute("ALTER TABLE scenarios DROP CONSTRAINT scenarios_diagram_present")
    op.execute("ALTER TABLE scenarios DROP COLUMN diagram")
