"""Library of system components for one-line diagrams.

This registry is the single source for component types, their ports, the
parameters shown in the edit box (with units, limits and defaults) and the
outputs each component reports. The frontend builds its palette and edit
forms from it via GET /api/components.

Ports have a kind. A wire may only join an output port to an input port of
the same kind. Kinds:
  solar_loop    collector fluid (collector -> heat exchanger or tank)
  heating_loop  heating water from a tank to a heat emitter
  hot_water     potable hot water drawn from a tank
  heat          heat delivered to a space
  electric      reserved for electrical components (PV, battery, EV); no
                component uses it yet

Wires have parameters only where the physics uses them; they are declared
on the port at either end of the wire (`wire_params`).
"""
from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel

from .constants import CP_WATER, DEFAULT_HOT_WATER_PROFILE

GPM_TO_KG_PER_HR = 8.3237 / 2.2046 * 60  # US gal water -> kg, per minute -> per hour
PORT_KINDS = ["solar_loop", "heating_loop", "hot_water", "heat", "electric"]


class Param(BaseModel):
    key: str
    label: str
    type: Literal["number", "integer", "select", "profile", "months"] = "number"
    unit: str | None = None
    default: Any
    min: float | None = None
    max: float | None = None
    step: float | None = None
    options: list[dict] | None = None  # for select: [{"value", "label"}]
    help: str | None = None


class Port(BaseModel):
    id: str
    label: str
    direction: Literal["in", "out"]
    kind: str
    max_connections: int | None = None  # None = any number
    wire_params: list[Param] = []


class Output(BaseModel):
    key: str
    label: str
    kind: Literal["energy", "temperature", "irradiance"]
    unit: str  # unit of daily/monthly values


class ComponentType(BaseModel):
    type: str
    title: str
    category: str
    description: str
    ports: list[Port]
    params: list[Param]
    outputs: list[Output]
    summary: str  # output shown on the diagram after a run

    def port(self, port_id: str) -> Port | None:
        return next((p for p in self.ports if p.id == port_id), None)

    def defaults(self) -> dict:
        return {p.key: p.default for p in self.params}


DEFAULT_LOOP_FLOW = 2 * 0.64 * GPM_TO_KG_PER_HR  # two collectors at 0.64 gpm each

LOOP_FLOW = Param(
    key="flow", label="Flow rate", unit="kg/hr", default=round(DEFAULT_LOOP_FLOW, 2), min=0, step=1,
    help="Mass flow in this loop. 1 gpm of water is about 227 kg/hr.",
)

E = "kJ/day"
COMPONENTS: dict[str, ComponentType] = {c.type: c for c in [
    ComponentType(
        type="collector", title="Solar collector array", category="Sources",
        description="Flat-plate collectors (Hottel–Whillier–Bliss) on a tilted surface. "
                    "Radiation uses the Erbs split and HDKR sky model.",
        ports=[Port(id="out", label="Collector loop", direction="out", kind="solar_loop", max_connections=1)],
        params=[
            Param(key="num_collectors", label="Number of collectors", type="integer", default=2, min=1, max=500),
            Param(key="area_per_collector", label="Area per collector", unit="m²", default=2.277, min=0.01, step=0.01),
            Param(key="tilt", label="Tilt (β)", unit="deg", default=40, min=0, max=90),
            Param(key="azimuth", label="Azimuth (γ, 0 = south, west positive)", unit="deg", default=0, min=-180, max=180),
            Param(key="ground_reflectance", label="Ground reflectance (ρg)", default=0.2, min=0, max=1, step=0.05),
            Param(key="fr_ta", label="FR(τα)", default=0.725, min=0, max=1, step=0.005),
            Param(key="fr_ul", label="FR·UL", unit="kJ/hr·m²·C", default=16.96, min=0, step=0.1),
            Param(key="num_glazings", label="Glazings", type="select", default=1,
                  options=[{"value": 1, "label": "1"}, {"value": 2, "label": "2"}]),
            Param(key="inlet_temp", label="Inlet temperature when not connected", unit="C", default=30,
                  help="Used only if the collector is not wired to a tank or heat exchanger."),
        ],
        outputs=[
            Output(key="It", label="Radiation on collector", kind="irradiance", unit="kJ/day-m2"),
            Output(key="Qu", label="Useful gain", kind="energy", unit=E),
            Output(key="Qts", label="Delivered to storage", kind="energy", unit=E),
        ],
        summary="Qts",
    ),
    ComponentType(
        type="heat_exchanger", title="Heat exchanger", category="Distribution",
        description="Separates the collector loop from the tank. The collector inlet temperature "
                    "follows from the exchanger effectiveness and the tank-side flow.",
        ports=[
            Port(id="in", label="From collector", direction="in", kind="solar_loop", max_connections=1),
            Port(id="out", label="To tank", direction="out", kind="solar_loop", max_connections=1,
                 wire_params=[LOOP_FLOW.model_copy(update={"label": "Tank-side flow rate"})]),
        ],
        params=[
            Param(key="effectiveness", label="Effectiveness", default=0.8, min=0, max=1, step=0.05),
            Param(key="tank_side_flow", label="Tank-side flow", type="select", default="loop",
                  options=[{"value": "loop", "label": "Loop flow (set on the wire to the tank)"},
                           {"value": "hot_water_draw", "label": "Hot water draw of the tank's loads"}],
                  help="The original program's DHW model used the hourly hot water draw here; "
                       "its other models used the loop flow."),
        ],
        outputs=[Output(key="Q", label="Heat transferred", kind="energy", unit=E)],
        summary="Q",
    ),
    ComponentType(
        type="tank", title="Storage tank", category="Storage",
        description="Fully mixed water tank, stepped hourly. Collection stops above the maximum temperature.",
        ports=[
            Port(id="in", label="Solar input", direction="in", kind="solar_loop"),
            Port(id="heating", label="Heating supply", direction="out", kind="heating_loop"),
            Port(id="hot_water", label="Hot water", direction="out", kind="hot_water"),
        ],
        params=[
            Param(key="mass", label="Water mass", unit="kg", default=1500, min=1),
            Param(key="initial_temp", label="Initial temperature", unit="C", default=22.22),
            Param(key="max_temp", label="Maximum temperature", unit="C", default=90),
        ],
        outputs=[
            Output(key="Ts", label="Tank temperature", kind="temperature", unit="C"),
            Output(key="Qin", label="Solar in", kind="energy", unit=E),
            Output(key="Qout", label="Delivered to loads", kind="energy", unit=E),
        ],
        summary="Ts",
    ),
    ComponentType(
        type="fan_coil", title="Fan coil", category="Heat emitters",
        description="Water-to-air coil. Runs when the tank is above its minimum supply temperature; "
                    "delivers effectiveness × capacitance rate × (tank − room) up to the load.",
        ports=[
            Port(id="in", label="From tank", direction="in", kind="heating_loop", max_connections=1),
            Port(id="out", label="To space", direction="out", kind="heat", max_connections=1),
        ],
        params=[
            Param(key="effectiveness", label="Effectiveness", default=0.25, min=0, max=1, step=0.05),
            Param(key="capacitance", label="Capacitance rate (C_min)", unit="kJ/hr·C", default=1274, min=0),
            Param(key="min_supply_temp", label="Minimum tank temperature to run", unit="C", default=48.9),
        ],
        outputs=[Output(key="Q", label="Heat delivered", kind="energy", unit=E)],
        summary="Q",
    ),
    ComponentType(
        type="direct_emitter", title="Direct supply emitter", category="Heat emitters",
        description="Tank water circulated straight to the space (e.g. radiant floor). "
                    "Delivers flow × cp × (tank − room) up to the load when the tank is warmer than the room.",
        ports=[
            Port(id="in", label="From tank", direction="in", kind="heating_loop", max_connections=1,
                 wire_params=[LOOP_FLOW.model_copy(update={"label": "Supply flow rate"})]),
            Port(id="out", label="To space", direction="out", kind="heat", max_connections=1),
        ],
        params=[],
        outputs=[Output(key="Q", label="Heat delivered", kind="energy", unit=E)],
        summary="Q",
    ),
    ComponentType(
        type="house", title="House heating load", category="Loads",
        description="Space heating load UA × (set point − ambient) − internal gains, in the heating months. "
                    "Whatever the emitter can't supply is auxiliary heat.",
        ports=[Port(id="in", label="Heat", direction="in", kind="heat", max_connections=1)],
        params=[
            Param(key="ua", label="Heat loss coefficient (UA)", unit="kJ/hr·C", default=1185, min=0),
            Param(key="internal_gains", label="Internal gains", unit="kJ/hr", default=2520, min=0),
            Param(key="room_temp", label="Set point", unit="C", default=22.22),
            Param(key="heating_months", label="Heating months", type="months", default=[10, 11, 12, 1, 2, 3, 4]),
        ],
        outputs=[
            Output(key="Qload", label="Heating load", kind="energy", unit=E),
            Output(key="Qsolar", label="Solar supplied", kind="energy", unit=E),
            Output(key="Qaux", label="Auxiliary heat", kind="energy", unit=E),
        ],
        summary="Qaux",
    ),
    ComponentType(
        type="dhw_load", title="Hot water load", category="Loads",
        description="Hourly hot water draw. Solar supplies (tank − mains) up to the delivery temperature; "
                    "a tempering valve mixes in cold water above it; auxiliary heat makes up the rest.",
        ports=[Port(id="in", label="Hot water", direction="in", kind="hot_water", max_connections=1)],
        params=[
            Param(key="ground_water_temp", label="Mains water temperature", unit="C", default=10),
            Param(key="supply_temp", label="Delivery temperature", unit="C", default=55),
            Param(key="profile", label="Draw by hour (1–24)", type="profile", unit="kg/hr",
                  default=list(DEFAULT_HOT_WATER_PROFILE)),
        ],
        outputs=[
            Output(key="Qload", label="Hot water load", kind="energy", unit=E),
            Output(key="Qsolar", label="Solar supplied", kind="energy", unit=E),
            Output(key="Qaux", label="Auxiliary heat", kind="energy", unit=E),
        ],
        summary="Qsolar",
    ),
]}


def wire_params(source: ComponentType, source_port: Port, target: ComponentType, target_port: Port) -> list[Param]:
    return [*source_port.wire_params, *target_port.wire_params]


def library() -> dict:
    return {"port_kinds": PORT_KINDS, "cp_water": CP_WATER,
            "components": [c.model_dump() for c in COMPONENTS.values()]}
