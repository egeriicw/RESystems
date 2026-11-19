from __future__ import annotations

from enum import Enum
from typing import Literal

from pydantic import BaseModel, Field

from .constants import DEFAULT_HOT_WATER_PROFILE


class ModelType(str, Enum):
    """The five simulations offered by the original System form's Model Type combo box."""

    q_useful = "q_useful"
    house_heating = "house_heating"
    dhw = "dhw"
    house_fan_coil = "house_fan_coil"
    house_fan_coil_dhw = "house_fan_coil_dhw"


class CollectorParams(BaseModel):
    """Collector + placement inputs (the original 'Collectors' dialog)."""

    tilt: float = Field(40, ge=0, le=90, description="Beta, collector slope from horizontal (deg)")
    azimuth: float = Field(0, ge=-180, le=180, description="Gamma, surface azimuth, 0 = south, west positive (deg)")
    ground_reflectance: float = Field(0.2, ge=0, le=1, description="RhoG, ground albedo")
    fr_ta: float = Field(0.725, ge=0, le=1, description="FR(tau alpha), optical efficiency")
    fr_ul: float = Field(16.96, ge=0, description="FR UL, loss coefficient (kJ/hr-m2-C)")
    num_glazings: int = Field(1, ge=1, le=2, description="Number of cover glazings")
    num_collectors: int = Field(2, ge=1, le=500)
    area_per_collector: float = Field(2.277, gt=0, description="Gross area per collector (m2); 24.5 ft2")

    @property
    def area(self) -> float:
        return self.num_collectors * self.area_per_collector


class HouseUA(BaseModel):
    """Optional itemized UA values; when supplied their sum replaces ua_house (HCheck2)."""

    walls: float = 0
    windows: float = 0
    ceiling: float = 0
    infiltration: float = 0

    @property
    def total(self) -> float:
        return self.walls + self.windows + self.ceiling + self.infiltration


class SystemParams(BaseModel):
    """System inputs (the original 'System Parameters' dialog)."""

    collector_inlet_temp: float = Field(30, description="Tc1, fluid temperature entering collector (C)")
    flow_per_collector_gpm: float = Field(0.64, ge=0, description="Collector flow per collector (gpm)")
    storage_initial_temp: float = Field(22.22, description="Ts, initial tank temperature (C)")
    storage_mass: float = Field(1500, gt=0, description="ms, storage water mass (kg)")
    storage_max_temp: float = Field(90, description="Collection stops above this tank temperature (C)")
    collector_hx_effectiveness: float = Field(0.8, ge=0, le=1, description="e_hx")
    ua_house: float = Field(1185, ge=0, description="UAh, total house UA (kJ/hr-C)")
    ua_components: HouseUA | None = None
    internal_gains: float = Field(2520, ge=0, description="Qelec, internal/electrical gains (kJ/hr)")
    room_temp: float = Field(22.22, description="Tr, house set point (C)")
    supply_air_temp: float = Field(48.9, description="TsupH, min tank temp to run the fan coil (C)")
    fan_coil_effectiveness: float = Field(0.25, ge=0, le=1, description="e_loadhx")
    fan_coil_capacitance: float = Field(1274, ge=0, description="C_min of fan coil (kJ/hr-C)")
    ground_water_temp: float = Field(10, description="Tg, mains water temperature (C)")
    hot_water_supply_temp: float = Field(55, description="TsupW, hot water delivery temperature (C)")
    heating_months: list[int] = Field(default_factory=lambda: [10, 11, 12, 1, 2, 3, 4])
    hot_water_profile: list[float] = Field(
        default_factory=lambda: list(DEFAULT_HOT_WATER_PROFILE),
        min_length=24, max_length=24, description="Hot water draw by hour 1..24 (kg/hr)",
    )

    @property
    def ua(self) -> float:
        if self.ua_components is not None and self.ua_components.total > 0:
            return self.ua_components.total
        return self.ua_house


class SimulationRequest(BaseModel):
    weather_id: str = "sample"
    model: ModelType = ModelType.dhw
    collector: CollectorParams = CollectorParams()
    system: SystemParams = SystemParams()
    include_hourly: bool = True


class SeriesInfo(BaseModel):
    key: str
    label: str
    kind: str  # "energy" | "temperature" | "irradiance"
    unit: str  # unit of daily/monthly values


class SimulationResult(BaseModel):
    model: ModelType | Literal["diagram"]
    title: str
    weather: dict
    inputs: dict
    series: list[SeriesInfo]
    monthly: list[dict]
    annual: dict
    metrics: dict
    daily: dict[str, list[float]]
    hourly: dict[str, list[float]] | None = None
    notes: list[str] = []
    # Diagram runs only: annual outputs per component and energy carried per wire.
    diagram: dict | None = None
