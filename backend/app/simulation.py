"""Hourly thermal simulations: port of REMain.CalcQUseful, CalcSimHouse, CalcDHW,
CalcHouseWFan and CalcHouseCombo.

Fixes relative to the VB code (each was a latent bug, not a modelling choice):
  * daily arrays are indexed by the actual day number (VB never set `n`);
  * `If Quaxw < 0` / `If QuaxH < 0` typos meant auxiliary energy was never floored at 0;
  * Q-useful efficiency divides by collector area (VB compared kJ/day to kJ/day-m2);
  * the incidence-angle modifier guards against cos(theta) <= 0 instead of dividing by zero;
  * the glazing-dependent b0 is used in every model (VB hard-coded -0.1 in two of them).
"""
from __future__ import annotations

from collections.abc import Callable

from .constants import CP_WATER, DAYS_PER_MONTH
from .schemas import CollectorParams, ModelType, SeriesInfo, SimulationRequest, SimulationResult, SystemParams
from .solar import tilted_radiation
from .weather import WeatherData

GPM_TO_KG_PER_HR = 8.3237 / 2.2046 * 60  # US gal water -> kg, per minute -> per hour

TITLES = {
    ModelType.q_useful: "Collector Useful Energy (Qu)",
    ModelType.house_heating: "House Heating",
    ModelType.dhw: "Domestic Hot Water",
    ModelType.house_fan_coil: "House Heating with Fan Coil",
    ModelType.house_fan_coil_dhw: "House Heating with Fan Coil and DHW",
}

S = SeriesInfo
SERIES = {
    "Ta": S(key="Ta", label="Ambient temperature", kind="temperature", unit="C"),
    "Ih": S(key="Ih", label="Horizontal radiation", kind="irradiance", unit="kJ/day-m2"),
    "It": S(key="It", label="Tilted radiation", kind="irradiance", unit="kJ/day-m2"),
    "Qu": S(key="Qu", label="Collector useful gain", kind="energy", unit="kJ/day"),
    "Qts": S(key="Qts", label="Solar to storage", kind="energy", unit="kJ/day"),
    "Qfs": S(key="Qfs", label="Solar to house load", kind="energy", unit="kJ/day"),
    "Qload": S(key="Qload", label="House heating load", kind="energy", unit="kJ/day"),
    "Qaux": S(key="Qaux", label="Auxiliary heating", kind="energy", unit="kJ/day"),
    "Qfsh": S(key="Qfsh", label="Solar to house (fan coil)", kind="energy", unit="kJ/day"),
    "Qloadh": S(key="Qloadh", label="House heating load", kind="energy", unit="kJ/day"),
    "Qauxh": S(key="Qauxh", label="Auxiliary heating", kind="energy", unit="kJ/day"),
    "Qfsw": S(key="Qfsw", label="Solar to hot water", kind="energy", unit="kJ/day"),
    "Qloadw": S(key="Qloadw", label="Hot water load", kind="energy", unit="kJ/day"),
    "Qauxw": S(key="Qauxw", label="Auxiliary hot water", kind="energy", unit="kJ/day"),
    "Ts": S(key="Ts", label="Storage temperature", kind="temperature", unit="C"),
}

MODEL_SERIES = {
    ModelType.q_useful: ["Ta", "Ih", "It", "Qu"],
    ModelType.house_heating: ["Ta", "It", "Qu", "Qts", "Qfs", "Qload", "Qaux", "Ts"],
    ModelType.dhw: ["Ta", "It", "Qu", "Qts", "Qfsw", "Qloadw", "Qauxw", "Ts"],
    ModelType.house_fan_coil: ["Ta", "It", "Qu", "Qts", "Qfsh", "Qloadh", "Qauxh", "Ts"],
    ModelType.house_fan_coil_dhw: ["Ta", "It", "Qu", "Qts", "Qfsh", "Qloadh", "Qauxh", "Qfsw", "Qloadw", "Qauxw", "Ts"],
}


def incidence_modifier(cos_theta: float, num_glazings: int) -> float:
    """K_tau-alpha = 1 + b0 (1/cos(theta) - 1), clipped to [0, 1]."""
    if cos_theta <= 0:
        return 0.0
    b0 = -0.1 if num_glazings == 1 else -0.17
    kta = 1 + b0 * (1 / cos_theta - 1)
    return kta if 0 <= kta <= 1 else 0.0


def _house_load(sys: SystemParams, ta: float, month: int) -> float:
    q = sys.ua * (sys.room_temp - ta) - sys.internal_gains
    return q if q > 0 and month in sys.heating_months else 0.0


def _run(model: ModelType, w: WeatherData, col: CollectorParams, sys: SystemParams) -> dict[str, list[float]]:
    solar = tilted_radiation(w, col.tilt, col.azimuth, col.ground_reflectance)
    area = col.area
    cp = CP_WATER
    mc = col.num_collectors * sys.flow_per_collector_gpm * GPM_TO_KG_PER_HR
    ml = mc
    e_hx = sys.collector_hx_effectiveness
    ts = sys.storage_initial_temp
    tc1 = sys.collector_inlet_temp if model == ModelType.q_useful else sys.storage_initial_temp
    keys = MODEL_SERIES[model]
    out: dict[str, list[float]] = {k: [] for k in keys}

    for j in range(len(w.ghi)):
        ta, it, mo = w.temp[j], solar.it[j], w.month[j]
        kta = incidence_modifier(solar.cos_theta[j], col.num_glazings)
        absorbed = col.fr_ta * kta * it
        mwater = sys.hot_water_profile[(w.hour[j] - 1) % 24]
        row: dict[str, float] = {"Ta": ta, "Ih": w.ghi[j], "It": it}

        if model == ModelType.q_useful:
            row["Qu"] = max(area * (absorbed - col.fr_ul * (tc1 - ta)), 0.0)

        elif model == ModelType.house_heating:
            # Collector loop feeds the tank directly (no HX); inlet = tank temperature.
            qu = area * (absorbed - col.fr_ul * (ts - ta))
            qts = qu if qu > 0 and ts <= sys.storage_max_temp else 0.0
            qload = _house_load(sys, ta, mo)
            qfs = min(ml * cp * (ts - sys.room_temp), qload) if ts > sys.room_temp and qload > 0 else 0.0
            ts += (qts - qfs) / (sys.storage_mass * cp)
            row |= {"Qu": max(qu, 0.0), "Qts": qts, "Qfs": qfs, "Qload": qload, "Qaux": qload - qfs, "Ts": ts}

        else:
            # Collector with heat exchanger; inlet temperature from the previous hour.
            qu = max(area * (absorbed - col.fr_ul * (tc1 - ta)), 0.0)
            hx_flow = mwater if model == ModelType.dhw else ml
            denom = e_hx * hx_flow * cp + (1 - e_hx) * area * col.fr_ul
            if denom > 0:
                tc1 = ((1 - e_hx) * area * (absorbed + col.fr_ul * ta) + e_hx * hx_flow * cp * ts) / denom
            qts = qu if qu > 0 and ts <= sys.storage_max_temp else 0.0
            row |= {"Qu": qu, "Qts": qts}
            draw = 0.0

            if model in (ModelType.house_fan_coil, ModelType.house_fan_coil_dhw):
                qloadh = _house_load(sys, ta, mo)
                qfsh = 0.0
                if qloadh > 0 and ts > sys.supply_air_temp:
                    qfsh = min(sys.fan_coil_effectiveness * sys.fan_coil_capacitance * (ts - sys.room_temp), qloadh)
                row |= {"Qfsh": qfsh, "Qloadh": qloadh, "Qauxh": max(qloadh - qfsh, 0.0)}
                draw += qfsh

            if model in (ModelType.dhw, ModelType.house_fan_coil_dhw):
                qfsw = mwater * cp * (ts - sys.ground_water_temp)
                qloadw = mwater * cp * (sys.hot_water_supply_temp - sys.ground_water_temp)
                row |= {"Qfsw": qfsw, "Qloadw": qloadw, "Qauxw": max(qloadw - qfsw, 0.0)}
                draw += qfsw

            ts += (qts - draw) / (sys.storage_mass * cp)
            row["Ts"] = ts

        for k in keys:
            out[k].append(row[k])
    return out


def _aggregate(hourly: dict[str, list[float]], w: WeatherData) -> tuple[dict, list[dict], dict]:
    n_days = len(w.ghi) // 24
    daily: dict[str, list[float]] = {}
    for key, values in hourly.items():
        is_temp = SERIES[key].kind == "temperature"
        days = []
        for d in range(n_days):
            chunk = values[d * 24 : (d + 1) * 24]
            days.append(sum(chunk) / 24 if is_temp else sum(chunk))
        daily[key] = days

    monthly = []
    for m in range(1, 13):
        idx = [d for d in range(n_days) if w.month[d * 24] == m]
        row = {"month": m}
        for key, days in daily.items():
            row[key] = sum(days[d] for d in idx) / len(idx) if idx else 0.0
        monthly.append(row)
    annual = {key: sum(days) / len(days) for key, days in daily.items()}
    return daily, monthly, annual


def _metrics(model: ModelType, annual: dict, area: float) -> dict:
    it_total = annual["It"] * area
    eff = lambda q: q / it_total if it_total > 0 else None  # noqa: E731
    ratio = lambda a, b: a / b if b > 0 else None  # noqa: E731
    if model == ModelType.q_useful:
        return {"collector_efficiency": eff(annual["Qu"])}
    if model == ModelType.house_heating:
        return {"collector_efficiency": eff(annual["Qts"]), "solar_fraction": ratio(annual["Qfs"], annual["Qload"])}
    if model == ModelType.dhw:
        return {"collector_efficiency": eff(annual["Qts"]),
                "solar_fraction": ratio(annual["Qloadw"] - annual["Qauxw"], annual["Qloadw"])}
    if model == ModelType.house_fan_coil:
        return {"collector_efficiency": eff(annual["Qfsh"]), "solar_fraction": ratio(annual["Qfsh"], annual["Qloadh"])}
    delivered = annual["Qfsh"] + annual["Qfsw"]
    return {"collector_efficiency": eff(delivered),
            "solar_fraction": ratio(delivered, annual["Qloadh"] + annual["Qloadw"])}


def _round(values: list[float], digits: int = 2) -> list[float]:
    return [round(v, digits) for v in values]


def simulate(req: SimulationRequest, weather_lookup: Callable[[str], WeatherData]) -> SimulationResult:
    w = weather_lookup(req.weather_id)
    hourly = _run(req.model, w, req.collector, req.system)
    daily, monthly, annual = _aggregate(hourly, w)
    notes = []
    if w.source == "synthetic":
        notes.append("Weather is a synthetic typical year; upload a TMY2 file for site-specific results.")
    return SimulationResult(
        model=req.model,
        title=TITLES[req.model],
        weather={"name": w.name, "source": w.source, "latitude": w.latitude,
                 "longitude": w.longitude, "time_zone": w.time_zone},
        inputs={"collector": req.collector.model_dump(), "system": req.system.model_dump(),
                "collector_area_m2": req.collector.area, "house_ua": req.system.ua},
        series=[SERIES[k] for k in MODEL_SERIES[req.model]],
        monthly=[{k: (round(v, 2) if isinstance(v, float) else v) for k, v in r.items()} for r in monthly],
        annual={k: round(v, 2) for k, v in annual.items()},
        metrics={k: (round(v, 4) if v is not None else None) for k, v in _metrics(req.model, annual, req.collector.area).items()},
        daily={k: _round(v) for k, v in daily.items()},
        hourly={k: _round(v, 1) for k, v in hourly.items()} if req.include_hourly else None,
        notes=notes,
    )
