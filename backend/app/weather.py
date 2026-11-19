"""Hourly weather data: TMY2 parsing (port of REMain.LoadTMY2_Click) and a
synthetic typical-year generator so the app works without a TMY2 file."""
from __future__ import annotations

import math
import random
from dataclasses import dataclass, field

from .constants import DAYS_PER_MONTH, DTOR, GSC, HOURS_PER_YEAR


@dataclass
class WeatherData:
    name: str
    source: str  # "tmy2" | "synthetic"
    latitude: float  # deg, north positive
    longitude: float  # deg, east positive (west is negative)
    time_zone: float  # hours from UTC, e.g. -5 for US Eastern
    month: list[int] = field(default_factory=list)
    day: list[int] = field(default_factory=list)
    hour: list[int] = field(default_factory=list)  # 1..24, hour ending
    ghi: list[float] = field(default_factory=list)  # Ih, kJ/hr-m2
    temp: list[float] = field(default_factory=list)  # Ta, C
    dew_point: list[float] = field(default_factory=list)  # Tdp, C


class WeatherParseError(ValueError):
    pass


def _field(line: str, start: int, length: int) -> str:
    """VB's Mid(s, start, length) with a 1-based start."""
    return line[start - 1 : start - 1 + length]


def _num(text: str) -> float:
    text = text.strip()
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError as exc:
        raise WeatherParseError(f"Bad numeric field {text!r}") from exc


def parse_tmy2(text: str, name: str = "TMY2 upload") -> WeatherData:
    """Parse an NREL TMY2 (.tm2) file using the same fixed columns as the VB program."""
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if len(lines) < HOURS_PER_YEAR + 1:
        raise WeatherParseError(
            f"TMY2 file needs a header plus {HOURS_PER_YEAR} hourly records; got {len(lines)} lines"
        )
    header = lines[0]
    time_zone = _num(_field(header, 34, 3))
    lat = _num(_field(header, 40, 2)) + _num(_field(header, 43, 2)) / 60
    if _field(header, 38, 1).upper() == "S":
        lat = -lat
    lon = _num(_field(header, 48, 3)) + _num(_field(header, 52, 2)) / 60
    if _field(header, 46, 1).upper() != "E":
        lon = -lon  # TMY2 US stations are west longitudes
    city, state = _field(header, 8, 22).strip(), _field(header, 31, 2).strip()
    label = f"{city}, {state}" if city else name

    w = WeatherData(name=label or name, source="tmy2", latitude=lat, longitude=lon, time_zone=time_zone)
    for line in lines[1 : HOURS_PER_YEAR + 1]:
        w.month.append(int(_num(_field(line, 4, 2))))
        w.day.append(int(_num(_field(line, 6, 2))))
        w.hour.append(int(_num(_field(line, 8, 2))))
        w.ghi.append(_num(_field(line, 18, 4)) * 3.6)  # Wh/m2 -> kJ/hr-m2
        w.temp.append(_num(_field(line, 68, 4)) / 10)
        w.dew_point.append(_num(_field(line, 74, 4)) / 10)
    return w


def to_tmy2(w: WeatherData) -> str:
    """Write the fields this app uses back out in TMY2 column layout (for tests/export)."""
    lat_d, lat_m = divmod(round(abs(w.latitude) * 60), 60)
    lon_d, lon_m = divmod(round(abs(w.longitude) * 60), 60)
    h = [" "] * 60

    def hput(start: int, text: str) -> None:
        h[start - 1 : start - 1 + len(text)] = list(text)

    hput(2, "00000")
    hput(8, w.name[:22])
    hput(31, "XX")
    hput(34, f"{int(w.time_zone):3d}")
    hput(38, "N" if w.latitude >= 0 else "S")
    hput(40, f"{lat_d:2d}")
    hput(43, f"{lat_m:02d}")
    hput(46, "W" if w.longitude < 0 else "E")
    hput(48, f"{lon_d:3d}")
    hput(52, f"{lon_m:02d}")
    hput(56, f"{0:4d}")
    header = "".join(h).rstrip()
    out = [header]
    for i in range(len(w.ghi)):
        row = [" "] * 120
        def put(start: int, text: str) -> None:
            row[start - 1 : start - 1 + len(text)] = list(text)
        put(2, "99")
        put(4, f"{w.month[i]:2d}")
        put(6, f"{w.day[i]:2d}")
        put(8, f"{w.hour[i]:2d}")
        put(18, f"{round(w.ghi[i] / 3.6):4d}")
        put(68, f"{round(w.temp[i] * 10):4d}")
        put(74, f"{round(w.dew_point[i] * 10):4d}")
        out.append("".join(row).rstrip())
    return "\n".join(out) + "\n"


def synthetic_year(
    name: str = "Dayton, OH (synthetic)",
    latitude: float = 39.9,
    longitude: float = -84.2,
    time_zone: float = -5,
    mean_temp: float = 11.5,
    temp_amplitude: float = 13.0,
    seed: int = 42,
) -> WeatherData:
    """Generate a plausible typical meteorological year.

    The original program loads c:/Engineering Software/TMY2/Dayton.tm2, which is
    not in the repository. This stand-in uses clear-sky irradiance scaled by a
    persistent random daily clearness index and a seasonal + diurnal temperature
    model. It is deterministic for a given seed. Upload a real TMY2 file for
    real analysis.
    """
    rng = random.Random(seed)
    w = WeatherData(name=name, source="synthetic", latitude=latitude, longitude=longitude, time_zone=time_zone)
    lon_west = -longitude
    std_meridian = -15 * time_zone  # west positive
    southern = latitude < 0
    clearness = 0.6
    for n in range(1, 366):
        clearness = 0.55 * clearness + 0.45 * rng.uniform(0.2, 1.0)
        daily_temp_offset = rng.gauss(0, 3)
        decl = 23.45 * math.sin(DTOR * 360 * (284 + n) / 365)
        b = (n - 1) * 360 / 365 * DTOR
        eot = 229.2 * (0.000075 + 0.001868 * math.cos(b) - 0.032077 * math.sin(b)
                       - 0.014615 * math.cos(2 * b) - 0.04089 * math.sin(2 * b))
        seasonal = math.cos(2 * math.pi * (n - 200) / 365)  # warmest ~ mid July
        if southern:
            seasonal = -seasonal
        month, dom = _month_day(n)
        for hr in range(1, 25):
            solar_hr = hr - 0.5 + (4 * (std_meridian - lon_west) + eot) / 60
            omega = (solar_hr - 12) * 15
            cosz = (math.cos(latitude * DTOR) * math.cos(decl * DTOR) * math.cos(omega * DTOR)
                    + math.sin(latitude * DTOR) * math.sin(decl * DTOR))
            i0 = GSC * (1 + 0.033 * math.cos(DTOR * 360 * n / 365)) * max(cosz, 0)
            ghi = i0 * 0.78 * clearness * rng.uniform(0.9, 1.0)
            diurnal = math.cos(2 * math.pi * (hr - 15) / 24)
            ta = mean_temp + temp_amplitude * seasonal + 5 * diurnal * (0.5 + clearness) + daily_temp_offset
            w.month.append(month)
            w.day.append(dom)
            w.hour.append(hr)
            w.ghi.append(round(ghi, 1))
            w.temp.append(round(ta, 1))
            w.dew_point.append(round(ta - 4 - 6 * clearness, 1))
    return w


def _month_day(n: int) -> tuple[int, int]:
    for m, days in enumerate(DAYS_PER_MONTH, start=1):
        if n <= days:
            return m, n
        n -= days
    raise ValueError("day of year out of range")
