"""Radiation on a tilted surface: port of REMainCommon.SolDat.

Erbs correlation splits global horizontal into beam/diffuse; the HDKR
(Hay-Davies-Klucher-Reindl) model gives total radiation on the tilted collector.
"""
from __future__ import annotations

import math
from dataclasses import dataclass

from .constants import DTOR, GSC
from .weather import WeatherData


@dataclass
class SolarResult:
    cos_theta: list[float]  # angle of incidence on collector
    cos_theta_z: list[float]  # zenith angle
    it: list[float]  # total radiation on tilted surface, kJ/hr-m2


def erbs_diffuse_fraction(kt: float) -> float:
    # The VB code used `(kT <= 0.8) Or (kT > 0.22)`, which made the kT > 0.8
    # branch unreachable. This is the intended Erbs et al. correlation.
    if kt <= 0.22:
        return 1 - 0.09 * kt
    if kt <= 0.8:
        return 0.9511 - 0.1604 * kt + 4.388 * kt**2 - 16.638 * kt**3 + 12.336 * kt**4
    return 0.165


def tilted_radiation(w: WeatherData, tilt: float, azimuth: float, ground_reflectance: float) -> SolarResult:
    """tilt = Beta (deg from horizontal), azimuth = Gamma (deg, 0 = due south, west positive)."""
    beta, gamma = tilt * DTOR, azimuth * DTOR
    lat = w.latitude * DTOR
    lon_west = -w.longitude  # VB convention: west longitude positive
    std_meridian = w.time_zone * -15
    sin_b, cos_b = math.sin(beta), math.cos(beta)
    n_hours = len(w.ghi)
    cos_t, cos_tz, it = [0.0] * n_hours, [0.0] * n_hours, [0.0] * n_hours

    for j in range(n_hours):
        n = 1 + j // 24  # day number
        ndeg = (n - 1) * 360 / 365
        eqn_time = 229.2 * (0.000075 + 0.001868 * math.cos(ndeg * DTOR) - 0.032077 * math.sin(ndeg * DTOR)
                            - 0.014615 * math.cos(2 * ndeg * DTOR) - 0.04089 * math.sin(2 * ndeg * DTOR))
        decl = 23.45 * math.sin(360 * (284 + n) / 365 * DTOR) * DTOR
        hr_sol = w.hour[j] + (4 * (std_meridian - lon_west) + eqn_time) / 60
        wsol = (((hr_sol - 12) * 15) - 7.5) * DTOR  # mid-point of the hour

        ct = (math.sin(decl) * math.sin(lat) * cos_b
              - math.sin(decl) * math.cos(lat) * sin_b * math.cos(gamma)
              + math.cos(decl) * math.cos(lat) * cos_b * math.cos(wsol)
              + math.cos(decl) * math.sin(lat) * sin_b * math.cos(gamma) * math.cos(wsol)
              + math.cos(decl) * sin_b * math.sin(gamma) * math.sin(wsol))
        ctz = math.cos(lat) * math.cos(decl) * math.cos(wsol) + math.sin(lat) * math.sin(decl)
        cos_t[j], cos_tz[j] = ct, ctz

        rb = ct / ctz if ctz != 0 else 0.0
        if ct < 0 or ctz < 0:
            rb = 0.0
        rb = min(rb, 2.8)

        i0 = max(GSC * (1 + 0.033 * math.cos((360 * n / 365) * DTOR)) * ctz, 0.0)
        ih = w.ghi[j]
        if ih <= 0:
            continue
        if i0 <= 0:
            # Sun below the horizon at the hour's midpoint but measured GHI in the
            # hour (sunrise/sunset): treat it all as isotropic diffuse.
            it[j] = ih * (1 + cos_b) / 2 + ih * ground_reflectance * (1 - cos_b) / 2
            continue
        # The VB code dropped hours with kT >= 1 entirely; cap kT and the
        # anisotropy index instead so measured radiation is never discarded.
        kt = min(ih / i0, 1.0)
        i_diffuse = erbs_diffuse_fraction(kt) * ih
        i_beam = ih - i_diffuse
        ai = min(max(i_beam / i0, 0.0), 1.0)  # anisotropy index
        dif_cs = i_diffuse * rb * ai
        dif_iso = i_diffuse * ((1 + cos_b) / 2) * (1 - ai)
        dif_hor = dif_iso * math.sqrt(max(i_beam / ih, 0.0)) * math.sin(beta / 2) ** 3
        dif_ref = ih * ground_reflectance * ((1 - cos_b) / 2)
        it[j] = i_beam * rb + dif_cs + dif_iso + dif_hor + dif_ref
    return SolarResult(cos_theta=cos_t, cos_theta_z=cos_tz, it=it)
