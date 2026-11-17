"""Physical constants and unit conventions carried over from REMain.Form_Load.

All energy quantities follow the original program's units:
  * irradiance: kJ/hr-m2 (hourly), kJ/day-m2 (daily/monthly averages)
  * heat flows: kJ/hr (hourly), kJ/day (daily/monthly averages)
  * temperatures: deg C
"""
import math

PI = math.pi
DTOR = PI / 180.0  # degrees -> radians
RTOD = 180.0 / PI  # radians -> degrees
GSC = 4920.0  # solar constant, kJ/hr-m2 (1367 W/m2)
CP_WATER = 4.19  # kJ/kg-C
HOURS_PER_YEAR = 8760
DAYS_PER_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

# Hourly domestic hot-water draw profile (kg/hr), index 0 = hour 1 ... index 23 = hour 24.
# From the MWater() array in REMain.CalcDHW.
DEFAULT_HOT_WATER_PROFILE = [
    1, 0.5, 0.5, 0.3, 1, 4, 17.5, 18, 22, 17.5, 14.5, 13,
    12.5, 9, 8, 9, 10, 13, 18, 16, 13, 11, 10, 3,
]
