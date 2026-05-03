# -*- coding: utf-8 -*-
"""
settings.py  —  Supply Stoplight v5
Static configuration for the dashboard. Adjust before each run.
"""

from datetime import date
from calendar import monthrange

# ───────────────────────────────────────────────────────────────────────────────
#  DATES & PERIOD
# ───────────────────────────────────────────────────────────────────────────────
FECHA_SALDO  = date(2026, 4, 24)   # Date of the Current Stock file
FECHA_B2WISE = date(2026, 4, 24)   # Date of the ERP cut
MES          = 4                   # Current month (April)
AÑO          = 2026
DIAS_MES     = 30                  # Days in the current month
N_SABADOS    = 16                  # Saturdays to project (16 weeks)

# ───────────────────────────────────────────────────────────────────────────────
#  MONTH LABELS
# ───────────────────────────────────────────────────────────────────────────────
MESES_ES = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr",
    5: "May", 6: "Jun", 7: "Jul", 8: "Aug",
    9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}
MESES_FULL = {
    1: "January",  2: "February",  3: "March",     4: "April",
    5: "May",      6: "June",      7: "July",       8: "August",
    9: "September", 10: "October", 11: "November",  12: "December",
}
MES_M0 = MESES_FULL[MES]   # Full name of the current month

# ───────────────────────────────────────────────────────────────────────────────
#  HISTORICAL ADU  (last 3 complete months)
# ───────────────────────────────────────────────────────────────────────────────
_HIST_MONTHS = []
for _i in range(1, 4):
    _m, _y = MES - _i, AÑO
    if _m <= 0:
        _m += 12
        _y -= 1
    _HIST_MONTHS.append((_y, _m))

HIST_DAYS = sum(monthrange(y, m)[1] for y, m in _HIST_MONTHS)  # total days (~90)

# ───────────────────────────────────────────────────────────────────────────────
#  UNIT CONVERSIONS  (PO unit → KG)
# ───────────────────────────────────────────────────────────────────────────────
UM_A_KG = {
    "KG":  1.0,
    "G":   0.001,
    "TO":  1000.0,
    "T":   1000.0,
    "MT":  1000.0,
    "TS":  1000.0,    # Dry Ton
    "ADT": 1000.0,    # Air Dry Ton  — overridden by MAT_ADT_KG when applicable
    "WTO": 1000.0,    # Wet Ton      — overridden by MAT_ADT_KG when applicable
    "LB":  0.453592,
    "GAL": 1.0,
    "GLL": 1.0,
    "L":   1.0,
    "UN":  1.0,
}

# Actual KG per ADT or WTO by material code (specific conversion factor per product).
# Takes precedence over UM_A_KG when the PO unit is ADT or WTO.
MAT_ADT_KG = {
    "6010109": 900,
    "6011346": 900,
    "6015138": 900,
    "6013463": 900,
    "6010048": 995,
    "6010034": 930,
    "6010627": 964,
    "6011807": 900,
    "6006466": 995,
    "6000663": 952,
    "6000685": 952,
    "6002672": 990,
    "6010717": 997,
    "6019945": 900,
    "6013612": 900,
    "6019937": 900,
    "6020197": 900,
}

# ───────────────────────────────────────────────────────────────────────────────
#  RUNTIME OVERRIDE
# ───────────────────────────────────────────────────────────────────────────────

def apply_fecha(d: date) -> None:
    """
    Override all date-derived module globals at runtime.
    Must be called BEFORE importing preprocess / material_builder,
    since those modules capture these values at import time.
    """
    import sys
    mod = sys.modules[__name__]
    mod.FECHA_SALDO  = d
    mod.FECHA_B2WISE = d
    mod.MES          = d.month
    mod.AÑO          = d.year
    mod.DIAS_MES     = monthrange(d.year, d.month)[1]
    mod.MES_M0       = MESES_FULL[d.month]
    hist_months = []
    for i in range(1, 4):
        m, y = d.month - i, d.year
        if m <= 0:
            m  += 12
            y  -= 1
        hist_months.append((y, m))
    mod.HIST_DAYS = sum(monthrange(y, m)[1] for y, m in hist_months)
