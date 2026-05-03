# -*- coding: utf-8 -*-
"""
core/adu.py  —  Supply Stoplight v5
ADU (Average Daily Usage) calculation.

Rules
-----
- adu_hist  = (MRP M+0 + M+1 + M+2) divided by the real calendar days of those 3 months.
- adu_plan  = hybrid:
    * Paint line at plant 4301: B2Wise ADU first, fall back to adu_hist.
    * All others: adu_hist first; fall back to B2Wise ONLY if there is real recent
      consumption (consumos_3m_val > 0); otherwise 0 to avoid inflating projections.
"""

from __future__ import annotations
from calendar import monthrange
from datetime import date


def calculate_adu_hist(mrp_months: list[float], fecha_saldo: date) -> float:
    """
    Compute historical ADU from the first 3 MRP projected months.

    Divides the total MRP volume of M+0, M+1, M+2 by the real calendar days
    of those same months so the rate is comparable to a daily consumption figure.

    Parameters
    ----------
    mrp_months : list[float]
        MRP projected consumption per month (index 0 = current month M+0).
        Must have at least 3 elements.
    fecha_saldo : date
        Reference date used to resolve which calendar months M+0..M+2 are.

    Returns
    -------
    float
        Daily ADU in the material's base unit.  Returns 0.0 if MRP is zero.
    """
    total_cons = sum(mrp_months[i] for i in range(3))
    if total_cons <= 0:
        return 0.0

    total_days = sum(
        monthrange(
            fecha_saldo.year + (fecha_saldo.month - 1 + i) // 12,
            (fecha_saldo.month - 1 + i) % 12 + 1,
        )[1]
        for i in range(3)
    )
    return round(total_cons / total_days, 3) if total_days > 0 else 0.0


def calculate_adu(
    adu_b2w: float,
    adu_hist: float,
    consumos_3m_val: float,
    linea_fab: str,
    centros: list[int],
) -> float:
    """
    Select the planning ADU using the hybrid MRP / ERP logic.

    Priority rules
    --------------
    Paint line at plant 4301
        1. ERP ADU  (adu_b2w)
        2. Historical ADU from MRP  (adu_hist)

    All other materials
        1. Historical ADU from MRP  (adu_hist)
        2. ERP ADU — ONLY when consumos_3m_val > 0 (prevents inflating
           projections with an ERP ADU when there is no real consumption)
        3. 0 — no known demand

    Parameters
    ----------
    adu_b2w : float
        ADU read directly from the ERP cut (field "ADU").
    adu_hist : float
        ADU derived from MRP M+0..M+2 (output of calculate_adu_hist).
    consumos_3m_val : float
        Total real consumption in the last 90 days across all plants.
    linea_fab : str
        Production line name (e.g. "Paint").
    centros : list[int]
        Plant codes for the material group being evaluated.

    Returns
    -------
    float
        Planning ADU to use for coverage and projection calculations.
    """
    if linea_fab == "Pintura" and 4301 in centros:
        return adu_b2w if adu_b2w > 0 else adu_hist

    if adu_hist > 0:
        return adu_hist
    if consumos_3m_val > 0:
        return adu_b2w   # ERP fallback, only when real history exists
    return 0.0
