# -*- coding: utf-8 -*-
"""
core/coverage.py  —  Supply Planner v5
Days-of-coverage simulation and DDMRP buffer zone helpers.
"""

from __future__ import annotations
from calendar import monthrange
from datetime import date, timedelta


def calculate_coverage(
    saldo: float,
    mrp_months: list[float],
    adu_fallback: float,
    ref_date: date,
) -> float:
    """
    Simulate month-by-month consumption to find exact days of coverage.

    Starting from ref_date + 1 day, the function burns stock at the MRP
    projected daily rate for each month.  When the MRP value for a month
    is zero it uses adu_fallback.  If adu (both MRP and fallback) is zero
    the material has no known demand → returns 9999.

    Parameters
    ----------
    saldo : float
        Current on-hand stock in the material's base unit.
    mrp_months : list[float]
        Monthly MRP projected consumption (index 0 = ref_date's month).
    adu_fallback : float
        Daily ADU used when the MRP month value is 0.
    ref_date : date
        Reference date (typically FECHA_SALDO).

    Returns
    -------
    float
        Days of coverage (0.0 if saldo ≤ 0, 9999 if demand is zero).
    """
    if saldo <= 0:
        return 0.0

    stock = saldo
    cur   = ref_date + timedelta(days=1)
    days  = 0.0

    for _ in range(250):   # safety cap ~20 years of months
        last  = monthrange(cur.year, cur.month)[1]
        off   = max(0, min(
            (cur.year - ref_date.year) * 12 + (cur.month - ref_date.month),
            len(mrp_months) - 1,
        ))
        daily = (mrp_months[off] / last) if mrp_months[off] > 0 else adu_fallback
        if daily <= 0:
            return 9999

        seg  = last - cur.day + 1   # remaining days in current month (inclusive)
        cons = daily * seg

        if cons >= stock:
            days += stock / daily
            return round(days, 1)

        stock -= cons
        days  += seg
        cur = date(
            cur.year + (cur.month == 12),
            cur.month % 12 + 1,
            1,
        )

    return 9999


def adjust_planning_zones(z_roja: float, z_am: float) -> tuple[float, float]:
    """
    Convert DDMRP execution zones to planning zones.

    The Top of Red (TOR) is split into two equal halves:
      - Lower half → new red zone top   (z_roja_plan = TOR / 2)
      - Upper half → new yellow zone top (z_am_plan  = TOR)

    This makes the planning semaphore more sensitive than the execution one.

    Parameters
    ----------
    z_roja : float
        Top of Red from B2Wise (execution buffer).
    z_am : float
        Top of Yellow from B2Wise (unused in this calculation, kept for
        signature symmetry and potential future use).

    Returns
    -------
    tuple[float, float]
        (z_roja_plan, z_am_plan)
    """
    return z_roja / 2, z_roja


def zones_to_days(
    z_roja: float,
    z_am: float,
    z_verde: float,
    adu_plan: float,
) -> tuple[float, float, float]:
    """
    Convert buffer zone quantities to days-of-coverage equivalents.

    Parameters
    ----------
    z_roja, z_am, z_verde : float
        Zone thresholds in the material's base unit.
    adu_plan : float
        Planning ADU (daily).

    Returns
    -------
    tuple[float, float, float]
        (z_roja_d, z_am_d, z_verde_d) — days; 0.0 when adu_plan is zero.
    """
    if adu_plan <= 0:
        return 0.0, 0.0, 0.0
    return (
        round(z_roja  / adu_plan, 1) if z_roja  > 0 else 0.0,
        round(z_am    / adu_plan, 1) if z_am    > 0 else 0.0,
        round(z_verde / adu_plan, 1) if z_verde > 0 else 0.0,
    )
