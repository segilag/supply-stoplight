# -*- coding: utf-8 -*-
"""
core/projection.py  —  Supply Planner v5
Saturday-based inventory projection and dot-color semaphore.
"""

from __future__ import annotations
from calendar import monthrange
from datetime import date, timedelta


def daily_adu_for_month(
    yr: int,
    mo: int,
    mrp_months: list[float],
    adu_plan: float,
    ref_date: date,
) -> float:
    """
    Return the daily ADU for a given calendar month.

    Uses the MRP projected value for that month if available and > 0;
    otherwise falls back to adu_plan.

    Parameters
    ----------
    yr, mo : int
        Target year and month (1-12).
    mrp_months : list[float]
        Monthly MRP projected consumption (index 0 = ref_date's month).
    adu_plan : float
        Planning ADU fallback.
    ref_date : date
        Reference date used to compute the month offset.

    Returns
    -------
    float
        Daily consumption rate for the requested month.
    """
    offset = (yr - ref_date.year) * 12 + (mo - ref_date.month)
    offset = max(0, min(offset, len(mrp_months) - 1))
    mrp_v  = mrp_months[offset]
    days_m = monthrange(yr, mo)[1]
    return (mrp_v / days_m) if mrp_v > 0 and days_m > 0 else adu_plan


def dot_color_projected(
    stock_sat: float,
    z_roja: float,
    z_am: float,
    cob_sat: float,
) -> str:
    """
    Semaphore color for a projected Saturday stock level.

    Color codes
    -----------
    K  — stockout (stock ≤ 0)
    R  — red zone  (stock ≤ Top of Red)
    Y  — yellow zone (stock ≤ Top of Yellow)
    B  — blue: excess coverage > 90 days
    G  — green (within normal range)

    Parameters
    ----------
    stock_sat : float
        Projected stock on that Saturday.
    z_roja : float
        Planning Top of Red (z_roja_plan from adjust_planning_zones).
    z_am : float
        Planning Top of Yellow (z_am_plan from adjust_planning_zones).
    cob_sat : float
        Projected days of coverage on that Saturday.

    Returns
    -------
    str  One of "K", "R", "Y", "B", "G".
    """
    if stock_sat <= 0:
        return "K"
    if z_roja > 0 and stock_sat <= z_roja:
        return "R"
    if z_am   > 0 and stock_sat <= z_am:
        return "Y"
    if 0 < cob_sat < 9999 and cob_sat > 90:
        return "B"
    return "G"


def project_inventory(
    saldo: float,
    mrp_months: list[float],
    adu_plan: float,
    inflows_map: dict[date, float],
    sabados: list[date],
    z_roja_plan: float,
    z_am_plan: float,
    ref_date: date,
) -> list[dict]:
    """
    Project on-hand inventory for each Saturday in sabados.

    Consumption is accumulated in monthly segments between Saturdays using
    the MRP-projected daily rate for each month.  Purchase order inflows
    (inflows_map) are added on the Saturday they fall on or before.

    Parameters
    ----------
    saldo : float
        Current on-hand stock at ref_date.
    mrp_months : list[float]
        Monthly MRP projected consumption (index 0 = ref_date's month).
    adu_plan : float
        Daily ADU fallback when MRP month is zero.
    inflows_map : dict[date, float]
        Mapping of {arrival_date: quantity} for open purchase orders.
        Quantities must already be converted to the material's base unit.
    sabados : list[date]
        Ordered list of Saturday dates to project.
    z_roja_plan : float
        Planning Top of Red (from adjust_planning_zones).
    z_am_plan : float
        Planning Top of Yellow (from adjust_planning_zones).
    ref_date : date
        Reference date (typically FECHA_SALDO).

    Returns
    -------
    list[dict]
        One dict per Saturday with keys:
        fecha (ISO str), label (DD/MM), stock (float), cob (float), color (str).
    """
    result    = []
    cum_cons  = 0.0
    prev_date = ref_date

    for sat in sabados:
        # Accumulate consumption in monthly segments from prev_date to sat
        cur = prev_date
        while cur < sat:
            last_day = monthrange(cur.year, cur.month)[1]
            seg_end  = min(sat, date(cur.year, cur.month, last_day))
            days_seg = (seg_end - cur).days
            cum_cons += daily_adu_for_month(cur.year, cur.month, mrp_months, adu_plan, ref_date) * days_seg
            cur = seg_end + timedelta(days=1)

        inflow_sat = sum(q for d, q in inflows_map.items() if ref_date < d <= sat)
        stock_sat  = saldo + inflow_sat - cum_cons
        daily_sat  = daily_adu_for_month(sat.year, sat.month, mrp_months, adu_plan, ref_date)
        cob_sat    = round(stock_sat / daily_sat, 1) if daily_sat > 0 else 9999

        result.append({
            "fecha": sat.strftime("%Y-%m-%d"),
            "label": sat.strftime("%d/%m"),
            "stock": round(stock_sat, 0),
            "cob":   cob_sat,
            "color": dot_color_projected(stock_sat, z_roja_plan, z_am_plan, cob_sat),
        })
        prev_date = sat

    return result
