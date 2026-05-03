# -*- coding: utf-8 -*-
"""
core/inventory_state.py  —  Supply Planner v5
Inventory state classification: CRITICO / RIESGO / ALERTA / OK / SIN_CONSUMO.
"""

from __future__ import annotations
from datetime import date


def semaforo_estado(cob_hoy: float) -> str:
    """
    Classify current coverage into a traffic-light state.

    Thresholds
    ----------
    CRÍTICO  : cob_hoy ≤ 8 days
    RIESGO   : 9 ≤ cob_hoy ≤ 21 days
    OK       : cob_hoy > 21 days  (or 9999 = no demand)

    Parameters
    ----------
    cob_hoy : float
        Days of coverage calculated from today's on-hand stock.

    Returns
    -------
    str  "CRITICO" | "RIESGO" | "OK"
    """
    if cob_hoy >= 9999:  return "OK"
    if cob_hoy <= 8:     return "CRITICO"
    if cob_hoy <= 21:    return "RIESGO"
    return "OK"


def detect_quiebre_info(
    sabs: list[dict],
    lt: int,
    ref_date: date,
) -> tuple[bool, bool, int | None, bool]:
    """
    Analyse the Saturday projection list to detect stockout risk.

    Definitions
    -----------
    quiebre_12sem  : any Saturday in the first 12 has color "K" (stockout).
    riesgo_12sem   : quiebre_12sem OR any Saturday in the first 12 has color "R".
    primer_quiebre_dias : calendar days from ref_date to the first stockout Saturday.
    sin_tiempo     : the first stockout occurs within the supplier lead time,
                     meaning there is no time to place a purchase order.

    Parameters
    ----------
    sabs : list[dict]
        Output of project_inventory (full Saturday list).
    lt : int
        Supplier lead time in days.
    ref_date : date
        Reference date (typically FECHA_SALDO).

    Returns
    -------
    tuple[bool, bool, int | None, bool]
        (quiebre_12sem, riesgo_12sem, primer_quiebre_dias, sin_tiempo)
    """
    horizon = sabs[:12]

    quiebre_12sem = any(s["color"] == "K" for s in horizon)
    riesgo_12sem  = quiebre_12sem or any(s["color"] == "R" for s in horizon)

    primer_quiebre_dias: int | None = None
    for s in horizon:
        if s["color"] == "K":
            try:
                sat_d = date.fromisoformat(s["fecha"])
                primer_quiebre_dias = (sat_d - ref_date).days
            except Exception:
                pass
            break

    sin_tiempo = (
        primer_quiebre_dias is not None
        and lt > 0
        and primer_quiebre_dias <= lt
    )

    return quiebre_12sem, riesgo_12sem, primer_quiebre_dias, sin_tiempo


def classify_inventory_state(
    saldo: float,
    z_roja_plan: float,
    z_am_plan: float,
    quiebre_12sem: bool,
    sin_tiempo: bool,
    sin_demanda: bool,
    net_flow_str: str,
) -> str:
    """
    Assign the final inventory state for a material.

    Decision tree
    -------------
    SIN_CONSUMO : no known demand but stock is positive.
    CRITICO     : stock already in red zone  OR  projected stockout within LT.
    RIESGO      : stock in yellow zone (DDMRP concept).
    ALERTA      : projected stockout in 12 weeks but LT still allows a PO,
                  OR B2Wise net-flow alert signals demand risk.
    OK          : everything else.

    Parameters
    ----------
    saldo : float
        Current on-hand stock.
    z_roja_plan : float
        Planning Top of Red (from adjust_planning_zones).
    z_am_plan : float
        Planning Top of Yellow.
    quiebre_12sem : bool
        Any projected stockout in the next 12 Saturdays.
    sin_tiempo : bool
        Stockout occurs within supplier lead time.
    sin_demanda : bool
        adu_plan == 0 and all MRP months are zero.
    net_flow_str : str
        B2Wise NetFlow alert color string (e.g. "10 - DR", "20 - R").

    Returns
    -------
    str  "CRITICO" | "RIESGO" | "ALERTA" | "OK" | "SIN_CONSUMO"
    """
    if sin_demanda and saldo > 0:
        return "SIN_CONSUMO"

    # Current buffer zone (stock vs planning thresholds)
    if saldo <= 0 or (z_roja_plan > 0 and saldo <= z_roja_plan):
        zona_act = "R"
    elif z_am_plan > 0 and saldo <= z_am_plan:
        zona_act = "Y"
    else:
        zona_act = "G"

    if zona_act == "R" or (quiebre_12sem and sin_tiempo):
        return "CRITICO"
    if zona_act == "Y":
        return "RIESGO"
    if quiebre_12sem:
        return "ALERTA"
    if any(v in net_flow_str for v in ["10 - DR", "20 - R"]):
        return "ALERTA"
    return "OK"
