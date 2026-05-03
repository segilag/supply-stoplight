# -*- coding: utf-8 -*-
"""
material_builder.py  —  Supply Planner v5
Builds the per-material record consumed by the HTML dashboard.

Public API
----------
    MaterialContext   — dataclass that bundles all shared DataFrames
    build_material()  — entry point; returns a dict or None

Internal pipeline (called in order by build_material)
------------------------------------------------------
    extract_material_data()     — read raw fields from B2Wise / MRP / lookups
    compute_inventory_metrics() — calculate stock, ADU, coverage, zones
    compute_projections()       — build OC/SOLPED lists + Saturday projection
    build_material_output()     — assemble the final output dict
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, timedelta

import pandas as pd

from settings import FECHA_SALDO, UM_A_KG, MAT_ADT_KG
from core import (
    calculate_adu_hist, calculate_adu,
    calculate_coverage, adjust_planning_zones, zones_to_days,
    project_inventory, detect_quiebre_info, classify_inventory_state,
)


# ───────────────────────────────────────────────────────────────────────────────
#  Shared utility helpers  (also imported by generar_tablero for build_oc/build_sol)
# ───────────────────────────────────────────────────────────────────────────────

def sf(v, default: float = 0.0) -> float:
    """Safe float conversion — returns default on None, NaN, or any error."""
    try:
        if v is None:
            return default
        if isinstance(v, float) and math.isnan(v):
            return default
        return float(v)
    except Exception:
        return default


def clean_mat_group(v) -> str:
    """'001002 OPC' → 'OPC'  |  strips leading numeric prefix."""
    return re.sub(r'^\d+\s*', '', str(v).strip()).strip()


def _oc_fecha_info(fd) -> tuple[str, str, bool, int]:
    """
    Parse a purchase-order arrival date field.

    Returns
    -------
    (fecha_original, fecha_proyectada, atrasada, dias_atraso)
        fecha_original   — SAP date as ISO string (shown in tables).
        fecha_proyectada — FECHA_SALDO + 5d for overdue orders (used in semaphore).
        atrasada         — True if arrival date is before FECHA_SALDO.
        dias_atraso      — Calendar days past due (0 if not overdue).
    """
    if not pd.notna(fd):
        return "", "", False, 0
    d               = fd.date()
    atrasada        = d < FECHA_SALDO
    dias_atraso     = (FECHA_SALDO - d).days if atrasada else 0
    fecha_original  = fd.strftime("%Y-%m-%d")
    fecha_proyectada = (
        (FECHA_SALDO + timedelta(days=5)).strftime("%Y-%m-%d")
        if atrasada else fecha_original
    )
    return fecha_original, fecha_proyectada, atrasada, dias_atraso


def _sol_lib_info(r) -> tuple[str, str, bool]:
    """
    Parse a SOLPED row for release indicator and delivery date.

    Returns
    -------
    (fecha_str, ind_lib, sin_liberar)
    """
    ind_lib   = str(r.get("Indicador liberación", "")).strip()
    fecha_ent = r.get("Fecha de entrega")
    try:
        fecha_str = (
            pd.to_datetime(str(int(fecha_ent)), format="%Y%m%d").strftime("%Y-%m-%d")
            if pd.notna(fecha_ent) else ""
        )
    except Exception:
        fecha_str = (
            pd.to_datetime(fecha_ent, errors="coerce").strftime("%Y-%m-%d")
            if pd.notna(fecha_ent) else ""
        )
    return fecha_str, ind_lib, ind_lib == "X"


# ───────────────────────────────────────────────────────────────────────────────
#  Context  — shared DataFrames passed once per centro group
# ───────────────────────────────────────────────────────────────────────────────

@dataclass
class MaterialContext:
    """
    Bundles all shared DataFrames needed to build every material record.
    Create one instance per run and pass it to every build_material call.
    """
    # Lookup tables (from preprocess)
    df_lin:        pd.DataFrame
    origen_lookup: dict

    # Pre-aggregated tables (from preprocess)
    cons_3m:  pd.DataFrame
    ini_agg:  pd.DataFrame
    ing_mes:  pd.DataFrame
    cons_mes: pd.DataFrame
    b2_soh:   pd.DataFrame
    sdo_agg:  pd.DataFrame

    # Open documents (from loaders)
    df_b2oc: pd.DataFrame
    df_sol:  pd.DataFrame

    # Calendar (computed from settings)
    sabados: list[date]


# ───────────────────────────────────────────────────────────────────────────────
#  Step 1 — extract_material_data
# ───────────────────────────────────────────────────────────────────────────────

def extract_material_data(
    mat_str: str,
    centros: list[int],
    mrp_df: pd.DataFrame,
    b2_df: pd.DataFrame,
    df_lin: pd.DataFrame,
    origen_lookup: dict,
) -> dict:
    """
    Extract raw planning fields for one material from B2Wise, MRP, and lookup tables.

    Precedence rules
    ----------------
    - desc  : B2Wise first, fallback MRP
    - lt    : B2Wise (DLT) first, fallback MRP (LT TOTAL)
    - moq   : B2Wise first, fallback MRP (MOQ2)
    - um    : B2Wise; defaults to "UN" if missing

    Parameters
    ----------
    mat_str : str
        Normalised material code (output of mat_key).
    centros : list[int]
        SAP plant codes for this material group.
    mrp_df  : pd.DataFrame
        MRP DataFrame filtered to the current centro group.
    b2_df   : pd.DataFrame
        B2Wise DataFrame filtered to the current centro group.
    df_lin  : pd.DataFrame
        Líneas de fabricación lookup (all centros).
    origen_lookup : dict
        {(centro, mat_str): bool (True=importado)} lookup.

    Returns
    -------
    dict with keys:
        desc, um, adu, lt, moq,
        z_roja, z_am, z_verde, zona_b2, net_flow_str, mat_group,
        mrp_m0, mrp_months,
        linea_fab, importado, lt_detalle
    """
    # — Línea de fabricación + origen —
    lin_r     = df_lin[df_lin["_mat"] == mat_str]
    linea_fab = ""
    importado = False

    _orig_match = next(
        (origen_lookup[(c, mat_str)] for c in centros if (c, mat_str) in origen_lookup),
        None,
    )
    if _orig_match is not None:
        importado = _orig_match
    if not lin_r.empty:
        linea_fab = str(lin_r.iloc[0].get("linea_fab", "")).strip()

    # — B2Wise fields —
    b2r  = b2_df[b2_df["_mat"] == mat_str]
    desc, um, adu, lt, moq        = "", "UN", 0.0, 0, 0.0
    z_roja, z_am, z_verde         = 0.0, 0.0, 0.0
    zona_b2, net_flow_str, mat_group = "", "", ""

    if not b2r.empty:
        r            = b2r.iloc[0]
        desc         = str(r.get("Part Description", "")).strip()
        um           = str(r.get("UOM", "UN")).strip()
        adu          = sf(r.get("ADU", 0))
        lt           = int(sf(r.get("DLT", 0)))
        moq          = sf(r.get("MOQ", 0))
        z_roja       = sf(r.get("Top of Red", 0))
        z_am         = sf(r.get("Top of Yellow", 0))
        z_verde      = sf(r.get("Top of Green", 0))
        zona_b2      = str(r.get("SOH__AlertColor_Hidden", "")).strip()
        net_flow_str = str(r.get("NetFlow__AlertColor_Hidden", "")).strip()
        raw_mg       = r.get("Mat Group", "")
        mat_group    = clean_mat_group(raw_mg) if pd.notna(raw_mg) and str(raw_mg).strip() else ""

    # — MRP fields (complete / override if B2Wise was missing) —
    mrpr   = mrp_df[mrp_df["_mat"] == mat_str]
    mrp_m0 = 0.0

    if not mrpr.empty:
        r2 = mrpr.iloc[0]
        if not desc:
            desc   = str(r2.get("Descripción", "")).strip()
        if not lt:
            lt     = int(sf(r2.get("LT TOTAL", 0)))
        if not moq:
            moq    = sf(r2.get("MOQ2", 0))
        if not um:
            um     = "UN"
        mrp_m0 = sf(r2.get("MRP M+0", 0))

    mrp_months: list[float] = (
        [sf(mrpr.iloc[0].get(f"MRP M+{i}", 0)) for i in range(19)]
        if not mrpr.empty else [0.0] * 19
    )

    return {
        "desc":         desc,
        "um":           um,
        "adu":          adu,
        "lt":           lt,
        "moq":          moq,
        "z_roja":       z_roja,
        "z_am":         z_am,
        "z_verde":      z_verde,
        "zona_b2":      zona_b2,
        "net_flow_str": net_flow_str,
        "mat_group":    mat_group,
        "mrp_m0":       mrp_m0,
        "mrp_months":   mrp_months,
        "linea_fab":    linea_fab,
        "importado":    importado,
        "lt_detalle":   {},
    }


# ───────────────────────────────────────────────────────────────────────────────
#  Step 2 — compute_inventory_metrics
# ───────────────────────────────────────────────────────────────────────────────

def compute_inventory_metrics(
    mat_str: str,
    centros: list[int],
    mat_data: dict,
    ctx: MaterialContext,
) -> dict:
    """
    Calculate all stock, ADU, coverage, and buffer-zone metrics.

    Parameters
    ----------
    mat_str  : str           Normalised material code.
    centros  : list[int]     SAP plant codes.
    mat_data : dict          Output of extract_material_data.
    ctx      : MaterialContext

    Returns
    -------
    dict with keys:
        inv_ini, ingresos, consumos, saldo,
        adu_hist, consumos_3m_val, adu_plan,
        z_roja_plan, z_am_plan,
        cob_hoy,
        z_roja_d, z_am_d, z_verde_d,
        sin_demanda
    """
    mrp_months = mat_data["mrp_months"]
    adu        = mat_data["adu"]
    linea_fab  = mat_data["linea_fab"]
    z_roja     = mat_data["z_roja"]
    z_am       = mat_data["z_am"]
    z_verde    = mat_data["z_verde"]

    # — Inventory position —
    inv_ini = sum(
        sf(ctx.ini_agg[(ctx.ini_agg["Centro_ini"] == c) & (ctx.ini_agg["_mat"] == mat_str)]["inv_ini"].sum())
        for c in centros
    )
    ingresos = sum(
        sf(ctx.ing_mes[(ctx.ing_mes["Centro_ing"] == c) & (ctx.ing_mes["_mat"] == mat_str)]["ingresos_mes"].sum())
        for c in centros
    )
    consumos = sum(
        sf(ctx.cons_mes[(ctx.cons_mes["Centro"] == c) & (ctx.cons_mes["_mat"] == mat_str)]["consumos_mes"].sum())
        for c in centros
    )

    # Saldo: B2Wise SOH first (one row per mat+centro); fallback to Saldo Actual
    saldo = 0.0
    for c in centros:
        b2_match = ctx.b2_soh[(ctx.b2_soh["_centro"] == c) & (ctx.b2_soh["_mat"] == mat_str)]
        if not b2_match.empty:
            saldo += sf(b2_match.iloc[0]["saldo_b2w"])
        else:
            saldo += sf(ctx.sdo_agg[(ctx.sdo_agg["Centro"] == c) & (ctx.sdo_agg["_mat"] == mat_str)]["saldo"].sum())

    # — ADU —
    adu_hist = calculate_adu_hist(mrp_months, FECHA_SALDO)
    consumos_3m_val = sum(
        sf(ctx.cons_3m[(ctx.cons_3m["Centro"] == c) & (ctx.cons_3m["_mat"] == mat_str)]["consumos_3m"].sum())
        for c in centros
    )
    adu_plan = calculate_adu(adu, adu_hist, consumos_3m_val, linea_fab, centros)

    # — Coverage and zones —
    z_roja_plan, z_am_plan      = adjust_planning_zones(z_roja, z_am)
    cob_hoy                     = calculate_coverage(saldo, mrp_months, adu_plan, FECHA_SALDO)
    z_roja_d, z_am_d, z_verde_d = zones_to_days(z_roja, z_am, z_verde, adu_plan)

    sin_demanda = adu_plan <= 0 and all(v == 0 for v in mrp_months)

    return {
        "inv_ini":        inv_ini,
        "ingresos":       ingresos,
        "consumos":       consumos,
        "saldo":          saldo,
        "adu_hist":       adu_hist,
        "consumos_3m_val": consumos_3m_val,
        "adu_plan":       adu_plan,
        "z_roja_plan":    z_roja_plan,
        "z_am_plan":      z_am_plan,
        "cob_hoy":        cob_hoy,
        "z_roja_d":       z_roja_d,
        "z_am_d":         z_am_d,
        "z_verde_d":      z_verde_d,
        "sin_demanda":    sin_demanda,
    }


# ───────────────────────────────────────────────────────────────────────────────
#  Step 3 — compute_projections
# ───────────────────────────────────────────────────────────────────────────────

def compute_projections(
    mat_str: str,
    centros: list[int],
    mat_data: dict,
    metrics: dict,
    ctx: MaterialContext,
) -> dict:
    """
    Build OC and SOLPED lists, project inventory, and classify final state.

    Parameters
    ----------
    mat_str  : str
    centros  : list[int]
    mat_data : dict   Output of extract_material_data.
    metrics  : dict   Output of compute_inventory_metrics.
    ctx      : MaterialContext

    Returns
    -------
    dict with keys:
        oc_list, sol_list,
        sabs, quiebre_12sem, riesgo_12sem, primer_quiebre_dias, sin_tiempo,
        estado, brk_sin
    """
    um         = mat_data["um"]
    lt         = mat_data["lt"]
    mrp_months = mat_data["mrp_months"]
    net_flow   = mat_data["net_flow_str"]
    saldo      = metrics["saldo"]
    adu_plan   = metrics["adu_plan"]
    z_roja_plan = metrics["z_roja_plan"]
    z_am_plan   = metrics["z_am_plan"]
    sin_demanda = metrics["sin_demanda"]

    # — OC inflows + oc_list —
    oc_rows    = ctx.df_b2oc[(ctx.df_b2oc["_centro"].isin(centros)) & (ctx.df_b2oc["_mat"] == mat_str)].copy()
    inflows_map: dict[date, float] = {}
    oc_list    = []
    mat_um_up  = um.upper()

    for _, row in oc_rows.iterrows():
        fd    = row["_fecha_ent"]
        qty   = float(row["_qty_pend"])
        oc_um = str(row["_um"]).strip().upper()
        if qty <= 0:
            continue

        # Convert OC quantity to the material's base unit
        if oc_um in ("ADT", "WTO") and oc_um != mat_um_up:
            kg_por_unit = MAT_ADT_KG.get(mat_str, 1000)
            conv = kg_por_unit / UM_A_KG.get(mat_um_up, 1.0)
        else:
            conv = (
                UM_A_KG.get(oc_um, 1.0) / UM_A_KG.get(mat_um_up, 1.0)
                if oc_um and oc_um != mat_um_up else 1.0
            )
        qty_base = qty * conv

        eta_orig, eta_proy, atrasada, _ = _oc_fecha_info(fd)
        if pd.notna(fd):
            map_date = (FECHA_SALDO + timedelta(days=5)) if atrasada else fd.date()
            inflows_map[map_date] = inflows_map.get(map_date, 0) + qty_base

        oc_list.append({
            "doc":       str(row.get("Supplier Order Number", "")),
            "qty":       round(qty, 0),
            "qty_base":  round(qty_base, 0),
            "oc_um":     oc_um,
            "eta":       eta_orig,
            "eta_proy":  eta_proy,
            "proveedor": str(row.get("SupplierDescription", "")).strip(),
            "atrasada":  atrasada,
            "pos":       str(row["_pos"]),
        })

    # — SOLPED list —
    sol_rows = ctx.df_sol[(ctx.df_sol["Centro"].isin(centros)) & (ctx.df_sol["_mat"] == mat_str)].copy()
    sol_list = []
    for _, row in sol_rows.iterrows():
        fecha_str, _, sin_lib = _sol_lib_info(row)
        sol_list.append({
            "doc":         str(row.get("Solicitud de pedido", "")),
            "qty":         round(sf(row.get("Cantidad solicitada", 0)), 0),
            "um":          str(row.get("Unidad de medida", "UN")).strip(),
            "fecha":       fecha_str,
            "sin_liberar": sin_lib,
        })

    # — Saturday projection + state —
    sabs = project_inventory(
        saldo, mrp_months, adu_plan, inflows_map,
        ctx.sabados, z_roja_plan, z_am_plan, FECHA_SALDO,
    )
    quiebre_12sem, riesgo_12sem, primer_quiebre_dias, sin_tiempo = detect_quiebre_info(
        sabs, lt, FECHA_SALDO,
    )
    estado = classify_inventory_state(
        saldo, z_roja_plan, z_am_plan,
        quiebre_12sem, sin_tiempo, sin_demanda, net_flow,
    )

    cob_hoy = metrics["cob_hoy"]
    brk_sin = round(cob_hoy, 1) if cob_hoy < 9999 and estado != "SIN_CONSUMO" else None

    return {
        "oc_list":             oc_list,
        "sol_list":            sol_list,
        "sabs":                sabs,
        "quiebre_12sem":       quiebre_12sem,
        "riesgo_12sem":        riesgo_12sem,
        "primer_quiebre_dias": primer_quiebre_dias,
        "sin_tiempo":          sin_tiempo,
        "estado":              estado,
        "brk_sin":             brk_sin,
    }


# ───────────────────────────────────────────────────────────────────────────────
#  Step 4 — build_material_output
# ───────────────────────────────────────────────────────────────────────────────

def build_material_output(
    mat_str: str,
    mat_data: dict,
    metrics: dict,
    projections: dict,
) -> dict:
    """
    Assemble the final material record consumed by the HTML dashboard.

    Parameters
    ----------
    mat_str     : str
    mat_data    : dict   Output of extract_material_data.
    metrics     : dict   Output of compute_inventory_metrics.
    projections : dict   Output of compute_projections.

    Returns
    -------
    dict  — exact same structure as the original build_material return value.
    """
    return {
        "mat":       mat_str,
        "desc":      mat_data["desc"],
        "um":        mat_data["um"],
        "lt":        mat_data["lt"],
        "moq":       round(mat_data["moq"], 2),
        "inv_ini":   round(metrics["inv_ini"], 0),
        "ingresos":  round(metrics["ingresos"], 0),
        "consumos":  round(metrics["consumos"], 0),
        "saldo":     round(metrics["saldo"], 0),
        "adu":       round(mat_data["adu"], 3),
        "adu_hist":  metrics["adu_hist"],
        "adu_plan":  round(metrics["adu_plan"], 3),
        "cob_hoy":   metrics["cob_hoy"],
        "z_roja":    round(mat_data["z_roja"], 0),
        "z_am":      round(mat_data["z_am"], 0),
        "z_verde":   round(mat_data["z_verde"], 0),
        "z_roja_d":  metrics["z_roja_d"],
        "z_am_d":    metrics["z_am_d"],
        "z_verde_d": metrics["z_verde_d"],
        "zona_b2":   mat_data["zona_b2"],
        "mat_group": mat_data["mat_group"],
        "estado":              projections["estado"],
        "brk_sin":             projections["brk_sin"],
        "riesgo_12sem":        projections["riesgo_12sem"],
        "quiebre_12sem":       projections["quiebre_12sem"],
        "sin_tiempo":          projections["sin_tiempo"],
        "primer_quiebre_dias": projections["primer_quiebre_dias"],
        "sabados":   projections["sabs"],
        "oc_list":   projections["oc_list"],
        "sol_list":  projections["sol_list"],
        "mrp_m0":    round(mat_data["mrp_m0"], 0),
        "linea_fab": mat_data["linea_fab"],
        "importado": mat_data["importado"],
        "lt_detalle": mat_data["lt_detalle"],
        "impact":    "",
        "action":    "",
    }


# ───────────────────────────────────────────────────────────────────────────────
#  Entry point
# ───────────────────────────────────────────────────────────────────────────────

def build_material(
    mat_str: str,
    centros: list[int],
    mrp_df: pd.DataFrame,
    b2_df: pd.DataFrame,
    ctx: MaterialContext,
) -> dict | None:
    """
    Build the complete planning record for one material × centro group.

    Returns None for materials with no demand AND no stock (not shown in dashboard).

    Parameters
    ----------
    mat_str : str         Normalised material code.
    centros : list[int]   SAP plant codes for this group.
    mrp_df  : pd.DataFrame  MRP filtered to this centro group.
    b2_df   : pd.DataFrame  B2Wise filtered to this centro group.
    ctx     : MaterialContext  Shared DataFrames and calendar.

    Returns
    -------
    dict | None
    """
    mat_data = extract_material_data(
        mat_str, centros, mrp_df, b2_df, ctx.df_lin, ctx.origen_lookup,
    )
    metrics = compute_inventory_metrics(mat_str, centros, mat_data, ctx)

    # Filter: no demand + no stock → skip
    if metrics["sin_demanda"] and metrics["saldo"] <= 0:
        return None

    projections = compute_projections(mat_str, centros, mat_data, metrics, ctx)

    return build_material_output(mat_str, mat_data, metrics, projections)
