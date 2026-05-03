# -*- coding: utf-8 -*-
"""
preprocess.py  —  Supply Stoplight v5
Normalizes keys and pre-aggregates the raw DataFrames returned by loaders.

Usage:
    from preprocess import preprocess_data
    clean = preprocess_data(raw_data)
"""

import pandas as pd
from datetime import timedelta

from loaders import mat_key
from settings import FECHA_SALDO, MES, AÑO


# ───────────────────────────────────────────────────────────────────────────────
#  Per-DataFrame normalization
# ───────────────────────────────────────────────────────────────────────────────

def _norm_mrp(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega _mat (código SAP normalizado)."""
    df = df.copy()
    df["_mat"] = df["Código SAP"].apply(mat_key)
    return df


def _norm_b2wise(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Agrega _centro y _mat a B2Wise.
    Devuelve (df_normalizado, b2_soh_deduplicado).

    SOH y zonas son por material+centro, no por lote de compra:
    se elimina duplicados tomando un único valor por _mat+_centro.
    """
    df = df.copy()
    df["_centro"] = df["Warehouse Code"].apply(lambda v: int(float(v)) if pd.notna(v) else 0)
    df["_mat"]    = df["Part Code"].apply(mat_key)

    b2_soh = (
        df.drop_duplicates(subset=["_mat", "_centro"])
          [["_centro", "_mat", "SOH"]]
          .rename(columns={"SOH": "saldo_b2w"})
    )
    return df, b2_soh


def _norm_inventario(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega _mat al inventario inicial."""
    df = df.copy()
    df["_mat"] = df["Material"].apply(mat_key)
    return df


def _norm_saldo(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega _mat al saldo actual."""
    df = df.copy()
    df["_mat"] = df["Material"].apply(mat_key)
    return df


def _norm_consumos(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega _mat y _fecha (Fe.contabilización) a consumos."""
    df = df.copy()
    df["_mat"]   = df["Material"].apply(mat_key)
    df["_fecha"] = pd.to_datetime(df["Fe.contabilización"], errors="coerce")
    return df


def _norm_ingresos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Excludes receipts from plant 4402 (warehouses starting with 'CH'),
    then adds _mat and _fecha.
    """
    df = df.copy()
    _ce  = pd.to_numeric(df["Ce."], errors="coerce").fillna(0).astype(int)
    _alm = df["Alm."].fillna("").astype(str).str.upper()
    df = df[~((_ce == 4402) | (_alm.str.startswith("CH")))].copy()
    df["_mat"]   = df["Material"].apply(mat_key)
    df["_fecha"] = pd.to_datetime(df["Fe.contab."], errors="coerce")
    return df


def _norm_solped(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega _mat a solicitudes de pedido."""
    df = df.copy()
    df["_mat"] = df["Material"].apply(mat_key)
    return df


# ───────────────────────────────────────────────────────────────────────────────
#  Pre-aggregations
# ───────────────────────────────────────────────────────────────────────────────

def _agg_consumos_mes(df_con: pd.DataFrame) -> pd.DataFrame:
    """Consumo total del mes en curso por Centro + material."""
    return (
        df_con[(df_con["_fecha"].dt.month == MES) & (df_con["_fecha"].dt.year == AÑO)]
        .groupby(["Centro", "_mat"])["Ctd.en UM entrada"]
        .sum()
        .abs()
        .reset_index()
        .rename(columns={"Ctd.en UM entrada": "consumos_mes"})
    )


def _agg_consumos_3m(df_con: pd.DataFrame) -> pd.DataFrame:
    """Consumo real de los últimos 90 días (valida fallback ADU B2Wise)."""
    fecha_3m = FECHA_SALDO - timedelta(days=90)
    return (
        df_con[df_con["_fecha"] >= pd.Timestamp(fecha_3m)]
        .groupby(["Centro", "_mat"])["Ctd.en UM entrada"]
        .sum()
        .abs()
        .reset_index()
        .rename(columns={"Ctd.en UM entrada": "consumos_3m"})
    )


def _agg_ingresos_mes(df_ing: pd.DataFrame) -> pd.DataFrame:
    """Ingresos totales del mes en curso por centro + material."""
    return (
        df_ing[(df_ing["_fecha"].dt.month == MES) & (df_ing["_fecha"].dt.year == AÑO)]
        .groupby(["Ce.", "_mat"])["Cantidad"]
        .sum()
        .reset_index()
        .rename(columns={"Ce.": "Centro_ing", "Cantidad": "ingresos_mes"})
    )


def _agg_inventario_inicial(df_ini: pd.DataFrame) -> pd.DataFrame:
    """Inventario inicial consolidado por centro + material."""
    return (
        df_ini
        .groupby(["Ce.", "_mat"])["Libre utiliz."]
        .sum()
        .reset_index()
        .rename(columns={"Ce.": "Centro_ini", "Libre utiliz.": "inv_ini"})
    )


def _agg_saldo_actual(df_sdo: pd.DataFrame) -> pd.DataFrame:
    """
    Saldo disponible = Libre utilización + Inspección de calidad,
    consolidado por Centro + material.
    """
    df = df_sdo.copy()
    df["_saldo_disp"] = df["Libre utilización"].fillna(0) + df["Inspecc.de calidad"].fillna(0)
    return (
        df
        .groupby(["Centro", "_mat"])["_saldo_disp"]
        .sum()
        .reset_index()
        .rename(columns={"_saldo_disp": "saldo"})
    )


# ───────────────────────────────────────────────────────────────────────────────
#  Main function
# ───────────────────────────────────────────────────────────────────────────────

def preprocess_data(raw_data: dict) -> dict:
    """
    Normaliza claves y pre-agrega todos los DataFrames crudos.

    Parameters
    ----------
    raw_data : dict
        Salida de loaders.load_all_data().

    Returns
    -------
    dict with keys:
        mrp, b2wise, b2_soh,
        inventario_inicial, saldo_actual, consumos, ingresos, solped,
        b2wise_oc, origen_lookup, lineas_fab,   ← passthrough de loaders
        cons_mes, cons_3m, ing_mes, ini_agg, sdo_agg
    """
    # — Normalización —
    df_mrp          = _norm_mrp(raw_data["mrp"])
    df_b2, b2_soh   = _norm_b2wise(raw_data["b2wise"])
    df_ini          = _norm_inventario(raw_data["inventario_inicial"])
    df_sdo          = _norm_saldo(raw_data["saldo_actual"])
    df_con          = _norm_consumos(raw_data["consumos"])
    df_ing          = _norm_ingresos(raw_data["ingresos"])
    df_sol          = _norm_solped(raw_data["solped"])

    # — Filtro Tipo_Material: solo conservar SOLPEDs de materiales activos —
    tipo_mat = raw_data.get("tipo_material_set", frozenset())
    if tipo_mat:
        _n0    = len(df_sol)
        df_sol = df_sol[df_sol.apply(
            lambda r: (r["_mat"], str(int(float(r["Centro"]))) if pd.notna(r.get("Centro")) else "0") in tipo_mat,
            axis=1
        )].copy()
        print(f"  PRs Material_Type filter: {_n0} -> {len(df_sol)} rows kept")

    # — Pre-agregaciones —
    cons_mes = _agg_consumos_mes(df_con)
    cons_3m  = _agg_consumos_3m(df_con)
    ing_mes  = _agg_ingresos_mes(df_ing)
    ini_agg  = _agg_inventario_inicial(df_ini)
    sdo_agg  = _agg_saldo_actual(df_sdo)

    return {
        # DataFrames normalizados
        "mrp":                df_mrp,
        "b2wise":             df_b2,
        "b2_soh":             b2_soh,
        "inventario_inicial": df_ini,
        "saldo_actual":       df_sdo,
        "consumos":           df_con,
        "ingresos":           df_ing,
        "solped":             df_sol,
        # Passthrough (ya normalizados en loaders)
        "b2wise_oc":          raw_data["b2wise_oc"],
        "origen_lookup":      raw_data["origen_lookup"],
        "lineas_fab":         raw_data["lineas_fab"],
        "impactos_lookup":    raw_data.get("impactos_lookup", {}),
        "centros_master":     raw_data.get("centros_master", {}),
        "tipo_material_set":  tipo_mat,
        # Pre-agregados
        "cons_mes":           cons_mes,
        "cons_3m":            cons_3m,
        "ing_mes":            ing_mes,
        "ini_agg":            ini_agg,
        "sdo_agg":            sdo_agg,
    }
