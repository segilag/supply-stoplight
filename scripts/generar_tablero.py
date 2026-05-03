# -*- coding: utf-8 -*-
"""
generar_tablero.py  —  Supply Stoplight v5
Generates supply_stoplight.html by reading Excel files from the Datos/ directory.

Usage:  python generar_tablero.py [--fecha YYYY-MM-DD]
"""

import argparse
from datetime import date, timedelta
from pathlib import Path

# ── Parse --fecha BEFORE importing downstream modules ─────────────────────────
# preprocess.py and material_builder.py capture settings values at import time,
# so the override must happen first.
def _parse_args():
    p = argparse.ArgumentParser(description="Supply Stoplight v5")
    p.add_argument(
        "--fecha", metavar="YYYY-MM-DD",
        help="Override FECHA_SALDO/FECHA_B2WISE (default: value in settings.py)"
    )
    return p.parse_args()

_args = _parse_args()

import settings as _settings          # import the module object (not its names)

if _args.fecha:
    try:
        _settings.apply_fecha(date.fromisoformat(_args.fecha))
        print(f"  Date override applied: {_args.fecha}")
    except ValueError:
        print(f"  WARNING: invalid date '{_args.fecha}' — using default date.")

# ── Now import downstream modules (they will capture the updated values) ───────
import pandas as pd

from settings import (
    FECHA_SALDO, FECHA_B2WISE,
    N_SABADOS,
    MESES_ES, MES_M0,
)
from loaders import load_all_data
from preprocess import preprocess_data
from material_builder import (
    MaterialContext, build_material,
    sf, _oc_fecha_info, _sol_lib_info,
)
from json_builder import build_json
from html_builder import write_html, build_impactos_xlsx_b64

BASE  = Path(__file__).parent.parent   # project root
DATOS = BASE / "Datos"
OUT   = BASE / "supply_stoplight.html"

# ═══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def next_saturdays(from_date: date, n: int):
    sats, d = [], from_date
    while len(sats) < n:
        d += timedelta(days=1)
        if d.weekday() == 5:
            sats.append(d)
    return sats


# ═══════════════════════════════════════════════════════════════════════════════
#  LOAD FILES
# ═══════════════════════════════════════════════════════════════════════════════
_raw          = load_all_data(DATOS)
_clean        = preprocess_data(_raw)
df_mrp        = _clean["mrp"]
df_b2         = _clean["b2wise"]
b2_soh        = _clean["b2_soh"]
df_sol        = _clean["solped"]
df_b2oc       = _clean["b2wise_oc"]
origen_lookup    = _clean["origen_lookup"]
df_lin           = _clean["lineas_fab"]
impactos_lookup  = _clean.get("impactos_lookup", {})
centros_master   = _clean.get("centros_master", {})
cons_mes      = _clean["cons_mes"]
cons_3m       = _clean["cons_3m"]
ing_mes       = _clean["ing_mes"]
ini_agg       = _clean["ini_agg"]
sdo_agg       = _clean["sdo_agg"]

SABADOS = next_saturdays(FECHA_SALDO, N_SABADOS)

SEM_WEEKS = [
    {"date":     s.strftime("%Y-%m-%d"),
     "label":    s.strftime("%d/%m"),
     "month":    MESES_ES[s.month],
     "week_num": int(s.strftime("%W"))}
    for s in SABADOS
]

# Shared context for all build_material calls (created once, reused per plant)
_CTX = MaterialContext(
    df_lin=df_lin,           origen_lookup=origen_lookup,
    cons_3m=cons_3m,         ini_agg=ini_agg,
    ing_mes=ing_mes,         cons_mes=cons_mes,
    b2_soh=b2_soh,           sdo_agg=sdo_agg,
    df_b2oc=df_b2oc,         df_sol=df_sol,
    sabados=SABADOS,
)

# ═══════════════════════════════════════════════════════════════════════════════
#  BUILD PO RECORDS
# ═══════════════════════════════════════════════════════════════════════════════

def build_oc(centros):
    rows = df_b2oc[df_b2oc["_centro"].isin(centros)].copy()
    result = []
    for _, r in rows.iterrows():
        qty_pend = float(r["_qty_pend"])
        if qty_pend <= 0:
            continue
        fd = r["_fecha_ent"]
        fecha_orig, fecha_proy, atrasada, dias_atraso = _oc_fecha_info(fd)
        fecha_doc = r.get("Order Creation Date")
        fecha_doc_str = pd.to_datetime(fecha_doc).strftime("%Y-%m-%d") if pd.notna(fecha_doc) else ""
        result.append({
            "doc":             str(r.get("Supplier Order Number", "")),
            "mat":             r["_mat"],
            "desc":            str(r.get("Part Description", "")).strip(),
            "proveedor":       str(r.get("SupplierDescription", "")).strip(),
            "qty_pedida":      round(float(r["_qty_ped"]), 0),
            "qty_pend":        round(qty_pend, 0),
            "um":              str(r["_um"]),
            "fecha_doc":       fecha_doc_str,
            "fecha_entrega":   fecha_orig,
            "fecha_proyectada": fecha_proy,
            "dias_atraso":     dias_atraso,
            "atrasada":        atrasada,
        })
    return result

# ═══════════════════════════════════════════════════════════════════════════════
#  BUILD PR RECORDS
# ═══════════════════════════════════════════════════════════════════════════════

def build_sol(centros):
    rows = df_sol[df_sol["Centro"].isin(centros)].copy()
    result = []
    for _, r in rows.iterrows():
        fecha_ent_str, ind_lib, sin_lib = _sol_lib_info(r)
        result.append({
            "doc":          str(r.get("Solicitud de pedido", "")),
            "mat":          r["_mat"],
            "desc":         str(r.get("Texto breve", "")).strip(),
            "qty":          round(sf(r.get("Cantidad solicitada", 0)), 0),
            "um":           str(r.get("Unidad de medida", "UN")).strip(),
            "fecha_entrega": fecha_ent_str,
            "ind_lib":      ind_lib,
            "sin_liberar":  sin_lib,
        })
    return result

# ═══════════════════════════════════════════════════════════════════════════════
#  BUILD PORTFOLIO FOR A SET OF SAP PLANTS
# ═══════════════════════════════════════════════════════════════════════════════

def build_centro_data(centros, label, nombre, flag):
    mrp_c = df_mrp[df_mrp["CENTRO"].isin(centros)].copy()
    b2_c  = df_b2[df_b2["_centro"].isin(centros)].copy()

    mats = sorted(set(mrp_c["_mat"].unique()) | set(b2_c["_mat"].unique()))
    print(f"  {label}: {len(mats)} materials in portfolio")

    materiales = [m for m in (build_material(m, centros, mrp_c, b2_c, _CTX) for m in mats)
                  if m is not None and not str(m.get("mat_group", "")).startswith(("600", "700"))]
    oc   = build_oc(centros)
    sol  = build_sol(centros)

    quiebran_30 = sum(1 for m in materiales if m["brk_sin"] is not None and m["brk_sin"] <= 30)
    quiebran_60 = sum(1 for m in materiales if m["brk_sin"] is not None and m["brk_sin"] <= 60)
    mat_groups  = sorted(set(m["mat_group"] for m in materiales if m["mat_group"]))
    lineas_fab  = sorted(set(m["linea_fab"] for m in materiales if m["linea_fab"]))

    return {
        "nombre":      nombre,
        "flag":        flag,
        "centros_sap": [str(c) for c in centros],
        "materiales":  materiales,
        "oc":          oc,
        "sol":         sol,
        "quiebran_30": quiebran_30,
        "quiebran_60": quiebran_60,
        "mat_groups":  mat_groups,
        "lineas_fab":  lineas_fab,
    }

def _apply_impactos(centro_data):
    """Injects impact/action from impactos_lookup into the plant's materials.
    Tries matching by plant name and by SAP code (in that order)."""
    nombre = centro_data["nombre"].upper()
    sap_codes = [str(c) for c in centro_data.get("centros_sap", [])]
    n = 0
    for m in centro_data["materiales"]:
        mat = m["mat"]
        inp = impactos_lookup.get((mat, nombre), {})
        if not inp:
            for code in sap_codes:
                inp = impactos_lookup.get((mat, code), {})
                if inp:
                    break
        m["impact"] = inp.get("impact", "")
        m["action"] = inp.get("action", "")
        if m["impact"] or m["action"]:
            n += 1
    return n

# ═══════════════════════════════════════════════════════════════════════════════
#  BUILD PORTFOLIOS DYNAMICALLY FROM centros_master
# ═══════════════════════════════════════════════════════════════════════════════

# Plants with actual data in MRP or ERP
_data_centros = set(df_mrp["CENTRO"].dropna().astype(int).unique()) | \
                set(df_b2["_centro"].dropna().astype(int).unique())

# If no centros_master, build fallback with discovered plants
if not centros_master:
    print("  WARNING: centros_master empty — using plants discovered in data")
    for _c in sorted(_data_centros):
        centros_master[_c] = {
            "pais": "Unknown", "descripcion": str(_c),
            "nombre": str(_c), "flag": "??",
        }

all_centro_data = {}   # key: str(code) → centro_data dict
paises_meta     = {}   # key: country → {flag, centros:[str_keys]}

for _code in sorted(_data_centros):
    _meta = centros_master.get(_code)
    if _meta is None:
        print(f"  WARNING: plant {_code} not found in centros_master — skipped")
        continue
    _label  = str(_code)
    _nombre = _meta["nombre"]
    _flag   = _meta["flag"]
    _pais   = _meta["pais"]

    print(f"Building {_nombre} ({_code})...")
    _cd = build_centro_data([_code], _label, _nombre, _flag)
    _cd["pais"] = _pais
    all_centro_data[_label] = _cd

    if _pais not in paises_meta:
        paises_meta[_pais] = {"flag": _flag, "centros": []}
    paises_meta[_pais]["centros"].append(_label)

if impactos_lookup:
    _n_imp = sum(_apply_impactos(d) for d in all_centro_data.values())
    print(f"  Impact alerts applied: {_n_imp} materials with data")
else:
    for d in all_centro_data.values():
        _apply_impactos(d)  # sets empty defaults

# ═══════════════════════════════════════════════════════════════════════════════
#  GENERATE OUTPUT
# ═══════════════════════════════════════════════════════════════════════════════

json_str = build_json(
    centros=all_centro_data,
    meta={
        "fecha_saldo":  FECHA_SALDO.strftime("%d-%b-%Y"),
        "fecha_b2wise": FECHA_B2WISE.strftime("%d-%b-%Y"),
        "mes_m0":       MES_M0,
        "sem_weeks":    SEM_WEEKS,
        "paises":       paises_meta,
    },
)

print("Generating Impact Report (Excel)...")
_country_centros = [
    {"pais": _pais.upper(), "centros": [all_centro_data[k] for k in _pmeta["centros"]]}
    for _pais, _pmeta in paises_meta.items()
    if _pmeta["centros"]
]
impactos_b64 = build_impactos_xlsx_b64(
    country_centros=_country_centros,
    fecha_saldo=FECHA_SALDO.strftime("%d-%b-%Y"),
)

write_html(json_str, OUT, impactos_b64=impactos_b64)
print(f"\nGenerated OK: {OUT}")
for _key, _d in all_centro_data.items():
    print(f"  {_d['nombre']} ({_key}): {len(_d['materiales'])} materials | {len(_d['oc'])} POs | {len(_d['sol'])} PRs")
_all_sol = [s for _d in all_centro_data.values() for s in _d["sol"]]
_all_oc  = [o for _d in all_centro_data.values() for o in _d["oc"]]
print(f"  Overdue POs: {sum(1 for o in _all_oc if o['atrasada'])} | Unreleased PRs: {sum(1 for s in _all_sol if s['sin_liberar'])}")
