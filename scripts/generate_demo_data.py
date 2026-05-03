# -*- coding: utf-8 -*-
"""
generate_demo_data.py  —  Supply Stoplight v5
Creates anonymized demo Excel files in Datos/demo/ so the dashboard
can run without any real company data.

Usage:  python scripts/generate_demo_data.py
Output: Datos/demo/  (all files expected by the dashboard)
"""

import random
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import date, timedelta

random.seed(42)
np.random.seed(42)

# ── Plant definitions (6 plants: 2 CO, 2 PE, 2 BR) ───────────────────────────
PLANTS = [
    {"code": 4301, "name": "Plant 1", "country": "CO", "flag": "CO"},
    {"code": 4321, "name": "Plant 2", "country": "CO", "flag": "CO"},
    {"code": 4401, "name": "Plant 3", "country": "PE", "flag": "PE"},
    {"code": 4403, "name": "Plant 4", "country": "PE", "flag": "PE"},
    {"code": 4202, "name": "Plant 5", "country": "BR", "flag": "BR"},
    {"code": 4208, "name": "Plant 6", "country": "BR", "flag": "BR"},
]

# ── Material catalogue ────────────────────────────────────────────────────────
RAW_MATS = [
    {"code": "1001001", "desc": "Raw Material 1",  "um": "KG",  "lt": 45, "moq": 20000},
    {"code": "1001002", "desc": "Raw Material 2",  "um": "KG",  "lt": 60, "moq": 15000},
    {"code": "1001003", "desc": "Raw Material 3",  "um": "TO",  "lt": 30, "moq": 500},
    {"code": "1001004", "desc": "Raw Material 4",  "um": "KG",  "lt": 90, "moq": 10000},
    {"code": "1001005", "desc": "Raw Material 5",  "um": "KG",  "lt": 45, "moq": 25000},
    {"code": "1001006", "desc": "Raw Material 6",  "um": "KG",  "lt": 60, "moq": 5000},
    {"code": "1001007", "desc": "Raw Material 7",  "um": "TO",  "lt": 45, "moq": 200},
    {"code": "1001008", "desc": "Raw Material 8",  "um": "KG",  "lt": 30, "moq": 8000},
    {"code": "1001009", "desc": "Raw Material 9",  "um": "KG",  "lt": 75, "moq": 12000},
    {"code": "1001010", "desc": "Raw Material 10", "um": "KG",  "lt": 45, "moq": 6000},
    {"code": "1001011", "desc": "Raw Material 11", "um": "KG",  "lt": 60, "moq": 4000},
    {"code": "1001012", "desc": "Raw Material 12", "um": "TO",  "lt": 90, "moq": 100},
    {"code": "1001013", "desc": "Raw Material 13", "um": "KG",  "lt": 30, "moq": 3000},
    {"code": "1001014", "desc": "Raw Material 14", "um": "KG",  "lt": 45, "moq": 7000},
    {"code": "1001015", "desc": "Raw Material 15", "um": "KG",  "lt": 60, "moq": 9000},
]

PKG_MATS = [
    {"code": "2001001", "desc": "Packaging Material 1",  "um": "UN",  "lt": 21, "moq": 10000},
    {"code": "2001002", "desc": "Packaging Material 2",  "um": "UN",  "lt": 14, "moq": 5000},
    {"code": "2001003", "desc": "Packaging Material 3",  "um": "UN",  "lt": 21, "moq": 20000},
    {"code": "2001004", "desc": "Packaging Material 4",  "um": "KG",  "lt": 30, "moq": 2000},
    {"code": "2001005", "desc": "Packaging Material 5",  "um": "UN",  "lt": 14, "moq": 15000},
    {"code": "2001006", "desc": "Packaging Material 6",  "um": "UN",  "lt": 21, "moq": 8000},
    {"code": "2001007", "desc": "Packaging Material 7",  "um": "UN",  "lt": 14, "moq": 12000},
    {"code": "2001008", "desc": "Packaging Material 8",  "um": "UN",  "lt": 30, "moq": 3000},
]

ALL_MATS = RAW_MATS + PKG_MATS

# Each plant has a subset of materials (some shared, some unique)
PLANT_MATS = {
    4301: [m["code"] for m in RAW_MATS[:10]] + [m["code"] for m in PKG_MATS[:5]],
    4321: [m["code"] for m in RAW_MATS[3:12]] + [m["code"] for m in PKG_MATS[2:7]],
    4401: [m["code"] for m in RAW_MATS[:8]] + [m["code"] for m in PKG_MATS[:4]],
    4403: [m["code"] for m in RAW_MATS[5:14]] + [m["code"] for m in PKG_MATS[3:8]],
    4202: [m["code"] for m in RAW_MATS[2:12]] + [m["code"] for m in PKG_MATS[:6]],
    4208: [m["code"] for m in RAW_MATS[6:15]] + [m["code"] for m in PKG_MATS[1:7]],
}

MAT_MAP = {m["code"]: m for m in ALL_MATS}

PROD_LINES = ["Line A", "Line B", "Line C", "Line D"]
SUPPLIERS = [
    "Alpha Chemicals Ltd", "Beta Supplies SA", "Gamma Materials Corp",
    "Delta Resources Inc", "Epsilon Industrial", "Zeta Packaging Co",
]

TODAY = date(2026, 4, 24)


def _adu(mat_code: str, scale: float = 1.0) -> float:
    m = MAT_MAP[mat_code]
    base = {"KG": 800, "TO": 0.8, "UN": 1500, "G": 800000}[m["um"]]
    return round(base * scale * random.uniform(0.6, 1.4), 1)


def _stock(adu: float, lt: int, factor: float = 1.5) -> float:
    return round(adu * lt * factor * random.uniform(0.3, 2.5), 0)


# ── 1. MRP.xlsx ───────────────────────────────────────────────────────────────
def make_mrp() -> pd.DataFrame:
    rows = []
    for plant in PLANTS:
        for mat_code in PLANT_MATS[plant["code"]]:
            m = MAT_MAP[mat_code]
            adu = _adu(mat_code)
            days = [round(adu * 30 * random.uniform(0.7, 1.3), 0) for _ in range(6)]
            rows.append({
                "PAÍS":         plant["country"],
                "CENTRO":       plant["code"],
                "PLANTA":       plant["name"],
                "Código SAP":   mat_code,
                "Descripción":  m["desc"],
                "LT TOTAL":     m["lt"],
                "MOQ":          m["moq"],
                "MULTIPLO":     m["moq"],
                "Frecuencia 12m": round(12 / random.randint(1, 4), 1),
                "Promedio 12m": round(adu * 30, 0),
                "MRP M+0":      days[0],
                "MRP M+1":      days[1],
                "MRP M+2":      days[2],
                "MRP M+3":      days[3],
                "MRP M+4":      days[4],
                "MRP M+5":      days[5],
            })
    return pd.DataFrame(rows)


# ── 2. Inventario Inicial.xlsx ────────────────────────────────────────────────
def make_inv_inicial() -> pd.DataFrame:
    rows = []
    for plant in PLANTS:
        for mat_code in PLANT_MATS[plant["code"]]:
            m = MAT_MAP[mat_code]
            adu = _adu(mat_code)
            rows.append({
                "Ce.":          plant["code"],
                "Material":     mat_code,
                "Descripción":  m["desc"],
                "Libre utiliz.": _stock(adu, m["lt"], random.uniform(0.8, 2.0)),
                "UM":           m["um"],
            })
    return pd.DataFrame(rows)


# ── 3. Saldo Actual.xlsx ──────────────────────────────────────────────────────
def make_saldo_actual() -> pd.DataFrame:
    rows = []
    for plant in PLANTS:
        for mat_code in PLANT_MATS[plant["code"]]:
            m = MAT_MAP[mat_code]
            adu = _adu(mat_code)
            libre = _stock(adu, m["lt"], random.uniform(0.5, 2.5))
            insp  = libre * random.uniform(0, 0.1) if random.random() > 0.8 else 0
            rows.append({
                "Centro":              plant["code"],
                "Material":            mat_code,
                "Texto breve":         m["desc"],
                "Almacén":             f"{plant['code']}01",
                "Libre utilización":   libre,
                "Inspecc.de calidad":  round(insp, 0),
                "Bloqueado":           0,
                "Stock tránsito":      0,
                "UM":                  m["um"],
            })
    return pd.DataFrame(rows)


# ── 4. Consumos materiales.xlsx ───────────────────────────────────────────────
def make_consumos() -> pd.DataFrame:
    rows = []
    for plant in PLANTS:
        for mat_code in PLANT_MATS[plant["code"]]:
            m = MAT_MAP[mat_code]
            adu = _adu(mat_code)
            for days_back in range(90, 0, -1):
                if random.random() > 0.3:
                    continue
                cons_date = TODAY - timedelta(days=days_back)
                qty = round(adu * random.uniform(0.5, 3.0), 1)
                rows.append({
                    "Centro":             plant["code"],
                    "Material":           mat_code,
                    "Almacén":            f"{plant['code']}01",
                    "Clase movimiento":   "261",
                    "Fe.contabilización": cons_date,
                    "Ctd.en UM entrada":  -qty,
                    "UM":                 m["um"],
                })
    return pd.DataFrame(rows)


# ── 5. BD_Sp_Pedidos_Compra.xlsx ─────────────────────────────────────────────
def make_solped() -> pd.DataFrame:
    rows = []
    pr_num = 1000000
    for plant in PLANTS:
        for mat_code in PLANT_MATS[plant["code"]]:
            if random.random() > 0.4:
                continue
            m = MAT_MAP[mat_code]
            adu = _adu(mat_code)
            lt_days = m["lt"]
            eta = TODAY + timedelta(days=lt_days + random.randint(-10, 20))
            qty = round(adu * lt_days * random.uniform(1.0, 2.5) / (m["moq"]), 0) * m["moq"]
            liberada = random.random() > 0.3
            pr_num += 1
            rows.append({
                "Solicitud de pedido":  str(pr_num),
                "Centro":               plant["code"],
                "Almacén":              f"{plant['code']}01",
                "Material":             mat_code,
                "Texto breve":          m["desc"],
                "Cantidad solicitada":  max(qty, m["moq"]),
                "Unidad de medida":     m["um"],
                "Fecha entrega":        eta,
                "Status":               "L" if liberada else "",
                "Indicador liberación":  "" if liberada else "X",
            })
    return pd.DataFrame(rows)


def make_ingresos() -> pd.DataFrame:
    rows = []
    for plant in PLANTS:
        for mat_code in PLANT_MATS[plant["code"]]:
            m = MAT_MAP[mat_code]
            adu = _adu(mat_code)
            for days_back in range(90, 0, -7):
                if random.random() > 0.4:
                    continue
                ing_date = TODAY - timedelta(days=days_back)
                qty = round(m["moq"] * random.uniform(0.8, 2.0), 0)
                rows.append({
                    "Ce.":        plant["code"],
                    "Alm.":       f"{plant['code']}01",
                    "Material":   mat_code,
                    "Fe.contab.": ing_date,
                    "Cantidad":   qty,
                })
    return pd.DataFrame(rows)


def make_tipo_material() -> pd.DataFrame:
    rows = []
    for plant in PLANTS:
        for mat_code in PLANT_MATS[plant["code"]]:
            rows.append({"Material": mat_code, "Ce.": plant["code"]})
    return pd.DataFrame(rows)


# ── 6. CORTE B2WISE.xlsx ──────────────────────────────────────────────────────
def make_b2wise() -> pd.DataFrame:
    rows = []
    for plant in PLANTS:
        for mat_code in PLANT_MATS[plant["code"]]:
            m = MAT_MAP[mat_code]
            adu = _adu(mat_code, scale=random.uniform(0.8, 1.2))
            soh = _stock(adu, m["lt"], random.uniform(0.5, 2.5))
            rows.append({
                "Part Code":        mat_code,
                "Part Description": m["desc"],
                "Warehouse Code":   plant["code"],
                "SOH":              soh,
                "ADU":              adu,
                "DLT":              m["lt"],
                "Days Cover":       round(soh / adu, 1) if adu > 0 else 999,
                "UOM":              m["um"],
                "Zone":             random.choice(["Green", "Yellow", "Red"]),
                "TOR":              round(m["lt"] * random.uniform(0.3, 0.5), 0),
                "TOY":              round(m["lt"] * random.uniform(0.5, 0.7), 0),
                "TOG":              round(m["lt"] * random.uniform(0.7, 1.0), 0),
            })
    return pd.DataFrame(rows)


# ── 7. Ordenes de compra B2WISE.xlsx ─────────────────────────────────────────
def make_oc() -> pd.DataFrame:
    rows = []
    po_num = 4000000
    for plant in PLANTS:
        for mat_code in PLANT_MATS[plant["code"]]:
            if random.random() > 0.45:
                continue
            m = MAT_MAP[mat_code]
            adu = _adu(mat_code)
            qty = round(m["moq"] * random.uniform(1.0, 3.0), 0)
            days_offset = random.randint(-15, m["lt"] + 30)
            eta = TODAY + timedelta(days=days_offset)
            po_num += random.randint(1, 5)
            rows.append({
                "Supplier Order Number":           str(po_num),
                "Line_Number":                     f"{po_num}0010",
                "Part Code":                       mat_code,
                "Part Description":                m["desc"],
                "Warehouse Code":                  plant["code"],
                "SupplierDescription":             random.choice(SUPPLIERS),
                "Supplier Order Arrival Date":     eta,
                "Supplier Order Qty Outstanding":  qty,
                "SupplierOrderQuantityOrdered":    qty,
                "Order Creation Date":             TODAY - timedelta(days=random.randint(10, 60)),
            })
    return pd.DataFrame(rows)


# ── 8. 0. Base / Lineas Fabricacion.xlsx ────────────────────────────────────
def make_lineas_fab() -> pd.DataFrame:
    rows = []
    for plant in PLANTS:
        for mat_code in PLANT_MATS[plant["code"]]:
            m = MAT_MAP[mat_code]
            line = random.choice(PROD_LINES)
            rows.append({
                "Centro":     plant["code"],
                "Codigo SAP": mat_code,
                "Descripcion Material": m["desc"],
                "Linea Fabricacion": line,
            })
    return pd.DataFrame(rows)


def make_centros() -> pd.DataFrame:
    rows = []
    for plant in PLANTS:
        rows.append({
            "Pais":        plant["country"],
            "Centro":      plant["code"],
            "Descripcion": f"{plant['country']} - {plant['name']}",
            "Flag":        plant["flag"],
        })
    return pd.DataFrame(rows)


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    base = Path(__file__).parent.parent / "Datos" / "demo"
    base.mkdir(parents=True, exist_ok=True)
    (base / "0. Base").mkdir(exist_ok=True)

    print("Generating demo data…")

    df_mrp = make_mrp()
    df_mrp.to_excel(base / "MRP.xlsx", sheet_name="MRP", index=False)
    print(f"  MRP.xlsx: {len(df_mrp)} rows")

    df_ini = make_inv_inicial()
    df_ini.to_excel(base / "Inventario Inicial.xlsx", sheet_name="Cierre_Inv", index=False)
    print(f"  Inventario Inicial.xlsx: {len(df_ini)} rows")

    df_sdo = make_saldo_actual()
    df_sdo.to_excel(base / "Saldo Actual.xlsx", index=False)
    print(f"  Saldo Actual.xlsx: {len(df_sdo)} rows")

    df_con = make_consumos()
    df_con.to_excel(base / "Consumos materiales.xlsx", index=False)
    print(f"  Consumos materiales.xlsx: {len(df_con)} rows")

    df_sol = make_solped()
    df_ing = make_ingresos()
    df_tipo = make_tipo_material()
    with pd.ExcelWriter(base / "BD_Sp_Pedidos_Compra.xlsx", engine="openpyxl") as w:
        df_sol.to_excel(w, sheet_name="SolPed (ME5A)", index=False)
        df_ing.to_excel(w, sheet_name="Ingresos (SQVI)", index=False)
        df_tipo.to_excel(w, sheet_name="Tipo_Material", index=False)
    print(f"  BD_Sp_Pedidos_Compra.xlsx: {len(df_sol)} PRs, {len(df_ing)} receipts")

    df_b2 = make_b2wise()
    df_b2.to_excel(base / "CORTE B2WISE.xlsx", index=False)
    print(f"  CORTE B2WISE.xlsx: {len(df_b2)} rows")

    df_oc = make_oc()
    df_oc.to_excel(base / "Ordenes de compra B2WISE.xlsx", sheet_name="Sheet1", index=False)
    print(f"  Ordenes de compra B2WISE.xlsx: {len(df_oc)} POs")

    df_lin = make_lineas_fab()
    df_cen = make_centros()
    with pd.ExcelWriter(base / "0. Base" / "Líneas Fabricación.xlsx", engine="openpyxl") as w:
        df_lin.to_excel(w, sheet_name="Lineas de fabricacion", index=False)
        df_cen.to_excel(w, sheet_name="Centros", index=False)
    print(f"  Lineas Fabricacion.xlsx: {len(df_lin)} lines, {len(df_cen)} plants")

    print(f"\nDemo data generated in: {base}")
    print("Run: python scripts/generar_tablero.py")
    print("(set DATOS = BASE / 'Datos' / 'demo' in generar_tablero.py to use demo data)")


if __name__ == "__main__":
    main()
