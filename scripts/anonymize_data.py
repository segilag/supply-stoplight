"""
anonymize_data.py
Replaces all company-identifying names in Datos/ Excel files.
Overwrites originals in place.

Plant mapping:
  4301 -> Plant 1 (Colombia)   4321 -> Plant 2 (Colombia)
  4401 -> Plant 3 (Peru)       4403 -> Plant 4 (Peru)
  4202 -> Plant 5 (Brazil)     4208 -> Plant 6 (Brazil)

Material naming (from B2Wise Mat Type):
  ZRO1  -> Raw Material 1, 2, 3...
  ZVE1  -> Packaging Material 1, 2, 3...
  other -> Material 1, 2, 3...

Supplier naming: Supplier 1, 2, 3...
"""
import sys
from pathlib import Path
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

BASE = Path(__file__).parent.parent / "Datos"

# ── Constants ─────────────────────────────────────────────────────────────────
TARGET_PLANTS = {4301, 4321, 4401, 4403, 4202, 4208}

CENTRO_TO_NAME = {
    4301: "Plant 1", 4321: "Plant 2",
    4401: "Plant 3", 4403: "Plant 4",
    4202: "Plant 5", 4208: "Plant 6",
}
CENTRO_TO_COUNTRY = {
    4301: "Colombia", 4321: "Colombia",
    4401: "Peru",     4403: "Peru",
    4202: "Brazil",   4208: "Brazil",
}


# ── Helpers ───────────────────────────────────────────────────────────────────
def _centro_int(val):
    try:
        return int(str(val).strip()[:4])
    except (ValueError, TypeError):
        return None


def _in_target(val):
    return _centro_int(val) in TARGET_PLANTS


def _update_excel(path, modifications):
    """
    Load all sheets, apply modifications dict {sheet_name: df}, write back.
    modifications: {sheet_name: DataFrame}  — only listed sheets are replaced.
    """
    all_sheets = pd.read_excel(path, sheet_name=None, dtype=str)
    for sheet, df in modifications.items():
        if sheet in all_sheets:
            all_sheets[sheet] = df
        else:
            print(f"  WARNING: sheet '{sheet}' not found in {path.name}")
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for name, df in all_sheets.items():
            df.to_excel(w, sheet_name=name, index=False)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Build material mapping from CORTE B2WISE
# ═══════════════════════════════════════════════════════════════════════════════
print("Building material mapping from B2Wise...")
b2_raw = pd.read_excel(BASE / "CORTE B2WISE.xlsx", sheet_name="Sheet1", dtype=str)
b2_raw["_centro"] = b2_raw["Warehouse Code"].apply(_centro_int)
b2_target = b2_raw[b2_raw["_centro"].isin(TARGET_PLANTS)].copy()

mat_map = {}   # str(SAP code) -> generic name
raw_n = pkg_n = oth_n = 1

for _, row in b2_target.drop_duplicates("Part Code").iterrows():
    code = str(row["Part Code"]).strip()
    mat_type = str(row.get("Mat Type", "")).strip()
    if code in mat_map:
        continue
    if mat_type == "ZRO1":
        mat_map[code] = f"Raw Material {raw_n}";      raw_n += 1
    elif mat_type == "ZVE1":
        mat_map[code] = f"Packaging Material {pkg_n}"; pkg_n += 1
    else:
        mat_map[code] = f"Material {oth_n}";           oth_n += 1

# Supplement with MRP codes not in B2Wise
mrp_all = pd.read_excel(BASE / "MRP.xlsx", sheet_name="MRP", dtype=str)
mrp_all["_centro"] = mrp_all["CENTRO"].apply(_centro_int)
mrp_target = mrp_all[mrp_all["_centro"].isin(TARGET_PLANTS)].copy()

for code in mrp_target["Código SAP"].dropna().unique():
    code = str(code).strip()
    if code not in mat_map:
        mat_map[code] = f"Material {oth_n}"; oth_n += 1

print(f"  {len(mat_map)} materials mapped "
      f"({raw_n-1} raw, {pkg_n-1} packaging, {oth_n-1} other)")


# ── PLANTA city mapping (derive from actual data) ─────────────────────────────
planta_map = {}
_invalid = {"", "nan", "none", "false", "true"}
for _, row in mrp_target.iterrows():
    centro = _centro_int(row["CENTRO"])
    city   = str(row.get("PLANTA", "")).strip()
    if city and city.lower() not in _invalid and centro in CENTRO_TO_NAME:
        planta_map[city] = CENTRO_TO_NAME[centro]
print(f"  City->Plant map: {planta_map}")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Build supplier mapping
# ═══════════════════════════════════════════════════════════════════════════════
print("\nBuilding supplier mapping...")
supp_map = {}
supp_n = 1

def _add_supplier(name):
    global supp_n
    key = str(name).strip()
    if key and key not in supp_map and key.lower() not in ("nan", "none", ""):
        supp_map[key] = f"Supplier {supp_n}"; supp_n += 1

oc_raw = pd.read_excel(BASE / "Ordenes de compra B2WISE.xlsx", sheet_name="Sheet1", dtype=str)
oc_raw["_centro"] = oc_raw["Warehouse Code"].apply(_centro_int)
oc_target = oc_raw[oc_raw["_centro"].isin(TARGET_PLANTS)].copy()
for v in sorted(oc_target["SupplierDescription"].dropna().unique()):
    _add_supplier(v)

po_raw = pd.read_excel(BASE / "BD_Sp_Pedidos_Compra.xlsx", sheet_name="PO (ME2W)", dtype=str)
po_raw["_centro"] = po_raw["Centro"].apply(_centro_int)
po_target = po_raw[po_raw["_centro"].isin(TARGET_PLANTS)].copy()
for v in sorted(po_target["Proveedor/Centro suministrador"].dropna().unique()):
    _add_supplier(v)

print(f"  {len(supp_map)} suppliers mapped")


# ── Lookup helpers ────────────────────────────────────────────────────────────
def _mat(code):
    return mat_map.get(str(code).strip(), f"Material {code}")

def _plant(code):
    c = _centro_int(code)
    return CENTRO_TO_NAME.get(c, str(code))

def _country(plant_name_val):
    mapping = {
        "Plant 1": "Colombia", "Plant 2": "Colombia",
        "Plant 3": "Peru",     "Plant 4": "Peru",
        "Plant 5": "Brazil",   "Plant 6": "Brazil",
    }
    return mapping.get(str(plant_name_val).strip(), str(plant_name_val))

def _supp(name):
    return supp_map.get(str(name).strip(), str(name))


# ═══════════════════════════════════════════════════════════════════════════════
# FILE 1 — MRP.xlsx  (sheet: MRP)
# ═══════════════════════════════════════════════════════════════════════════════
print("\nProcessing MRP.xlsx...")
mrp = mrp_target.copy()
mrp["Descripción"] = mrp["Código SAP"].apply(_mat)
mrp["PLANTA"]      = mrp["PLANTA"].map(planta_map).fillna(mrp["PLANTA"])
# CENTRO and PAÍS kept as numeric/original — pipeline uses them as join keys
mrp = mrp.drop(columns=["_centro"])
_update_excel(BASE / "MRP.xlsx", {"MRP": mrp})
print(f"  {len(mrp)} rows written")


# ═══════════════════════════════════════════════════════════════════════════════
# FILE 2 — CORTE B2WISE.xlsx  (sheet: Sheet1)
# ═══════════════════════════════════════════════════════════════════════════════
print("\nProcessing CORTE B2WISE.xlsx...")
b2 = b2_target.copy()
b2["Part Description"] = b2["Part Code"].apply(_mat)
b2 = b2.drop(columns=["_centro"])
_update_excel(BASE / "CORTE B2WISE.xlsx", {"Sheet1": b2})
print(f"  {len(b2)} rows written")


# ═══════════════════════════════════════════════════════════════════════════════
# FILE 3 — Saldo Actual.xlsx  (sheet: Stock actual (MB52))
# ═══════════════════════════════════════════════════════════════════════════════
print("\nProcessing Saldo Actual.xlsx...")
sa_raw = pd.read_excel(BASE / "Saldo Actual.xlsx",
                       sheet_name="Stock actual (MB52)", dtype=str)
sa_raw["_centro"] = sa_raw["Centro"].apply(_centro_int)
sa = sa_raw[sa_raw["_centro"].isin(TARGET_PLANTS)].copy()
sa["Texto breve de material"] = sa["Material"].apply(_mat)
# Centro kept numeric
sa = sa.drop(columns=["_centro"])
_update_excel(BASE / "Saldo Actual.xlsx", {"Stock actual (MB52)": sa})
print(f"  {len(sa)} rows written")


# ═══════════════════════════════════════════════════════════════════════════════
# FILE 4 — Consumos materiales.xlsx  (sheet: Consumos)
# ═══════════════════════════════════════════════════════════════════════════════
print("\nProcessing Consumos materiales.xlsx...")
con_raw = pd.read_excel(BASE / "Consumos materiales.xlsx",
                        sheet_name="Consumos", dtype=str)
con_raw["_centro"] = con_raw["Centro"].apply(_centro_int)
con = con_raw[con_raw["_centro"].isin(TARGET_PLANTS)].copy()
con["Texto breve de material"] = con["Material"].apply(_mat)
# Centro kept numeric
# Anonymise SAP usernames
if "Nombre del usuario" in con.columns:
    users = {u: f"User {i+1}"
             for i, u in enumerate(sorted(con["Nombre del usuario"].dropna().unique()))}
    con["Nombre del usuario"] = con["Nombre del usuario"].map(users).fillna("User ?")
con = con.drop(columns=["_centro"])
_update_excel(BASE / "Consumos materiales.xlsx", {"Consumos": con})
print(f"  {len(con)} rows written")


# ═══════════════════════════════════════════════════════════════════════════════
# FILE 5 — BD_Sp_Pedidos_Compra.xlsx  (sheets: SolPed, PO, Ingresos)
# ═══════════════════════════════════════════════════════════════════════════════
print("\nProcessing BD_Sp_Pedidos_Compra.xlsx...")

# SolPed (ME5A)
sol_raw = pd.read_excel(BASE / "BD_Sp_Pedidos_Compra.xlsx",
                        sheet_name="SolPed (ME5A)", dtype=str)
sol_raw["_centro"] = sol_raw["Centro"].apply(_centro_int)
sol = sol_raw[sol_raw["_centro"].isin(TARGET_PLANTS)].copy()
sol["Texto breve"] = sol["Material"].apply(_mat)
# Centro kept numeric
sol = sol.drop(columns=["_centro"])

# PO (ME2W)
po = po_target.copy()
po["Texto breve"] = po["Material"].apply(_mat)
po["Proveedor/Centro suministrador"] = po["Proveedor/Centro suministrador"].apply(_supp)
# Centro kept numeric
po = po.drop(columns=["_centro"])

# Ingresos (SQVI)
ing_raw = pd.read_excel(BASE / "BD_Sp_Pedidos_Compra.xlsx",
                        sheet_name="Ingresos (SQVI)", dtype=str)
ing_raw["_centro"] = ing_raw["Ce."].apply(_centro_int)
ing = ing_raw[ing_raw["_centro"].isin(TARGET_PLANTS)].copy()
# Ce. kept numeric
ing = ing.drop(columns=["_centro"])

_update_excel(BASE / "BD_Sp_Pedidos_Compra.xlsx", {
    "SolPed (ME5A)": sol,
    "PO (ME2W)":     po,
    "Ingresos (SQVI)": ing,
})
print(f"  SolPed {len(sol)} rows | PO {len(po)} rows | Ingresos {len(ing)} rows")


# ═══════════════════════════════════════════════════════════════════════════════
# FILE 6 — Ordenes de compra B2WISE.xlsx  (sheet: Sheet1)
# ═══════════════════════════════════════════════════════════════════════════════
print("\nProcessing Ordenes de compra B2WISE.xlsx...")
oc = oc_target.copy()
oc["Part Description"]    = oc["Part Code"].apply(_mat)
oc["SupplierDescription"] = oc["SupplierDescription"].apply(_supp)
if "Supplier" in oc.columns:
    oc["Supplier"] = oc["Supplier"].apply(_supp)
if "Supplier Code" in oc.columns:
    # Map supplier codes to same anonymous label via description
    oc["Supplier Code"] = oc["SupplierDescription"]
oc = oc.drop(columns=["_centro"])
_update_excel(BASE / "Ordenes de compra B2WISE.xlsx", {"Sheet1": oc})
print(f"  {len(oc)} rows written")


# ═══════════════════════════════════════════════════════════════════════════════
# FILE 7 — Inventario Inicial.xlsx  (sheet: Cierre_Inv)
# ═══════════════════════════════════════════════════════════════════════════════
print("\nProcessing Inventario Inicial.xlsx...")
inv_raw = pd.read_excel(BASE / "Inventario Inicial.xlsx",
                        sheet_name="Cierre_Inv", dtype=str)
inv_raw["_centro"] = inv_raw["Ce."].apply(_centro_int)
inv = inv_raw[inv_raw["_centro"].isin(TARGET_PLANTS)].copy()
# Ce. kept numeric
# Jquía.productos = product hierarchy description — anonymise
if "Jquía.productos" in inv.columns:
    inv["Jquía.productos"] = inv["Material"].apply(_mat)
inv = inv.drop(columns=["_centro"])
_update_excel(BASE / "Inventario Inicial.xlsx", {"Cierre_Inv": inv})
print(f"  {len(inv)} rows written")


# ═══════════════════════════════════════════════════════════════════════════════
print("\n✓ All files anonymised successfully.")
print(f"  {raw_n-1} raw materials | {pkg_n-1} packaging | {oth_n-1} other")
print(f"  {len(supp_map)} suppliers | 6 plants")
