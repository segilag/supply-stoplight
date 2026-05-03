"""
fix_centro_codes.py
Reverses the Centro column from "Plant X" back to numeric SAP codes.
Run once to fix anonymize_data.py v1 output.
"""
import sys
from pathlib import Path
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

BASE = Path(__file__).parent.parent / "Datos"

NAME_TO_CENTRO = {
    "Plant 1": "4301", "Plant 2": "4321",
    "Plant 3": "4401", "Plant 4": "4403",
    "Plant 5": "4202", "Plant 6": "4208",
}


def _rev(val):
    return NAME_TO_CENTRO.get(str(val).strip(), str(val))


def _update(path, mods):
    all_sheets = pd.read_excel(path, sheet_name=None, dtype=str)
    for sheet, df in mods.items():
        if sheet in all_sheets:
            all_sheets[sheet] = df
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for name, df in all_sheets.items():
            df.to_excel(w, sheet_name=name, index=False)


# MRP.xlsx
print("Fixing MRP.xlsx...")
mrp = pd.read_excel(BASE / "MRP.xlsx", sheet_name="MRP", dtype=str)
mrp["CENTRO"] = mrp["CENTRO"].apply(_rev)
mrp["LLAVE"]  = mrp["CENTRO"].astype(str) + mrp["Código SAP"].astype(str)
_update(BASE / "MRP.xlsx", {"MRP": mrp})
print(f"  {len(mrp)} rows")

# Saldo Actual.xlsx
print("Fixing Saldo Actual.xlsx...")
sa = pd.read_excel(BASE / "Saldo Actual.xlsx",
                   sheet_name="Stock actual (MB52)", dtype=str)
sa["Centro"] = sa["Centro"].apply(_rev)
_update(BASE / "Saldo Actual.xlsx", {"Stock actual (MB52)": sa})
print(f"  {len(sa)} rows")

# Consumos materiales.xlsx
print("Fixing Consumos materiales.xlsx...")
con = pd.read_excel(BASE / "Consumos materiales.xlsx",
                    sheet_name="Consumos", dtype=str)
con["Centro"] = con["Centro"].apply(_rev)
_update(BASE / "Consumos materiales.xlsx", {"Consumos": con})
print(f"  {len(con)} rows")

# BD_Sp_Pedidos_Compra.xlsx
print("Fixing BD_Sp_Pedidos_Compra.xlsx...")
sol = pd.read_excel(BASE / "BD_Sp_Pedidos_Compra.xlsx",
                    sheet_name="SolPed (ME5A)", dtype=str)
sol["Centro"] = sol["Centro"].apply(_rev)

po = pd.read_excel(BASE / "BD_Sp_Pedidos_Compra.xlsx",
                   sheet_name="PO (ME2W)", dtype=str)
po["Centro"] = po["Centro"].apply(_rev)

ing = pd.read_excel(BASE / "BD_Sp_Pedidos_Compra.xlsx",
                    sheet_name="Ingresos (SQVI)", dtype=str)
ing["Ce."] = ing["Ce."].apply(_rev)

_update(BASE / "BD_Sp_Pedidos_Compra.xlsx", {
    "SolPed (ME5A)":  sol,
    "PO (ME2W)":      po,
    "Ingresos (SQVI)": ing,
})
print(f"  SolPed {len(sol)} | PO {len(po)} | Ingresos {len(ing)}")

# Inventario Inicial.xlsx
print("Fixing Inventario Inicial.xlsx...")
inv = pd.read_excel(BASE / "Inventario Inicial.xlsx",
                    sheet_name="Cierre_Inv", dtype=str)
inv["Ce."] = inv["Ce."].apply(_rev)
_update(BASE / "Inventario Inicial.xlsx", {"Cierre_Inv": inv})
print(f"  {len(inv)} rows")

print("\nDone — Centro codes restored to numeric SAP codes.")
