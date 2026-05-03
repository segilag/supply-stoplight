# -*- coding: utf-8 -*-
"""
loaders.py  —  Supply Planner v5
Carga todos los archivos Excel del directorio Datos/ y devuelve los DataFrames
en un diccionario estructurado.  No modifica lógica de negocio.
"""

import pandas as pd
from pathlib import Path


# ───────────────────────────────────────────────────────────────────────────────
#  Utilidad compartida  (también importada por generar_tablero.py)
# ───────────────────────────────────────────────────────────────────────────────

def mat_key(v):
    """Normaliza código de material a string entero (ej. '6010117')."""
    try:
        return str(int(float(v)))
    except Exception:
        return str(v).strip()


# ───────────────────────────────────────────────────────────────────────────────
#  Función principal de ingesta
# ───────────────────────────────────────────────────────────────────────────────

def load_all_data(datos_path: Path) -> dict:
    """
    Lee todos los archivos Excel del directorio datos_path.

    Returns
    -------
    dict con claves:
        mrp, inventario_inicial, saldo_actual, consumos,
        solped, ingresos, b2wise, b2wise_oc,
        origen_lookup, lineas_fab
    """
    datos = Path(datos_path)
    print("Leyendo archivos Excel…")

    # ── Archivos obligatorios ────────────────────────────────────────────────
    df_mrp = pd.read_excel(datos / "MRP.xlsx",                    sheet_name="MRP")
    df_ini = pd.read_excel(datos / "Inventario Inicial.xlsx",     sheet_name="Cierre_Inv")
    df_sdo = pd.read_excel(datos / "Saldo Actual.xlsx")
    df_con = pd.read_excel(datos / "Consumos materiales.xlsx")
    df_sol = pd.read_excel(datos / "BD_Sp_Pedidos_Compra.xlsx",   sheet_name="SolPed (ME5A)")
    df_ing = pd.read_excel(datos / "BD_Sp_Pedidos_Compra.xlsx",   sheet_name="Ingresos (SQVI)")

    # ── Tipo Material (filtro materiales activos en SOLPEDs) ─────────────────
    try:
        _df_tipo = pd.read_excel(datos / "BD_Sp_Pedidos_Compra.xlsx", sheet_name="Tipo_Material")
        tipo_material_set = frozenset(
            (mat_key(r["Material"]), str(int(float(r["Ce."]))) if pd.notna(r["Ce."]) else "0")
            for _, r in _df_tipo.iterrows()
            if pd.notna(r["Material"])
        )
        print(f"  Tipo_Material: {len(tipo_material_set)} combinaciones material+centro cargadas")
    except Exception as _e:
        tipo_material_set = frozenset()
        print(f"  Tipo_Material: hoja no encontrada ({_e}) — sin filtro activo")
    df_b2  = pd.read_excel(datos / "CORTE B2WISE.xlsx")

    # ── Órdenes de compra B2Wise ─────────────────────────────────────────────
    # Solo PO (inician con 4); las que inician con 1 son SOLPEDs
    df_b2oc = pd.read_excel(datos / "Ordenes de compra B2WISE.xlsx", sheet_name="Sheet1")
    df_b2oc = df_b2oc[df_b2oc["Supplier Order Number"].astype(str).str.startswith("4")].copy()
    df_b2oc["_mat"]       = df_b2oc["Part Code"].apply(mat_key)
    df_b2oc["_centro"]    = df_b2oc["Warehouse Code"].apply(lambda v: int(float(v)) if pd.notna(v) else 0)
    df_b2oc["_fecha_ent"] = pd.to_datetime(df_b2oc["Supplier Order Arrival Date"], errors="coerce")
    df_b2oc["_qty_pend"]  = pd.to_numeric(df_b2oc["Supplier Order Quantity Outstanding"], errors="coerce").fillna(0)
    df_b2oc["_qty_ped"]   = pd.to_numeric(df_b2oc["SupplierOrderQuantityOrdered"],        errors="coerce").fillna(0)
    df_b2oc["_pos"]       = df_b2oc["Line_Number"].fillna("").astype(str).str[:6].str[-2:]
    _um_map = (
        df_b2.assign(_k=df_b2["Part Code"].apply(mat_key)).set_index("_k")["UOM"].to_dict()
        if "UOM" in df_b2.columns else {}
    )
    df_b2oc["_um"] = df_b2oc["_mat"].map(_um_map).fillna("UN")
    print(f"  OC B2Wise: {len(df_b2oc)} filas | {df_b2oc['_centro'].nunique()} centros")

    # ── Origen materiales (opcional) ─────────────────────────────────────────
    _orig_path = datos / "0. Base" / "Origen materiales.xlsx"
    if _orig_path.exists():
        _df_orig = pd.read_excel(_orig_path, sheet_name="Origen")
        _df_orig["_centro"] = _df_orig["Warehouse Code"].apply(lambda v: int(float(v)) if pd.notna(v) else 0)
        _df_orig["_mat"]    = _df_orig["Part Code"].apply(mat_key)
        _df_orig["_imp"]    = _df_orig["Descripción 2"].fillna("").str.strip().str.upper() == "IMPORTADO"
        origen_lookup = {(r["_centro"], r["_mat"]): r["_imp"] for _, r in _df_orig.iterrows()}
        print(f"  Origen materiales: {len(origen_lookup)} registros cargados")
    else:
        origen_lookup = {}
        print("  Origen materiales: archivo no encontrado — se omite")

    # ── Impactos Planners (opcional) ─────────────────────────────────────────
    _imp_path = datos / "Reporte_Impactos_Planners.xlsx"
    impactos_lookup = {}
    if _imp_path.exists():
        try:
            _xl_imp = pd.ExcelFile(_imp_path)
            _rows = []
            for _sh in _xl_imp.sheet_names:
                _df_sh = _xl_imp.parse(_sh)
                # Normalizar columnas (pueden venir con acentos/mayúsculas)
                _df_sh.columns = [str(c).strip().upper() for c in _df_sh.columns]
                if "CÓDIGO" not in _df_sh.columns and "CODIGO" in _df_sh.columns:
                    _df_sh = _df_sh.rename(columns={"CODIGO": "CÓDIGO"})
                needed = {"CENTRO", "CÓDIGO", "IMPACTO", "ACCIÓN"}
                if not needed.issubset(set(_df_sh.columns)):
                    continue
                _df_sh["_mat"]    = _df_sh["CÓDIGO"].apply(mat_key)
                _df_sh["_centro"] = _df_sh["CENTRO"].fillna("").astype(str).str.strip().str.upper()
                _df_sh["_impact"] = _df_sh["IMPACTO"].fillna("").astype(str).str.strip()
                _df_sh["_action"] = _df_sh["ACCIÓN"].fillna("").astype(str).str.strip()
                _rows.append(_df_sh[["_mat","_centro","_impact","_action"]])
            if _rows:
                _df_all = pd.concat(_rows, ignore_index=True)
                _df_all = _df_all[(_df_all["_impact"] != "") | (_df_all["_action"] != "")]
                impactos_lookup = {
                    (r["_mat"], r["_centro"]): {"impact": r["_impact"], "action": r["_action"]}
                    for _, r in _df_all.iterrows()
                }
                print(f"  Impactos planners: {len(impactos_lookup)} registros cargados")
        except PermissionError:
            print("  Impactos planners: archivo en uso (Excel abierto) — se omite")
        except Exception as _e:
            print(f"  Impactos planners: error al leer ({_e}) — se omite")
    else:
        print("  Impactos planners: archivo no encontrado — se omite")

    # ── Líneas de fabricación + Centros master (opcional, búsqueda por glob) ──
    _lin_paths = list((datos / "0. Base").glob("*neas*")) or list(datos.glob("*neas*"))
    if _lin_paths:
        _lin_file = _lin_paths[0]
        _df_lin_raw = pd.read_excel(_lin_file, sheet_name="Lineas de fabricacion")
        _cols = list(_df_lin_raw.columns)
        df_lin = _df_lin_raw[[_cols[0], _cols[1], _cols[3]]].copy()
        df_lin.columns = ["centro", "mat_sap", "linea_fab"]
        df_lin["_mat"]    = df_lin["mat_sap"].apply(mat_key)
        df_lin["_centro"] = df_lin["centro"].apply(lambda v: int(float(v)) if pd.notna(v) else 0)
        print(f"  Lineas fab.: {len(df_lin)} filas | centros: {sorted(df_lin['_centro'].unique().tolist())}")

        # ── Centros master (hoja "Centros" del mismo archivo) ────────────────
        # Fallback flags por nombre de país (para Excels sin columna Flag)
        _FLAG_FALLBACK = {
            "Argentina": "AR", "Brasil": "BR", "Colombia": "CO",
            "Perú": "PE", "Peru": "PE", "Chile": "CL",
        }
        try:
            _df_c = pd.read_excel(_lin_file, sheet_name="Centros")
            # Detect if Flag column exists (col D / index 3)
            _has_flag_col = _df_c.shape[1] >= 4
            centros_master = {}
            for _, _row in _df_c.iterrows():
                try:
                    _code = int(float(_row.iloc[1]))
                except Exception:
                    continue
                _pais = str(_row.iloc[0]).strip()
                _desc = str(_row.iloc[2]).strip()
                _short = _desc.split(" - ")[-1].strip() if " - " in _desc else _desc
                if _has_flag_col and pd.notna(_row.iloc[3]) and str(_row.iloc[3]).strip():
                    _flag = str(_row.iloc[3]).strip().upper()
                else:
                    _flag = _FLAG_FALLBACK.get(_pais, "??")
                centros_master[_code] = {
                    "pais":        _pais,
                    "descripcion": _desc,
                    "nombre":      _short,
                    "flag":        _flag,
                }
            print(f"  Centros master: {len(centros_master)} centros ({', '.join(sorted(set(v['pais'] for v in centros_master.values())))})")
        except Exception as _e:
            centros_master = {}
            print(f"  Centros master: hoja 'Centros' no encontrada ({_e}) — se omite")
    else:
        df_lin = pd.DataFrame(columns=["centro", "mat_sap", "linea_fab", "_mat", "_centro"])
        centros_master = {}
        print("  Lineas fab.: archivo no encontrado — se omite")

    print("Archivos leídos.")

    return {
        "mrp":                df_mrp,
        "inventario_inicial": df_ini,
        "saldo_actual":       df_sdo,
        "consumos":           df_con,
        "solped":             df_sol,
        "ingresos":           df_ing,
        "b2wise":             df_b2,
        "b2wise_oc":          df_b2oc,
        "origen_lookup":      origen_lookup,
        "lineas_fab":         df_lin,
        "impactos_lookup":    impactos_lookup,
        "centros_master":     centros_master,
        "tipo_material_set":  tipo_material_set,
    }
