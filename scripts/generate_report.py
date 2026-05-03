# -*- coding: utf-8 -*-
"""
generate_report.py  —  Supply Planner v5
Extrae métricas del pipeline de datos y genera el cuerpo HTML del email.
"""

import sys
import logging
from datetime import date, timedelta
from pathlib import Path

logger = logging.getLogger(__name__)

BASE     = Path(__file__).parent.parent
DATOS    = BASE / "Datos"
HTML_OUT = BASE / "supply_planner_v5.html"

# ── Colores inline para clientes de email (no soportan <style>) ───────────────
C_BG    = "#12161f"
C_BG1   = "#1a1f2e"
C_BG2   = "#232838"
C_BRD   = "#2e3448"
C_TX1   = "#e8ecf4"
C_TX2   = "#8892aa"
C_TX3   = "#555e75"
C_RED   = "#ff4d4d"
C_RED2  = "#2a0a0a"
C_YEL   = "#f5a623"
C_YEL2  = "#2a1a00"
C_GRN   = "#2ecc71"
C_GRN2  = "#0a2010"
C_ORA   = "#F26522"
C_BLUE  = "#4fa3e0"
C_BLUE2 = "#0a1a2a"

STATUS_COLOR = {
    "CRITICO":     (C_RED,  C_RED2),
    "RIESGO":      (C_YEL,  C_YEL2),
    "ALERTA":      (C_ORA,  "#1a0e00"),
    "OK":          (C_GRN,  C_GRN2),
    "SIN_CONSUMO": (C_TX2,  C_BG2),
}


# ═══════════════════════════════════════════════════════════════════════════════
#  Pipeline data loader
# ═══════════════════════════════════════════════════════════════════════════════

def _load_pipeline_data():
    """Ejecuta el mismo pipeline de carga que generar_tablero.py y devuelve los datos."""
    scripts_dir = str(Path(__file__).parent)
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)

    import settings
    from loaders import load_all_data
    from preprocess import preprocess_data
    from material_builder import MaterialContext, build_material, _oc_fecha_info
    import pandas as pd

    raw   = load_all_data(DATOS)
    clean = preprocess_data(raw)

    from settings import FECHA_SALDO, N_SABADOS
    import pandas as pd

    def next_saturdays(from_date, n):
        sats, d = [], from_date
        while len(sats) < n:
            d += timedelta(days=1)
            if d.weekday() == 5:
                sats.append(d)
        return sats

    sabados = next_saturdays(FECHA_SALDO, N_SABADOS)

    ctx = MaterialContext(
        df_lin=clean["lineas_fab"],       origen_lookup=clean["origen_lookup"],
        cons_3m=clean["cons_3m"],         ini_agg=clean["ini_agg"],
        ing_mes=clean["ing_mes"],         cons_mes=clean["cons_mes"],
        b2_soh=clean["b2_soh"],           sdo_agg=clean["sdo_agg"],
        df_b2oc=clean["b2wise_oc"],       df_sol=clean["solped"],
        sabados=sabados,
    )

    # Build centros_map dynamically from centros_master
    centros_master = clean.get("centros_master", {})
    _data_centros  = (
        set(clean["mrp"]["CENTRO"].dropna().astype(int).unique()) |
        set(clean["b2wise"]["_centro"].dropna().astype(int).unique())
    )
    centros_map = {}
    for _code in sorted(_data_centros):
        _meta = centros_master.get(_code)
        if _meta is None:
            continue
        _nombre = _meta["nombre"]
        centros_map[_nombre] = ([_code], _meta["flag"])

    impactos_lookup = clean.get("impactos_lookup", {})

    result = {}
    for nombre, (centros, flag) in centros_map.items():
        mrp_c = clean["mrp"][clean["mrp"]["CENTRO"].isin(centros)].copy()
        b2_c  = clean["b2wise"][clean["b2wise"]["_centro"].isin(centros)].copy()
        mats  = sorted(set(mrp_c["_mat"].unique()) | set(b2_c["_mat"].unique()))

        materiales = [
            m for m in (build_material(mat, centros, mrp_c, b2_c, ctx) for mat in mats)
            if m is not None and not str(m.get("mat_group", "")).startswith(("600", "700"))
        ]

        for m in materiales:
            key = (m["mat"], nombre.upper())
            inp = impactos_lookup.get(key, {})
            m["impact"] = inp.get("impact", "")
            m["action"] = inp.get("action", "")

        oc_rows = clean["b2wise_oc"][clean["b2wise_oc"]["_centro"].isin(centros)]
        oc_atrasadas = sum(
            1 for _, r in oc_rows.iterrows()
            if float(r["_qty_pend"]) > 0 and _oc_fecha_info(r["_fecha_ent"])[2]
        )

        sol_rows = clean["solped"][clean["solped"]["Centro"].isin(centros)]
        sol_sin_lib = sum(
            1 for _, r in sol_rows.iterrows()
            if pd.notna(r.get("Liberación")) and str(r.get("Liberación", "")).strip()
        )

        result[nombre] = {
            "flag":         flag,
            "materiales":   materiales,
            "oc_atrasadas": oc_atrasadas,
            "sol_sin_lib":  sol_sin_lib,
        }

    return result, FECHA_SALDO


# ═══════════════════════════════════════════════════════════════════════════════
#  Supply & impact helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _compute_next_supply(m: dict) -> tuple:
    """
    Returns (supply_date, qty, um) for the earliest confirmed supply.
    Considers all OC lines and released SOLPEDs (sin_liberar=False).
    Returns (None, 0.0, '') when no supply is found.
    """
    candidates = []

    for oc in m.get("oc_list", []):
        eta = oc.get("eta_proy") or oc.get("eta")
        if eta:
            try:
                candidates.append((
                    date.fromisoformat(str(eta)[:10]),
                    float(oc.get("qty_base") or oc.get("qty", 0)),
                    m.get("um", ""),
                ))
            except Exception:
                pass

    for sol in m.get("sol_list", []):
        if not sol.get("sin_liberar"):
            fecha = sol.get("fecha", "")
            if fecha:
                try:
                    candidates.append((
                        date.fromisoformat(str(fecha)[:10]),
                        float(sol.get("qty", 0)),
                        sol.get("um", m.get("um", "")),
                    ))
                except Exception:
                    pass

    if not candidates:
        return None, 0.0, ""

    candidates.sort(key=lambda x: x[0])
    return candidates[0]


def _is_parada_real(m: dict, next_supply_date, fecha_saldo: date) -> bool:
    """
    True if production stoppage is projected:
    stockout occurs before (or on the same day as) next confirmed supply,
    or stockout is projected with zero supply.
    """
    quiebre_dias = m.get("primer_quiebre_dias")
    if quiebre_dias is None:
        return False

    stockout = fecha_saldo + timedelta(days=int(quiebre_dias))
    if next_supply_date is None:
        return True

    return stockout <= next_supply_date


def _accion(m: dict, next_date_str: str) -> str:
    """Recommended action: use impactos_lookup value if set, else derive from data."""
    if m.get("action", "").strip():
        return m["action"]
    if m.get("sin_tiempo") and next_date_str == "N/P":
        return "EMITIR OC URGENTE"
    if any(s.get("sin_liberar") for s in m.get("sol_list", [])):
        return "LIBERAR SOLPED"
    if any(o.get("atrasada") for o in m.get("oc_list", [])):
        return "GESTIONAR OC"
    return "MONITOREAR"


# ═══════════════════════════════════════════════════════════════════════════════
#  Build metrics
# ═══════════════════════════════════════════════════════════════════════════════

def _build_metrics(centros_data: dict, fecha_saldo: date) -> dict:
    metrics = {}
    for nombre, d in centros_data.items():
        ms = d["materiales"]

        by_estado: dict[str, int] = {}
        for m in ms:
            e = m.get("estado", "SIN_CONSUMO")
            by_estado[e] = by_estado.get(e, 0) + 1

        # Enriched table: CRITICO + RIESGO materials with supply & impact data
        materiales_tabla = []
        for m in ms:
            if m.get("estado") not in ("CRITICO", "RIESGO"):
                continue
            next_date, next_qty, next_um = _compute_next_supply(m)
            parada        = _is_parada_real(m, next_date, fecha_saldo)
            next_date_str = next_date.strftime("%d-%b-%Y") if next_date else "N/P"
            next_qty_str  = f"{next_qty:,.0f} {next_um}".strip() if next_qty else "—"
            materiales_tabla.append({
                **m,
                "next_date": next_date_str,
                "next_qty":  next_qty_str,
                "parada":    parada,
                "_accion":   _accion(m, next_date_str),
            })
        # PARADA first, then coverage ascending
        materiales_tabla.sort(key=lambda x: (not x["parada"], x.get("cob_hoy", 9999)))

        # All SOLPEDs across materials for this centro (unreleased first)
        sol_all = []
        for m in ms:
            for s in m.get("sol_list", []):
                sol_all.append({**s, "mat": m["mat"], "desc": m.get("desc", "")})
        sol_all.sort(key=lambda x: (not x.get("sin_liberar", False), x.get("fecha", "")))

        metrics[nombre] = {
            "flag":             d["flag"],
            "total":            len(ms),
            "criticos":         by_estado.get("CRITICO",     0),
            "riesgo":           by_estado.get("RIESGO",      0),
            "alerta":           by_estado.get("ALERTA",      0),
            "ok":               by_estado.get("OK",          0),
            "inv_sc":           by_estado.get("SIN_CONSUMO", 0),
            "oc_atrasadas":     d["oc_atrasadas"],
            "sol_sin_lib":      d["sol_sin_lib"],
            "materiales_tabla": materiales_tabla,
            "sol_all":          sol_all,
        }
    return metrics


# ═══════════════════════════════════════════════════════════════════════════════
#  HTML helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _flag_img(flag: str) -> str:
    return {"CO": "🇨🇴", "PE": "🇵🇪"}.get(flag, "")


def _estado_badge(label: str, count: int, display: str = "") -> str:
    if count == 0:
        return ""
    color, bg = STATUS_COLOR.get(label, (C_TX2, C_BG2))
    text = display or label
    return (
        f'<span style="display:inline-block;padding:3px 10px;border-radius:12px;'
        f'background:{bg};color:{color};font-size:11px;font-weight:700;'
        f'border:1px solid {color};margin:2px 3px">{text}&nbsp;{count}</span>'
    )


def _centro_card(nombre: str, m: dict) -> str:
    flag = _flag_img(m["flag"])
    total_critico_riesgo = m["criticos"] + m["riesgo"]
    border_color = C_RED if m["criticos"] > 0 else (C_YEL if m["riesgo"] > 0 else C_BRD)

    # ── A: Estado de inventario ───────────────────────────────────────────────
    badges = "".join([
        _estado_badge("CRITICO",     m["criticos"]),
        _estado_badge("RIESGO",      m["riesgo"]),
        _estado_badge("ALERTA",      m["alerta"]),
        _estado_badge("OK",          m["ok"]),
        _estado_badge("SIN_CONSUMO", m["inv_sc"], "INV.SC"),
    ])
    section_a = f"""
      <div style="background:{C_BLUE2};border:1px solid {C_BLUE}33;border-radius:4px;
                  padding:8px 12px;margin-bottom:8px">
        <div style="font-size:9px;font-weight:700;color:{C_BLUE};letter-spacing:1px;margin-bottom:6px">
          ESTADO DE INVENTARIO &mdash; {m["total"]} materiales en portafolio
        </div>
        <div>{badges}</div>
      </div>"""

    # ── B: Posiciones de abastecimiento ──────────────────────────────────────
    oc_v  = (f'<b style="color:{C_RED}">{m["oc_atrasadas"]}</b>'
             if m["oc_atrasadas"] > 0 else f'<b style="color:{C_TX1}">0</b>')
    sol_v = (f'<b style="color:{C_YEL}">{m["sol_sin_lib"]}</b>'
             if m["sol_sin_lib"] > 0 else f'<b style="color:{C_TX1}">0</b>')
    section_b = f"""
      <div style="background:{C_YEL2};border:1px solid {C_YEL}33;border-radius:4px;
                  padding:8px 12px;margin-bottom:10px">
        <div style="font-size:9px;font-weight:700;color:{C_YEL};letter-spacing:1px;margin-bottom:4px">
          POSICIONES DE ABASTECIMIENTO
        </div>
        <div style="font-size:12px;color:{C_TX1}">
          {oc_v} posiciones OC atrasadas
          &nbsp;&middot;&nbsp;
          {sol_v} posiciones SOLPED sin liberar
        </div>
      </div>"""

    # ── C: Tabla CRÍTICOS + RIESGO ────────────────────────────────────────────
    section_c = ""
    if m["materiales_tabla"]:
        def _mat_row(p):
            origen = "IMP" if p.get("importado") else "NAC"
            cob    = f'{p["cob_hoy"]:.1f}d' if p.get("cob_hoy") is not None else "—"
            estado_color = STATUS_COLOR.get(p.get("estado", "OK"), (C_TX1, C_BG2))[0]
            parada_html = (
                f'<span style="color:{C_RED};font-weight:700;font-size:10px">🔴 PARADA</span>'
                if p["parada"] else
                f'<span style="color:{C_GRN};font-size:10px">🟢 NO PARADA</span>'
            )
            return (
                f'<tr style="border-bottom:1px solid {C_BRD}">'
                f'<td style="padding:4px 6px;font-family:monospace;font-size:11px;color:{C_TX1}">{p["mat"]}</td>'
                f'<td style="padding:4px 6px;font-size:11px;color:{C_TX2}">{p.get("desc","")[:35]}</td>'
                f'<td style="padding:4px 6px;font-size:10px;color:{C_TX2};text-align:center">{origen}</td>'
                f'<td style="padding:4px 6px;font-size:11px;color:{estado_color};text-align:center">{cob}</td>'
                f'<td style="padding:4px 6px;text-align:center">{parada_html}</td>'
                f'<td style="padding:4px 6px;font-size:11px;color:{C_TX1};text-align:center">{p["next_date"]}</td>'
                f'<td style="padding:4px 6px;font-size:11px;color:{C_TX2};text-align:right">{p["next_qty"]}</td>'
                f'<td style="padding:4px 6px;font-size:11px;color:{C_ORA};font-weight:600">{p["_accion"]}</td>'
                f'</tr>'
            )

        th = (f'<th style="padding:5px 6px;text-align:{{a}};font-weight:600;'
              f'font-size:9px;color:{C_TX3};white-space:nowrap">{{t}}</th>')
        headers = "".join(
            th.format(a=a, t=t) for a, t in [
                ("left",   "Código"),
                ("left",   "Descripción"),
                ("center", "Origen"),
                ("center", "Cob."),
                ("center", "Impacto"),
                ("center", "Próx. suministro"),
                ("right",  "Cantidad"),
                ("left",   "Acción"),
            ]
        )
        rows = "".join(_mat_row(p) for p in m["materiales_tabla"])
        section_c = f"""
      <div style="margin-bottom:10px">
        <div style="font-size:9px;font-weight:700;color:{C_TX3};letter-spacing:.8px;
                    margin-bottom:5px">MATERIALES CR&Iacute;TICOS Y EN RIESGO</div>
        <table style="width:100%;border-collapse:collapse">
          <thead><tr style="background:{C_BG2}">{headers}</tr></thead>
          <tbody>{rows}</tbody>
        </table>
      </div>"""

    # ── D: Tabla SOLPEDs ──────────────────────────────────────────────────────
    section_d = ""
    if m["sol_all"]:
        def _sol_row(s):
            lib = (
                f'<span style="color:{C_RED};font-size:9px;font-weight:700">SIN LIBERAR</span>'
                if s.get("sin_liberar") else
                f'<span style="color:{C_GRN};font-size:9px">LIBERADA</span>'
            )
            qty_str = f'{s["qty"]:,.0f}' if isinstance(s.get("qty"), (int, float)) else str(s.get("qty", ""))
            return (
                f'<tr style="border-bottom:1px solid {C_BRD}">'
                f'<td style="padding:3px 6px;font-family:monospace;font-size:10px;color:{C_TX2}">{s["doc"]}</td>'
                f'<td style="padding:3px 6px;font-family:monospace;font-size:10px;color:{C_TX1}">{s["mat"]}</td>'
                f'<td style="padding:3px 6px;font-size:10px;color:{C_TX2}">{s["desc"][:30]}</td>'
                f'<td style="padding:3px 6px;font-size:10px;color:{C_TX1};text-align:right">{qty_str}</td>'
                f'<td style="padding:3px 6px;font-size:10px;color:{C_TX2}">{s["um"]}</td>'
                f'<td style="padding:3px 6px;font-size:10px;color:{C_TX2}">{s["fecha"]}</td>'
                f'<td style="padding:3px 6px;text-align:center">{lib}</td>'
                f'</tr>'
            )

        th2 = (f'<th style="padding:4px 6px;text-align:{{a}};font-weight:600;'
               f'font-size:9px;color:{C_TX3};white-space:nowrap">{{t}}</th>')
        sol_headers = "".join(
            th2.format(a=a, t=t) for a, t in [
                ("left",   "N° SOLPED"),
                ("left",   "Material"),
                ("left",   "Descripción"),
                ("right",  "Cantidad"),
                ("left",   "UM"),
                ("left",   "Fecha esperada"),
                ("center", "Estado"),
            ]
        )
        sol_rows = "".join(_sol_row(s) for s in m["sol_all"])
        section_d = f"""
      <div style="margin-bottom:8px">
        <div style="font-size:9px;font-weight:700;color:{C_TX3};letter-spacing:.8px;
                    margin-bottom:5px">SOLPEDs &mdash; {len(m["sol_all"])} posiciones</div>
        <table style="width:100%;border-collapse:collapse">
          <thead><tr style="background:{C_BG2}">{sol_headers}</tr></thead>
          <tbody>{sol_rows}</tbody>
        </table>
      </div>"""

    # ── Footer note ───────────────────────────────────────────────────────────
    footer = (
        f'<div style="font-size:10px;color:{C_TX3};margin-top:10px;'
        f'padding-top:8px;border-top:1px solid {C_BRD}">'
        f'&#128206; Detalle de &Oacute;rdenes de Compra disponible en el PDF adjunto.</div>'
    )

    return f"""
    <div style="background:{C_BG1};border:1px solid {border_color};
                border-radius:8px;margin-bottom:12px;overflow:hidden">
      <div style="background:{C_BG2};padding:10px 14px;border-bottom:1px solid {C_BRD}">
        <span style="font-size:15px;font-weight:700;color:{C_TX1}">{flag} {nombre}</span>
      </div>
      <div style="padding:12px 14px">
        {section_a}
        {section_b}
        {section_c}
        {section_d}
        {footer}
      </div>
    </div>"""


# ═══════════════════════════════════════════════════════════════════════════════
#  Build full email HTML
# ═══════════════════════════════════════════════════════════════════════════════

def build_email_html(centros_data: dict, fecha_saldo: date, estados_highlight: list) -> str:
    metrics = _build_metrics(centros_data, fecha_saldo)

    fecha_str = fecha_saldo.strftime("%d %b %Y").upper()

    total_criticos = sum(m["criticos"] for m in metrics.values())
    total_riesgo   = sum(m["riesgo"]   for m in metrics.values())
    total_paradas  = sum(
        sum(1 for t in m["materiales_tabla"] if t["parada"])
        for m in metrics.values()
    )

    if total_criticos > 0:
        summary_color = C_RED
        summary_icon  = "🔴"
        summary_text  = f"{total_criticos} CRITICAL materials"
    elif total_riesgo > 0:
        summary_color = C_YEL
        summary_icon  = "🟡"
        summary_text  = f"No criticals · {total_riesgo} at risk"
    else:
        summary_color = C_GRN
        summary_icon  = "🟢"
        summary_text  = "No criticals or risks"

    paradas_alert = ""
    if total_paradas > 0:
        paradas_alert = f"""
        <div style="background:{C_RED2};border:1px solid {C_RED};border-radius:6px;
                    padding:10px 14px;margin-bottom:14px;font-size:12px;color:{C_RED}">
          &#9888;&#65039; <b>{total_paradas} material(s) at risk of LINE STOPPAGE</b>
          &mdash; review detail per plant
        </div>"""

    cards = "".join(_centro_card(n, m) for n, m in metrics.items())

    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:{C_BG};font-family:Arial,Helvetica,sans-serif;color:{C_TX1}">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:{C_BG}">
    <tr><td align="center" style="padding:20px 12px">
      <table width="640" cellpadding="0" cellspacing="0" style="max-width:640px;width:100%">

        <!-- Header -->
        <tr><td style="background:{C_BG2};border-radius:8px 8px 0 0;
                        border:1px solid {C_BRD};border-bottom:none;padding:18px 20px">
          <div style="font-size:11px;color:{C_TX2};letter-spacing:1px">SUPPLY STOPLIGHT</div>
          <div style="font-size:20px;font-weight:700;color:{C_TX1};margin:4px 0">
            Weekly Inventory Report
          </div>
          <div style="font-size:12px;color:{C_TX2}">Corte: {fecha_str}</div>
        </td></tr>

        <!-- Summary banner -->
        <tr><td style="background:{summary_color}22;border-left:4px solid {summary_color};
                        border:1px solid {C_BRD};border-top:none;border-bottom:none;
                        padding:12px 20px">
          <span style="font-size:14px;font-weight:700;color:{summary_color}">
            {summary_icon} {summary_text}
          </span>
        </td></tr>

        <!-- Body -->
        <tr><td style="background:{C_BG1};border:1px solid {C_BRD};border-top:none;
                        border-radius:0 0 8px 8px;padding:16px 20px">
          {paradas_alert}
          {cards}
          <div style="margin-top:16px;font-size:11px;color:{C_TX2};text-align:center;
                      border-top:1px solid {C_BRD};padding-top:12px">
            Generado autom&aacute;ticamente por Supply Planner v5 &middot; Supply Stoplight
          </div>
        </td></tr>

      </table>
    </td></tr>
  </table>
</body>
</html>"""


def generate(estados_highlight: list | None = None) -> tuple[str, date]:
    """Punto de entrada. Retorna (html_body, fecha_saldo)."""
    if estados_highlight is None:
        estados_highlight = ["CRITICO", "RIESGO"]

    logger.info("Cargando datos del pipeline...")
    centros_data, fecha_saldo = _load_pipeline_data()
    logger.info(f"  Fecha saldo: {fecha_saldo}")

    html = build_email_html(centros_data, fecha_saldo, estados_highlight)
    return html, fecha_saldo
