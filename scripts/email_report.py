# -*- coding: utf-8 -*-
"""
email_report.py  —  Supply Planner v5
Punto de entrada para el envío automático del reporte diario.

Uso:
    python email_report.py                   # envía con config por defecto
    python email_report.py --dry-run         # genera HTML pero NO envía
    python email_report.py --to extra@e.com  # añade destinatario adicional

Configuración:
    config/email_config.json   — SMTP, destinatarios, opciones
    Variable de entorno SUPPLY_EMAIL_PASS — contraseña del remitente
"""

import argparse
import json
import logging
import sys
from datetime import date, datetime
from pathlib import Path

# ── Rutas ──────────────────────────────────────────────────────────────────────
BASE        = Path(__file__).parent.parent
CONFIG_PATH = BASE / "config" / "email_config.json"
LOG_DIR     = BASE / "logs"
LOG_DIR.mkdir(exist_ok=True)
SCRIPTS_DIR = Path(__file__).parent

# ── Logging ───────────────────────────────────────────────────────────────────
def _setup_logging(debug: bool = False) -> None:
    log_file = LOG_DIR / f"email_{datetime.now():%Y%m%d}.log"
    fmt = "%(asctime)s  %(levelname)-8s  %(message)s"
    handlers = [
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
    logging.basicConfig(
        level=logging.DEBUG if debug else logging.INFO,
        format=fmt,
        handlers=handlers,
    )

# ── Argparse ──────────────────────────────────────────────────────────────────
def _parse_args():
    p = argparse.ArgumentParser(description="Supply Planner v5 — Email Report")
    p.add_argument("--dry-run", action="store_true",
                   help="Genera el HTML pero no envía el email")
    p.add_argument("--to", metavar="EMAIL", action="append", default=[],
                   help="Añadir destinatario adicional (puede repetirse)")
    p.add_argument("--debug", action="store_true",
                   help="Logging detallado")
    return p.parse_args()


# ── Principal ─────────────────────────────────────────────────────────────────
def main():
    args = _parse_args()
    _setup_logging(args.debug)
    log = logging.getLogger(__name__)

    log.info("=" * 60)
    log.info("Supply Planner v5 — Reporte Email")
    log.info("=" * 60)

    # 1. Cargar config
    if not CONFIG_PATH.exists():
        log.error(f"Config no encontrada: {CONFIG_PATH}")
        sys.exit(1)

    with open(CONFIG_PATH, encoding="utf-8") as f:
        cfg = json.load(f)

    smtp_cfg   = cfg["smtp"]
    rpt_cfg    = cfg["report"]
    sched_cfg  = cfg["schedule"]

    # 2. Verificar día de envío
    today_name = date.today().strftime("%A")
    if today_name not in sched_cfg.get("send_days", []):
        log.info(f"Hoy es {today_name} — no está en send_days. Saliendo.")
        sys.exit(0)

    # 3. Construir destinatarios
    if str(SCRIPTS_DIR) not in sys.path:
        sys.path.insert(0, str(SCRIPTS_DIR))

    from send_email import build_recipients, send_report
    from generate_report import generate

    recipients = build_recipients(cfg["recipients"])
    for extra in args.to:
        if extra not in recipients:
            recipients.append(extra)

    log.info(f"Destinatarios: {recipients or ['(ninguno — dry-run)']}")

    # 4. Generar HTML
    log.info("Generando cuerpo del reporte...")
    html_body, fecha_saldo = generate(
        estados_highlight=rpt_cfg.get("estados_highlight", ["CRITICO", "RIESGO"])
    )

    if args.dry_run:
        preview_path = LOG_DIR / f"email_preview_{fecha_saldo:%Y%m%d}.html"
        preview_path.write_text(html_body, encoding="utf-8")
        log.info(f"Dry-run: HTML guardado en {preview_path}")
        log.info("Dry-run completado — email NO enviado.")
        return

    if not recipients:
        log.warning("Lista de destinatarios vacía — configura config/email_config.json")
        sys.exit(1)

    # 5. Opcional: adjuntar PDF (requiere pdfkit + wkhtmltopdf)
    attachments = []
    if rpt_cfg.get("attach_pdf", False):
        try:
            import pdfkit
            pdf_bytes = pdfkit.from_string(html_body, False)
            fname = f"Supply_Report_{fecha_saldo:%Y%m%d}.pdf"
            attachments.append((fname, pdf_bytes))
            log.info(f"PDF adjunto generado: {fname}")
        except Exception as e:
            log.warning(f"No se pudo generar PDF: {e} — se envía sin adjunto")

    # 6. Enviar
    subject = (
        f"📊 Supply Stoplight — {fecha_saldo:%d %b %Y} "
        f"| Daily Inventory Report"
    )
    log.info(f"Asunto: {subject}")
    send_report(smtp_cfg, recipients, subject, html_body, attachments or None)

    log.info("Proceso completado exitosamente.")


if __name__ == "__main__":
    main()
