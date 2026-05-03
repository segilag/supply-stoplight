# -*- coding: utf-8 -*-
"""
send_email.py  —  Supply Planner v5
Envía un email HTML con resumen de supply planning vía SMTP (Office 365).

Uso desde generate_report o email_report — no ejecutar directamente.
"""

import os
import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
from pathlib import Path

logger = logging.getLogger(__name__)


def _get_password(env_var: str) -> str:
    pwd = os.environ.get(env_var, "")
    if not pwd:
        raise EnvironmentError(
            f"Variable de entorno '{env_var}' no definida. "
            f"Defínela con: setx {env_var} \"tu-password\""
        )
    return pwd


def build_recipients(cfg_recipients: dict) -> list[str]:
    """Aplana todos los grupos de destinatarios en una lista única."""
    result = []
    for group, emails in cfg_recipients.items():
        for addr in emails:
            addr = addr.strip()
            if addr and addr not in result:
                result.append(addr)
    return result


def send_report(
    smtp_cfg: dict,
    recipients: list[str],
    subject: str,
    html_body: str,
    attachments: list[tuple[str, bytes]] | None = None,
) -> None:
    """
    Envía el email de reporte.

    Parameters
    ----------
    smtp_cfg    : bloque "smtp" del email_config.json
    recipients  : lista de direcciones destino
    subject     : asunto del email
    html_body   : contenido HTML del email
    attachments : lista de (filename, bytes) opcionales
    """
    if not recipients:
        logger.warning("No hay destinatarios configurados — email no enviado.")
        return

    password = _get_password(smtp_cfg["password_env_var"])

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = formataddr((smtp_cfg["sender_name"], smtp_cfg["sender_email"]))
    msg["To"]      = ", ".join(recipients)

    msg.attach(MIMEText(html_body, "html", "utf-8"))

    if attachments:
        outer = MIMEMultipart("mixed")
        outer["Subject"] = msg["Subject"]
        outer["From"]    = msg["From"]
        outer["To"]      = msg["To"]
        outer.attach(msg)
        for fname, data in attachments:
            part = MIMEApplication(data, Name=fname)
            part["Content-Disposition"] = f'attachment; filename="{fname}"'
            outer.attach(part)
        msg = outer

    host = smtp_cfg["host"]
    port = smtp_cfg["port"]
    sender = smtp_cfg["sender_email"]

    logger.info(f"Conectando a {host}:{port} como {sender} ...")
    with smtplib.SMTP(host, port, timeout=30) as server:
        server.ehlo()
        if smtp_cfg.get("use_tls", True):
            server.starttls()
            server.ehlo()
        server.login(sender, password)
        server.sendmail(sender, recipients, msg.as_string())

    logger.info(f"Email enviado a {len(recipients)} destinatario(s).")
