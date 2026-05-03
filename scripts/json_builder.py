# -*- coding: utf-8 -*-
"""
json_builder.py  —  Supply Planner v5
Assembles the planning DATA structure and serialises it to compact JSON.

The resulting string is consumed by html_builder.write_html() to inject
the data into the dashboard template.
"""

import json


def build_json(centros: dict, meta: dict) -> str:
    """
    Assemble the top-level DATA object and return it as a compact JSON string.

    Parameters
    ----------
    centros : dict
        Mapping of centro key → centro data dict, e.g.:
        {
            "4301": col_data,
            "4321": car_data,
            "PERU": per_data,
            "4403": hua_data,
        }
        Each value is the dict returned by build_centro_data().

    meta : dict
        Dashboard metadata, e.g.:
        {
            "fecha_saldo":  "15-Apr-2026",
            "fecha_b2wise": "15-Apr-2026",
            "mes_m0":       "Abril",
            "sem_weeks":    [...],
        }

    Returns
    -------
    str
        Compact JSON string (no spaces, ensure_ascii=False) ready to be
        injected into the HTML template placeholder.
    """
    data = {
        "meta":    meta,
        "centros": centros,
    }
    return json.dumps(data, ensure_ascii=False, separators=(",", ":"))
