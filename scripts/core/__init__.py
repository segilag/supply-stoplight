# -*- coding: utf-8 -*-
"""
core/  —  Supply Planner v5
Business logic modules (no I/O, no pandas, pure functions).

    adu              → ADU calculation (MRP + B2Wise hybrid)
    coverage         → Days-of-coverage simulation
    projection       → Saturday-based inventory projection
    inventory_state  → State classification (CRITICO / RIESGO / ALERTA / OK)
"""
from .adu             import calculate_adu_hist, calculate_adu
from .coverage        import calculate_coverage, adjust_planning_zones, zones_to_days
from .projection      import daily_adu_for_month, dot_color_projected, project_inventory
from .inventory_state import semaforo_estado, detect_quiebre_info, classify_inventory_state
