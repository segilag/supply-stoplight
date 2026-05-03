# Comandos del proyecto — Supply Planner LATAM

Comandos slash disponibles para este proyecto. Se ejecutan escribiendo `/nombre-comando [argumentos]` en el chat.

---

## Análisis & Planeación

| Comando | Argumentos | Qué hace |
|---|---|---|
| `/analisis-centro` | `[centro]` ej. `Manizales`, `Lima`, o vacío para todos | Análisis completo por centro: KPIs, prioridades 1-4, top OC atrasadas, SOLPEDs y checklist de trabajo 5 días hábiles |

---

## Tablero & Código

| Comando | Argumentos | Qué hace |
|---|---|---|
| *(próximamente)* | | |

---

## Reportes & Exportación

| Comando | Argumentos | Qué hace |
|---|---|---|
| `/reporte-ejecutivo` | `[planta(s)]` o vacío para LATAM completo | Reporte gerencial LATAM: semáforo regional, posición de inventario, alertas de quiebre, pipeline OC/SOLPEDs y plan de acción con checklist. Genera `reporte_ejecutivo_LATAM_YYYYMMDD.html` listo para imprimir como PDF (A4, estilo Etex) |

---

## Notas
- Los comandos leen siempre el `supply_planner_v5.html` más reciente. Ejecutar `python generar_tablero.py` antes si los datos cambiaron.
- Agregar nuevos comandos: crear archivo `.md` en esta misma carpeta `.claude/commands/`.
