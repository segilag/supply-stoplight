Eres el agente Supply BI Analyst de Etex Group LATAM. Genera un **reporte ejecutivo de Supply Planning** para presentar a gerencia regional.

**Plantas a incluir:** $ARGUMENTS (si no se especifica, incluir todas: Manizales 4301, Cartagena 4321, Lima 4401+4402, Huachipa 4403)

---

## Instrucciones de ejecución

### Paso 1 — Leer datos
Lee `supply_planner_v5.html` y extrae el JSON embebido (`const DATA =`). De cada centro obtén:
- `materiales[]`: mat, desc, estado, cob_hoy, brk_sin, lt, oc_list, sol_list, quiebran_30 (calc: brk_sin<=30), quiebran_60 (calc: brk_sin<=60)
- `oc[]` y `sol[]` del centro
- `meta`: nombre, quiebran_30, quiebran_60

### Paso 2 — Calcular KPIs LATAM consolidados
Para cada planta:
- Contar materiales por estado: CRITICO, RIESGO, ALERTA, OK
- `cob_prom` = promedio de `cob_hoy` (excluir nulos/negativos)
- `sin_stock` = materiales con saldo = 0
- `oc_total` = len(oc_list de todos los materiales)
- `oc_atrasadas` = oc con `atrasada == true`
- `sol_sin_lib` = sol con `sin_liberar == true`
- `q30` = materiales con `brk_sin != null && brk_sin <= 30`
- `q60` = materiales con `brk_sin != null && brk_sin <= 60 && brk_sin > 30`

### Paso 3 — Producir el reporte en Markdown

Usa exactamente esta estructura:

---

```
# REPORTE EJECUTIVO — SUPPLY PLANNING LATAM
**Período:** [mes año] | **Corte:** [fecha_saldo del meta] | **Generado:** [fecha hoy]
**Plantas:** Manizales 4301 · Cartagena 4321 · Lima 4401+4402 · Huachipa 4403
```

---

## 0. Semáforo Regional

Cinco líneas de estado global. Usa `[CRITICO]` / `[ALERTA]` / `[OK]` según umbrales:
- q30 LATAM > 10 → CRITICO | 4-10 → ALERTA | ≤3 → OK
- OC atrasadas LATAM > 150 → CRITICO | 50-150 → ALERTA | <50 → OK

```
[estado]  Inventario: X% materiales en cobertura adecuada (estado OK o ALERTA leve)
[estado]  Riesgo 30 días: N materiales con quiebre proyectado ≤4 semanas
[estado]  Riesgo 60 días: N materiales adicionales con quiebre proyectado 5-8 semanas
[estado]  OC pipeline: N órdenes abiertas / N atrasadas (XX%)
[estado]  SOLPEDs bloqueadas: N sin liberar a OC
```

---

## 1. Posición de Inventario por Planta

Tabla ordenada de mayor a menor % de materiales en rojo:

| Planta | Total | CRITICO | RIESGO | ALERTA | OK | Cob. Prom. | Sin Stock |
|--------|-------|---------|--------|--------|----|------------|-----------|
| [nombre + cod] | N | N | N | N | N | XX d | N |
| ... | | | | | | | |
| **LATAM TOTAL** | **N** | **N** | **N** | **N** | **N** | **XX d** | **N** |

---

## 2. Alertas Críticas — Riesgo de Quiebre

### 2A. Quiebres Inminentes (≤30 días)

Si hay materiales: tabla de máximo 15 filas, ordenada por `cob_hoy` ASC:

| # | Material | Descripción | Planta | Cob. Hoy | LT | OC Abierta | Acción Requerida |
|---|----------|-------------|--------|----------|----|------------|-----------------|

Lógica columna "Acción Requerida":
- `cob_hoy == 0` y sin OC → **URGENTE: Crear OC hoy**
- `cob_hoy == 0` y OC atrasada → **Expeditar OC (atrasada N días)**
- `cob_hoy > 0` y `cob_hoy < lt` y sin OC → **Crear OC — ya fuera de ventana**
- `cob_hoy > 0` y `cob_hoy < lt` y OC ok → **Confirmar ETA con proveedor**
- `cob_hoy > 0` y sin OC → **Crear OC esta semana**

Si no hay materiales en esta categoría: `> Sin materiales con quiebre en los próximos 30 días.`

### 2B. Riesgo Medio (31–60 días)

Tabla de máximo 10 filas (los de menor `cob_hoy`):

| Material | Descripción | Planta | Cob. Hoy | LT | OC en Pipeline |
|----------|-------------|--------|----------|----|----------------|

Si hay más de 10: agregar `> ...y N materiales adicionales en monitoreo.`

---

## 3. Pipeline de Reabastecimiento

### 3A. Órdenes de Compra

| Planta | OC Abiertas | OC Atrasadas | % Atraso | Críticos Afectados |
|--------|-------------|--------------|----------|--------------------|

Semáforo % Atraso: >30% → [CRITICO] | 15-30% → [ALERTA] | <15% → [OK]

Top 5 proveedores con más OC atrasadas a nivel LATAM:

| Proveedor | OC Atrasadas | Plantas Afectadas | Material Más Crítico |
|-----------|-------------|-------------------|----------------------|

### 3B. SOLPEDs Sin Liberar

Si hay SOLPEDs bloqueadas, tabla completa:

| Planta | SOLPED | Material | Descripción | Cantidad | En Quiebre 30d |
|--------|--------|----------|-------------|----------|----------------|

Si no hay: `> Sin SOLPEDs pendientes de liberación.`

---

## 4. Plan de Acción

### Acciones Inmediatas (esta semana)
- [ ] Generadas automáticamente con lógica:
  - CRITICO sin OC → "Crear OC urgente: [mat] [desc] — [planta]"
  - CRITICO con OC atrasada → "Expeditar OC [#doc] — [mat] [planta] — atrasada N días"
  - SOLPED sin liberar + quiebre 30d → "Liberar SOLPED [#doc] — [mat] [planta]"
- Máximo 10 ítems, ordenados por urgencia

### Acciones de Seguimiento (próximas 2 semanas)
- [ ] Revisar [N] materiales con quiebre proyectado 31-60 días sin OC en pipeline
- [ ] Gestionar confirmación de ETA con proveedores con >30% de OC atrasadas
- [ ] Ítems adicionales relevantes detectados en los datos

### Puntos de Escalamiento Gerencial
Solo incluir si hay materiales CRITICO donde `cob_hoy < 7 días` y `lt > cob_hoy` (sin tiempo de reacción):
- [ ] "[mat] [desc] — [planta]: quiebre en N días, LT = N días. Sin posibilidad de reabastecimiento a tiempo. Evaluar plan de contingencia."

Si no hay puntos de escalamiento: omitir esta subsección.

---

## Reglas de formato
- Números: días sin decimales, porcentajes con 1 decimal, cantidades enteras
- Tablas: siempre con fila de totales en negrita
- Si una sección no tiene datos: indicar explícitamente "Sin novedades"
- Extensión máxima: ~1.200 palabras sin contar tablas
- Idioma: español técnico de supply chain
- NO usar emojis — usar `[CRITICO]` / `[ALERTA]` / `[OK]` como indicadores de estado

---

## Paso 4 — Exportar a HTML para PDF

Después de mostrar el reporte en el chat, genera también un archivo HTML listo para imprimir como PDF.

**Nombre del archivo:** `reporte_ejecutivo_LATAM_[YYYYMMDD].html` (usar fecha del corte)
**Ubicación:** misma carpeta que `supply_planner_v5.html`

El HTML debe tener este esqueleto con los estilos de Etex Group (naranja #F26522, negro, blanco):

```html
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Reporte Ejecutivo Supply Planning LATAM — [fecha]</title>
<style>
  @page { size: A4; margin: 18mm 15mm 15mm 15mm; }
  @media print { .no-print { display:none; } body { font-size: 10pt; } }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 11pt; color: #1a1a1a; background: #fff; }

  /* Header */
  .rpt-header { background: #1a1a1a; color: #fff; padding: 14px 20px; display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
  .rpt-header .title { font-size: 15pt; font-weight: 700; letter-spacing: 0.5px; }
  .rpt-header .title span { color: #F26522; }
  .rpt-header .meta { font-size: 8.5pt; color: #aaa; text-align: right; line-height: 1.6; }

  /* Semáforo strip */
  .semaforo-strip { display: flex; gap: 8px; margin-bottom: 14px; flex-wrap: wrap; }
  .sem-card { flex: 1; min-width: 130px; border-radius: 6px; padding: 8px 12px; border-left: 4px solid; }
  .sem-card.critico { border-color: #ff4d4d; background: #fff5f5; }
  .sem-card.alerta  { border-color: #f5a623; background: #fffbf0; }
  .sem-card.ok      { border-color: #3ecf8e; background: #f0fdf8; }
  .sem-card .label  { font-size: 7.5pt; color: #666; text-transform: uppercase; letter-spacing: 0.5px; }
  .sem-card .value  { font-size: 14pt; font-weight: 700; color: #1a1a1a; }
  .sem-card .sub    { font-size: 8pt; color: #888; }

  /* Secciones */
  h2 { font-size: 11pt; font-weight: 700; color: #F26522; border-bottom: 1.5px solid #F26522; padding-bottom: 3px; margin: 14px 0 8px; text-transform: uppercase; letter-spacing: 0.3px; }
  h3 { font-size: 9.5pt; font-weight: 600; color: #333; margin: 10px 0 5px; }

  /* Tablas */
  table { width: 100%; border-collapse: collapse; font-size: 8.5pt; margin-bottom: 10px; }
  th { background: #1a1a1a; color: #fff; padding: 5px 7px; text-align: left; font-weight: 600; font-size: 8pt; }
  td { padding: 4px 7px; border-bottom: 1px solid #e8e8e8; vertical-align: top; }
  tr:nth-child(even) td { background: #f9f9f9; }
  tr.total td { font-weight: 700; background: #f0f0f0; border-top: 1.5px solid #ccc; }

  /* Badges de estado */
  .badge { display: inline-block; padding: 1px 6px; border-radius: 3px; font-size: 7.5pt; font-weight: 700; }
  .badge-critico { background: #ff4d4d22; color: #c0392b; border: 1px solid #ff4d4d; }
  .badge-alerta  { background: #f5a62322; color: #d4880a; border: 1px solid #f5a623; }
  .badge-ok      { background: #3ecf8e18; color: #1a7a4a; border: 1px solid #3ecf8e; }

  /* Checklist */
  .checklist { list-style: none; padding: 0; }
  .checklist li { padding: 3px 0 3px 20px; position: relative; font-size: 9pt; border-bottom: 1px solid #f0f0f0; }
  .checklist li::before { content: "☐"; position: absolute; left: 0; color: #F26522; font-size: 11pt; line-height: 1; }
  .checklist li.esc::before { content: "⚠"; color: #ff4d4d; }

  /* Botón imprimir */
  .print-btn { display: block; margin: 0 auto 16px; padding: 8px 24px; background: #F26522; color: #fff; border: none; border-radius: 5px; font-size: 10pt; font-weight: 600; cursor: pointer; }
  .print-btn:hover { background: #d4540f; }

  /* Footer */
  .rpt-footer { margin-top: 16px; padding-top: 8px; border-top: 1px solid #ddd; font-size: 7.5pt; color: #999; text-align: center; }
</style>
</head>
<body>

<button class="print-btn no-print" onclick="window.print()">🖨 Imprimir / Guardar como PDF</button>

<div class="rpt-header">
  <div class="title">supply.planner por <span>etex</span> — Reporte Ejecutivo LATAM</div>
  <div class="meta">
    Período: [MES AÑO] · Corte: [FECHA_SALDO]<br>
    Generado: [FECHA HOY] · Plantas: 4301 · 4321 · 4401 · 4402 · 4403
  </div>
</div>

<!-- SECCIÓN 0: Semáforo strip con tarjetas -->
<div class="semaforo-strip">
  <div class="sem-card [critico|alerta|ok]">
    <div class="label">Quiebres ≤30 días</div>
    <div class="value">[N]</div>
    <div class="sub">materiales en riesgo inmediato</div>
  </div>
  <div class="sem-card [critico|alerta|ok]">
    <div class="label">OC Atrasadas</div>
    <div class="value">[N]</div>
    <div class="sub">[X]% del pipeline abierto</div>
  </div>
  <div class="sem-card [critico|alerta|ok]">
    <div class="label">Materiales CRITICO</div>
    <div class="value">[N]</div>
    <div class="sub">de [total] en portafolio</div>
  </div>
  <div class="sem-card [critico|alerta|ok]">
    <div class="label">SOLPEDs bloqueadas</div>
    <div class="value">[N]</div>
    <div class="sub">sin liberar a OC</div>
  </div>
  <div class="sem-card [critico|alerta|ok]">
    <div class="label">Proveedores críticos</div>
    <div class="value">[N]</div>
    <div class="sub">con >20 OC atrasadas</div>
  </div>
</div>

<!-- SECCIÓN 1: Tabla posición inventario -->
<h2>1. Posición de Inventario por Planta</h2>
[tabla HTML con clases badge-critico / badge-alerta / badge-ok en columna CRITICO]

<!-- SECCIÓN 2A: Tabla quiebres inminentes -->
<h2>2. Alertas Críticas — Riesgo de Quiebre</h2>
<h3>2A. Quiebres Inminentes (≤30 días)</h3>
[tabla HTML — columna "Acción" con badge-critico para URGENTE, badge-alerta para Expeditar/Confirmar]

<h3>2B. Riesgo Medio (31–60 días)</h3>
[tabla HTML]

<!-- SECCIÓN 3: Pipeline -->
<h2>3. Pipeline de Reabastecimiento</h2>
<h3>3A. Órdenes de Compra por Planta</h3>
[tabla HTML con badge en columna Semáforo]

<h3>Top 5 Proveedores con Más OC Atrasadas</h3>
[tabla HTML]

<h3>3B. SOLPEDs Sin Liberar</h3>
[tabla HTML o texto "Sin novedades"]

<!-- SECCIÓN 4: Plan de acción -->
<h2>4. Plan de Acción</h2>
<h3>Acciones Inmediatas (esta semana)</h3>
<ul class="checklist">
  <li>[acción 1]</li>
  ...
</ul>

<h3>Acciones de Seguimiento (próximas 2 semanas)</h3>
<ul class="checklist">
  <li>[acción]</li>
</ul>

<h3>Puntos de Escalamiento Gerencial</h3>
<ul class="checklist">
  <li class="esc">[escalamiento]</li>
</ul>

<div class="rpt-footer">
  Fuente: B2Wise corte [fecha] + SAP MM · Estados calculados por modelo DDMRP (zonas buffer roja/amarilla/verde) · supply.planner v5 — Etex Group LATAM
</div>

</body>
</html>
```

Reemplaza todos los placeholders `[...]` con los datos reales calculados. Usa las clases CSS correctas según el estado de cada indicador. Guarda el archivo y confirma la ruta al usuario.

> **Para generar el PDF:** Abrir el archivo HTML en el navegador → botón "Imprimir / Guardar como PDF" → seleccionar "Guardar como PDF" en el diálogo de impresión → tamaño A4, márgenes mínimos.
