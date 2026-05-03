# Supply Stoplight v1.0

A supply chain inventory dashboard that applies DDMRP (Demand-Driven MRP) buffer logic to give planners an instant visual snapshot of material coverage across multiple plants.

## What it does

Reads data from SAP/ERP exports (MRP, purchase orders, consumption history, on-hand inventory) and a B2Wise DDMRP system, then generates a single self-contained HTML dashboard with:

- **Traffic-light status** per material per plant: CRITICAL / RISK / ALERT / OK / NO_CONSUMPTION
- **Coverage in days** vs. buffer targets (Red / Yellow / Green zones)
- **Purchase order tracking**: open POs, overdue POs, unreleased PRs
- **ADU (Average Daily Usage)** calculation — hybrid logic using both MRP-projected demand and historical consumption
- **Multi-plant view**: filter by country, plant, or material type
- **Dark/light mode**, search, sort, export to CSV

## Stack

| Layer | Technology |
|-------|-----------|
| Data pipeline | Python 3 + pandas + openpyxl |
| Dashboard output | Single-file HTML (vanilla JS, no frameworks) |
| Scheduling | Windows `.bat` launcher |
| Email reports | SMTP via `smtplib` |

## Project structure

```
supply-stoplight/
├── scripts/
│   ├── generar_tablero.py      # Entry point — builds the HTML
│   ├── settings.py             # Dates, constants, column mappings
│   ├── loaders.py              # Excel → DataFrame loaders
│   ├── preprocess.py           # Normalisation and aggregation
│   ├── material_builder.py     # Per-material state calculation
│   ├── json_builder.py         # Serialises data for the HTML
│   ├── html_builder.py         # HTML/JS template (embedded)
│   ├── generate_demo_data.py   # Creates anonymised demo Excel files
│   └── core/
│       ├── adu.py              # ADU calculation engine
│       ├── coverage.py         # Days-of-cover logic
│       ├── inventory_state.py  # DDMRP buffer zone classification
│       └── projection.py       # Forward coverage projection
├── Datos/
│   └── demo/                   # Demo Excel files (safe to commit)
├── config/
│   └── email_config.json       # SMTP config (gitignored — see template below)
├── 🔄 Actualizar Tablero.bat   # Windows launcher
└── .gitignore
```

## Quick start

### 1. Install dependencies

```bash
pip install pandas openpyxl
```

### 2. Generate demo data

```bash
python scripts/generate_demo_data.py
```

This creates anonymised Excel files in `Datos/demo/` — 6 plants, 23 materials (Raw Material 1–15, Packaging Material 1–8).

### 3. Build the dashboard

```bash
python scripts/generar_tablero.py
```

Opens `supply_stoplight.html` in your browser.

Or double-click `🔄 Actualizar Tablero.bat` on Windows.

### 4. (Optional) Email reports

Copy `config/email_config.json` template, fill in your SMTP credentials, and run:

```bash
python scripts/send_email.py
```

**`config/email_config.json` template:**
```json
{
  "smtp": {
    "host": "smtp.office365.com",
    "port": 587,
    "use_tls": true,
    "sender_name": "Supply Stoplight",
    "sender_email": "your-email@company.com",
    "password_env_var": "SUPPLY_EMAIL_PASS"
  },
  "recipients": {
    "production_managers": [],
    "planning_managers": [],
    "purchasing_national": [],
    "comex": []
  }
}
```

## DDMRP Buffer logic

| Status | Condition |
|--------|-----------|
| CRITICAL | Coverage ≤ Red zone top |
| RISK | Coverage ≤ Yellow zone top |
| ALERT | Coverage ≤ Green zone top |
| OK | Coverage > Green zone top |
| NO_CONSUMPTION | ADU = 0 |

Buffer zones are sourced from the B2Wise system (`Top of Red`, `Top of Yellow`, `Top of Green` fields).

## Data sources expected

| File | Source system | Key columns |
|------|--------------|-------------|
| `MRP.xlsx` | SAP ME5A | Plant, Material, LT, MRP M+0…M+18 |
| `CORTE B2WISE.xlsx` | B2Wise | Part Code, SOH, ADU, DLT, Top of Red/Yellow/Green |
| `Saldo Actual.xlsx` | SAP MB52 | Material, Plant, On-hand qty |
| `Consumos materiales.xlsx` | SAP MB51 | Material, Plant, Qty, Date |
| `BD_Sp_Pedidos_Compra.xlsx` | SAP ME5A | PO/PR data |
| `Ordenes de compra B2WISE.xlsx` | B2Wise | Open supplier orders |
| `Inventario Inicial.xlsx` | Manual | Opening inventory for projections |

---

Built by Sergio Gil — Supply Planner | [LinkedIn](https://www.linkedin.com/in/sergiogilaguirre)
