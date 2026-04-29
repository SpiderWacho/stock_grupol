# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the script

```bash
# Via batch file
actualizarStock.bat

# Directly
python actualizarStock.py
```

The script is interactive — it prompts for one of three options:
- `1` — Download stock snapshot from Metabase and write to `Stock depositos.xlsx`
- `2` — Send low-stock alert emails via Outlook
- `3` — Show stockout KPIs (`mostrar_kpis()` is referenced but not yet implemented)

## System architecture

This is a supply-chain tracking system for disposable supplies (descartables) across two business units:

- **VO** — Grupo L warehouses: CIUDADELA and MORENO (CD Moreno, CDM02, CDM03)
- **ATLANTICO** — SERVICIOS ATLANTICO SA, co-located in the same warehouses

The data pipeline has three layers:

### Layer 1 — ERP raw data (Metabase / SAP)
`actualizar_stock()` authenticates to Metabase, downloads card 201 as Excel, and splits rows by company/location into three sheets of `Stock depositos.xlsx` (STOCK CIUDADELA, STOCK MORENO, STOCK ATLANTICO). Raw columns: `CtroDistrib`, `Empresa`, `Articulo`, `GrpAbr`, `UM`, `Fisico`, `PendGuardar`, `PedSinOP`, `PendPreparar`, `StockDisp`.

### Layer 2 — Excel workbooks (where the real logic lives)
Two workbooks in `Files/` maintain the operational dashboards. The Python script is glue around them; the formulas and structure inside these files are the core of the system.

**`Files/Seguimiento stock descartables VO.xlsx`** (5 sheets):
- **Tabla Articulos** — SKU master / normalization table. Maps many raw EANs (different pack sizes, brands, supplier variants) to a single `Familia` (product family) + `UM` (pack-to-unit multiplier). This collapses ~300 SKUs into ~60 families.
- **Stock** — Raw stock rows from `Stock depositos.xlsx`, enriched with `FAMILIA`, `UM2` (multiplier from Tabla Articulos), and `TOTAL = StockDisp × UM2` to normalize everything into base units.
- **Calcular necesidad** — Groups Stock rows by Familia, summing `TOTAL` per DC (Moreno, Ciudadela, CDM02, CDM03). Validates each article is registered in Tabla Articulos.
- **MRP** — Demand side. Weekly consumption per SKU (columns = week numbers, 6 weeks of history). `TOTAL` is a weighted forecast. `DIFERENCIA = Stock - TOTAL`. `STATUS = "COMPRAR"` when understocked.
- **Necesidad por familia** — Operational dashboard. Columns: `Familia`, `UM (ultima compra)`, `Stock total` (from Calcular necesidad), `Consumo` (from MRP), `Diferencia`, `Cobertura = Stock / Consumo`, `Estado` (manual notes), `Ultima OC` (last PO number), `Proveedor`.

**`Files/Seguimiento stock descartables ATLANTICO.xlsx`** (3 sheets):
- **Consumos** — Per-SKU demand table: EAN, COD SAP, `Familia`, `UM` (pack size), weekly `Consumo`, `Consumo total = Consumo × UM` in base units.
- **Stock** — Raw stock enriched with `Ean Normalizado`, `Familia`, `UM producto`, `Total = StockDisp × UM producto`.
- **Necesidad por familia** — Operational dashboard, same structure as VO: `Stock`, `Consumo`, `Diferencia`, `Cobertura`, `Estado`, `Ultima OC`, `Proveedor`.

### Layer 3 — Python automation
Reads the `Necesidad por familia` sheets (where `Cobertura` is already computed by Excel formulas) and:
- Sends HTML alert emails via the local Outlook COM object for products with `Cobertura < 0.50`
- Persists crash/broke/resolved events to the `stockouts` SQLite table

## The Cobertura metric (central KPI)

`Cobertura = Total Stock (base units) / Weekly Consumption (base units)`

Represents weeks of stock coverage. The Excel shows it as a decimal ratio; the Python alert threshold is `< 0.50` (less than half a week). The stockout status lifecycle is:

- **crash** — Cobertura < 0.50
- **broke** — Cobertura ≤ 0 (zero or negative stock)
- **resolved** — Cobertura recovered to ≥ 0.50

## What is manual vs automated

The Stock sheets in both workbooks are currently **updated manually** by pasting data from `Stock depositos.xlsx`. `write_to_excel()` in the Python script was written to automate this step, but all call sites are commented out. Weekly consumption numbers (MRP / Consumos sheets) are also maintained manually.

## Key constants and paths

| Name | Value |
|------|-------|
| `TARGET_EXCEL` | `Stock depositos.xlsx` (repo root) |
| VO tracking workbook | `Files/Seguimiento stock descartables VO.xlsx` |
| ATLANTICO tracking workbook | `Files/Seguimiento stock descartables ATLANTICO.xlsx` |
| `inventory.db` | `C:\Users\GastonVecchio\Documents\Code\inventory.db` (outside repo) |
| `PRODUCT_COL` | `"Familia"` |
| `STOCK_COL` | `"Cobertura"` |
| Metabase card | `201` at `https://metabase-new.grupol.ar` |

Note: the Python script still has old hardcoded SharePoint paths for the tracking workbooks — these are outdated and should be updated to the local `Files/` paths.

## Environment

Credentials in `.env` (gitignored): `METABASE_USER`, `PASSWORD`.

Email sending requires Outlook to be installed and open (uses `win32com.client` COM automation — Windows-only).

## Known gaps

- `mostrar_kpis()` is called on menu option `3` but is not defined (was in the reverted `add stockouts db` branch).
- In `alertar_faltantes()`, `rows_html` for the VO email is never assigned — `build_rows_html()` appends to a local `rows` variable but doesn't return it, causing a `NameError` at runtime.
- `_sync_stockouts()` is never called from `alertar_faltantes()`; stockout DB tracking is wired up but disconnected.
- `_migrate_db()` is commented out — uncomment to add missing columns (`cobertura`, `status`, `date_resolved`) to an older DB.
