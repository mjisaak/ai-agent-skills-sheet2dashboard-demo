---
name: sheet2dashboard
description: Transforms an Excel file (.xlsx) into (1) a cleaned/standardized sanitized-data.xlsx and (2) a polished, interactive HTML dashboard with KPIs, charts, filters, and cross-filtering. Use this skill when the user provides an Excel file and wants to analyze, visualize, or report on employee/sales data - especially when the data contains revenue columns in the pattern Umsatz_YYYY-MM, plus fields like Name, Stadt, Abteilung, Beruf, Teilzeit, and Alter.
---

# Sheet to Dashboard

Convert an Excel file into a clean dataset and an executive-ready interactive HTML dashboard.

## Workflow

1. **Sanitize data** - run `scripts/sanitize.py` to clean, enrich, and export `sanitized-data.xlsx`
2. **Generate dashboard** - run `scripts/generate_dashboard.py` to produce `dashboard.html`
3. **Deliver outputs** - share both files and print the text summary

## Step 1: Sanitize

```bash
python sanitize.py <input.xlsx> [sanitized-data.xlsx]
```

The script:
- Validates required columns: Name (or Vorname/Nachname), Stadt, Beruf, Abteilung, Teilzeit, Alter, and Umsatz_YYYY-MM columns. Exits with a clear error if required columns are missing.
- Normalizes data types: Alter -> int, Teilzeit -> Ja/Nein enum, Umsatz_* -> numeric (negatives -> 0, blanks -> 0)
- Splits a combined `Name` column into `Vorname` and `Nachname`
- Derives `Bundesland` from `Stadt` using a built-in DE/AT/CH city mapping (unknown -> "Unbekannt")
- Computes `Umsatz_Gesamt` and `Umsatz_OE_Monat` (average per month)
- Sorts columns: Vorname, Nachname, Stadt, Bundesland, Abteilung, Beruf, Teilzeit, Alter, Umsatz_* (chronological), Umsatz_Gesamt, Umsatz_OE_Monat
- Sorts rows: Abteilung -> Beruf -> Nachname
- Writes two sheets: `data` (wide format) and `facts_long` (tidy/long format)
- Prints a summary: record count, month range, warnings

## Step 2: Generate Dashboard

```bash
python generate_dashboard.py <sanitized-data.xlsx> [dashboard.html]
```

The script reads `sanitized-data.xlsx` (sheet: `data`) and writes a single self-contained HTML file with:

**KPIs:** Gesamtumsatz, Avg Monatsumsatz/Person, Top-Abteilung + Anteil, Avg Alter + Median, Teilzeitquote, Headcount

**Charts:**
1. Umsatzentwicklung (stacked area/line per Abteilung)
2. Umsatz nach Abteilung (horizontal bar, click to cross-filter)
3. Top 10 Berufe (horizontal bar, tooltip: HC + Avg Umsatz, click to cross-filter)
4. Umsatzverteilung (histogram)
5. Heatmap Monat x Abteilung (seasonality)

**Filters:** Zeitraum (start/end month + "Letzte 12 Monate" button), Abteilung, Bundesland, Stadt, Beruf, Teilzeit/Vollzeit, Altersbereich (dual slider)

**Default view:** Last 12 months, all departments.

## Step 3: Deliver Outputs

Save both artifacts:
- `sanitized-data.xlsx`
- `dashboard.html`

Print a text summary with:
- Number of records processed
- Revenue month range (from/to)
- Number of warnings (e.g., unknown cities)
- Active default filters

## Dependencies

```bash
pip install pandas openpyxl
```

The dashboard HTML is self-contained and requires an internet connection to load Chart.js from CDN (cdn.jsdelivr.net).

## Department Color Palette

Colors are defined in `generate_dashboard.py` and can be extended:

| Abteilung        | Hex       |
|------------------|-----------|
| Vertrieb         | #3B82F6   |
| Marketing        | #F59E0B   |
| IT               | #10B981   |
| HR               | #EC4899   |
| Finance          | #8B5CF6   |
| Operations       | #EF4444   |
| Customer Support | #06B6D4   |
| Produkt          | #84CC16   |
| Einkauf          | #F97316   |
| Recht            | #6366F1   |

To add new departments, update the `DEPT_COLORS` dict in `generate_dashboard.py`.

## Extending the City Mapping

The `CITY_BUNDESLAND` dict in `sanitize.py` covers major DE/AT/CH cities. To add cities, append entries:

```python
"stadtname": "Bundesland/Kanton",
```

Unknown cities are flagged as warnings and set to "Unbekannt".
