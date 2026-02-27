#!/usr/bin/env python3
"""
Sheet-to-Dashboard: HTML Dashboard Generator
Reads sanitized-data.xlsx and produces a self-contained dashboard.html.

Usage:
    python generate_dashboard.py <sanitized-data.xlsx> [output.html]
"""

import sys
import json
import re
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("Missing dependencies. Run: pip install pandas openpyxl")
    sys.exit(1)


DEPT_COLORS = {
    "Vertrieb":         "#3B82F6",
    "Marketing":        "#F59E0B",
    "IT":               "#10B981",
    "HR":               "#EC4899",
    "Finance":          "#8B5CF6",
    "Operations":       "#EF4444",
    "Customer Support": "#06B6D4",
    "Produkt":          "#84CC16",
    "Einkauf":          "#F97316",
    "Recht":            "#6366F1",
    "Sonstige":         "#9CA3AF",
}

DEFAULT_COLOR = "#9CA3AF"


def get_dept_color(dept: str) -> str:
    return DEPT_COLORS.get(dept, DEFAULT_COLOR)


def detect_umsatz_cols(df: pd.DataFrame):
    pattern = re.compile(r"^Umsatz_\d{4}-\d{2}$", re.IGNORECASE)
    cols = [c for c in df.columns if pattern.match(str(c))]
    cols.sort()
    return cols


def build_dashboard(input_path: str, output_path: str):
    print(f"Lese: {input_path}")
    df = pd.read_excel(input_path, sheet_name="data")
    df.columns = [str(c).strip() for c in df.columns]

    umsatz_cols = detect_umsatz_cols(df)
    months = [c.replace("Umsatz_", "") for c in umsatz_cols]

    # Ensure numerics
    for col in umsatz_cols + ["Alter"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "Umsatz_Gesamt" in df.columns:
        df["Umsatz_Gesamt"] = pd.to_numeric(df["Umsatz_Gesamt"], errors="coerce").fillna(0)

    # Collect unique values for filter dropdowns
    abteilungen = sorted(df["Abteilung"].dropna().unique().tolist())
    berufe = sorted(df["Beruf"].dropna().unique().tolist())
    staedte = sorted(df["Stadt"].dropna().unique().tolist())
    bundeslaender = sorted(df["Bundesland"].dropna().unique().tolist()) if "Bundesland" in df.columns else []

    dept_colors_js = {d: get_dept_color(d) for d in abteilungen}

    # Prepare row data for JS
    rows = []
    for _, row in df.iterrows():
        r = {
            "vorname": str(row.get("Vorname", "")),
            "nachname": str(row.get("Nachname", "")),
            "stadt": str(row.get("Stadt", "")),
            "bundesland": str(row.get("Bundesland", "")) if "Bundesland" in df.columns else "",
            "abteilung": str(row.get("Abteilung", "")),
            "beruf": str(row.get("Beruf", "")),
            "teilzeit": str(row.get("Teilzeit", "Nein")),
            "alter": int(row.get("Alter", 0)),
            "umsatz_gesamt": float(row.get("Umsatz_Gesamt", 0)),
            "monatsumsatz": [float(row.get(c, 0)) for c in umsatz_cols],
        }
        rows.append(r)

    data_json = json.dumps(rows, ensure_ascii=False)
    months_json = json.dumps(months)
    abteilungen_json = json.dumps(abteilungen)
    berufe_json = json.dumps(berufe)
    staedte_json = json.dumps(staedte)
    bundeslaender_json = json.dumps(bundeslaender)
    dept_colors_json = json.dumps(dept_colors_js)

    # Default: last 12 months
    default_start = months[-12] if len(months) >= 12 else months[0]
    default_end = months[-1]

    html = f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sales Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --bg: #0f172a;
    --surface: #1e293b;
    --surface2: #263044;
    --border: #334155;
    --text: #e2e8f0;
    --text-muted: #94a3b8;
    --accent: #3b82f6;
    --accent2: #60a5fa;
    --success: #10b981;
    --warn: #f59e0b;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ background: var(--bg); color: var(--text); font-family: 'Segoe UI', system-ui, sans-serif; font-size: 14px; }}
  a {{ color: var(--accent2); }}
  .app {{ display: flex; flex-direction: column; min-height: 100vh; }}

  /* Header */
  .header {{ background: var(--surface); border-bottom: 1px solid var(--border); padding: 16px 24px; display: flex; align-items: center; gap: 16px; }}
  .header h1 {{ font-size: 20px; font-weight: 700; color: var(--text); }}
  .header .subtitle {{ color: var(--text-muted); font-size: 13px; }}

  /* Layout */
  .main {{ display: flex; flex: 1; }}
  .sidebar {{ width: 280px; min-width: 280px; background: var(--surface); border-right: 1px solid var(--border); padding: 20px 16px; display: flex; flex-direction: column; gap: 20px; overflow-y: auto; }}
  .content {{ flex: 1; padding: 24px; overflow-y: auto; display: flex; flex-direction: column; gap: 24px; }}

  /* Filter sidebar */
  .filter-section h3 {{ font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .06em; color: var(--text-muted); margin-bottom: 10px; }}
  .filter-section select, .filter-section input[type=range] {{ width: 100%; }}
  select {{ background: var(--surface2); border: 1px solid var(--border); color: var(--text); padding: 6px 8px; border-radius: 6px; font-size: 13px; cursor: pointer; }}
  select:focus {{ outline: none; border-color: var(--accent); }}
  select[multiple] {{ height: 110px; }}

  .range-row {{ display: flex; justify-content: space-between; font-size: 12px; color: var(--text-muted); margin-top: 4px; }}
  input[type=range] {{ accent-color: var(--accent); cursor: pointer; }}

  .btn {{ display: inline-flex; align-items: center; justify-content: center; gap: 6px; padding: 7px 14px; border-radius: 6px; border: 1px solid var(--border); background: var(--surface2); color: var(--text); font-size: 12px; cursor: pointer; transition: background .15s; }}
  .btn:hover {{ background: var(--border); }}
  .btn.primary {{ background: var(--accent); border-color: var(--accent); color: #fff; }}
  .btn.primary:hover {{ background: var(--accent2); }}
  .btn.active {{ background: var(--accent); border-color: var(--accent); color: #fff; }}

  .filter-actions {{ display: flex; gap: 8px; flex-wrap: wrap; }}

  /* KPI row */
  .kpi-row {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 16px; }}
  .kpi-card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 18px 20px; }}
  .kpi-label {{ font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .06em; color: var(--text-muted); margin-bottom: 8px; }}
  .kpi-value {{ font-size: 26px; font-weight: 700; color: var(--text); line-height: 1; }}
  .kpi-sub {{ font-size: 12px; color: var(--text-muted); margin-top: 6px; }}

  /* Charts grid */
  .charts-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
  .chart-card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 20px; }}
  .chart-card.wide {{ grid-column: 1 / -1; }}
  .chart-title {{ font-size: 13px; font-weight: 600; color: var(--text); margin-bottom: 16px; }}
  .chart-wrap {{ position: relative; height: 240px; }}
  .chart-wrap canvas {{ max-height: 240px !important; }}

  /* No data */
  .no-data {{ display: flex; align-items: center; justify-content: center; height: 160px; color: var(--text-muted); font-size: 14px; }}

  /* Summary bar */
  .summary-bar {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 14px 20px; font-size: 13px; color: var(--text-muted); }}
  .summary-bar strong {{ color: var(--text); }}

  /* Active dept chip */
  .dept-filter-row {{ display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; }}
  .dept-chip {{ padding: 4px 10px; border-radius: 99px; font-size: 11px; font-weight: 600; cursor: pointer; border: 2px solid transparent; opacity: .6; transition: opacity .15s, border-color .15s; }}
  .dept-chip.active {{ opacity: 1; border-color: rgba(255,255,255,.35); }}

  /* Responsive */
  @media (max-width: 900px) {{
    .sidebar {{ width: 220px; min-width: 220px; }}
    .charts-grid {{ grid-template-columns: 1fr; }}
    .chart-card.wide {{ grid-column: 1; }}
  }}
</style>
</head>
<body>
<div class="app">
  <div class="header">
    <div>
      <h1>&#x1F4CA; Sales Dashboard</h1>
      <div class="subtitle">Interaktiver Report &mdash; Umsatz &amp; Personaldaten</div>
    </div>
  </div>
  <div class="main">
    <!-- Sidebar filters -->
    <aside class="sidebar">
      <div class="filter-section">
        <h3>Zeitraum</h3>
        <label style="font-size:12px;color:var(--text-muted)">Von</label>
        <select id="f-start"></select>
        <label style="font-size:12px;color:var(--text-muted);margin-top:8px;display:block">Bis</label>
        <select id="f-end"></select>
        <div style="margin-top:10px">
          <button class="btn" id="btn-last12">Letzte 12 Monate</button>
        </div>
      </div>

      <div class="filter-section">
        <h3>Abteilung</h3>
        <select id="f-abteilung" multiple title="Mehrfachauswahl: Strg+Klick"></select>
        <div style="margin-top:6px;font-size:11px;color:var(--text-muted)">Strg+Klick f&uuml;r Mehrfachauswahl</div>
      </div>

      <div class="filter-section">
        <h3>Bundesland</h3>
        <select id="f-bundesland" multiple></select>
      </div>

      <div class="filter-section">
        <h3>Stadt</h3>
        <select id="f-stadt" multiple></select>
      </div>

      <div class="filter-section">
        <h3>Beruf</h3>
        <select id="f-beruf" multiple></select>
      </div>

      <div class="filter-section">
        <h3>Besch&auml;ftigung</h3>
        <select id="f-teilzeit">
          <option value="all">Alle</option>
          <option value="Nein">Vollzeit</option>
          <option value="Ja">Teilzeit</option>
        </select>
      </div>

      <div class="filter-section">
        <h3>Alter</h3>
        <div style="display:flex;gap:8px">
          <div style="flex:1">
            <label style="font-size:11px;color:var(--text-muted)">Min</label>
            <input type="range" id="f-alter-min" min="0" max="100" value="0">
          </div>
          <div style="flex:1">
            <label style="font-size:11px;color:var(--text-muted)">Max</label>
            <input type="range" id="f-alter-max" min="0" max="100" value="100">
          </div>
        </div>
        <div class="range-row"><span id="alter-min-lbl">0</span><span id="alter-max-lbl">100</span></div>
      </div>

      <div class="filter-actions">
        <button class="btn" id="btn-reset">&#x21BA; Zur&uuml;cksetzen</button>
      </div>
    </aside>

    <!-- Main content -->
    <main class="content">
      <!-- KPIs -->
      <div class="kpi-row">
        <div class="kpi-card">
          <div class="kpi-label">Gesamtumsatz</div>
          <div class="kpi-value" id="kpi-total">&#8212;</div>
          <div class="kpi-sub" id="kpi-total-sub"></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">&Oslash; Monatsumsatz / Person</div>
          <div class="kpi-value" id="kpi-avg-monthly">&#8212;</div>
          <div class="kpi-sub" id="kpi-avg-monthly-sub"></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">Top-Abteilung</div>
          <div class="kpi-value" id="kpi-top-dept">&#8212;</div>
          <div class="kpi-sub" id="kpi-top-dept-sub"></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">&Oslash; Alter</div>
          <div class="kpi-value" id="kpi-avg-age">&#8212;</div>
          <div class="kpi-sub" id="kpi-avg-age-sub"></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">Teilzeitquote</div>
          <div class="kpi-value" id="kpi-tz">&#8212;</div>
          <div class="kpi-sub" id="kpi-tz-sub"></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">Headcount</div>
          <div class="kpi-value" id="kpi-headcount">&#8212;</div>
          <div class="kpi-sub" id="kpi-headcount-sub"></div>
        </div>
      </div>

      <!-- Charts -->
      <div class="charts-grid">
        <!-- Timeline -->
        <div class="chart-card wide">
          <div class="chart-title">&#x1F4C8; Umsatzentwicklung &uuml;ber die Zeit</div>
          <div class="chart-wrap"><canvas id="chart-timeline"></canvas></div>
        </div>
        <!-- Dept bar -->
        <div class="chart-card">
          <div class="chart-title">&#x1F3E2; Umsatz nach Abteilung</div>
          <div class="chart-wrap"><canvas id="chart-dept"></canvas></div>
        </div>
        <!-- Top Berufe -->
        <div class="chart-card">
          <div class="chart-title">&#x1F3C6; Top 10 Berufe nach Umsatz</div>
          <div class="chart-wrap"><canvas id="chart-berufe"></canvas></div>
        </div>
        <!-- Distribution -->
        <div class="chart-card">
          <div class="chart-title">&#x1F4CA; Umsatzverteilung (Histogramm)</div>
          <div class="chart-wrap"><canvas id="chart-dist"></canvas></div>
        </div>
        <!-- Heatmap Monat x Abteilung -->
        <div class="chart-card">
          <div class="chart-title">&#x1F525; Saisonalitat: Umsatz-Heatmap</div>
          <div id="chart-heatmap-wrap" style="overflow:auto;max-height:260px;"></div>
        </div>
      </div>

      <!-- Summary -->
      <div class="summary-bar" id="summary-bar">Lade Daten...</div>
    </main>
  </div>
</div>

<script>
const ALL_ROWS = {data_json};
const MONTHS   = {months_json};
const ABTEILUNGEN = {abteilungen_json};
const DEPT_COLORS = {dept_colors_json};
const DEFAULT_START = "{default_start}";
const DEFAULT_END   = "{default_end}";

// ---- Helpers ----
function eur(v) {{
  return new Intl.NumberFormat('de-DE', {{style:'currency',currency:'EUR',maximumFractionDigits:0}}).format(v);
}}
function num(v) {{
  return new Intl.NumberFormat('de-DE').format(Math.round(v));
}}

// ---- Filter State ----
let state = {{
  start: DEFAULT_START,
  end:   DEFAULT_END,
  abteilungen: [],
  bundeslaender: [],
  staedte: [],
  berufe: [],
  teilzeit: 'all',
  alterMin: 0,
  alterMax: 100,
  clickDept: null,
  clickBeruf: null,
}};

// ---- Populate selects ----
function populateSelect(id, items) {{
  const sel = document.getElementById(id);
  sel.innerHTML = '';
  items.forEach(v => {{
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    sel.appendChild(opt);
  }});
}}

function populateTimeSel() {{
  ['f-start','f-end'].forEach(id => {{
    const sel = document.getElementById(id);
    sel.innerHTML = '';
    MONTHS.forEach(m => {{
      const opt = document.createElement('option');
      opt.value = m;
      opt.textContent = m;
      sel.appendChild(opt);
    }});
  }});
  document.getElementById('f-start').value = DEFAULT_START;
  document.getElementById('f-end').value   = DEFAULT_END;
}}

populateTimeSel();
populateSelect('f-abteilung', {abteilungen_json});
populateSelect('f-bundesland', {bundeslaender_json});
populateSelect('f-stadt', {staedte_json});
populateSelect('f-beruf', {berufe_json});

const ages = ALL_ROWS.map(r => r.alter).filter(a => a > 0);
const minAge = Math.min(...ages);
const maxAge = Math.max(...ages);
document.getElementById('f-alter-min').min = minAge;
document.getElementById('f-alter-min').max = maxAge;
document.getElementById('f-alter-min').value = minAge;
document.getElementById('f-alter-max').min = minAge;
document.getElementById('f-alter-max').max = maxAge;
document.getElementById('f-alter-max').value = maxAge;
state.alterMin = minAge;
state.alterMax = maxAge;
document.getElementById('alter-min-lbl').textContent = minAge;
document.getElementById('alter-max-lbl').textContent = maxAge;

// ---- Filter logic ----
function getSelectedValues(selectId) {{
  const sel = document.getElementById(selectId);
  return Array.from(sel.selectedOptions).map(o => o.value);
}}

function getActiveMths() {{
  const si = MONTHS.indexOf(state.start);
  const ei = MONTHS.indexOf(state.end);
  if (si < 0 || ei < 0) return MONTHS;
  return MONTHS.slice(si, ei + 1);
}}

function filterRows() {{
  const activeMths = getActiveMths();
  return ALL_ROWS.filter(r => {{
    if (state.abteilungen.length && !state.abteilungen.includes(r.abteilung)) return false;
    if (state.bundeslaender.length && !state.bundeslaender.includes(r.bundesland)) return false;
    if (state.staedte.length && !state.staedte.includes(r.stadt)) return false;
    if (state.berufe.length && !state.berufe.includes(r.beruf)) return false;
    if (state.teilzeit !== 'all' && r.teilzeit !== state.teilzeit) return false;
    if (r.alter < state.alterMin || r.alter > state.alterMax) return false;
    if (state.clickDept && r.abteilung !== state.clickDept) return false;
    if (state.clickBeruf && r.beruf !== state.clickBeruf) return false;
    return true;
  }}).map(r => {{
    // Slice monatsumsatz to active months
    const mIdxs = activeMths.map(m => MONTHS.indexOf(m));
    const sliced = mIdxs.map(i => (i >= 0 && i < r.monatsumsatz.length) ? r.monatsumsatz[i] : 0);
    return {{ ...r, umsatz_period: sliced.reduce((a,b) => a+b, 0), monatsumsatz_sliced: sliced }};
  }});
}}

// ---- Chart instances ----
let charts = {{}};

function destroyChart(id) {{
  if (charts[id]) {{ charts[id].destroy(); delete charts[id]; }}
}}

// ---- Update ----
function update() {{
  const rows = filterRows();
  const activeMths = getActiveMths();
  updateKPIs(rows, activeMths);
  updateTimeline(rows, activeMths);
  updateDeptChart(rows);
  updateBerufeChart(rows);
  updateDistChart(rows);
  updateHeatmap(rows, activeMths);
  updateSummary(rows, activeMths);
}}

// ---- KPIs ----
function updateKPIs(rows, activeMths) {{
  if (!rows.length) {{
    ['kpi-total','kpi-avg-monthly','kpi-top-dept','kpi-avg-age','kpi-tz','kpi-headcount'].forEach(id => {{
      document.getElementById(id).textContent = '–';
    }});
    return;
  }}
  const totalUmsatz = rows.reduce((s,r) => s + r.umsatz_period, 0);
  document.getElementById('kpi-total').textContent = eur(totalUmsatz);
  document.getElementById('kpi-total-sub').textContent = `${{activeMths.length}} Monate`;

  const avgMonthly = rows.length ? (totalUmsatz / rows.length / activeMths.length) : 0;
  document.getElementById('kpi-avg-monthly').textContent = eur(avgMonthly);
  document.getElementById('kpi-avg-monthly-sub').textContent = `${{rows.length}} Personen`;

  // Top dept
  const deptTotals = {{}};
  rows.forEach(r => {{ deptTotals[r.abteilung] = (deptTotals[r.abteilung] || 0) + r.umsatz_period; }});
  const topDept = Object.entries(deptTotals).sort((a,b) => b[1]-a[1])[0];
  if (topDept) {{
    document.getElementById('kpi-top-dept').textContent = topDept[0];
    const pct = (topDept[1] / totalUmsatz * 100).toFixed(1);
    document.getElementById('kpi-top-dept-sub').textContent = `${{eur(topDept[1])}} (${{pct}}%)`;
  }}

  // Avg age
  const ages = rows.map(r => r.alter).filter(a => a > 0);
  const avgAge = ages.length ? (ages.reduce((a,b)=>a+b,0)/ages.length) : 0;
  const sorted = [...ages].sort((a,b)=>a-b);
  const median = sorted.length ? (sorted.length % 2 === 0 ? (sorted[sorted.length/2-1]+sorted[sorted.length/2])/2 : sorted[Math.floor(sorted.length/2)]) : 0;
  document.getElementById('kpi-avg-age').textContent = avgAge.toFixed(1);
  document.getElementById('kpi-avg-age-sub').textContent = `Median: ${{median.toFixed(1)}}`;

  // Teilzeit
  const tzCount = rows.filter(r => r.teilzeit === 'Ja').length;
  const tzPct = rows.length ? (tzCount / rows.length * 100).toFixed(1) : 0;
  document.getElementById('kpi-tz').textContent = `${{tzPct}}%`;
  const tzUmsatz = rows.filter(r=>r.teilzeit==='Ja').reduce((s,r)=>s+r.umsatz_period,0);
  const fzUmsatz = rows.filter(r=>r.teilzeit==='Nein').reduce((s,r)=>s+r.umsatz_period,0);
  document.getElementById('kpi-tz-sub').textContent = `TZ: ${{eur(tzUmsatz/Math.max(rows.filter(r=>r.teilzeit==='Ja').length,1))}} | VZ: ${{eur(fzUmsatz/Math.max(rows.filter(r=>r.teilzeit==='Nein').length,1))}} Ø`;

  // Headcount
  document.getElementById('kpi-headcount').textContent = rows.length;
  document.getElementById('kpi-headcount-sub').textContent = `von ${{ALL_ROWS.length}} gesamt`;
}}

// ---- Timeline chart ----
function updateTimeline(rows, activeMths) {{
  destroyChart('timeline');
  const wrap = document.querySelector('#chart-timeline').parentElement;
  if (!rows.length) {{ wrap.innerHTML = '<div class="no-data">Keine Daten f&uuml;r den gew&auml;hlten Filter</div>'; return; }}
  if (!wrap.querySelector('canvas')) {{ wrap.innerHTML = '<canvas id="chart-timeline"></canvas>'; }}

  const mData = activeMths.map((m, mi) => {{
    const idx = MONTHS.indexOf(m);
    return rows.reduce((s, r) => s + (idx >= 0 && idx < r.monatsumsatz.length ? r.monatsumsatz[idx] : 0), 0);
  }});

  // Per-dept breakdown stacked
  const depts = [...new Set(rows.map(r=>r.abteilung))].sort();
  const datasets = depts.map(dept => {{
    const color = DEPT_COLORS[dept] || '#9CA3AF';
    const data = activeMths.map((m,mi) => {{
      const idx = MONTHS.indexOf(m);
      return rows.filter(r=>r.abteilung===dept).reduce((s,r) => s+(idx>=0&&idx<r.monatsumsatz.length?r.monatsumsatz[idx]:0),0);
    }});
    return {{ label: dept, data, backgroundColor: color+'88', borderColor: color, borderWidth: 1.5, fill: true, tension: 0.3, pointRadius: 0, pointHoverRadius: 4 }};
  }});

  charts['timeline'] = new Chart(document.getElementById('chart-timeline'), {{
    type: 'line',
    data: {{ labels: activeMths, datasets }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      interaction: {{ mode: 'index', intersect: false }},
      scales: {{
        x: {{ ticks: {{ color: '#94a3b8', maxTicksLimit: 12 }}, grid: {{ color: '#1e293b' }} }},
        y: {{ ticks: {{ color: '#94a3b8', callback: v => eur(v) }}, grid: {{ color: '#334155' }}, stacked: true }}
      }},
      plugins: {{
        legend: {{ labels: {{ color: '#e2e8f0', boxWidth: 12 }} }},
        tooltip: {{ callbacks: {{ label: ctx => ` ${{ctx.dataset.label}}: ${{eur(ctx.raw)}}` }} }}
      }}
    }}
  }});
}}

// ---- Dept bar chart ----
function updateDeptChart(rows) {{
  destroyChart('dept');
  const wrap = document.querySelector('#chart-dept').parentElement;
  if (!rows.length) {{ wrap.innerHTML = '<div class="no-data">Keine Daten</div>'; return; }}
  if (!wrap.querySelector('canvas')) {{ wrap.innerHTML = '<canvas id="chart-dept"></canvas>'; }}

  const totals = {{}};
  const counts = {{}};
  rows.forEach(r => {{ totals[r.abteilung] = (totals[r.abteilung]||0)+r.umsatz_period; counts[r.abteilung]=(counts[r.abteilung]||0)+1; }});
  const sorted = Object.entries(totals).sort((a,b)=>b[1]-a[1]);

  charts['dept'] = new Chart(document.getElementById('chart-dept'), {{
    type: 'bar',
    data: {{
      labels: sorted.map(e=>e[0]),
      datasets: [{{ data: sorted.map(e=>e[1]), backgroundColor: sorted.map(e=>DEPT_COLORS[e[0]]||'#9CA3AF'), borderRadius: 4, borderSkipped: false }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      onClick: (e, el) => {{
        if (!el.length) {{ state.clickDept = null; }} else {{
          const d = sorted[el[0].index][0];
          state.clickDept = (state.clickDept === d) ? null : d;
        }}
        update();
      }},
      scales: {{
        x: {{ ticks: {{ color:'#94a3b8', callback: v=>eur(v) }}, grid: {{ color:'#334155' }} }},
        y: {{ ticks: {{ color:'#e2e8f0' }}, grid: {{ display:false }} }}
      }},
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{ callbacks: {{ label: ctx => ` ${{eur(ctx.raw)}} (${{counts[ctx.label]}} Personen)` }} }}
      }}
    }}
  }});
}}

// ---- Top Berufe ----
function updateBerufeChart(rows) {{
  destroyChart('berufe');
  const wrap = document.querySelector('#chart-berufe').parentElement;
  if (!rows.length) {{ wrap.innerHTML = '<div class="no-data">Keine Daten</div>'; return; }}
  if (!wrap.querySelector('canvas')) {{ wrap.innerHTML = '<canvas id="chart-berufe"></canvas>'; }}

  const totals = {{}}; const counts = {{}};
  rows.forEach(r => {{ totals[r.beruf]=(totals[r.beruf]||0)+r.umsatz_period; counts[r.beruf]=(counts[r.beruf]||0)+1; }});
  const sorted = Object.entries(totals).sort((a,b)=>b[1]-a[1]).slice(0,10);

  charts['berufe'] = new Chart(document.getElementById('chart-berufe'), {{
    type: 'bar',
    data: {{
      labels: sorted.map(e=>e[0]),
      datasets: [{{ data: sorted.map(e=>e[1]), backgroundColor: '#3b82f688', borderColor: '#3b82f6', borderWidth: 1, borderRadius: 4, borderSkipped: false }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      onClick: (e, el) => {{
        if (!el.length) {{ state.clickBeruf = null; }} else {{
          const b = sorted[el[0].index][0];
          state.clickBeruf = (state.clickBeruf === b) ? null : b;
        }}
        update();
      }},
      scales: {{
        x: {{ ticks: {{ color:'#94a3b8', callback: v=>eur(v) }}, grid: {{ color:'#334155' }} }},
        y: {{ ticks: {{ color:'#e2e8f0', font:{{size:11}} }}, grid: {{ display:false }} }}
      }},
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{ callbacks: {{ label: ctx => ` ${{eur(ctx.raw)}} | HC: ${{counts[ctx.label]}} | Ø: ${{eur(ctx.raw/counts[ctx.label])}}` }} }}
      }}
    }}
  }});
}}

// ---- Distribution histogram ----
function updateDistChart(rows) {{
  destroyChart('dist');
  const wrap = document.querySelector('#chart-dist').parentElement;
  if (!rows.length) {{ wrap.innerHTML = '<div class="no-data">Keine Daten</div>'; return; }}
  if (!wrap.querySelector('canvas')) {{ wrap.innerHTML = '<canvas id="chart-dist"></canvas>'; }}

  const values = rows.map(r=>r.umsatz_period);
  const max = Math.max(...values);
  const bins = 15;
  const step = max / bins || 1;
  const counts = Array(bins).fill(0);
  values.forEach(v => {{ const b = Math.min(Math.floor(v/step), bins-1); counts[b]++; }});
  const labels = counts.map((_,i) => eur(i*step) + ' – ' + eur((i+1)*step));

  charts['dist'] = new Chart(document.getElementById('chart-dist'), {{
    type: 'bar',
    data: {{
      labels,
      datasets: [{{ data: counts, backgroundColor: '#8b5cf688', borderColor: '#8b5cf6', borderWidth: 1, borderRadius: 3 }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      scales: {{
        x: {{ ticks: {{ color:'#94a3b8', maxTicksLimit:6, maxRotation:30 }}, grid: {{ display:false }} }},
        y: {{ ticks: {{ color:'#94a3b8' }}, grid: {{ color:'#334155' }} }}
      }},
      plugins: {{ legend: {{ display:false }}, tooltip: {{ callbacks: {{ title: ctx => ctx[0].label, label: ctx => ` ${{ctx.raw}} Personen` }} }} }}
    }}
  }});
}}

// ---- Heatmap (table-based) ----
function updateHeatmap(rows, activeMths) {{
  const wrap = document.getElementById('chart-heatmap-wrap');
  if (!rows.length) {{ wrap.innerHTML = '<div class="no-data">Keine Daten</div>'; return; }}

  const depts = [...new Set(rows.map(r=>r.abteilung))].sort();
  // Build dept x month matrix
  const matrix = {{}};
  depts.forEach(d => {{ matrix[d] = {{}}; activeMths.forEach(m => {{ matrix[d][m] = 0; }}); }});
  rows.forEach(r => {{
    activeMths.forEach((m,mi) => {{
      const idx = MONTHS.indexOf(m);
      matrix[r.abteilung][m] = (matrix[r.abteilung][m]||0) + (idx>=0&&idx<r.monatsumsatz.length?r.monatsumsatz[idx]:0);
    }});
  }});

  const allVals = depts.flatMap(d => activeMths.map(m => matrix[d][m]));
  const maxVal = Math.max(...allVals) || 1;

  let html = '<table style="border-collapse:collapse;font-size:10px;width:100%">';
  html += '<tr><th style="padding:3px 6px;text-align:left;color:#94a3b8;position:sticky;left:0;background:#1e293b">Abteilung</th>';
  // Show at most every 2nd or 3rd month label to avoid crowding
  const step = Math.ceil(activeMths.length / 12);
  activeMths.forEach((m,i) => {{
    html += `<th style="padding:3px 4px;color:#94a3b8;font-weight:400;min-width:28px;text-align:center">${{i % step === 0 ? m.slice(5) : ''}}</th>`;
  }});
  html += '</tr>';

  depts.forEach(d => {{
    const color = DEPT_COLORS[d] || '#9CA3AF';
    html += `<tr><td style="padding:3px 6px;color:#e2e8f0;white-space:nowrap;position:sticky;left:0;background:#1e293b">${{d}}</td>`;
    activeMths.forEach(m => {{
      const v = matrix[d][m];
      const intensity = Math.round((v / maxVal) * 180);
      const bg = color + intensity.toString(16).padStart(2,'0');
      html += `<td title="${{d}} / ${{m}}: ${{eur(v)}}" style="background:${{bg}};width:24px;height:20px;"></td>`;
    }});
    html += '</tr>';
  }});
  html += '</table>';
  wrap.innerHTML = html;
}}

// ---- Summary ----
function updateSummary(rows, activeMths) {{
  const bar = document.getElementById('summary-bar');
  const filters = [];
  if (state.abteilungen.length) filters.push(`Abteilung: <strong>${{state.abteilungen.join(', ')}}</strong>`);
  if (state.bundeslaender.length) filters.push(`Bundesland: <strong>${{state.bundeslaender.join(', ')}}</strong>`);
  if (state.staedte.length) filters.push(`Stadt: <strong>${{state.staedte.join(', ')}}</strong>`);
  if (state.berufe.length) filters.push(`Beruf: <strong>${{state.berufe.join(', ')}}</strong>`);
  if (state.teilzeit !== 'all') filters.push(`<strong>${{state.teilzeit==='Ja'?'Nur Teilzeit':'Nur Vollzeit'}}</strong>`);
  if (state.clickDept) filters.push(`Chart-Filter Abteilung: <strong>${{state.clickDept}}</strong>`);
  if (state.clickBeruf) filters.push(`Chart-Filter Beruf: <strong>${{state.clickBeruf}}</strong>`);
  const filterStr = filters.length ? filters.join(' &middot; ') : '<strong>Alle Abteilungen</strong>';
  bar.innerHTML = `Zeige <strong>${{rows.length}}</strong> von <strong>${{ALL_ROWS.length}}</strong> Personen &mdash; Zeitraum <strong>${{state.start}}</strong> bis <strong>${{state.end}}</strong> &mdash; ${{filterStr}}.`;
}}

// ---- Event wiring ----
document.getElementById('f-start').addEventListener('change', e => {{ state.start = e.target.value; update(); }});
document.getElementById('f-end').addEventListener('change', e => {{ state.end = e.target.value; update(); }});

document.getElementById('btn-last12').addEventListener('click', () => {{
  state.start = MONTHS.length >= 12 ? MONTHS[MONTHS.length-12] : MONTHS[0];
  state.end   = MONTHS[MONTHS.length-1];
  document.getElementById('f-start').value = state.start;
  document.getElementById('f-end').value   = state.end;
  update();
}});

['f-abteilung','f-bundesland','f-stadt','f-beruf'].forEach(id => {{
  const key = {{
    'f-abteilung':'abteilungen','f-bundesland':'bundeslaender','f-stadt':'staedte','f-beruf':'berufe'
  }}[id];
  document.getElementById(id).addEventListener('change', e => {{
    state[key] = Array.from(e.target.selectedOptions).map(o=>o.value);
    update();
  }});
}});

document.getElementById('f-teilzeit').addEventListener('change', e => {{ state.teilzeit = e.target.value; update(); }});

document.getElementById('f-alter-min').addEventListener('input', e => {{
  state.alterMin = +e.target.value;
  if (state.alterMin > state.alterMax) {{ state.alterMax = state.alterMin; document.getElementById('f-alter-max').value = state.alterMax; }}
  document.getElementById('alter-min-lbl').textContent = state.alterMin;
  update();
}});
document.getElementById('f-alter-max').addEventListener('input', e => {{
  state.alterMax = +e.target.value;
  if (state.alterMax < state.alterMin) {{ state.alterMin = state.alterMax; document.getElementById('f-alter-min').value = state.alterMin; }}
  document.getElementById('alter-max-lbl').textContent = state.alterMax;
  update();
}});

document.getElementById('btn-reset').addEventListener('click', () => {{
  state = {{
    start: DEFAULT_START, end: DEFAULT_END,
    abteilungen: [], bundeslaender: [], staedte: [], berufe: [],
    teilzeit: 'all', alterMin: minAge, alterMax: maxAge,
    clickDept: null, clickBeruf: null,
  }};
  populateTimeSel();
  ['f-abteilung','f-bundesland','f-stadt','f-beruf'].forEach(id => {{
    Array.from(document.getElementById(id).options).forEach(o => o.selected = false);
  }});
  document.getElementById('f-teilzeit').value = 'all';
  document.getElementById('f-alter-min').value = minAge;
  document.getElementById('f-alter-max').value = maxAge;
  document.getElementById('alter-min-lbl').textContent = minAge;
  document.getElementById('alter-max-lbl').textContent = maxAge;
  update();
}});

// Initial render
update();
</script>
</body>
</html>"""

    Path(output_path).write_text(html, encoding="utf-8")
    print(f"Gespeichert: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_dashboard.py <sanitized-data.xlsx> [output.html]")
        sys.exit(1)
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "dashboard.html"
    build_dashboard(inp, out)
