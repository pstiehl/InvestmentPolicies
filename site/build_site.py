#!/usr/bin/env python3
"""
Build the static HTML site + Excel export from data/extracted/*.json.

Outputs:
  site/dist/index.html             — main table: clients × asset classes
  site/dist/client/<slug>.html     — per-client detail page
  site/dist/policies.xlsx          — OMS-ready Excel export
  site/dist/policies.json          — flat JSON used by the in-browser table
  site/dist/assets/                — palette.css + style.css copied here
"""
from __future__ import annotations

import json
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
EXTRACTED_DIR = REPO_ROOT / "data" / "extracted"
SITE_DIR = REPO_ROOT / "site"
DIST = SITE_DIR / "dist"
ASSETS_SRC = SITE_DIR / "assets"


def slugify(s: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]+", "_", s).strip("_").lower() or "client"


def fmt_pct(v):
    return "" if v is None else f"{v:g}%"


def fmt_years(v):
    return "" if v is None else f"{v:g} yr" + ("s" if v != 1 else "")


def load_policies():
    out = []
    if not EXTRACTED_DIR.exists():
        return out
    for p in sorted(EXTRACTED_DIR.glob("*.json")):
        try:
            out.append(json.loads(p.read_text()))
        except Exception as e:
            print(f"[skip] {p}: {e}")
    return out


def flatten(policies):
    """One row per (client, asset_class) for the index table + Excel export."""
    rows = []
    for pol in policies:
        client = pol.get("client", "")
        date = pol.get("policy_date") or ""
        for ac in pol.get("asset_classes", []) or []:
            rows.append({
                "client": client,
                "policy_date": date,
                "category": ac.get("category", ""),
                "subcategory": ac.get("subcategory") or "",
                "permitted": bool(ac.get("permitted", True)),
                "max_maturity_years": ac.get("max_maturity_years"),
                "max_wal_years": ac.get("max_weighted_average_life_years"),
                "max_portfolio_pct": ac.get("max_portfolio_pct"),
                "max_net_worth_pct": ac.get("max_net_worth_pct"),
                "single_issuer_pct": ac.get("single_issuer_pct"),
                "single_security_limit": ac.get("single_security_limit") or "",
                "min_credit_rating": ac.get("min_credit_rating") or "",
                "min_call_protection_years": ac.get("min_call_protection_years"),
                "structures_permitted": ", ".join(ac.get("structures_permitted") or []),
                "structures_prohibited": ", ".join(ac.get("structures_prohibited") or []),
            })
    return rows


def write_xlsx(rows, out_path: Path):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        print("[warn] openpyxl not installed; skipping xlsx export")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Permissible Trades"
    headers = [
        "Client", "Policy Date", "Asset Class", "Sub-Structure",
        "Permitted", "Max Maturity (yrs)", "Max WAL (yrs)",
        "Max Portfolio %", "Max Net Worth %", "Single Issuer %",
        "Single Security Limit", "Min Credit Rating", "Min Call Protection (yrs)",
        "Structures Permitted", "Structures Prohibited",
    ]
    ws.append(headers)
    header_fill = PatternFill("solid", fgColor="06314C")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx, _ in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center")

    for r in rows:
        ws.append([
            r["client"], r["policy_date"], r["category"], r["subcategory"],
            "Yes" if r["permitted"] else "No",
            r["max_maturity_years"], r["max_wal_years"],
            r["max_portfolio_pct"], r["max_net_worth_pct"], r["single_issuer_pct"],
            r["single_security_limit"], r["min_credit_rating"], r["min_call_protection_years"],
            r["structures_permitted"], r["structures_prohibited"],
        ])

    # auto-width-ish
    for col_idx, h in enumerate(headers, 1):
        max_len = max(len(str(h)),
                      *(len(str(ws.cell(row=i, column=col_idx).value or "")) for i in range(2, ws.max_row + 1)))
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max(max_len + 2, 10), 40)

    ws.freeze_panes = "A2"
    wb.save(out_path)


# ---------- HTML ----------

INDEX_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>ALM First — Client Investment Policies</title>
<link rel="stylesheet" href="assets/style.css">
</head>
<body>
<header>
  <h1>ALM First — Client Investment Policies</h1>
  <div class="meta">{n_clients} clients · {n_rows} asset-class rows · Generated {generated}</div>
</header>
<main>
  <div class="toolbar">
    <input id="search" type="text" placeholder="Search client, asset class, structure…">
    <select id="filterCategory"><option value="">All asset classes</option></select>
    <select id="filterClient"><option value="">All clients</option></select>
    <select id="filterPermitted">
      <option value="">All</option>
      <option value="yes">Permitted only</option>
      <option value="no">Prohibited only</option>
    </select>
    <div class="spacer"></div>
    <span class="stat" id="rowCount"></span>
    <a class="btn secondary" href="policies.xlsx" download>Download Excel</a>
  </div>
  <div class="legend">Click a column header to sort. Click a client to see their full policy detail.</div>
  <table id="t">
    <thead>
      <tr>
        <th data-k="client">Client</th>
        <th data-k="category">Asset Class</th>
        <th data-k="subcategory">Sub-Structure</th>
        <th data-k="permitted">Permitted</th>
        <th data-k="max_maturity_years" class="num">Max Maturity</th>
        <th data-k="max_portfolio_pct" class="num">Max Portfolio</th>
        <th data-k="max_net_worth_pct" class="num">Max Net Worth</th>
        <th data-k="single_issuer_pct" class="num">Single Issuer</th>
        <th data-k="single_security_limit">Single Security Limit</th>
        <th data-k="min_credit_rating">Min Rating</th>
        <th data-k="structures_prohibited">Prohibited Structures</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>
</main>
<footer>
  <span class="brand">ALM First</span> · Internal use only · Do not distribute externally
</footer>
<script src="assets/app.js"></script>
</body>
</html>
"""

CLIENT_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{client} — Investment Policy</title>
<link rel="stylesheet" href="../assets/style.css">
</head>
<body>
<header>
  <h1>{client}</h1>
  <div class="meta">Policy {policy_date} · Source: {policy_file} · Extracted {extracted_at}</div>
</header>
<main>
  <a class="back" href="../index.html">← Back to all clients</a>
  <div class="client-card">
    <table>
      <thead><tr>
        <th>Asset Class</th><th>Sub</th><th>OK?</th>
        <th class="num">Max Maturity</th><th class="num">Max Portfolio</th>
        <th class="num">Max Net Worth</th><th class="num">Single Issuer</th>
        <th>Single Security Limit</th><th>Min Rating</th>
        <th>Prohibited Structures</th>
      </tr></thead>
      <tbody>{rows}</tbody>
    </table>
    {unmapped_block}
  </div>
</main>
<footer><span class="brand">ALM First</span> · Internal use only</footer>
</body>
</html>
"""


APP_JS = r"""
const ROWS = window.__POLICY_ROWS__;

const search = document.getElementById('search');
const fCat = document.getElementById('filterCategory');
const fCli = document.getElementById('filterClient');
const fOk  = document.getElementById('filterPermitted');
const tbody = document.querySelector('#t tbody');
const rowCount = document.getElementById('rowCount');

const uniq = (arr) => [...new Set(arr)].sort();
uniq(ROWS.map(r => r.category)).forEach(c => fCat.add(new Option(c, c)));
uniq(ROWS.map(r => r.client)).forEach(c => fCli.add(new Option(c, c)));

let sortKey = 'client', sortDir = 1;
document.querySelectorAll('#t thead th').forEach(th => {
  th.addEventListener('click', () => {
    const k = th.dataset.k;
    if (sortKey === k) sortDir *= -1; else { sortKey = k; sortDir = 1; }
    render();
  });
});
[search, fCat, fCli, fOk].forEach(el => el.addEventListener('input', render));

function fmtNum(v, suffix='') { return v == null || v === '' ? '' : (v + suffix); }
function fmtPct(v) { return v == null || v === '' ? '' : (v + '%'); }
function fmtYrs(v) { return v == null || v === '' ? '' : (v + ' yrs'); }
function slug(s) { return (s || '').replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, '').toLowerCase(); }

function render() {
  const q = search.value.trim().toLowerCase();
  const cat = fCat.value, cli = fCli.value, ok = fOk.value;

  let rows = ROWS.filter(r => {
    if (cat && r.category !== cat) return false;
    if (cli && r.client !== cli) return false;
    if (ok === 'yes' && !r.permitted) return false;
    if (ok === 'no' && r.permitted) return false;
    if (q) {
      const hay = [r.client, r.category, r.subcategory, r.structures_permitted, r.structures_prohibited, r.min_credit_rating].join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  rows.sort((a, b) => {
    const av = a[sortKey] ?? '';
    const bv = b[sortKey] ?? '';
    if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * sortDir;
    return String(av).localeCompare(String(bv)) * sortDir;
  });

  tbody.innerHTML = rows.map(r => `
    <tr>
      <td><a class="client-link" href="client/${slug(r.client)}.html">${r.client}</a></td>
      <td>${r.category}</td>
      <td>${r.subcategory || ''}</td>
      <td>${r.permitted ? '<span class="pill ok">Yes</span>' : '<span class="pill bad">No</span>'}</td>
      <td class="num">${fmtYrs(r.max_maturity_years)}</td>
      <td class="num">${fmtPct(r.max_portfolio_pct)}</td>
      <td class="num">${fmtPct(r.max_net_worth_pct)}</td>
      <td class="num">${fmtPct(r.single_issuer_pct)}</td>
      <td>${r.single_security_limit || ''}</td>
      <td>${r.min_credit_rating || ''}</td>
      <td>${r.structures_prohibited || ''}</td>
    </tr>
  `).join('');
  rowCount.textContent = `${rows.length} rows`;
}
render();
"""


def render_client_page(pol):
    rows_html = []
    for ac in pol.get("asset_classes", []) or []:
        pill = '<span class="pill ok">Yes</span>' if ac.get('permitted', True) else '<span class="pill bad">No</span>'
        prohibited = ', '.join(ac.get('structures_prohibited') or [])
        rows_html.append(
            "<tr>"
            f"<td>{ac.get('category','')}</td>"
            f"<td>{ac.get('subcategory') or ''}</td>"
            f"<td>{pill}</td>"
            f"<td class='num'>{fmt_years(ac.get('max_maturity_years'))}</td>"
            f"<td class='num'>{fmt_pct(ac.get('max_portfolio_pct'))}</td>"
            f"<td class='num'>{fmt_pct(ac.get('max_net_worth_pct'))}</td>"
            f"<td class='num'>{fmt_pct(ac.get('single_issuer_pct'))}</td>"
            f"<td>{ac.get('single_security_limit') or ''}</td>"
            f"<td>{ac.get('min_credit_rating') or ''}</td>"
            f"<td>{prohibited}</td>"
            "</tr>"
        )
    unmapped = pol.get("unmapped_clauses") or []
    unmapped_block = ""
    if unmapped:
        items = "".join(f"<li>{c}</li>" for c in unmapped)
        unmapped_block = (
            '<div class="unmapped"><strong>Unmapped clauses (human review):</strong>'
            f'<ul>{items}</ul></div>'
        )
    return CLIENT_HTML.format(
        client=pol.get("client", "Unknown"),
        policy_date=pol.get("policy_date") or "—",
        policy_file=pol.get("policy_file") or "—",
        extracted_at=pol.get("extracted_at", ""),
        rows="".join(rows_html),
        unmapped_block=unmapped_block,
    )


def main():
    DIST.mkdir(parents=True, exist_ok=True)
    (DIST / "client").mkdir(exist_ok=True)
    (DIST / "assets").mkdir(exist_ok=True)
    for f in ("palette.css", "style.css"):
        shutil.copy(ASSETS_SRC / f, DIST / "assets" / f)

    policies = load_policies()
    rows = flatten(policies)

    # policies.json (table fuel)
    (DIST / "policies.json").write_text(json.dumps(rows, indent=2), encoding="utf-8")

    # app.js with rows inlined for zero network dependency
    app_js = f"window.__POLICY_ROWS__ = {json.dumps(rows)};\n{APP_JS}"
    (DIST / "assets" / "app.js").write_text(app_js, encoding="utf-8")

    # index.html
    n_clients = len({r["client"] for r in rows})
    generated = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    (DIST / "index.html").write_text(
        INDEX_HTML.format(n_clients=n_clients, n_rows=len(rows), generated=generated),
        encoding="utf-8",
    )

    # per-client detail pages
    for pol in policies:
        slug = slugify(pol.get("client", "client"))
        (DIST / "client" / f"{slug}.html").write_text(render_client_page(pol), encoding="utf-8")

    # excel export
    write_xlsx(rows, DIST / "policies.xlsx")

    print(f"[ok] built site at {DIST} ({n_clients} clients, {len(rows)} rows)")


if __name__ == "__main__":
    main()
