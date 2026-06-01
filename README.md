# ALM-First-Policies

Aggregates every ALM First client's investment policy into a single queryable
database — "can client X do trade Y today?" — and publishes a token-gated,
internal-only HTML dashboard plus an Excel export for OMS integration.

## Pipeline (end-to-end)

```
S:\Clients\<ClientName>\policies\<YYYY>\*investment*policy*.{docx,xlsx}
        │                          (Windows network share)
        │  1. scraper/scan_clients.py   ← runs weekly on Phil's work laptop
        ▼                                   via Task Scheduler
   data/inbox/<ClientName>/<file>.{docx,xlsx}  (local cache, .gitignored)
        │
        │  2. extractor/extract_policy.py   ← LLM-assisted, redacts before API
        ▼
   data/extracted/<ClientName>.json   (canonical schema, committed)
        │
        │  3. site/build_site.py       ← static HTML + xlsx export
        ▼
   site/dist/                          (published by GH Actions to Pages)
```

## Architecture decisions

1. **Scraper runs on Phil's work laptop** (only machine that can mount
   `S:\Clients`). Pushes structured JSON — *not raw policy docs* — to GitHub.
   Raw `.docx`/`.xlsx` stay on Phil's laptop in `data/inbox/` (`.gitignored`).
2. **Extractor uses Anthropic API** with the same redaction discipline as
   `~/work/alm-first-reporting`: client names pseudonymized before every
   API call, no prompt bodies persisted, only metadata logged.
3. **Site is GitHub Pages on a private repo**, served behind a token-gated
   link. ~50 internal viewers. **Never public.**
4. **Schema is derived from the CSP preferred-language doc** so every client
   policy lands in the same shape. Variations in client docs map onto the
   canonical fields; anything the extractor can't confidently map gets
   surfaced as `unmapped_clauses[]` for human review.

## Schema

See [`schema/policy.schema.json`](schema/policy.schema.json). Top-level:

```json
{
  "client": "Merck Sharp & Dohme",
  "policy_file": "Investment Policy 2026-03.docx",
  "policy_date": "2026-03-15",
  "extracted_at": "2026-06-01T15:00:00Z",
  "asset_classes": [
    {
      "category": "US Treasuries",
      "permitted": true,
      "max_maturity_years": 15,
      "max_portfolio_pct": 100,
      "max_net_worth_pct": null,
      "single_issuer_pct": null,
      "single_security_limit": null,
      "min_credit_rating": null,
      "structures_permitted": ["Bullet"],
      "structures_prohibited": []
    },
    ...
  ],
  "unmapped_clauses": ["..."]
}
```

## Repo layout

- `scraper/` — Windows-side scanner + Task Scheduler installer (Phil runs)
- `extractor/` — docx/xlsx → canonical JSON (LLM-assisted, with redaction)
- `schema/` — JSON Schema for the canonical policy shape
- `site/` — static HTML viewer (ALM color palette), build script, Excel export
- `data/extracted/` — committed JSON, one file per client
- `data/inbox/` — *gitignored*; raw policy docs cached on Phil's laptop only
- `samples/` — anonymized test fixtures
- `.github/workflows/` — publish-to-Pages on push to `main`
- `docs/` — color palette PDF + ops runbook

## Color palette (ALM First brand)

| Name        | Hex        |
| ----------- | ---------- |
| Peacock     | `#06314c`  |
| Chartreuse  | `#b0bc23`  |
| Moss        | `#63802b`  |
| Ocean       | `#006F98`  |
| Apricot     | `#d59600`  |
| Cerulean    | `#0a95b3`  |
| Emerald     | `#0d6440`  |
| Grape       | `#673dac`  |
| Raspberry   | `#b52d6a`  |

CSS variables live in [`site/assets/palette.css`](site/assets/palette.css).

## Setup (Phil's work laptop, one-time)

See [`scraper/INSTALL_WINDOWS.md`](scraper/INSTALL_WINDOWS.md).

## Local dev (extractor + site)

```bash
pip install --user anthropic python-docx openpyxl jinja2 jsonschema
export ANTHROPIC_API_KEY=...
python extractor/extract_policy.py samples/CSP_preferred_language.docx \
    --client "CSP Sample" --out data/extracted/csp_sample.json
python site/build_site.py
open site/dist/index.html
```
