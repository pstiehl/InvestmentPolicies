# Windows install — weekly policy scraper

This sets up the policy scraper to run **once a week** on Phil's work laptop
against `S:\Clients`. Total time: ~15 minutes.

## What you need

- Python 3.11 or 3.12 (already installed? check: `py --version`)
  - If not: <https://www.python.org/downloads/windows/> → tick "Add Python to PATH"
- Git for Windows: <https://git-scm.com/download/win>
- Read access to `S:\Clients` (you have this)
- A GitHub Personal Access Token for `pstiehl` with `repo` scope:
  <https://github.com/settings/tokens?type=beta> → "Fine-grained" → repo
  `pstiehl/ALM-First-Policies` → Read & Write on Contents
- Your Anthropic API key

## 1. Clone the repo

Open PowerShell:

```powershell
cd $env:USERPROFILE
git clone https://github.com/pstiehl/InvestmentPolicies.git
cd InvestmentPolicies
```

## 2. Install Python deps

```powershell
py -m pip install anthropic python-docx openpyxl jsonschema
```

> If you previously ran with `--user` and got `ModuleNotFoundError: No module
> named 'docx'`, that's the Microsoft Store Python quirk where `--user`
> installs land in a path the `py` launcher can't see. Drop `--user` (the
> command above) and re-run.

## 3. Configure the scraper

> ⚠️ *The block below is the **contents of a file**, not commands to run.*
> Do not paste these lines into PowerShell.

From inside `C:\Users\pstiehl\InvestmentPolicies`, run:

```powershell
Copy-Item scraper\.env.example scraper\.env
notepad scraper\.env
```

Notepad opens. Replace `sk-ant-PASTE-YOUR-REAL-KEY-HERE` with your real
Anthropic API key. Save and close. The file's final contents should look like
this (your key on line 3):

```text
CLIENTS_ROOT=S:\Clients
REPO_ROOT=C:\Users\pstiehl\InvestmentPolicies
ANTHROPIC_API_KEY=sk-ant-...your real key...
GITHUB_PUSH=true
```

Verify it looks right:

```powershell
Get-Content scraper\.env
```

Configure git to use your PAT (one-time):

```powershell
git config --global credential.helper manager
# On first push it will prompt for your token; Windows stores it.
```

## 4. Smoke test (manual run)

```powershell
py scraper\scan_clients.py
```

Expected: it walks every client folder under `S:\Clients`, finds the latest
policy file, calls Claude to extract structured JSON for any client whose
policy hash changed, regenerates `site\dist\`, commits, pushes.

First run will hit every client (≈ minutes per client × number of clients).
Future runs only re-extract changed files.

## 5. Schedule it weekly

Run as your normal user account (not SYSTEM — needs your `S:\` mount):

```powershell
$action = New-ScheduledTaskAction `
    -Execute "py" `
    -Argument "scraper\scan_clients.py" `
    -WorkingDirectory "C:\Users\pstiehl\InvestmentPolicies"

# Every Monday at 7:00 AM (laptop must be awake; pick a time you're online)
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 7:00AM

$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive

Register-ScheduledTask `
    -TaskName "ALM-First-Policies Weekly Scrape" `
    -Action $action -Trigger $trigger -Principal $principal `
    -Description "Scans S:\Clients for latest investment policies and publishes ALM-First-Policies dashboard."
```

To confirm:

```powershell
Get-ScheduledTask -TaskName "ALM-First-Policies Weekly Scrape" | Get-ScheduledTaskInfo
```

To run it on demand:

```powershell
Start-ScheduledTask -TaskName "ALM-First-Policies Weekly Scrape"
```

Logs land in `scraper\last_run.log` (the task action redirects stdout there).

## 6. View the published site

After a successful push, GitHub Actions rebuilds and publishes Pages. The
token-gated URL is wired up in the repo's Pages settings (Settings → Pages
→ "Private" + token). I'll add the URL to `docs/INTERNAL_LINK.md` once it's
deployed.

## Troubleshooting

- **"S:\Clients not found"** — you need to be on VPN / connected to the office
  share before the task runs. Adjust the trigger time to when you're online.
- **"git push failed"** — your PAT expired. Generate a new one and run any
  `git pull` interactively to re-prompt for credentials.
- **"ANTHROPIC_API_KEY missing"** — verify `scraper\.env` is at
  `InvestmentPolicies\scraper\.env` (not at the repo root).
- **A specific client extraction fails** — check `data\llm_audit.jsonl`. You
  can re-run a single client manually:
  ```powershell
  py extractor\extract_policy.py "S:\Clients\<Name>\policies\2026\Investment Policy.docx" --client "<Name>"
  ```
