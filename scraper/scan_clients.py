#!/usr/bin/env python3
r"""
Walk S:\Clients\* and find the latest investment policy file for each client.

For every client folder under CLIENTS_ROOT:
  1. Enter <client>\policies\<latest year>\
  2. Find the most recent file whose name fuzzy-matches "investment policy"
     (case-insensitive, allows extra words, allows .docx or .xlsx)
  3. Hash the file; if hash differs from data\manifest.json's last seen,
     enqueue it for extraction.

For each new/changed file:
  4. Call extractor/extract_policy.py to produce data\extracted\<slug>.json
  5. Commit + push the JSON (and the updated manifest) to the GitHub repo.
     Raw policy docs are NEVER pushed (they're .gitignored).

This script is intended to run on Phil's work laptop on a weekly schedule via
Windows Task Scheduler. See INSTALL_WINDOWS.md.

Environment (in .env next to this file):
  CLIENTS_ROOT=S:\Clients          # the share path
  REPO_ROOT=C:\Users\<you>\ALM-First-Policies
  ANTHROPIC_API_KEY=sk-ant-...
  GITHUB_PUSH=true                 # set false to skip git push (dry runs)
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

# Broad fuzzy match for filenames that look like investment-policy documents.
# Strategy: a list of patterns, tried in order. First one that matches anywhere
# in the filename wins. Order matters — most specific first so we don't pick
# up unrelated docs like "investment_committee_minutes".
POLICY_NAME_PATTERNS = [
    re.compile(r"invest(?:ment)?\s*[-_ ]*polic", re.IGNORECASE),   # "Investment Policy", "InvestmentPolicy", "Investment-Policy"
    re.compile(r"\bIPS\b", re.IGNORECASE),                          # IPS = Investment Policy Statement (industry shorthand)
    re.compile(r"polic(?:y|ies)", re.IGNORECASE),                   # last resort: anything with "policy" in the name
]
ACCEPTED_EXTS = {".docx", ".xlsx", ".xlsm", ".doc", ".xls"}


def load_env(script_dir: Path):
    env = {}
    env_file = script_dir / ".env"
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    # OS env overrides
    for k in ("CLIENTS_ROOT", "REPO_ROOT", "ANTHROPIC_API_KEY", "GITHUB_PUSH"):
        if k in os.environ:
            env[k] = os.environ[k]
    return env


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def find_year_dirs_newest_first(policies_dir: Path) -> list[Path]:
    """Return all year-named subdirs newest-first. Fallback to [policies_dir]
    when no year-named subdir exists."""
    if not policies_dir.exists():
        return []
    year_dirs = []
    for child in policies_dir.iterdir():
        if not child.is_dir():
            continue
        m = re.search(r"\b(20\d{2}|19\d{2})\b", child.name)
        if m:
            year_dirs.append((int(m.group(1)), child))
    if not year_dirs:
        return [policies_dir]
    year_dirs.sort(reverse=True)
    return [c for _, c in year_dirs]


def find_latest_policy_file(folder: Path, verbose: bool = False) -> Path | None:
    """Try each pattern in POLICY_NAME_PATTERNS in order; first that matches wins.

    If `verbose` is True and no match is found, print every candidate filename
    in the folder so the operator can see what's actually there.
    """
    all_docs = []
    for f in folder.rglob("*"):
        if not f.is_file():
            continue
        if f.suffix.lower() not in ACCEPTED_EXTS:
            continue
        if f.name.startswith("~$"):  # Office lock files
            continue
        all_docs.append(f)

    for pat in POLICY_NAME_PATTERNS:
        hits = [f for f in all_docs if pat.search(f.name)]
        if hits:
            hits.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            return hits[0]

    if verbose and all_docs:
        print(f"  [debug] no policy-name match in {folder}; saw {len(all_docs)} doc(s):")
        for f in all_docs[:20]:
            print(f"    - {f.name}")
    return None


def slugify(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]+", "_", name).strip("_").lower() or "client"


def run(cmd, **kw):
    print(f"$ {' '.join(map(str, cmd))}", flush=True)
    # Stream child stdout/stderr live so the operator sees progress in real time.
    return subprocess.run(cmd, check=True, **kw)


def log(msg: str):
    """Print + flush so PowerShell sees output as it happens, not at the end."""
    print(msg, flush=True)


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--max-clients", type=int, default=int(os.environ.get("ALMP_MAX_CLIENTS", "75")),
                    help="Refuse to extract more than this many clients in one run (safety belt against runaway loops).")
    ap.add_argument("--only-client", default=None,
                    help="Limit to a single client folder name. Useful for testing one policy.")
    args = ap.parse_args()

    script_dir = Path(__file__).resolve().parent
    env = load_env(script_dir)

    clients_root = Path(env.get("CLIENTS_ROOT", r"S:\Clients"))
    repo_root = Path(env.get("REPO_ROOT", script_dir.parent))
    push = env.get("GITHUB_PUSH", "true").lower() != "false"

    if not clients_root.exists():
        sys.exit(f"CLIENTS_ROOT not found: {clients_root}")

    log(f"[start] {datetime.now().strftime('%H:%M:%S')} scanning {clients_root}")
    log(f"[start] repo_root={repo_root}")
    if args.only_client:
        log(f"[start] limiting to client: {args.only_client}")
    log(f"[start] max_clients={args.max_clients}")

    manifest_path = repo_root / "data" / "manifest.json"
    extracted_dir = repo_root / "data" / "extracted"
    extracted_dir.mkdir(parents=True, exist_ok=True)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)

    if manifest_path.exists():
        manifest = json.loads(manifest_path.read_text())
    else:
        manifest = {"updated_at": None, "clients": {}}

    changed_clients = []
    seen_clients = set()
    extracted_this_run = 0

    all_client_dirs = sorted([d for d in clients_root.iterdir() if d.is_dir()])
    log(f"[scan] found {len(all_client_dirs)} client folder(s) under {clients_root}")

    for idx, client_dir in enumerate(all_client_dirs, 1):
        client_name = client_dir.name
        seen_clients.add(client_name)
        if args.only_client and client_name != args.only_client:
            continue
        log(f"\n[{idx}/{len(all_client_dirs)}] === {client_name} ===")

        policies_dir = client_dir / "policies"
        if not policies_dir.exists():
            # case-insensitive: some shares have "Policies"
            alt = next((c for c in client_dir.iterdir() if c.is_dir() and c.name.lower() == "policies"), None)
            if not alt:
                log(f"  [skip] no policies/ folder")
                continue
            policies_dir = alt

        year_dirs = find_year_dirs_newest_first(policies_dir)
        if not year_dirs:
            print(f"[skip] {client_name}: no year folder under {policies_dir}")
            continue

        # Walk year folders newest-first; fall back to older years if current
        # year has no matching policy file (common when the new year's policy
        # hasn't been uploaded yet).
        policy_file = None
        searched_year = None
        for yd in year_dirs:
            policy_file = find_latest_policy_file(yd, verbose=(yd == year_dirs[0]))
            if policy_file:
                searched_year = yd
                break

        if not policy_file:
            log(f"  [skip] no policy file found under any of {[y.name for y in year_dirs]}")
            continue
        if searched_year != year_dirs[0]:
            log(f"  [fallback] using {searched_year.name}/ (newest year {year_dirs[0].name}/ had no match)")
        year_dir = searched_year
        log(f"  [found] {policy_file.relative_to(client_dir)}")

        h = file_sha256(policy_file)
        prior = manifest["clients"].get(client_name, {})
        if prior.get("file_hash") == h:
            log(f"  [unchanged] same hash as last run, skipping LLM call")
            # update path/year in case the file moved
            manifest["clients"][client_name] = {
                **prior,
                "policy_file": str(policy_file),
                "year_folder": year_dir.name,
            }
            continue

        # Safety belt: refuse to exceed max-clients in one run
        if extracted_this_run >= args.max_clients:
            print(f"[stop] Hit --max-clients={args.max_clients} cap. Remaining clients will run next invocation.")
            break

        # Run extractor
        slug = slugify(client_name)
        out_json = extracted_dir / f"{slug}.json"
        log(f"  [extract] calling Claude (typically 20-60s)…")
        t_start = datetime.now()
        try:
            extracted_this_run += 1
            run([
                sys.executable,
                str(repo_root / "extractor" / "extract_policy.py"),
                str(policy_file),
                "--client", client_name,
                "--policy-year", year_dir.name,
                "--out", str(out_json),
            ])
            manifest["clients"][client_name] = {
                "policy_file": str(policy_file),
                "year_folder": year_dir.name,
                "file_hash": h,
                "extracted_at": datetime.now(timezone.utc).isoformat(),
                "extracted_json": str(out_json.relative_to(repo_root)),
            }
            elapsed = (datetime.now() - t_start).total_seconds()
            log(f"  [done]  {elapsed:.1f}s → {out_json.relative_to(repo_root)}")
            changed_clients.append(client_name)
        except subprocess.CalledProcessError as e:
            log(f"  [error] extraction failed: {e}")

    # Prune clients that have disappeared from S:\Clients
    for vanished in [c for c in manifest["clients"] if c not in seen_clients]:
        print(f"[gone] {vanished} no longer present in {clients_root}")
        del manifest["clients"][vanished]

    manifest["updated_at"] = datetime.now(timezone.utc).isoformat()
    manifest_path.write_text(json.dumps(manifest, indent=2))

    # Rebuild site
    print("[build] regenerating site/dist/")
    run([sys.executable, str(repo_root / "site" / "build_site.py")])

    if push and changed_clients:
        print(f"[git] pushing changes for {len(changed_clients)} client(s)")
        try:
            run(["git", "-C", str(repo_root), "add",
                 "data/extracted", "data/manifest.json", "site/dist"])
            msg = f"weekly: {len(changed_clients)} client policy update(s) — {', '.join(changed_clients[:5])}"
            run(["git", "-C", str(repo_root), "commit", "-m", msg])
            run(["git", "-C", str(repo_root), "push", "origin", "main"])
        except subprocess.CalledProcessError as e:
            print(f"[error] git push failed: {e}. Site updated locally; please push manually.")
    elif not changed_clients:
        log("\n[ok] no policy changes this run")

    log(f"\n[finish] {datetime.now().strftime('%H:%M:%S')} extracted={extracted_this_run} changed={len(changed_clients)}")


if __name__ == "__main__":
    main()
