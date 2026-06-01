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

POLICY_NAME_PAT = re.compile(r"invest(?:ment)?\s*polic", re.IGNORECASE)
ACCEPTED_EXTS = {".docx", ".xlsx", ".xlsm"}


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


def find_latest_year_dir(policies_dir: Path) -> Path | None:
    if not policies_dir.exists():
        return None
    year_dirs = []
    for child in policies_dir.iterdir():
        if not child.is_dir():
            continue
        m = re.search(r"\b(20\d{2}|19\d{2})\b", child.name)
        if m:
            year_dirs.append((int(m.group(1)), child))
    if not year_dirs:
        # No year subdir? fall back to policies_dir itself
        return policies_dir
    year_dirs.sort(reverse=True)
    return year_dirs[0][1]


def find_latest_policy_file(folder: Path) -> Path | None:
    candidates = []
    for f in folder.rglob("*"):
        if not f.is_file():
            continue
        if f.suffix.lower() not in ACCEPTED_EXTS:
            continue
        if not POLICY_NAME_PAT.search(f.name):
            continue
        # Skip lock files / temp
        if f.name.startswith("~$"):
            continue
        candidates.append(f)
    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


def slugify(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]+", "_", name).strip("_").lower() or "client"


def run(cmd, **kw):
    print(f"$ {' '.join(map(str, cmd))}")
    return subprocess.run(cmd, check=True, **kw)


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

    for client_dir in sorted(clients_root.iterdir()):
        if not client_dir.is_dir():
            continue
        client_name = client_dir.name
        seen_clients.add(client_name)
        if args.only_client and client_name != args.only_client:
            continue

        policies_dir = client_dir / "policies"
        if not policies_dir.exists():
            # case-insensitive: some shares have "Policies"
            alt = next((c for c in client_dir.iterdir() if c.is_dir() and c.name.lower() == "policies"), None)
            if not alt:
                print(f"[skip] {client_name}: no policies/ folder")
                continue
            policies_dir = alt

        year_dir = find_latest_year_dir(policies_dir)
        if not year_dir:
            print(f"[skip] {client_name}: no year folder under {policies_dir}")
            continue

        policy_file = find_latest_policy_file(year_dir)
        if not policy_file:
            print(f"[skip] {client_name}: no investment-policy file under {year_dir}")
            continue

        h = file_sha256(policy_file)
        prior = manifest["clients"].get(client_name, {})
        if prior.get("file_hash") == h:
            print(f"[unchanged] {client_name}: {policy_file.name}")
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
        print(f"[extract] {client_name}: {policy_file.name}")
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
            changed_clients.append(client_name)
        except subprocess.CalledProcessError as e:
            print(f"[error] {client_name} extraction failed: {e}")

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
        print("[ok] no policy changes this run")


if __name__ == "__main__":
    main()
