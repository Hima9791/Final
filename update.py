#!/usr/bin/env python3
"""
Update/Delete RequestedSeries in master Excel by key (NO MPN in key).

INPUT FILE (CSV or XLSX) must have columns (exact names):
  VariantID, ManufacturerName, Category, Family, RequestedSeries, is delete

RULES
- "is delete": 1/true/yes/y => delete; 0/blank/anything-else => update mode
- Keys: VariantID, ManufacturerName, Category, Family
- Update only if a master row exists and RequestedSeries is different (case-sensitive).
- If RequestedSeries is identical (case-sensitive), do NOT update; add comment.
- If no match in master, do NOT insert; add comment.
- If any required header is missing in the INPUT file, reject whole file (no changes).

USAGE
  python update.py --input path/to/input.xlsx
  optional:
    --master ./MasterSeriesHistory.xlsx
    --sheet Master
    --dry-run
    --log-dir ./logs
"""

import argparse
import sys
from pathlib import Path
from datetime import datetime
import pandas as pd

REQUIRED_COLS = [
    "VariantID",
    "ManufacturerName",
    "Category",
    "Family",
    "RequestedSeries",
    "is delete",
]

KEY_COLS = [
    "VariantID",
    "ManufacturerName",
    "Category",
    "Family",
]

TRUTHY = {"1", "true", "yes", "y", "t"}

def read_any(path: Path) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Input file not found: {p}")
    suf = p.suffix.lower()
    if suf in {".xlsx", ".xls"}:
        return pd.read_excel(p)
    if suf == ".csv":
        try:
            return pd.read_csv(p)
        except Exception:
            return pd.read_csv(p, sep=",", engine="python")
    raise ValueError(f"Unsupported input extension: {p.suffix}")

def load_master(master_path: Path, sheet_name: str) -> pd.DataFrame:
    """
    Load the master. If the file exists but the requested sheet is missing,
    fall back to the first sheet and print a warning (do NOT crash).
    If the file does not exist, return an empty master with required columns.
    """
    if not master_path.exists():
        cols = KEY_COLS + ["RequestedSeries"]
        print(f"[WARN] Master not found, creating a new empty one at {master_path} with sheet '{sheet_name}'.")
        return pd.DataFrame(columns=cols)

    try:
        return pd.read_excel(master_path, sheet_name=sheet_name)
    except ValueError as e:
        # Likely "Worksheet named 'X' not found"
        try:
            xls = pd.ExcelFile(master_path)
            names = xls.sheet_names
            if not names:
                print(f"[WARN] Master workbook has no sheets; creating empty master with columns {KEY_COLS + ['RequestedSeries']}.")
                return pd.DataFrame(columns=KEY_COLS + ["RequestedSeries"])
            fallback = names[0]
            print(f"[WARN] Sheet '{sheet_name}' not found in {master_path.name}. Using first sheet '{fallback}' instead.")
            return pd.read_excel(master_path, sheet_name=fallback)
        except Exception as e2:
            # If even listing sheets fails, return empty to proceed safely
            print(f"[WARN] Could not inspect sheets in {master_path.name}: {e2}. Creating empty master.")
            return pd.DataFrame(columns=KEY_COLS + ["RequestedSeries"])

def save_master(df: pd.DataFrame, master_path: Path, sheet_name: str):
    master_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(master_path, engine="openpyxl", mode="w") as w:
        df.to_excel(w, sheet_name=sheet_name, index=False)

def normalize_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def row_has_all_keys(row) -> bool:
    return all(normalize_str(row[k]) != "" for k in KEY_COLS)

def to_bool_delete(v) -> bool:
    s = normalize_str(v).lower()
    return s in TRUTHY

def main():
    ap = argparse.ArgumentParser(description="Update/Delete RequestedSeries by key in master Excel (NO MPN in key).")
    ap.add_argument("--input", required=True, type=Path, help="Input CSV/XLSX with required columns.")
    ap.add_argument("--master", type=Path, default=Path("./MasterSeriesHistory.xlsx"),
                    help="Path to master Excel in your repo (will be updated in place).")
    ap.add_argument("--sheet", default="Master", help="Sheet name in the master workbook.")
    ap.add_argument("--dry-run", action="store_true", help="Do everything except writing the master file.")
    ap.add_argument("--log-dir", type=Path, default=Path("./logs"), help="Where to write the audit log CSV.")
    args = ap.parse_args()

    # Load input
    try:
        df_in = read_any(args.input)
    except Exception as e:
        print(f"[ERROR] Failed to read input: {e}", file=sys.stderr)
        sys.exit(2)

    # Validate headers
    missing = [c for c in REQUIRED_COLS if c not in df_in.columns]
    if missing:
        print(f"[REJECT] Missing required columns: {missing}. Master not changed.")
        audit = pd.DataFrame([{
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "action": "reject_file",
            "key": "",
            "old_value": "",
            "new_value": "",
            "comment": f"Rejected: missing required columns {missing}",
        }])
        args.log_dir.mkdir(parents=True, exist_ok=True)
        audit_path = args.log_dir / f"update_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        audit.to_csv(audit_path, index=False)
        print(f"[INFO] Wrote audit log: {audit_path}")
        sys.exit(0)

    # Normalize input
    df_in = df_in.copy()
    for c in REQUIRED_COLS:
        df_in[c] = df_in[c].apply(normalize_str)

    # Load master (robust to missing sheet)
    try:
        df_master = load_master(args.master, args.sheet)
    except Exception as e:
        print(f"[ERROR] Failed to load master workbook: {e}", file=sys.stderr)
        sys.exit(2)

    # Ensure master has necessary columns
    if "RequestedSeries" not in df_master.columns:
        df_master["RequestedSeries"] = ""
    for c in KEY_COLS:
        if c not in df_master.columns:
            df_master[c] = ""

    # Normalize master strings
    for c in KEY_COLS + ["RequestedSeries"]:
        df_master[c] = df_master[c].apply(normalize_str)

    # Build index: tuple(key) -> list of row indices
    def make_key_tuple(sr: pd.Series):
        return tuple(sr[k] for k in KEY_COLS)

    master_index = {}
    for i, row in df_master.iterrows():
        key = make_key_tuple(row)
        master_index.setdefault(key, []).append(i)

    # Process rows
    audits = []
    updates = 0
    deletions = 0
    skipped = 0

    for _, r in df_in.iterrows():
        if not row_has_all_keys(r):
            audits.append({
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "action": "skip_row",
                "key": "|".join([r[k] for k in KEY_COLS]),
                "old_value": "",
                "new_value": "",
                "comment": "Row skipped: one or more key fields are blank.",
            })
            skipped += 1
            continue

        key = tuple(r[k] for k in KEY_COLS)
        delete_flag = to_bool_delete(r["is delete"])
        req_series_in = r["RequestedSeries"]

        if key not in master_index:
            audits.append({
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "action": "no_match",
                "key": "|".join(key),
                "old_value": "",
                "new_value": req_series_in,
                "comment": "No matching row in master; no action taken.",
            })
            skipped += 1
            continue

        idxs = master_index[key]  # may be multiple → apply to ALL
        if delete_flag:
            df_master.loc[idxs, "_to_delete"] = True
            deletions += len(idxs)
            audits.append({
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "action": "delete",
                "key": "|".join(key),
                "old_value": ";".join(df_master.loc[idxs, "RequestedSeries"].astype(str).tolist()),
                "new_value": "",
                "comment": f"Deleted {len(idxs)} row(s) by key.",
            })
        else:
            old_values = df_master.loc[idxs, "RequestedSeries"].astype(str).tolist()
            if len(set(old_values)) == 1 and old_values[0] == req_series_in:
                audits.append({
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "action": "no_change",
                    "key": "|".join(key),
                    "old_value": old_values[0],
                    "new_value": req_series_in,
                    "comment": "RequestedSeries identical (case-sensitive); no update.",
                })
                skipped += 1
            else:
                df_master.loc[idxs, "RequestedSeries"] = req_series_in
                updates += len(idxs)
                audits.append({
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "action": "update",
                    "key": "|".join(key),
                    "old_value": ";".join(old_values),
                    "new_value": req_series_in,
                    "comment": f"Updated RequestedSeries on {len(idxs)} row(s).",
                })

    # Apply deletions
    if "_to_delete" in df_master.columns:
        df_master = df_master[df_master["_to_delete"] != True].drop(columns=["_to_delete"])

    # Save audit (and master if not dry-run)
    args.log_dir.mkdir(parents=True, exist_ok=True)
    audit_df = pd.DataFrame(audits, columns=["timestamp", "action", "key", "old_value", "new_value", "comment"])
    audit_path = args.log_dir / f"update_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    audit_df.to_csv(audit_path, index=False)

    if args.dry_run:
        print(f"[DRY‑RUN] Would apply: updates={updates}, deletions={deletions}, skipped={skipped}")
        print(f"[DRY‑RUN] Audit log written → {audit_path}")
        sys.exit(0)

    try:
        save_master(df_master, args.master, args.sheet)
    except Exception as e:
        print(f"[ERROR] Failed to write master: {e}", file=sys.stderr)
        sys.exit(2)

    print(f"[OK] Master updated → {args.master} (sheet '{args.sheet}')")
    print(f"[OK] Actions: updates={updates}, deletions={deletions}, skipped={skipped}")
    print(f"[OK] Audit log → {audit_path}")

if __name__ == "__main__":
    main()
