# utils.py
import os
import io
import json
import subprocess
from pathlib import Path
from datetime import datetime

import pandas as pd
import requests
from io import BytesIO
import streamlit as st

# --- Constants / defaults ---
REPO_MASTER_PATH = Path("./MasterSeriesHistory.xlsx")   # local master in repo
REPO_SHEET_NAME  = "Master"

TRUTHY = {"1","true","yes","y","t"}

# ---------- Existing ----------
def load_from_github(url):
    """Load an Excel file from a GitHub raw link."""
    resp = requests.get(url)
    resp.raise_for_status()
    return pd.read_excel(BytesIO(resp.content), engine="openpyxl")

def match_series(comparison_df, master_df, rules_df, top_n):
    """Match requested series with master series using rules."""
    results = []
    unique_requests = comparison_df["RequestedSeries"].dropna().unique()
    total = len(unique_requests)

    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, requested in enumerate(unique_requests, start=1):
        status_text.text(f"Processing {i}/{total}: {requested}")
        progress_bar.progress(i / total)

        matches = master_df[master_df["SeriesName"].str.contains(str(requested), case=False, na=False)].copy()

        if not matches.empty:
            # Guard against zero/NaN
            denom = matches["UsageCount"].sum()
            if not pd.isna(denom) and float(denom) != 0.0:
                matches["UsagePercent"] = matches["UsageCount"] / denom * 100.0
            else:
                matches["UsagePercent"] = 0.0

            matches["RequestedSeries"] = requested
            matches = matches.sort_values("UsagePercent", ascending=False).head(top_n)
            results.append(matches)
        else:
            results.append(pd.DataFrame([{
                "RequestedSeries": requested,
                "SeriesName": None,
                "UsagePercent": 0.0
            }]))

    result_df = pd.concat(results, ignore_index=True)

    # Optional threshold
    if "MinUsagePercent" in rules_df.columns and not rules_df["MinUsagePercent"].empty:
        try:
            min_threshold = float(rules_df["MinUsagePercent"].max())
            result_df = result_df[result_df["UsagePercent"] >= min_threshold]
        except Exception:
            pass

    progress_bar.empty()
    status_text.empty()

    return result_df

# ---------- New helpers for master I/O ----------
def read_master(source="repo", path: str | None = None, sheet_name: str = REPO_SHEET_NAME) -> pd.DataFrame:
    """
    Read the master Excel.
      - source="repo": read from REPO_MASTER_PATH
      - source is a path string: read from that path
      - source="github": NOT used here (you already download templates via URLs in app)
    """
    if source == "repo":
        p = REPO_MASTER_PATH
    elif isinstance(source, str) and source not in {"repo", "github"}:
        p = Path(source)
    else:
        # If you later want to support a GitHub master, add your raw URL here.
        raise ValueError("Unsupported source for read_master. Use 'repo' or pass an explicit path.")

    if not p.exists():
        # Create minimal empty master if missing (safe default)
        cols = [
            "VariantID",
            "ManufacturerName",
            "Manufacturer Part Number",
            "Category",
            "Family",
            "RequestedSeries",
        ]
        st.warning(f"Master not found at {p}. Returning empty master with required columns.")
        return pd.DataFrame(columns=cols)

    return pd.read_excel(p, sheet_name=sheet_name, engine="openpyxl")

def write_master(df: pd.DataFrame, mode: str = "update_py", sheet_name: str = REPO_SHEET_NAME, input_rows: pd.DataFrame | None = None):
    """
    Save master:
      - mode="update_py": route through update.py (recommended). Expects an INPUT file that carries actions.
        If 'input_rows' is provided, we write a temp INPUT.xlsx for update.py to consume.
      - mode="local": write the DataFrame directly to REPO_MASTER_PATH (no rules/audits).
    """
    if mode == "local":
        REPO_MASTER_PATH.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(REPO_MASTER_PATH, engine="openpyxl", mode="w") as w:
            df.to_excel(w, sheet_name=sheet_name, index=False)
        return

    if mode == "update_py":
        # We need an INPUT file that encodes desired actions (update/delete).
        # If the admin UI gave us preview-changes in-memory (rename/append), synthesize input_rows accordingly.
        if input_rows is None or input_rows.empty:
            # No explicit rows = nothing to apply.
            st.info("No staged INPUT rows to apply via update.py.")
            return

        tmp_input = Path("._tmp_input_update.xlsx")
        input_rows.to_excel(tmp_input, index=False, engine="openpyxl")

        # Call update.py pointing to the repo master
        cmd = [
            "python", "update.py",
            "--input", str(tmp_input),
            "--master", str(REPO_MASTER_PATH),
            "--sheet", sheet_name
        ]
        st.write("Running:", " ".join(cmd))
        try:
            out = subprocess.run(cmd, capture_output=True, text=True, check=False)
            st.code(out.stdout or "", language="bash")
            if out.stderr:
                st.error(out.stderr)
            if out.returncode != 0:
                raise RuntimeError(f"update.py failed with code {out.returncode}")
        finally:
            try:
                tmp_input.unlink(missing_ok=True)
            except Exception:
                pass
        return

    raise ValueError("write_master: mode must be 'update_py' or 'local'")

# ---------- New: small in-memory preview ops for Admin panel ----------
def apply_update(dfm: pd.DataFrame, action: str, **kwargs) -> pd.DataFrame:
    """
    Preview-only transformations in the Admin panel (do NOT write master here).
    To actually persist, Admin panel calls write_master() later.

    Supported:
      - action="rename_series", kwargs: old_name, new_name
      - action="append_rows", kwargs: rows (DataFrame with same headers)
    """
    dfm = dfm.copy()

    if action == "rename_series":
        old = str(kwargs.get("old_name", "")).strip()
        new = str(kwargs.get("new_name", "")).strip()
        if not old or not new:
            st.warning("Please provide both Old and New SeriesName.")
            return dfm

        if "RequestedSeries" in dfm.columns:
            # If your master stores series in RequestedSeries (or SeriesName), adjust accordingly.
            target_col = "RequestedSeries"
        elif "SeriesName" in dfm.columns:
            target_col = "SeriesName"
        else:
            st.error("Could not find SeriesName/RequestedSeries column in master.")
            return dfm

        mask = dfm[target_col] == old
        cnt = int(mask.sum())
        dfm.loc[mask, target_col] = new
        st.info(f"Preview: renamed {cnt} row(s) from '{old}' → '{new}'.")
        return dfm

    if action == "append_rows":
        rows = kwargs.get("rows")
        if rows is None or rows.empty:
            st.info("No rows provided to append (preview).")
            return dfm
        # Align columns: keep only known columns, add missing as empty
        missing = [c for c in dfm.columns if c not in rows.columns]
        for c in missing:
            rows[c] = pd.NA
        rows = rows[dfm.columns]
        dfm = pd.concat([dfm, rows], ignore_index=True)
        st.info(f"Preview: appended {len(rows)} row(s).")
        return dfm

    st.warning(f"Unknown action '{action}'. No changes.")
    return dfm
