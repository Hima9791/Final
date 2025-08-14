# app.py
import streamlit as st
import pandas as pd
import requests, subprocess, tempfile, os, base64, time
from io import BytesIO
from pathlib import Path
from datetime import datetime

# ========= Paths & sheet =========
BASE_DIR = Path(__file__).resolve().parent
REPO_MASTER_PATH = BASE_DIR / "MasterSeriesHistory.xlsx"   # adjust if stored elsewhere
REPO_SHEET_NAME  = "Master"

# ========= Templates =========
TEMPLATE_MASTER_URL = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/1c93e405525d5480fd43c46e15c3a1b12872d1ee/Serise/TempleteMasterSeriesHistory.xlsx"
TEMPLATE_INPUT_URL  = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/1c93e405525d5480fd43c46e15c3a1b12872d1ee/Serise/TempleteInput_series.xlsx"
TEMPLATE_RULES_URL  = "https://raw.githubusercontent.com/AbdallahHesham44/z2Tools/1c93e405525d5480fd43c46e15c3a1b12872d1ee/Serise/TempleteSampleSeriesRules.xlsx"

# ========= Required headers for update.py (NO MPN) =========
REQUIRED_UPDATE_COLS = [
    "VariantID",
    "ManufacturerName",
    "Category",
    "Family",
    "RequestedSeries",
    "is delete",
]

TRUTHY = {"1", "true", "yes", "y", "t"}

import sys
UPDATE_PY = BASE_DIR / "update.py"

# ========= Utils from your repo =========
from utils import match_series  # keep using your current matching util

# ---------- Helper: df -> Excel bytes ----------
def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    return output.getvalue()

# ---------- Cached template bytes (fast; avoids re-encoding) ----------
@st.cache_data(show_spinner=False)
def get_template_bytes(url: str) -> bytes:
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.content

# ---------- Run update.py (always real update; no dry-run) ----------
def run_update_py(input_df: pd.DataFrame):
    with tempfile.TemporaryDirectory() as td:
        tmp_in = Path(td) / "input.xlsx"
        input_df.to_excel(tmp_in, index=False, engine="openpyxl")
        cmd = [
            sys.executable,
            str(UPDATE_PY),
            "--input", str(tmp_in),
            "--master", str(REPO_MASTER_PATH),
            "--sheet", REPO_SHEET_NAME,
            # NOTE: no "--dry-run"
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True, cwd=str(BASE_DIR))
        return proc.returncode, proc.stdout, proc.stderr

# ---------- Master readers (with fallback) ----------
def read_repo_master() -> pd.DataFrame:
    """Full read of the master. Falls back to first sheet if 'Master' is missing."""
    if not REPO_MASTER_PATH.exists():
        st.warning(f"Master not found at {REPO_MASTER_PATH}. Returning empty frame.")
        return pd.DataFrame(columns=[
            "VariantID","ManufacturerName","Manufacturer Part Number",
            "Category","Family","RequestedSeries","SeriesName","UsageCount"
        ])
    try:
        return pd.read_excel(REPO_MASTER_PATH, sheet_name=REPO_SHEET_NAME, engine="openpyxl")
    except ValueError:
        try:
            xls = pd.ExcelFile(REPO_MASTER_PATH, engine="openpyxl")
            first = xls.sheet_names[0]
            st.warning(f"Sheet '{REPO_SHEET_NAME}' not found. Using first sheet '{first}' for reading.")
            return pd.read_excel(REPO_MASTER_PATH, sheet_name=first, engine="openpyxl")
        except Exception as e2:
            st.error(f"Failed to read master workbook: {e2}")
            return pd.DataFrame(columns=[
                "VariantID","ManufacturerName","Manufacturer Part Number",
                "Category","Family","RequestedSeries","SeriesName","UsageCount"
            ])

@st.cache_data(show_spinner=False)
def _read_excel_slice(path_str: str, sheet_name: str, nrows: int, mtime: float):
    """Cached partial read; mtime invalidates cache when file changes."""
    return pd.read_excel(path_str, sheet_name=sheet_name, engine="openpyxl", nrows=nrows)

def read_repo_master_preview(nrows: int = 100) -> pd.DataFrame:
    """Fast preview (first nrows). Falls back if 'Master' missing."""
    if not REPO_MASTER_PATH.exists():
        st.warning(f"Master not found at {REPO_MASTER_PATH}.")
        return pd.DataFrame(columns=[
            "VariantID","ManufacturerName","Manufacturer Part Number",
            "Category","Family","RequestedSeries","SeriesName","UsageCount"
        ])
    mtime = os.path.getmtime(REPO_MASTER_PATH)
    try:
        return _read_excel_slice(str(REPO_MASTER_PATH), REPO_SHEET_NAME, nrows, mtime)
    except ValueError:
        try:
            xls = pd.ExcelFile(REPO_MASTER_PATH, engine="openpyxl")
            first = xls.sheet_names[0]
            st.warning(f"Sheet '{REPO_SHEET_NAME}' not found. Previewing '{first}' instead.")
            return _read_excel_slice(str(REPO_MASTER_PATH), first, nrows, mtime)
        except Exception as e2:
            st.error(f"Failed to read master: {e2}")
            return pd.DataFrame()

def validate_required_headers(df: pd.DataFrame, required: list[str]) -> list[str]:
    return [c for c in required if c not in df.columns]

# ---------- GitHub push helper (always used after update) ----------
def github_upsert_file(
    repo_full_name: str,
    repo_path: str,
    local_file: Path,
    branch: str = "main",
    commit_message: str = "Update MasterSeriesHistory.xlsx via app",
    token: str | None = None,
):
    """
    Create/update a file in GitHub via the Contents API.
    repo_full_name: "owner/repo"
    repo_path: path in repo, e.g. "MasterSeriesHistory.xlsx" or "Serise/MasterSeriesHistory.xlsx"
    """
    if token is None:
        token = st.secrets.get("GITHUB_TOKEN")
    if not token:
        return (401, "Missing GITHUB_TOKEN in secrets")

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
    }
    base_url = f"https://api.github.com/repos/{repo_full_name}/contents/{repo_path}"

    # Read and base64-encode file content
    b = local_file.read_bytes()
    content_b64 = base64.b64encode(b).decode("utf-8")

    # Try to get existing file SHA on the target branch
    sha = None
    try:
        r_get = requests.get(base_url, headers=headers, params={"ref": branch}, timeout=30)
        if r_get.status_code == 200:
            sha = r_get.json().get("sha")
    except Exception:
        pass  # non-fatal

    payload = {
        "message": commit_message,
        "content": content_b64,
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha  # required for updates

    r_put = requests.put(base_url, headers=headers, json=payload, timeout=60)
    return (r_put.status_code, r_put.text)

# ========= UI =========
st.set_page_config(page_title="Series Matcher & Updater", layout="wide")
st.title("📊 Series Matcher & 🔧 Master Updater")

with st.sidebar:
    # Hard refresh to clear stale JS chunks if needed
    if st.button("🔄 Hard refresh app"):
        st.experimental_set_query_params(_cb=str(int(time.time())))
        st.rerun()

# --- Sidebar: templates (cached) ---
st.sidebar.header("📥 Download Templates")
st.sidebar.download_button(
    label="Download Master Template",
    data=get_template_bytes(TEMPLATE_MASTER_URL),
    file_name="TempleteMasterSeriesHistory.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.sidebar.download_button(
    label="Download Input Template",
    data=get_template_bytes(TEMPLATE_INPUT_URL),
    file_name="TempleteInput_series.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.sidebar.download_button(
    label="Download Rules Template",
    data=get_template_bytes(TEMPLATE_RULES_URL),
    file_name="TempleteSampleSeriesRules.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Sidebar: Input template (NO MPN)
template_df = pd.DataFrame(columns=REQUIRED_UPDATE_COLS)
st.sidebar.download_button(
    "⬇️ Download Update Input Template (no MPN)",
    data=df_to_excel_bytes(template_df),
    file_name="TempleteInput_series_NO_MPN.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# --- GitHub settings (auto-push uses these) ---
st.sidebar.header("🌐 GitHub Push Settings")
gh_repo   = st.sidebar.text_input("Repo (owner/name)", value=st.secrets.get("GH_REPO", ""), placeholder="e.g. AbdallahHesham44/z2Tools")
gh_branch = st.sidebar.text_input("Branch", value=st.secrets.get("GH_BRANCH", "main"))
# IMPORTANT: set this to where the Excel lives relative to repo root:
gh_path   = st.sidebar.text_input("Path in repo", value=st.secrets.get("GH_PATH", "MasterSeriesHistory.xlsx"))

with st.sidebar.expander("🔎 Preview repo master"):
    load_prev = st.checkbox("Load preview (first 100 rows)", value=False, key="preview_toggle")
    if load_prev:
        try:
            mdf = read_repo_master_preview(nrows=100)
            st.caption(f"Rows (showing first 100): {len(mdf):,}")
            st.dataframe(mdf, use_container_width=True)
            if REPO_MASTER_PATH.exists():
                st.download_button(
                    "Download current master.xlsx",
                    data=REPO_MASTER_PATH.read_bytes(),
                    file_name="MasterSeriesHistory__current.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as e:
            st.error(f"Failed to read master: {e}")

# --- Main: Two tabs: Compare | Update ---
tab_compare, tab_update = st.tabs(["🔍 Compare (Rules)", "🛠️ Update Master"])

# -------------------- Compare Tab --------------------
with tab_compare:
    st.subheader("Compare requested series vs repo master (using rules)")
    st.caption("Master is always loaded from your repo; no master upload needed.")

    col1, col2 = st.columns(2)
    with col1:
        comparison_file = st.file_uploader("Upload Comparison File (Input Series)", type=["xlsx","csv"], key="cmp_in")
    with col2:
        rules_file = st.file_uploader("Upload Rules File", type=["xlsx","csv"], key="cmp_rules")

    top_n = st.number_input("Top N Matches", min_value=1, max_value=50, value=5)
    run_compare = st.button("▶️ Run Compare")

    if run_compare:
        if not comparison_file or not rules_file:
            st.error("Please upload both Comparison and Rules files.")
        else:
            try:
                # load input
                if comparison_file.name.lower().endswith(".csv"):
                    comparison_df = pd.read_csv(comparison_file)
                else:
                    comparison_df = pd.read_excel(comparison_file)
                if rules_file.name.lower().endswith(".csv"):
                    rules_df = pd.read_csv(rules_file)
                else:
                    rules_df = pd.read_excel(rules_file)

                master_df = read_repo_master()

                # run matching (your existing util)
                result_df = match_series(comparison_df, master_df, rules_df, top_n)
                st.success("✅ Matching completed!")
                st.dataframe(result_df, use_container_width=True)

                st.download_button(
                    label="📥 Download Results",
                    data=df_to_excel_bytes(result_df),
                    file_name=f"MatchedSeriesResults_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Compare failed: {e}")

# -------------------- Update Tab --------------------
with tab_update:
    st.subheader("Apply updates/deletes to repo master (no dry-run; auto-push to GitHub)")
    st.caption("Upload an **Update Input** file with required headers. The update will be applied immediately and then pushed to GitHub.")

    up_file = st.file_uploader("Upload Update Input (xlsx/csv)", type=["xlsx","csv"], key="upd_in")
    colu1, colu2 = st.columns([1,1])
    with colu1:
        run_update = st.button("🚀 Run Update now (apply + push)")
    with colu2:
        show_master_after = st.checkbox("Reload master after run", value=True)

    if run_update:
        if not up_file:
            st.error("Please upload the Update Input file.")
        else:
            try:
                # load input with required headers
                if up_file.name.lower().endswith(".csv"):
                    in_df = pd.read_csv(up_file)
                else:
                    in_df = pd.read_excel(up_file)

                missing = validate_required_headers(in_df, REQUIRED_UPDATE_COLS)
                if missing:
                    st.error(f"Rejected: missing required columns: {missing}")
                    st.stop()

                # summarize action mix
                st.write("### Input summary")
                total_rows = len(in_df)
                del_mask = in_df["is delete"].astype(str).str.strip().str.lower().isin(TRUTHY)
                n_delete = int(del_mask.sum())
                n_update = total_rows - n_delete
                st.metric("Total rows", f"{total_rows:,}")
                st.metric("Deletes", f"{n_delete:,}")
                st.metric("Updates", f"{n_update:,}")

                with st.status("Applying update.py ...", expanded=True) as status:
                    code, out, err = run_update_py(in_df)
                    st.write("**update.py stdout:**")
                    st.code(out or "(no stdout)", language="bash")
                    if err:
                        st.write("**update.py stderr:**")
                        st.code(err, language="bash")

                    if code == 0:
                        status.update(label="update.py completed", state="complete")
                        st.success("Master updated locally.")

                        # ---------- Auto GitHub push (no confirmation) ----------
                        if REPO_MASTER_PATH.exists():
                            with st.status("Pushing updated master to GitHub...", expanded=True) as gh_status:
                                ok = True
                                if not gh_repo.strip():
                                    ok = False
                                    st.error("Missing repo (owner/name). Set it in the sidebar.")
                                if not gh_path.strip():
                                    ok = False
                                    st.error("Missing path in repo. Set it in the sidebar.")
                                if ok:
                                    try:
                                        code_push, body = github_upsert_file(
                                            repo_full_name=gh_repo.strip(),
                                            repo_path=gh_path.strip(),
                                            local_file=REPO_MASTER_PATH,
                                            branch=gh_branch.strip() or "main",
                                            commit_message=f"Update {gh_path.strip()} via Streamlit app",
                                        )
                                        st.code(body, language="json")
                                        if 200 <= code_push < 300:
                                            gh_status.update(label="GitHub push completed", state="complete")
                                            st.success(f"Pushed to GitHub: {gh_repo}:{gh_branch}/{gh_path}")
                                        else:
                                            gh_status.update(label="GitHub push failed", state="error")
                                            st.error(f"GitHub push failed with status {code_push}")
                                    except Exception as e:
                                        gh_status.update(label="GitHub push error", state="error")
                                        st.error(f"GitHub push error: {e}")
                        else:
                            st.error(f"Local master not found at {REPO_MASTER_PATH}; cannot push.")

                    else:
                        status.update(label="update.py failed", state="error")
                        st.error(f"update.py failed with return code {code}")

                if show_master_after and code == 0:
                    try:
                        mdf2 = read_repo_master()
                        st.write("### Master (after update)")
                        st.caption(f"Rows: {len(mdf2):,}")
                        st.dataframe(mdf2.head(100), use_container_width=True)
                    except Exception as e:
                        st.error(f"Could not reload master: {e}")

            except Exception as e:
                st.error(f"Update failed: {e}")
