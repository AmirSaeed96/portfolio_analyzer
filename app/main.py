import os
import shutil
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import pandas as pd
import streamlit as st

from updater import (
    EXCEL_PATH, MONEY_MARKETS,
    fetch_one, get_tickers,
    write_to_excel, read_sheet1_grouped,
)

st.set_page_config(page_title="ETF Updater", page_icon="📈", layout="wide")

# ── Formatters ────────────────────────────────────────────────────────────────
def fp(v):
    return f"{v * 100:+.2f}%" if v is not None else "—"

def fpr(v):
    return f"${v:,.2f}" if v is not None else "—"

def fw(v):
    return f"{v * 100:.2f}%" if v is not None else "—"

def trend(v):
    if v is None:
        return "⬜"
    return "🟢" if v >= 0 else "🔴"

# ── Header ────────────────────────────────────────────────────────────────────
st.title("📈 ETF Market Data Updater")
st.caption(datetime.now().strftime("%A, %B %d, %Y"))

# ── Resolve Excel path ────────────────────────────────────────────────────────
# Docker mode: EXCEL_PATH env var is set — skip the uploader entirely.
# Cloud mode: no env var — show a file uploader and work from /tmp/.

if EXCEL_PATH and os.path.exists(EXCEL_PATH):
    # Docker / local: file is already on disk
    working_path = EXCEL_PATH
else:
    # Streamlit Cloud: need user to upload the file
    uploaded = st.file_uploader(
        "Upload your ETF Excel file to get started",
        type=["xlsx"],
        help="Upload `ETF_DETAILS_daily_update_interface.xlsx`. "
             "The file is processed locally — nothing is stored on any server.",
    )

    if uploaded is None:
        st.info("Please upload your Excel file above to continue.")
        st.stop()

    # Write uploaded bytes to a stable /tmp/ path for this session
    tmp_path = os.path.join(tempfile.gettempdir(), "ETF_DETAILS_daily_update_interface.xlsx")

    # Only overwrite if a new file was uploaded (compare name+size as a cheap heuristic)
    upload_key = f"{uploaded.name}_{uploaded.size}"
    if st.session_state.get("_upload_key") != upload_key:
        with open(tmp_path, "wb") as f:
            f.write(uploaded.getvalue())
        st.session_state["_upload_key"] = upload_key
        # Clear stale results when a new file is uploaded
        st.session_state.pop("results", None)

    working_path = tmp_path

# ── Load tickers ──────────────────────────────────────────────────────────────
try:
    tickers = get_tickers(working_path)
except Exception as exc:
    st.error(f"Could not read Excel file: {exc}\n\nMake sure the file is a valid ETF tracker workbook.")
    st.stop()

tradeable  = [(s, n) for s, n in tickers if s not in MONEY_MARKETS]
mm_tickers = [(s, n) for s, n in tickers if s in MONEY_MARKETS]

# ── Action bar ────────────────────────────────────────────────────────────────
btn_col, dl_col, m1, m2, m3 = st.columns([2, 2, 1, 1, 1])

with btn_col:
    refresh = st.button("🔄  Refresh Market Data", type="primary", use_container_width=True)

with dl_col:
    # Offer download of the (possibly updated) working file
    try:
        excel_bytes = open(working_path, "rb").read()
        st.download_button(
            label="⬇️  Download Excel",
            data=excel_bytes,
            file_name="ETF_DETAILS_daily_update_interface.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception:
        pass

m1.metric("Total Tickers",  len(tickers))
m2.metric("Market Tickers", len(tradeable))
m3.metric("Money Markets",  len(mm_tickers))

st.divider()

# ── Refresh ───────────────────────────────────────────────────────────────────
if refresh:
    results: dict = {}

    for sym, _ in mm_tickers:
        results[sym] = {
            "price": 1.00, "volume": None,
            "1d": None, "5d": None, "1mo": None,
            "1yr": None, "2yr": None, "5yr": None,
            "money_market": True,
        }

    tradeable_syms = [s for s, _ in tradeable]
    bar = st.progress(0, text="Connecting to Yahoo Finance…")

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(fetch_one, sym): sym for sym in tradeable_syms}
        done = 0
        for future in as_completed(futures):
            sym, data = future.result()
            results[sym] = data
            done += 1
            icon = "✓" if "error" not in data else "✗"
            bar.progress(done / len(tradeable_syms), text=f"{icon} {sym}  ({done}/{len(tradeable_syms)})")

    bar.progress(1.0, text="Saving to Excel…")

    try:
        write_to_excel(results, working_path)
    except PermissionError:
        bar.empty()
        st.error("❌ Cannot save — please close the Excel file first, then try again.")
        st.stop()
    except Exception as exc:
        bar.empty()
        st.error(f"❌ Error saving: {exc}")
        st.stop()

    bar.empty()

    st.session_state["results"] = results

    ok     = [s for s in results if "error" not in results[s]]
    failed = [(s, results[s]["error"]) for s in results if "error" in results[s]]

    st.success(f"✅  Updated {len(ok)} of {len(tickers)} tickers — {datetime.now().strftime('%I:%M %p')}")

    if failed:
        with st.expander(f"⚠️  {len(failed)} ticker(s) failed"):
            for sym, err in failed:
                st.write(f"**{sym}**: {err}")

    st.rerun()  # rerender so Download button serves the freshly written file

# ── Portfolio Visualization ───────────────────────────────────────────────────
st.subheader("Portfolio Holdings")

try:
    by_account = read_sheet1_grouped(working_path, results=st.session_state.get("results"))
except Exception as exc:
    st.warning(f"Could not load portfolio data: {exc}")
    st.stop()

has_data = bool(st.session_state.get("results"))

if not has_data:
    st.info("Hit **Refresh Market Data** above to populate prices and performance data.")

ACCOUNT_ORDER = ["IRA", "ROTH", "Brokerage"]
tab_labels = [a for a in ACCOUNT_ORDER if a in by_account] + \
             [a for a in by_account if a not in ACCOUNT_ORDER]

tabs = st.tabs(tab_labels)

for tab, acct in zip(tabs, tab_labels):
    with tab:
        for g in by_account.get(acct, []):
            header = (
                f"{trend(g['d1'])} **{g['etf']}**  —  {g['name']}  "
                f"|  {fpr(g['price'])}  |  1D: {fp(g['d1'])}"
            )
            with st.expander(header, expanded=False):
                if not g["holdings"]:
                    st.caption("No holdings listed.")
                    continue

                rows = [
                    {
                        "Name":   h["name"],
                        "Ticker": h["ticker"] or "—",
                        "Weight": fw(h["weight"]),
                        "Price":  fpr(h["price"]),
                        "1D %":   fp(h["d1"]),
                        "5D %":   fp(h["d5"]),
                        "1Mo %":  fp(h["mo1"]),
                        "1Yr %":  fp(h["yr1"]),
                        "2Yr %":  fp(h["yr2"]),
                        "5Yr %":  fp(h["yr5"]),
                    }
                    for h in g["holdings"]
                ]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
