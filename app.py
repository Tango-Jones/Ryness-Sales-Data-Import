"""
Ryness Weekly Import - Streamlit app
Run:  streamlit run app.py
"""
from __future__ import annotations

import io
import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import sales_db as sdb
from ryness_parser import parse_pdf

st.set_page_config(page_title="Ryness Weekly Import", layout="wide")

DEFAULT_DB = os.environ.get("RYNESS_DB", str(Path(__file__).with_name("New_Home_Sales_Database.xlsx")))


def copy_db(db):
    return {k: (v.copy() if isinstance(v, pd.DataFrame) else v) for k, v in db.items()}


def db_bytes(db) -> bytes:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        tmp = f.name
    sdb.save_db(db, tmp)
    data = Path(tmp).read_bytes()
    os.unlink(tmp)
    return data


EVENT_COLORS = {"NEW": "#C6EFCE", "ADDED": "#DDEBF7", "RETURNED": "#FFEB9C",
                "DROPPED": "#FFC7CE", "POSSIBLE RENAME": "#E4DFEC"}


def style_events(df):
    def f(v):
        for k, c in EVENT_COLORS.items():
            if str(v).startswith(k):
                return f"background-color: {c}"
        return ""
    return df.style.map(f, subset=["event"]) if "event" in df else df


# ----------------------------------------------------------------- sidebar: database
st.sidebar.header("1. Database workbook")
mode = st.sidebar.radio("Where is the database?", ["File on this computer", "Upload a copy"],
                        help="'File on this computer' reads and saves the workbook in place.")
if mode == "File on this computer":
    db_path = st.sidebar.text_input("Workbook path", DEFAULT_DB)
    if st.sidebar.button("Load / reload", width="stretch") or "db" not in st.session_state:
        if Path(db_path).exists():
            st.session_state.db = sdb.load_db(db_path)
            st.session_state.db_name = Path(db_path).name
        else:
            st.session_state.db = sdb.empty_db()
            st.session_state.db_name = Path(db_path).name
            st.sidebar.info("No workbook at that path yet - a new one will be created on import.")
else:
    db_path = None
    up = st.sidebar.file_uploader("Database workbook (.xlsx)", type=["xlsx"])
    if up is not None and st.session_state.get("db_upload_id") != up.file_id:
        st.session_state.db = sdb.load_db(io.BytesIO(up.getvalue()))
        st.session_state.db_name = up.name
        st.session_state.db_upload_id = up.file_id
    if "db" not in st.session_state:
        st.session_state.db = sdb.empty_db()
        st.session_state.db_name = "New_Home_Sales_Database.xlsx"

db = st.session_state.db
comm, weekly = db["communities"], db["weekly"]
pdf_weeks = sorted(set(weekly.loc[weekly["source"] == sdb.SRC_PDF, "week_ending"]))
st.sidebar.metric("Communities tracked", len(comm))
st.sidebar.caption(f"Ryness weeks imported: {len(pdf_weeks)}"
                   + (f" (latest {pdf_weeks[-1]:%m/%d/%Y})" if pdf_weeks else ""))

st.title("Ryness Weekly Import")
tab_imp, tab_db, tab_legacy = st.tabs(["Import PDF", "Database", "One-time legacy import"])

# ----------------------------------------------------------------- import tab
with tab_imp:
    st.subheader("2. Upload this week's Ryness Report PDF")
    pdfs = st.file_uploader("Ryness Report PDF(s) - several weeks can be loaded at once",
                            type=["pdf"], accept_multiple_files=True)
    parsed = []
    for f in pdfs or []:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as t:
            t.write(f.getvalue())
            path = t.name
        res = parse_pdf(path)
        os.unlink(path)
        parsed.append((f.name, res))

    if parsed:
        parsed.sort(key=lambda x: x[1].week_ending or pd.Timestamp.min.date())
        preview = copy_db(db)
        summaries = []
        for name, res in parsed:
            df = res.to_frame()
            st.markdown(f"#### {name} &mdash; week ending **{res.week_ending:%A, %B %d, %Y}**")
            if res.week_ending in pdf_weeks:
                st.warning("This week is already in the database. Importing again replaces it.")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Communities", len(df))
            c2.metric("Gross sales", int(df.gross_sales.sum()))
            c3.metric("Cancellations", int(df.cancellations.sum()))
            c4.metric("Net sales", int(df.net_sales.sum()))
            checks = pd.DataFrame(res.section_checks)
            ok = int(checks["ok"].sum()) if len(checks) else 0
            if len(checks) and ok == len(checks):
                st.success(f"All {len(checks)} sections match the PDF's own TOTALS rows "
                           "(projects reporting, Wk's Sales, Wk's Cans).")
            else:
                st.error(f"{len(checks) - ok} of {len(checks)} sections do not match the PDF totals - review before importing.")
            for w in res.warnings:
                st.warning(w)
            with st.expander("Market totals and section checks"):
                st.dataframe(df.groupby("market", sort=False)[["gross_sales", "cancellations", "net_sales"]].sum(),
                             width="stretch")
                st.dataframe(checks, width="stretch", hide_index=True)
            with st.expander("Parsed rows"):
                st.dataframe(df.drop(columns=["week_ending"]), width="stretch", hide_index=True)
            s = sdb.import_report(preview, res, name)
            summaries.append(s)
            ch = s["changes"][["event", "market", "submarket", "community", "builder", "detail"]]
            st.markdown(f"**What changes:** {s['new']} new, {s['added']} added (first tracked), "
                        f"{s['dropped']} dropped (likely sold out), {s['returned']} returned")
            if len(ch):
                st.dataframe(style_events(ch), width="stretch", hide_index=True)
            st.divider()

        st.subheader("3. Import")
        if st.button("Import into database", type="primary", width="stretch"):
            st.session_state.db = preview
            st.session_state.out_bytes = db_bytes(preview)
            if db_path:
                Path(db_path).parent.mkdir(parents=True, exist_ok=True)
                sdb.save_db(preview, db_path)
                st.success(f"Saved to {db_path}")
            st.session_state.just_imported = True
            st.rerun()

    if st.session_state.pop("just_imported", False):
        st.success("Import complete. Download the updated workbook below (or open it from its path).")

    if "out_bytes" in st.session_state:
        st.download_button("Download updated workbook", data=st.session_state.out_bytes,
                           file_name=st.session_state.get("db_name", "New_Home_Sales_Database.xlsx"),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch")

# ----------------------------------------------------------------- database tab
with tab_db:
    if not len(comm):
        st.info("Database is empty - import a PDF first.")
    else:
        st.subheader("Status of all communities")
        st.dataframe(comm["status"].value_counts().rename("communities"), width="content")
        chg = db["change_log"]
        if len(chg):
            wk = st.selectbox("Change log for week", sorted(chg["week_ending"].unique(), reverse=True))
            st.dataframe(style_events(chg[chg["week_ending"] == wk]
                                      [["event", "market", "submarket", "community", "builder", "detail"]]),
                         width="stretch", hide_index=True)
        q = st.text_input("Find a community")
        view = comm if not q else comm[comm["community"].str.contains(q, case=False, na=False)]
        st.dataframe(view[["community_id", "market", "submarket", "community", "builder", "city", "units",
                           "status", "tracking", "first_week", "last_reported"]],
                     width="stretch", hide_index=True)
        if q and len(view) == 1:
            cid = view.iloc[0]["community_id"]
            hist = weekly[weekly["community_id"] == cid].sort_values("week_ending")
            st.line_chart(hist.set_index("week_ending")[["net_sales"]])
        st.subheader("Export for an online database")
        c1, c2 = st.columns(2)
        c1.download_button("Weekly_Data.csv", weekly.to_csv(index=False).encode(), "weekly_data.csv", "text/csv")
        c2.download_button("Communities.csv", comm.to_csv(index=False).encode(), "communities.csv", "text/csv")

# ----------------------------------------------------------------- legacy tab
with tab_legacy:
    st.markdown("Merge net-sales history from the old **New Home Sales Trends** workbook. "
                "Only needed once - the delivered database already includes v4.")
    lf = st.file_uploader("Old trends workbook (.xlsx)", type=["xlsx"], key="legacy")
    if lf is not None and st.button("Merge legacy history"):
        rep, cmp_ = sdb.import_legacy(st.session_state.db, io.BytesIO(lf.getvalue()), lf.name)
        st.session_state.db["legacy_report"], st.session_state.db["legacy_compare"] = rep, cmp_
        if db_path:
            sdb.save_db(st.session_state.db, db_path)
        st.session_state.out_bytes = db_bytes(st.session_state.db)
        st.dataframe(rep["method"].value_counts())
        st.dataframe(rep, width="stretch", hide_index=True)
