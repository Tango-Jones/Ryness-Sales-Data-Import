"""
sales_db.py
The workbook *is* the database. It holds two DB-ready tables plus views:

  Communities   one row per community (master list; never deleted)
  Weekly_Data   one row per community per week (long format -> online DB later)
  Name_Map      optional aliases: an alternate spelling -> an existing community_id
  Import_Log    one row per PDF import
  Change_Log    derived: NEW / ADDED / DROPPED / RETURNED / S/O / TSO events

  Gross Sales, Cancellations, Net Sales   wide weekly views (rebuilt every save)

Workflow:  db = load_db(path)  ->  summary = import_report(db, parse_result, file)
           ->  save_db(db, path)
"""
from __future__ import annotations

import difflib
import re
from datetime import datetime, timedelta, date

import pandas as pd

COMM_COLS = ["community_id", "market", "county_group", "submarket", "community", "builder",
             "city", "product_type", "units", "sold_to_date", "remaining", "total_lots",
             "grand_opening", "tracking", "status", "first_week", "last_reported", "weeks_off_report",
             "ryness_flag", "legacy_name", "legacy_group", "m_ord", "s_ord", "r_ord",
             "last_pdf_week", "user_notes"]
WEEK_COLS = ["week_ending", "community_id", "market", "gross_sales", "cancellations",
             "net_sales", "report_flag", "ryness_notes", "units", "sold_to_date",
             "source", "legacy_text", "source_file"]
MAP_COLS = ["alias", "market", "community_id"]
LOG_COLS = ["imported_at", "week_ending", "source_file", "markets", "communities_parsed",
            "gross_sales", "cancellations", "net_sales", "new", "added", "dropped",
            "returned", "totals_check", "warnings"]
CHG_COLS = ["week_ending", "market", "submarket", "community", "builder", "event", "detail",
            "community_id"]

SRC_PDF, SRC_OFF, SRC_LEGACY = "Ryness PDF", "Not on report", "Legacy v4 (net only)"
SRC_MANUAL = "Manual entry"
USER_FIELDS = ["total_lots", "grand_opening", "tracking", "user_notes"]
MANUAL_COLS = ["week_ending", "community_id", "gross_sales", "cancellations", "note"]


# --------------------------------------------------------------------------- names
_BUILDER_FIX = {"bonnani": "bonanni", "lenanr": "lennar", "tripointe": "tripointe",
                "baldwinsons": "baldwinandsons", "khovnanian": "khovnanian"}


def norm_name(s) -> str:
    s = str(s or "").strip().lower().replace("&", " and ")
    m = re.match(r"^(.*),\s*the$", s)          # "Gables, The" -> "gables"
    if m:
        s = m.group(1)
    s = re.sub(r"^the\s+", "", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


_STOP = {"at", "the", "of", "and", "in"}


def name_tokens(s) -> frozenset:
    return frozenset(t for t in norm_name(s).split() if t not in _STOP)


def name_similarity(a, b) -> float:
    ta, tb = name_tokens(a), name_tokens(b)
    if ta and ta == tb:
        return 1.0
    sa, sb = " ".join(sorted(ta)), " ".join(sorted(tb))
    return max(difflib.SequenceMatcher(None, norm_name(a), norm_name(b)).ratio(),
               difflib.SequenceMatcher(None, sa, sb).ratio())


def norm_builder(s) -> str:
    s = re.sub(r"[^a-z0-9]", "", str(s or "").lower())
    s = s.replace("development", "").replace("homes", "").replace("home", "")
    s = s.replace("communities", "").replace("company", "")
    return _BUILDER_FIX.get(s, s)


def builder_match(a, b) -> bool:
    a, b = norm_builder(a), norm_builder(b)
    if not a or not b:
        return True
    return a == b or a in b or b in a or difflib.SequenceMatcher(None, a, b).ratio() >= 0.8


def _slug(s):
    return re.sub(r"[^a-z0-9]+", "-", str(s).lower()).strip("-")


MARKET_CODES = {"LA-Orange-North": "LAON", "Inland Empire": "IE", "San Diego-Imperial": "SDI"}


def make_id(market, community, builder, existing: set) -> str:
    code = MARKET_CODES.get(market, _slug(market)[:6].upper())
    base = f"{code}-{_slug(norm_name(community))}"
    cid = base if base not in existing else f"{base}-{_slug(norm_builder(builder))}"
    n = 2
    while cid in existing:
        cid = f"{base}-{n}"
        n += 1
    return cid


# --------------------------------------------------------------------------- db io
def empty_db():
    return {"communities": pd.DataFrame(columns=COMM_COLS),
            "weekly": pd.DataFrame(columns=WEEK_COLS),
            "name_map": pd.DataFrame(columns=MAP_COLS),
            "import_log": pd.DataFrame(columns=LOG_COLS),
            "change_log": pd.DataFrame(columns=CHG_COLS),
            "manual": pd.DataFrame(columns=MANUAL_COLS)}


def load_db(path) -> dict:
    db = empty_db()
    if path is None:
        return db
    xl = pd.ExcelFile(path)
    sheets = {"communities": ("Communities", COMM_COLS, 2), "weekly": ("Weekly_Data", WEEK_COLS, 0),
              "name_map": ("Name_Map", MAP_COLS, 2), "import_log": ("Import_Log", LOG_COLS, 0),
              "manual": ("Manual_Entries", MANUAL_COLS, 2)}
    for key, (sheet, cols, hdr) in sheets.items():
        if sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet, header=hdr, dtype=object)
            df = df.rename(columns=_HEADER_TO_COL).reindex(columns=cols)
            df = df.dropna(how="all")
            db[key] = df
    for key, sheet in (("legacy_report", "Legacy_Match"), ("legacy_compare", "Legacy_Check")):
        if sheet in xl.sheet_names:
            db[key] = pd.read_excel(xl, sheet, dtype=object)
    w = db["weekly"]
    if len(w):
        w["week_ending"] = pd.to_datetime(w["week_ending"]).dt.date
        for col in ("gross_sales", "cancellations", "net_sales", "units", "sold_to_date"):
            w[col] = pd.to_numeric(w[col], errors="coerce")
    finalize(db)
    return db


# --------------------------------------------------------------------------- import
def _match(db, market, community, builder):
    """Return community_id for a parsed record, or None if new."""
    c = db["communities"]
    nn = norm_name(community)
    nm = db["name_map"]
    if len(nm):
        hit = nm[(nm["market"].fillna(market) == market) &
                 (nm["alias"].map(norm_name) == nn)]
        if len(hit):
            return hit.iloc[0]["community_id"]
    cand = c[(c["market"] == market) & (c["community"].map(norm_name) == nn)]
    if len(cand) == 1:
        return cand.iloc[0]["community_id"]
    if len(cand) > 1:
        for _, r in cand.iterrows():
            if norm_builder(r["builder"]) == norm_builder(builder):
                return r["community_id"]
        for _, r in cand.iterrows():
            if builder_match(r["builder"], builder):
                return r["community_id"]
    return None


def import_report(db: dict, parsed, source_file: str = "") -> dict:
    """Merge one parsed PDF into db (in place). Re-importing a week replaces it."""
    df = parsed.to_frame()
    week = parsed.week_ending
    if df.empty or week is None:
        raise ValueError("No community rows / week-ending date found in the PDF.")
    comm = db["communities"]
    markets = list(dict.fromkeys(df["market"]))

    # Section order from this PDF
    m_ord = {m: i for i, m in enumerate(markets)}
    sec_order = {}
    for (m, cg, sm) in df[["market", "county_group", "submarket"]].drop_duplicates().itertuples(index=False):
        sec_order[(m, sm)] = len(sec_order)

    existing_ids = set(comm["community_id"].dropna())
    new_rows, week_rows, upd = [], [], {}
    for i, r in df.iterrows():
        cid = _match(db, r.market, r.community, r.builder)
        if cid is None:
            cid = make_id(r.market, r.community, r.builder, existing_ids)
            existing_ids.add(cid)
            new_rows.append({"community_id": cid, "market": r.market})
        upd[cid] = {"market": r.market, "county_group": r.county_group, "submarket": r.submarket,
                    "community": r.community, "builder": r.builder, "city": r.city,
                    "product_type": r.product_type, "units": r.units,
                    "sold_to_date": r.sold_to_date,
                    "remaining": (r.units - r.sold_to_date) if pd.notna(r.units) and pd.notna(r.sold_to_date) else None,
                    "ryness_flag": r.status_flag,
                    "m_ord": m_ord[r.market], "s_ord": sec_order[(r.market, r.submarket)],
                    "r_ord": i, "last_pdf_week": week}
        week_rows.append({"week_ending": week, "community_id": cid, "market": r.market,
                          "gross_sales": r.gross_sales, "cancellations": r.cancellations,
                          "net_sales": r.net_sales, "report_flag": r.status_flag,
                          "ryness_notes": r.notes, "units": r.units,
                          "sold_to_date": r.sold_to_date, "source": SRC_PDF,
                          "legacy_text": None, "source_file": source_file})

    if new_rows:
        comm = pd.concat([comm, pd.DataFrame(new_rows).reindex(columns=COMM_COLS)], ignore_index=True)
    comm = comm.set_index("community_id")
    for cid, vals in upd.items():
        prev = comm.at[cid, "last_pdf_week"]
        prev = pd.to_datetime(prev).date() if pd.notna(prev) and prev != "" else None
        if prev is None or week >= prev:              # older imports don't overwrite attributes
            for k, v in vals.items():
                comm.at[cid, k] = v
    db["communities"] = comm.reset_index()

    # Replace this week's rows for every community in the markets covered by this PDF
    w = db["weekly"]
    mk_ids = set(db["communities"].loc[db["communities"]["market"].isin(markets), "community_id"])
    keep = ~((w["week_ending"] == week) & (w["community_id"].isin(mk_ids)))
    replaced = int((~keep & (w["source"] == SRC_PDF)).sum())
    db["weekly"] = pd.concat([w[keep], pd.DataFrame(week_rows)], ignore_index=True)

    # Drop orphan communities (e.g. created under a name later mapped via Name_Map)
    c = db["communities"]
    has_rows = set(db["weekly"]["community_id"])
    man = db.get("manual")
    if man is not None and len(man):
        has_rows |= set(man["community_id"].dropna())
    orphan = ~c["community_id"].isin(has_rows) & c["legacy_name"].isna() & \
        ~c["tracking"].astype(str).str.lower().str.startswith("manual")
    db["communities"] = c[~orphan].reset_index(drop=True)

    finalize(db)

    chg = db["change_log"]
    wk = chg[chg["week_ending"] == week]
    checks = parsed.section_checks
    bad = [c for c in checks if not c["ok"]]
    summary = {
        "imported_at": datetime.now().strftime("%Y-%m-%d %H:%M"), "week_ending": week,
        "source_file": source_file, "markets": ", ".join(markets),
        "communities_parsed": len(df), "gross_sales": int(df.gross_sales.sum()),
        "cancellations": int(df.cancellations.sum()), "net_sales": int(df.net_sales.sum()),
        "new": int((wk["event"] == "NEW").sum()), "added": int((wk["event"] == "ADDED").sum()),
        "dropped": int((wk["event"] == "DROPPED").sum()),
        "returned": int((wk["event"] == "RETURNED").sum()),
        "totals_check": f"{len(checks) - len(bad)}/{len(checks)} sections match PDF totals",
        "warnings": "; ".join(parsed.warnings + ([f"Replaced {replaced} previously imported rows for this week"] if replaced else [])),
    }
    db["import_log"] = pd.concat([db["import_log"], pd.DataFrame([summary])], ignore_index=True)
    summary["changes"] = wk.copy()
    return summary


# --------------------------------------------------------------------------- derive
def _present(w):
    return w["net_sales"].notna() & w["source"].isin([SRC_PDF, SRC_LEGACY, SRC_MANUAL])


def _manual_rows(db, existing_pdf_keys):
    m = db.get("manual")
    if m is None or not len(m):
        return pd.DataFrame(columns=WEEK_COLS)
    m = m.dropna(subset=["week_ending", "community_id"]).copy()
    m["week_ending"] = pd.to_datetime(m["week_ending"], errors="coerce").dt.date
    m = m.dropna(subset=["week_ending"])
    for col in ("gross_sales", "cancellations"):
        m[col] = pd.to_numeric(m[col], errors="coerce").fillna(0).astype(int)
    m = m[[k not in existing_pdf_keys for k in zip(m["community_id"], m["week_ending"])]]
    return pd.DataFrame({"week_ending": m["week_ending"], "community_id": m["community_id"],
                         "gross_sales": m["gross_sales"], "cancellations": m["cancellations"],
                         "net_sales": m["gross_sales"] - m["cancellations"],
                         "ryness_notes": m["note"], "source": SRC_MANUAL,
                         "source_file": "Manual_Entries"}).reindex(columns=WEEK_COLS)


def finalize(db):
    """Regenerate 'Not on report' rows, statuses, and the change log."""
    comm = db["communities"].copy()
    w = db["weekly"]
    w = w[~w["source"].isin([SRC_OFF, SRC_MANUAL])].copy()
    w["week_ending"] = pd.to_datetime(w["week_ending"]).dt.date
    pdf_keys_cw = set(zip(w.loc[w["source"] == SRC_PDF, "community_id"], w.loc[w["source"] == SRC_PDF, "week_ending"]))
    mrows = _manual_rows(db, pdf_keys_cw)
    if len(mrows):
        w = pd.concat([w, mrows], ignore_index=True)
    for col in ("gross_sales", "cancellations", "net_sales"):
        w[col] = pd.to_numeric(w[col], errors="coerce")
    comm["tracking"] = comm["tracking"].fillna("Ryness").replace("", "Ryness")
    cmap = comm.set_index("community_id")
    manual_ids = set(cmap.index[cmap["tracking"].astype(str).str.lower().str.startswith("manual")])

    pres = w[_present(w)]
    first_present = pres.groupby("community_id")["week_ending"].min()

    # 1) Not-on-report rows: every PDF week, every community in that market that
    #    had already appeared before that week but isn't in the PDF.
    pdf = w[w["source"] == SRC_PDF]
    so_txt = w[w["legacy_text"].astype(str).str.upper().str.strip() == "SOLD OUT"]
    first_so = so_txt.groupby("community_id")["week_ending"].min()
    off_rows = []
    for (wk, mkt), g in pdf.groupby(["week_ending", "market"]):
        listed = set(g["community_id"])
        ids = cmap.index[cmap["market"] == mkt]
        for cid in ids:
            if cid in listed or cid in manual_ids:
                continue
            fp, fs = first_present.get(cid), first_so.get(cid)
            if (fp is not None and fp < wk) or (fs is not None and fs < wk):
                off_rows.append({"week_ending": wk, "community_id": cid, "market": mkt,
                                 "source": SRC_OFF, "legacy_text": "SOLD OUT"})
    # PDF weeks supersede legacy rows for the same market/week
    pdf_keys = set(zip(pdf["week_ending"], pdf["market"]))
    w["_mkt"] = w["community_id"].map(cmap["market"])
    drop_legacy = (w["source"] == SRC_LEGACY) & ~w["community_id"].isin(manual_ids) & \
        pd.Series([k in pdf_keys for k in zip(w["week_ending"], w["_mkt"])], index=w.index, dtype=bool)
    w = w[~drop_legacy].drop(columns="_mkt")
    w = pd.concat([w, pd.DataFrame(off_rows, columns=WEEK_COLS)], ignore_index=True)
    w["market"] = w["community_id"].map(cmap["market"]).fillna(w["market"])
    w = w.sort_values(["week_ending", "community_id"]).reset_index(drop=True)
    db["weekly"] = w

    # 2) Change log, by submarket (so partial legacy coverage doesn't create noise)
    w["_sub"] = w["community_id"].map(cmap["submarket"])
    pres = w[_present(w)]
    events = []
    for sub, gs in pres.groupby("_sub"):
        weeks = sorted(gs["week_ending"].unique())
        pdf_weeks = set(gs.loc[gs["source"] == SRC_PDF, "week_ending"])
        seen: set = set()
        prev_set: set = set()
        prev_flags: dict = {}
        for wk in weeks:
            cur = gs[gs["week_ending"] == wk]
            cur_set = set(cur["community_id"])
            flags = dict(zip(cur["community_id"], cur["report_flag"].fillna("")))
            if wk in pdf_weeks:
                for cid in cur_set - prev_set:
                    if cid not in seen:
                        ev = "NEW" if seen else "ADDED"
                        note = cur.loc[cur["community_id"] == cid, "ryness_notes"].fillna("").iloc[0]
                        if str(note).strip().lower() == "new":
                            ev = "NEW"
                        detail = ("First appearance on the Ryness report" if ev == "NEW"
                                  else "First week tracked in this database")
                        if str(note).strip():
                            detail += f" (Ryness note: {note})"
                        events.append((wk, cid, ev, detail))
                    else:
                        events.append((wk, cid, "RETURNED", "Back on the report after being absent"))
                for cid in prev_set - cur_set - manual_ids:
                    lastp = max(x for x in weeks if x < wk and cid in set(gs.loc[gs["week_ending"] == x, "community_id"]))
                    events.append((wk, cid, "DROPPED", "No longer on the Ryness report - likely sold out / "
                                   f"temporarily sold out (last reported {lastp:%m/%d/%y})"))
                for cid in cur_set & set(flags):
                    f = flags[cid]
                    if f and prev_flags.get(cid, "") != f:
                        events.append((wk, cid, f"FLAG {f}",
                                       "Ryness marks this community " +
                                       ("Sold Out" if f == "S/O" else "Temporarily Sold Out")))
            seen |= cur_set
            prev_set, prev_flags = cur_set, flags
    w.drop(columns="_sub", inplace=True)
    chg = pd.DataFrame(events, columns=["week_ending", "community_id", "event", "detail"])
    chg["market"] = chg["community_id"].map(cmap["market"])
    chg["submarket"] = chg["community_id"].map(cmap["submarket"])
    chg["community"] = chg["community_id"].map(cmap["community"])
    chg["builder"] = chg["community_id"].map(cmap["builder"])
    # Possible renames: NEW + DROPPED in same submarket/week, same builder, similar name
    hints = []
    for (wk, sub), g in chg.groupby(["week_ending", "submarket"]):
        news, drops = g[g.event.isin(["NEW", "ADDED"])], g[g.event == "DROPPED"]
        for _, n in news.iterrows():
            for _, d in drops.iterrows():
                if builder_match(n.builder, d.builder) and name_similarity(n.community, d.community) >= 0.75:
                    hints.append({"week_ending": wk, "community_id": n.community_id,
                                  "event": "POSSIBLE RENAME", "market": n.market,
                                  "submarket": sub, "community": n.community, "builder": n.builder,
                                  "detail": f"May be the same project as '{d.community}' ({d.community_id}). "
                                            f"If so, add a Name_Map row: alias '{n.community}' -> {d.community_id}"})
    if hints:
        chg = pd.concat([chg, pd.DataFrame(hints)], ignore_index=True)
    order = {"NEW": 0, "ADDED": 1, "RETURNED": 2, "DROPPED": 3, "POSSIBLE RENAME": 4}
    chg["_o"] = chg["event"].map(order).fillna(5)
    chg = chg.sort_values(["week_ending", "_o", "market", "submarket", "community"],
                          ascending=[False, True, True, True, True]).drop(columns="_o")
    db["change_log"] = chg.reindex(columns=CHG_COLS).reset_index(drop=True)

    # 3) Community status (relative to the latest week its submarket has data)
    sub_last = {}
    for sub, gs in pres.assign(_sub=pres["community_id"].map(cmap["submarket"])).groupby("_sub"):
        sub_last[sub] = max(gs["week_ending"])
    sub_weeks = {}
    for sub, gs in pres.assign(_sub=pres["community_id"].map(cmap["submarket"])).groupby("_sub"):
        sub_weeks[sub] = sorted(gs["week_ending"].unique())
    last_present = pres.groupby("community_id")["week_ending"].max()
    ev_at = {(r.community_id, r.week_ending): r.event for r in chg.itertuples()
             if r.event in ("NEW", "ADDED", "RETURNED", "DROPPED")}
    latest_flag = (w[w["source"] == SRC_PDF].sort_values("week_ending")
                   .groupby("community_id")["report_flag"].last())

    statuses, firsts, lasts, offs = [], [], [], []
    for r in comm.itertuples():
        cid, sub = r.community_id, r.submarket
        fp, lp = first_present.get(cid), last_present.get(cid)
        L = sub_last.get(sub)
        firsts.append(fp)
        lasts.append(lp)
        if cid in manual_ids:
            statuses.append("MANUAL (not on Ryness)")
            offs.append(None)
            continue
        if lp is None:
            statuses.append("SOLD OUT")
            offs.append(None)
            continue
        if L is not None and lp == L:
            ev = ev_at.get((cid, L))
            flag = latest_flag.get(cid, "")
            if ev in ("NEW", "ADDED", "RETURNED"):
                st = ev
            elif isinstance(flag, str) and flag in ("S/O", "TSO"):
                st = f"ACTIVE - {flag} per Ryness"
            else:
                st = "ACTIVE"
            statuses.append(st)
            offs.append(0)
        else:
            n = len([x for x in sub_weeks.get(sub, []) if x > lp])
            statuses.append("DROPPED - likely sold out" if n == 1 else "SOLD OUT")
            offs.append(n)
    comm["status"], comm["first_week"], comm["last_reported"], comm["weeks_off_report"] = \
        statuses, firsts, lasts, offs
    db["communities"] = comm.reindex(columns=COMM_COLS)


# --------------------------------------------------------------------------- legacy
LEGACY_GROUP_MARKET = {
    "CENTRAL OC": "LA-Orange-North", "THE REST OF OC": "LA-Orange-North",
    "SOUTH BAY LOS ANGELES": "LA-Orange-North",
    "EAST SAN DIEGO": "San Diego-Imperial", "NORTH COASTAL SAN DIEGO": "San Diego-Imperial",
    "NO INLAND SAN DIEGO": "San Diego-Imperial",
    "DESERT SAN BERNARDINO": "Inland Empire", "CENTRAL RIVERSIDE": "Inland Empire",
    "NORTH CENTRAL RIVERSIDE": "Inland Empire", "SOUTH RIVERSIDE": "Inland Empire",
    "DESERT RIVERSIDE": "Inland Empire",
}


def read_legacy(path):
    """Read 'New Home Sales Trends' v4 layout -> (communities list, weekly rows)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    hdr_row = next(r for r in range(1, 10) if ws.cell(r, 1).value == "Community")
    date_row = hdr_row - 1
    week_cols = {c: ws.cell(date_row, c).value for c in range(1, ws.max_column + 1)
                 if isinstance(ws.cell(date_row, c).value, datetime)}
    group, comms, rows = None, [], []
    for r in range(hdr_row + 1, ws.max_row + 1):
        name, builder = ws.cell(r, 1).value, ws.cell(r, 2).value
        if name and not builder:
            group = str(name).strip()
            continue
        if not name:
            continue
        go = ws.cell(r, 4).value
        comms.append({"legacy_name": str(name).strip(), "builder": builder,
                      "city": ws.cell(r, 3).value, "legacy_group": group,
                      "grand_opening": go.date() if isinstance(go, datetime) else go,
                      "total_lots": ws.cell(r, 5).value, "row": r})
        for c, d in week_cols.items():
            v = ws.cell(r, c).value
            if v is None or (isinstance(v, str) and v.strip() in ("", "-")):
                continue
            rows.append({"legacy_row": r, "week_ending": d.date(),
                         "net_sales": v if isinstance(v, (int, float)) else None,
                         "legacy_text": v.strip() if isinstance(v, str) else None})
    return comms, rows


def import_legacy(db, path, source_file="New Home Sales Trends_v4.xlsx"):
    """One-time merge of the v4 net-sales history. Returns a match report."""
    comms, rows = read_legacy(path)
    c = db["communities"]
    taken = set()
    report = []
    by_row = {}
    for lc in comms:
        mkt = LEGACY_GROUP_MARKET.get((lc["legacy_group"] or "").upper())
        pool = c[(c["market"] == mkt) & ~c["community_id"].isin(taken)] if mkt else c[~c["community_id"].isin(taken)]
        ln = norm_name(lc["legacy_name"])
        cid, how = None, ""
        exact = pool[pool["community"].map(norm_name) == ln]
        if len(exact) == 1:
            cid, how = exact.iloc[0]["community_id"], "exact name"
        elif len(exact) > 1:
            b = exact[exact["builder"].map(lambda x: builder_match(x, lc["builder"]))]
            if len(b):
                cid, how = b.iloc[0]["community_id"], "exact name + builder"
        if cid is None:   # one name contains the other, same builder ("Isla at Luna" / "Isla at Luna Park")
            cands = pool[pool.apply(lambda p: builder_match(p["builder"], lc["builder"]) and (
                norm_name(p["community"]).startswith(ln + " ") or ln.startswith(norm_name(p["community"]) + " ")), axis=1)]
            if len(cands) == 1:
                cid, how = cands.iloc[0]["community_id"], "partial name + builder"
        if cid is None:   # same words, different order ("Regal at Canterwood" / "Canterwood Regal")
            lt = name_tokens(lc["legacy_name"])
            cands = pool[pool.apply(lambda p: builder_match(p["builder"], lc["builder"]) and
                                    name_tokens(p["community"]) == lt, axis=1)]
            if len(cands) == 1:
                cid, how = cands.iloc[0]["community_id"], "same words reordered + builder"
        if cid is None:
            best, score = None, 0
            for p in pool.itertuples():
                if not builder_match(p.builder, lc["builder"]):
                    continue
                s = name_similarity(lc["legacy_name"], p.community)
                if s > score:
                    best, score = p, s
            if best is not None and score >= 0.85:
                cid, how = best.community_id, f"fuzzy name {score:.2f} + builder"
        if cid is not None:
            taken.add(cid)
        report.append({"legacy_name": lc["legacy_name"], "legacy_builder": lc["builder"],
                       "legacy_group": lc["legacy_group"], "community_id": cid,
                       "matched_to": c.loc[c.community_id == cid, "community"].iloc[0] if cid else None,
                       "method": how or "no match"})
        lc["cid"] = cid
        lc["market"] = mkt
        by_row[lc["row"]] = lc

    # Infer section for unmatched: majority section of matched rows in same legacy group
    cidx = c.set_index("community_id")
    grp_sec = {}
    for lc in comms:
        if lc["cid"]:
            k = lc["legacy_group"]
            grp_sec.setdefault(k, []).append(tuple(cidx.loc[lc["cid"], ["market", "county_group", "submarket", "m_ord", "s_ord"]]))
    grp_sec = {k: max(set(v), key=v.count) for k, v in grp_sec.items()}

    # Unmatched communities that still had a number in a week covered by the PDF
    # are tracked from another source (e.g. Toll Brothers doesn't report to Ryness).
    pdf_weeks = set(db["weekly"].loc[db["weekly"]["source"] == SRC_PDF, "week_ending"])
    manual_rows = {lr["legacy_row"] for lr in rows
                   if lr["week_ending"] in pdf_weeks and lr["net_sales"] is not None}
    existing = set(c["community_id"])
    new_c = []
    for lc in comms:
        if lc["cid"]:
            cid = lc["cid"]
            i = c.index[c["community_id"] == cid][0]
            c.at[i, "legacy_name"], c.at[i, "legacy_group"] = lc["legacy_name"], lc["legacy_group"]
            for f in ("grand_opening", "total_lots"):
                if lc[f] is not None and (pd.isna(c.at[i, f]) or c.at[i, f] == ""):
                    c.at[i, f] = lc[f]
            continue
        sec = grp_sec.get(lc["legacy_group"], (lc["market"] or "Unassigned", lc["legacy_group"], lc["legacy_group"], 9, 999))
        cid = make_id(sec[0], lc["legacy_name"], lc["builder"], existing)
        existing.add(cid)
        lc["cid"] = cid
        new_c.append({"community_id": cid, "market": sec[0], "county_group": sec[1], "submarket": sec[2],
                      "community": lc["legacy_name"], "builder": lc["builder"], "city": lc["city"],
                      "units": lc["total_lots"], "total_lots": lc["total_lots"],
                      "grand_opening": lc["grand_opening"], "legacy_name": lc["legacy_name"],
                      "legacy_group": lc["legacy_group"], "m_ord": sec[3], "s_ord": sec[4],
                      "r_ord": 10000 + lc["row"],
                      "tracking": "Manual" if lc["row"] in manual_rows else "Ryness"})
    if new_c:
        c = pd.concat([c, pd.DataFrame(new_c).reindex(columns=COMM_COLS)], ignore_index=True)
    db["communities"] = c
    for rr in report:
        if rr["community_id"] is None:
            lc = next(x for x in comms if x["legacy_name"] == rr["legacy_name"] and x["legacy_group"] == rr["legacy_group"])
            rr["community_id"] = lc["cid"]
            rr["method"] = ("no match - not on Ryness, still reporting in v4 -> tracking = Manual"
                            if lc["row"] in manual_rows else
                            "no match - not on current Ryness report -> kept as SOLD OUT history")

    # Weekly rows; PDF weeks win (finalize drops legacy rows in PDF weeks) but keep a comparison
    w = db["weekly"]
    pdf = w[w["source"] == SRC_PDF].set_index(["community_id", "week_ending"])["net_sales"]
    compare = []
    new_w = []
    for lr in rows:
        lc = by_row[lr["legacy_row"]]
        cid = lc["cid"]
        key = (cid, lr["week_ending"])
        if key in pdf.index:
            compare.append({"community": lc["legacy_name"], "week_ending": lr["week_ending"],
                            "legacy_value": lr["net_sales"] if lr["net_sales"] is not None else lr["legacy_text"],
                            "pdf_net": pdf.loc[key],
                            "same": str(lr["net_sales"] if lr["net_sales"] is not None else lr["legacy_text"]) == str(pdf.loc[key])})
        new_w.append({"week_ending": lr["week_ending"], "community_id": cid, "market": lc["market"],
                      "net_sales": lr["net_sales"], "legacy_text": lr["legacy_text"],
                      "source": SRC_LEGACY, "source_file": source_file})
    w = w[w["source"] != SRC_LEGACY]
    db["weekly"] = pd.concat([w, pd.DataFrame(new_w, columns=WEEK_COLS)], ignore_index=True)
    finalize(db)
    db["import_log"] = pd.concat([db["import_log"], pd.DataFrame([{
        "imported_at": datetime.now().strftime("%Y-%m-%d %H:%M"), "source_file": source_file,
        "communities_parsed": len(comms),
        "warnings": f"Legacy net-only history: {sum(1 for r in report if not r['method'].startswith('no match'))} matched to Ryness, "
                    f"{sum(1 for r in report if 'Manual' in r['method'])} set to Manual tracking, "
                    f"{sum(1 for r in report if 'SOLD OUT' in r['method'])} kept as sold-out history"}])],
        ignore_index=True)
    return pd.DataFrame(report), pd.DataFrame(compare)


# --------------------------------------------------------------------------- save
_COL_TO_HEADER = {
    "community_id": "Community ID", "market": "Market", "county_group": "County / Group",
    "submarket": "Submarket", "community": "Community", "builder": "Builder", "city": "City",
    "product_type": "Type", "units": "Units (Ryness)", "sold_to_date": "Sold to Date",
    "remaining": "Remaining", "total_lots": "Total Lots", "grand_opening": "Grand Opening", "tracking": "Tracking",
    "status": "Status", "first_week": "First Week", "last_reported": "Last Reported",
    "weeks_off_report": "Weeks Off Report", "ryness_flag": "Ryness Flag",
    "legacy_name": "Legacy Name (v4)", "legacy_group": "Legacy Group (v4)",
    "m_ord": "m_ord", "s_ord": "s_ord", "r_ord": "r_ord", "last_pdf_week": "Last PDF Week",
    "user_notes": "My Notes", "alias": "Alias (name as printed)",
}
_HEADER_TO_COL = {v: k for k, v in _COL_TO_HEADER.items()}
_HEADER_TO_COL.update({"Week Ending": "week_ending", "Gross Sales": "gross_sales",
                       "Cancellations": "cancellations", "Note": "note"})


def save_db(db, path):
    from workbook_writer import write_workbook
    write_workbook(db, path, _COL_TO_HEADER)
