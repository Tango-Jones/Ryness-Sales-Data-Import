"""
ryness_parser.py
Reads a weekly Ryness Report PDF and returns one record per community.

Each record: week_ending, market, county_group, submarket, community, builder,
city_code, city, notes, status_flag (S/O, TSO, or ''), product_type, units,
gross_sales (Wk's Sales), cancellations (Wk's Cans), net_sales, source_page.

The parser reads word positions (not plain text), so a red "S/O" next to a
builder, a "New" note, or a blank column never shifts the numbers. Every
section is checked against the PDF's own TOTALS row (projects reporting,
sales, cans); mismatches are returned as warnings.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, asdict
from datetime import datetime, date

import pdfplumber

# Numeric columns in left-to-right order, keyed by the header word that sits
# lowest in each column's (multi-line) header label.
NUM_COLS = ["units", "new_rel", "remaining", "traffic", "gross_sales",
            "cancellations", "sold_to_date", "sold_ytd", "avg_week", "avg_ytd"]
HEADER_KEYS = {  # header word -> column (bottom line of each column label)
    "Units": "units", "Rel.": "new_rel", "Rm'g": "remaining", "Traffic": "traffic",
    "Sales": "gross_sales", "Cans": "cancellations", "Date": "sold_to_date",
    "YTD": "sold_ytd", "/Week": "avg_week", "/YTD": "avg_ytd",
}
# Left-side text columns are found from the "Development Name / Developer /
# City Code / Notes / Type" header on each page.
STATUS_TOKENS = {"S/O": "S/O", "TSO": "TSO", "SO": "S/O"}
DATE_RE = re.compile(r"(January|February|March|April|May|June|July|August|"
                     r"September|October|November|December)\s+(\d{1,2}),\s+(\d{4})")


@dataclass
class Record:
    week_ending: date
    market: str
    county_group: str
    submarket: str
    community: str
    builder: str
    city_code: str
    city: str = ""
    notes: str = ""
    status_flag: str = ""
    product_type: str = ""
    units: int | None = None
    gross_sales: int = 0
    cancellations: int = 0
    net_sales: int = 0
    sold_to_date: int | None = None
    source_page: int = 0


@dataclass
class ParseResult:
    week_ending: date | None
    records: list[Record] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    section_checks: list[dict] = field(default_factory=list)

    def to_frame(self):
        import pandas as pd
        return pd.DataFrame([asdict(r) for r in self.records])


def _to_int(tok: str | None):
    if tok is None:
        return None
    tok = tok.replace(",", "").strip()
    if re.fullmatch(r"-?\d+", tok):
        return int(tok)
    return None


def _lines(words, tol=2.5):
    """Group words into visual lines by their top coordinate."""
    words = sorted(words, key=lambda w: (round(w["top"]), w["x0"]))
    lines: list[list[dict]] = []
    for w in words:
        if lines and abs(lines[-1][0]["top"] - w["top"]) <= tol:
            lines[-1].append(w)
        else:
            lines.append([w])
    return [sorted(l, key=lambda w: w["x0"]) for l in lines]


def _is_red(w) -> bool:
    c = w.get("non_stroking_color")
    if isinstance(c, (list, tuple)) and len(c) == 3:
        r, g, b = c
        return r > 0.6 and g < 0.4 and b < 0.4
    return False


def parse_pdf(path) -> ParseResult:
    result = ParseResult(week_ending=None)
    with pdfplumber.open(path) as pdf:
        market = ""
        county_group = submarket = ""
        section = None          # current section accumulator for checks
        page_maps: dict[int, dict[str, str]] = {}
        for pno, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(extra_attrs=["non_stroking_color", "size"])
            text = page.extract_text() or ""

            # Week-ending date (first one found wins; all pages should agree)
            m = DATE_RE.search(text)
            if m:
                d = datetime.strptime(" ".join(m.groups()), "%B %d %Y").date()
                if result.week_ending is None:
                    result.week_ending = d
                elif d != result.week_ending:
                    result.warnings.append(
                        f"Page {pno}: week ending {d} differs from {result.week_ending}")

            if "Development Name" not in text:
                # Market summary page (e.g. 'LA-Orange-North' cover). Grab market name.
                for line in text.split("\n")[:6]:
                    s = line.strip()
                    if s and "RYNESS" not in s.upper() and "Research Company" not in s \
                            and not s.startswith("Week"):
                        market = s
                        break
                continue

            # Market name sits top-right on detail pages, left of "Page"
            top_words = [w for w in words if w["top"] < 35 and 380 < w["x0"] < 560]
            if top_words:
                market = " ".join(w["text"] for w in sorted(top_words, key=lambda w: w["x0"]))

            lines = _merge_wrapped(_lines(words))

            # Left text-column boundaries from the "Development Name ... Type" header
            hdr = next((l for l in lines if l and l[0]["text"] == "Development"), None)
            if hdr is None:
                result.warnings.append(f"Page {pno}: header row not found; skipped")
                continue
            hx = {w["text"]: w["x0"] for w in hdr}
            x_dev, x_city, x_notes, x_type = (hx.get("Developer", 158), hx.get("City", 237),
                                              hx.get("Notes", 288) - 6, hx.get("Type", 334) - 6)

            num_centers: dict[str, float] = {}
            city_map: dict[str, str] = {}
            page_maps[pno] = city_map
            pending_name_prefix = ""

            for line in lines:
                first = line[0]["text"]
                joined = " ".join(w["text"] for w in line)

                # Section header, e.g. "Central-North Orange | Central Orange (Continued ...)"
                if "|" in joined and line[0]["x0"] < 60 and line[0]["x0"] > 28:
                    left = " ".join(w["text"] for w in line if w["x0"] < 380)
                    left = left.replace("(Continued ...)", "").strip()
                    cg, _, sm = left.partition("|")
                    cg, sm = cg.strip(), sm.replace("Sold Av. Av.", "").strip()
                    cg = re.sub(r"\s+Sold.*$", "", cg)
                    sm = re.sub(r"\s+Sold.*$", "", sm)
                    if section is None or (cg, sm) != (section["county_group"], section["submarket"]):
                        section = {"market": market, "county_group": cg, "submarket": sm,
                                   "rows": 0, "sales": 0, "cans": 0, "page": pno}
                    county_group, submarket = cg, sm
                    continue

                # Numeric column header line(s): record x-centers
                if any(w["text"] in HEADER_KEYS for w in line) and line[0]["x0"] > 370:
                    for w in line:
                        col = HEADER_KEYS.get(w["text"])
                        if col and not (col == "gross_sales" and w["x0"] < 440):
                            num_centers[col] = (w["x0"] + w["x1"]) / 2
                    continue

                if first in ("TOTALS:", "GRAND"):
                    if first == "TOTALS:" and section is not None:
                        m = re.search(r"No\. Reporting:\s*(\d+)", joined)
                        nums = [w for w in line if w["x0"] > 370 and _to_int(w["text"]) is not None]
                        tot = _assign(nums, num_centers)
                        chk = {"market": section["market"], "county_group": section["county_group"],
                               "submarket": section["submarket"], "page": pno,
                               "reporting_pdf": int(m.group(1)) if m else None,
                               "reporting_parsed": section["rows"],
                               "sales_pdf": tot.get("gross_sales"), "sales_parsed": section["sales"],
                               "cans_pdf": tot.get("cancellations"), "cans_parsed": section["cans"]}
                        chk["ok"] = (chk["reporting_pdf"] == chk["reporting_parsed"]
                                     and chk["sales_pdf"] == chk["sales_parsed"]
                                     and chk["cans_pdf"] == chk["cans_parsed"])
                        result.section_checks.append(chk)
                        if not chk["ok"]:
                            result.warnings.append(
                                f"Totals mismatch: {chk['county_group']} | {chk['submarket']} (p{pno}) "
                                f"reporting {chk['reporting_parsed']}/{chk['reporting_pdf']}, "
                                f"sales {chk['sales_parsed']}/{chk['sales_pdf']}, "
                                f"cans {chk['cans_parsed']}/{chk['cans_pdf']}")
                    continue

                if first == "City" and len(line) > 1 and line[1]["text"] == "Codes:":
                    body = joined.split("Codes:", 1)[1]
                    for part in body.split(","):
                        if "=" in part:
                            k, v = part.split("=", 1)
                            city_map[k.strip()] = v.strip()
                    continue

                # Candidate data row: numbers at right (name may have wrapped)
                if not num_centers:
                    continue
                left = sorted([w for w in line if w["x1"] <= x_dev - 2],
                              key=lambda w: (round(w["top"]), w["x0"]))
                nums = [w for w in line if w["x0"] > 370]
                ints = [w for w in nums if _to_int(w["text"]) is not None]
                if not left:
                    continue
                if len(ints) < 4:
                    # Possibly a wrapped community name with no numbers
                    if not nums and section is not None:
                        pending_name_prefix = (pending_name_prefix + " " +
                                               " ".join(w["text"] for w in left)).strip()
                    continue

                name = " ".join(w["text"] for w in left)
                if pending_name_prefix:
                    name = f"{pending_name_prefix} {name}"
                    pending_name_prefix = ""

                dev_words = [w for w in line if x_dev - 2 <= w["x0"] < x_city - 2]
                status = ""
                builder_parts = []
                for w in dev_words:
                    if w["text"] in STATUS_TOKENS and (_is_red(w) or w["text"] in ("S/O", "TSO")):
                        status = STATUS_TOKENS[w["text"]]
                    else:
                        builder_parts.append(w["text"])
                city_code = " ".join(w["text"] for w in line if x_city - 2 <= w["x0"] < x_notes - 2)
                notes = " ".join(w["text"] for w in line if x_notes - 2 <= w["x0"] < x_type - 2)
                ptype = " ".join(w["text"] for w in line if x_type - 2 <= w["x0"] < 375)
                vals = _assign(nums, num_centers, raw=True)
                if not status:
                    rem = vals.get("remaining")
                    if isinstance(rem, str) and rem in STATUS_TOKENS:
                        status = STATUS_TOKENS[rem]

                gross = _to_int(vals.get("gross_sales")) or 0
                cans = _to_int(vals.get("cancellations")) or 0
                rec = Record(
                    week_ending=result.week_ending, market=market,
                    county_group=county_group, submarket=submarket,
                    community=name.strip(), builder=" ".join(builder_parts).strip(),
                    city_code=city_code.strip(), notes=notes.strip(), status_flag=status,
                    product_type=ptype.strip(), units=_to_int(vals.get("units")),
                    gross_sales=gross, cancellations=cans, net_sales=gross - cans,
                    sold_to_date=_to_int(vals.get("sold_to_date")), source_page=pno)
                result.records.append(rec)
                if section is not None:
                    section["rows"] += 1
                    section["sales"] += gross
                    section["cans"] += cans


        # Resolve city codes: same page first, then the following pages (a
        # section continued across pages prints its legend at the end).
        for r in result.records:
            for pg in [r.source_page] + sorted(p for p in page_maps if p > r.source_page) \
                    + sorted((p for p in page_maps if p < r.source_page), reverse=True):
                if r.city_code in page_maps[pg]:
                    r.city = page_maps[pg][r.city_code]
                    break
            else:
                r.city = r.city_code
                result.warnings.append(f"City code '{r.city_code}' not in any legend ({r.community})")

    # Back-fill week date for records parsed before date seen (shouldn't happen)
    for r in result.records:
        r.week_ending = r.week_ending or result.week_ending
    if result.week_ending is None:
        result.warnings.append("Week-ending date not found in PDF")
    return result


def _merge_wrapped(lines):
    """Attach text-only fragments of a wrapped community name to the numeric
    row they belong to (the numbers sit vertically centred between them)."""
    def has_nums(l):
        return sum(1 for w in l if w["x0"] > 370 and _to_int(w["text"]) is not None) >= 4
    out = [list(l) for l in lines]
    data_idx = [i for i, l in enumerate(out) if has_nums(l)]
    drop = set()
    for i, l in enumerate(out):
        if i in data_idx or not l or l[0]["x0"] > 30:
            continue
        txt = " ".join(w["text"] for w in l)
        if txt.startswith(("City Codes", "Development", "TOTALS", "GRAND")) or "|" in txt:
            continue
        if any(w["x0"] > 370 for w in l):
            continue
        near = [j for j in data_idx if abs(out[j][0]["top"] - l[0]["top"]) <= 10
                and out[j][0]["x0"] > 30]       # numeric line missing its name
        if near:
            j = min(near, key=lambda j: abs(out[j][0]["top"] - l[0]["top"]))
            out[j].extend(l)
            drop.add(i)
    return [sorted(l, key=lambda w: w["x0"]) for i, l in enumerate(out) if i not in drop]


def _assign(num_words, centers: dict[str, float], raw=False):
    """Map each numeric word to the nearest column center."""
    out = {}
    if not centers:
        return out
    for w in num_words:
        cx = (w["x0"] + w["x1"]) / 2
        col = min(centers, key=lambda c: abs(centers[c] - cx))
        if abs(centers[col] - cx) > 14:
            continue
        out[col] = w["text"] if raw else _to_int(w["text"])
    return out


if __name__ == "__main__":
    import sys
    res = parse_pdf(sys.argv[1])
    df = res.to_frame()
    print("Week ending:", res.week_ending, "| communities:", len(df))
    print(df.groupby(["market"])[["gross_sales", "cancellations", "net_sales"]].sum())
    bad = [c for c in res.section_checks if not c["ok"]]
    print(f"Section checks: {len(res.section_checks)} sections, {len(bad)} mismatches")
    for w in res.warnings:
        print("WARN:", w)
