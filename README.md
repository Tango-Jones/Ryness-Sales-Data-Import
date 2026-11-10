# Ryness Weekly Import

Reads the weekly SoCal Ryness Report PDF and updates `New_Home_Sales_Database.xlsx`:
**Gross Sales** (Wk's Sales), **Cancellations** (Wk's Cans) and **Net Sales** (Gross - Cancellations, live formulas).

## Setup (once)
1. Install Python 3.10+ (python.org).
2. Keep this folder together; the database workbook sits next to `app.py`.
3. Double-click **Run Ryness Import (Windows).bat** or **Run Ryness Import (Mac).command**
   (or run `pip install -r requirements.txt` then `streamlit run app.py`). A browser tab opens.

## Every week
1. In the app, the database path defaults to `New_Home_Sales_Database.xlsx` in this folder.
2. Drop in the new Ryness PDF. Review: totals check (every section vs the PDF's TOTALS row), market totals, and the change list.
3. Click **Import into database**. The workbook is saved in place; a download copy is also offered.
   Close the workbook in Excel before importing so it can be saved.

## What gets flagged
| Event | Meaning |
|---|---|
| NEW | First appearance on the Ryness report (also when Ryness notes "New") |
| ADDED | First week tracked here, in a submarket with no earlier data |
| DROPPED | On last week's report, missing this week - likely sold out / temporarily sold out. Shows SOLD OUT from then on; never deleted |
| RETURNED | Back on the report after being absent |
| FLAG S/O, FLAG TSO | Still listed, but Ryness marks it Sold Out / Temporarily Sold Out |
| POSSIBLE RENAME | A new and a dropped project with the same builder and similar names. Add a Name_Map row and re-import |

## Your editable parts (kept on every import)
- **Communities** tab, yellow columns: Total Lots, Grand Opening, Tracking (Ryness / Manual), My Notes.
- **Name_Map**: alias -> existing Community ID.
- **Manual_Entries**: weekly gross/cans for communities not on Ryness (Tracking = Manual, e.g. Toll Brothers at Great Park).

The three weekly grids are regenerated on each import, so don't type into them.

## Online database later
`Weekly_Data` (one row per community per week) and `Communities` are already table-shaped; the app's Database tab exports both as CSV. `sales_db.py` is the only module that touches storage - swap `load_db`/`save_db` for database calls.

## Files
- `ryness_parser.py` - PDF -> records (word-position based; validates every section's totals)
- `sales_db.py` - matching, new/dropped/returned logic, legacy import
- `workbook_writer.py` - builds the Excel tabs
- `app.py` - Streamlit front end
- `build_initial.py` - one-time rebuild from a PDF + the old v4 trends workbook
