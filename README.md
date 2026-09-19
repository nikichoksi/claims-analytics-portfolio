# Claims Analytics

A synthetic school-district workers' compensation claims analysis, built end-to-end:
messy raw data → a documented Power Query cleaning pipeline → an Excel BI dashboard →
a findings memo. All data is synthetic (Mountain Vista School District is a fictional
construct); no real students, employees, or incidents are represented.

## How to review (open in this order)

1. **raw_claims_extract.xlsx**
   The "before" - a 400-row, intentionally messy RMIS-style export (`RMIS_Export`
   sheet): duplicate claim IDs, missing IDs, three different date formats,
   uncontrolled vocabulary (school names, departments, injury body parts/causes,
   claim status), and dirty incurred-loss values (currency text, negatives, blanks,
   extreme outliers).

   This workbook also holds the **live Power Query pipeline** itself, not just the
   raw data: open it in Excel and check the Queries pane to see `Claims_Raw` →
   `Claims_Prepared` → `Claims_Clean` / `Exceptions`, built on 5 controlled-vocabulary
   mapping tables (`Map_School`, `Map_Department`, `Map_BodyPart`, `Map_Cause`,
   `Map_Status`). The `Claims_Clean` (375 rows) and `Exceptions` (10 rows) sheets are
   the pipeline's actual output.

2. **claims_cleaning_log.xlsx**
   The audit trail - a 13-step cleaning log (Action / Method / Rows Affected / Rule
   Applied / Rationale) documenting every transformation and *why* it was made, plus
   a data dictionary for the final `Claims_Clean` schema and a naming-convention note
   (PascalCase).

3. **claims_dashboard.xlsx**
   The "after" - built from the real `Claims_Clean` output, 4 sheets:
   - *Executive_KPIs*: total claims, open/closed/re-opened, total & average incurred
     loss, a pending-loss-estimate count, quarter-over-quarter comparison, a
     data-quality reconciliation strip (400 raw → 375 clean, with every step's
     count), and a real **PivotTable + 3 connected slicers** (School Level,
     Department, Claim Status) on the `Pivot_Base` sheet - the KPI cards themselves
     are a full-district snapshot; the slicers filter the interactive drill-down
     view, not the cards.
   - *Trends*: monthly claim volume, top injury causes, loss by body part, a month ×
     school-level heatmap, and a Pareto chart showing which sites concentrate the
     most incurred loss. Chart titles state findings directly (e.g., "Slip/Fall Is
     the #1 Injury Cause"), and every chart carries data labels.
   - *Drilldown*: Site → Department → Cause with incurred loss, top-10%-loss cells
     highlighted; claims with no loss estimate yet show "Pending" (with a count),
     not a misleading $0.
   - *Data*: the full clean dataset as an Excel Table, the source for the PivotTable
     and any further analysis.

4. **findings_memo.pdf**
   A one-page memo addressed to a Director of Risk Management: summary, key findings
   (e.g., slip/fall claims are 31% of volume but 42% of incurred loss, concentrated in
   winter at elementary sites), an exhibit chart of claims by month with winter
   months highlighted, recommended actions, data-quality notes, and a
   methodology/privacy footer referencing CORA.

## Supplementary reference (optional, for a technical reviewer)

- **power_query_M_code.txt** - the complete M code for the 4-query pipeline, with
  setup instructions, for anyone who wants to inspect the transformation logic
  directly without opening Excel.
- **dashboard_pivot_slicer_guide.txt** - the manual, click-by-click steps for adding
  a live, slicer-driven PivotTable on top of the dashboard's `Data` sheet.
- **build_live_dashboard.bas** - the VBA macro actually used to build the PivotTable
  and slicers already in `claims_dashboard.xlsx` (Excel's own PivotTable/Slicer
  objects can't be created by any file-writing library, so this was the automated
  path instead of the fully manual one above). To rebuild it from scratch: open
  `claims_dashboard.xlsx` in Excel, Tools → Macro → Visual Basic Editor (Option+F11),
  Insert → Module, paste this file's contents, then F5 to run `BuildLiveDashboard`.

## Notes

- Reporting period: school year August 2024 – July 2025.
- All monetary figures, dates, and names are synthetic and randomly generated.
- No PII is present anywhere in this portfolio.
