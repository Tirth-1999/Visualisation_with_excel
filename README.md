# Visualisation with Excel: A Week-by-Week Learning Story

This repository is not just a collection of Excel files. It is a learning journey where each workbook solves a specific analytics and reporting problem using formulas, PivotTables, charts, and VBA macros.

The core idea was simple: move from raw data to decision-ready dashboards, and automate repetitive work step by step.

## What We Were Trying to Solve

The project focused on these practical problems:

1. How do we convert raw retail data into clean, structured analysis?
2. How do we split one big dataset into meaningful business views (region/category/team)?
3. How do we create dashboards that are easy to present to stakeholders?
4. How do we reduce manual effort with macros and make reporting repeatable?
5. How do we write business-facing observations (growth, decline, contribution) from dashboard outputs?

## Learning Journey (Week 1 to Week 4)

### Week 1: First End-to-End Dashboard Build

Workbook: [Week1_Alphaa Superstore Nov 2020.xlsm](Week1_Alphaa%20Superstore%20Nov%202020.xlsm)

Problem solved:
- Build a complete pipeline from dataset to calculations to dashboard in one workbook.

How it was solved:
- `Alphaa-Superstore-Nov-2020` was used as the source dataset layer.
- `CALCULATIONS` became the transformation and metric layer.
- `DASHBOARD` became the final presentation layer.

Learning outcome:
- Understood dashboard architecture: `Data -> Compute -> Visualize`.

### Week 2-3: Multi-View Regional Analysis and Drilldowns

Workbook: [Week3_1.xlsm](Week3_1.xlsm)

Problem solved:
- Move from a single summary dashboard to region-specific analysis and focused product-level drilldowns.

How it was solved:
- Built separate regional computation sheets: `Central`, `East`, `West`, `South`.
- Added targeted sheets for high-interest combinations like `Central_Phones`, `East_Chairs`, `West_Copiers`, `South_Phones`.
- Consolidated outputs into `DashBoard` for a manager-level view.

Learning outcome:
- Learned segmentation strategy and how to design dashboards that support both overview and deep-dive questions.

### Week 4: Business Storytelling and Performance Communication

Reference brief: [Week_4_task.pdf](Week_4_task.pdf)

Problem solved:
- Convert numerical outputs into a stakeholder-ready narrative.

How it was solved:
- Compared sales periods and quantified increase/decrease percentages.
- Communicated contribution share (Team A, S, P) and growth rates in clear business language.
- Framed insights as recommendations and investigation points.

Learning outcome:
- Learned that dashboards are only half the job; executive communication and insight framing are equally important.

## Supporting Workbooks and Progression

### Foundations Template

Workbook: [Week0_macros.xltm](Week0_macros.xltm)

Purpose:
- Practice base macro patterns and visualization setups in `BASIC VISUALIZATION` and `ADVANCE VISUALIZATION` before applying them to project workbooks.

### Sharable and Advanced Versions

Workbooks:
- [WEEK-5_Sharable.xlsm](WEEK-5_Sharable.xlsm)
- [Superstore-Retail-Sep-2020_2_new.xlsm](Superstore-Retail-Sep-2020_2_new.xlsm)

Purpose:
- Package prior learnings into cleaner, presentation-ready files with stronger dashboard coverage and structured compute sheets.

## Manual Inspection Summary (What Exists Inside the Files)

I manually inspected all Excel macro files in this repository by reading workbook internals (sheets, chart parts, PivotTable parts, and VBA project presence).

### Confirmed VBA Presence

All workbook/template files contain `xl/vbaProject.bin`, meaning macro code exists in every one of them:

- [Week0_macros.xltm](Week0_macros.xltm)
- [Week1_Alphaa Superstore Nov 2020.xlsm](Week1_Alphaa%20Superstore%20Nov%202020.xlsm)
- [Week3_1.xlsm](Week3_1.xlsm)
- [WEEK-5_Sharable.xlsm](WEEK-5_Sharable.xlsm)
- [Superstore-Retail-Sep-2020_2_new.xlsm](Superstore-Retail-Sep-2020_2_new.xlsm)

### Structural Complexity by Workbook

| Workbook | Worksheets | Charts | PivotTables | Tables | VBA |
|---|---:|---:|---:|---:|---:|
| Week0_macros.xltm | 3 | 11 | 0 | 3 | Yes |
| Week1_Alphaa Superstore Nov 2020.xlsm | 3 | 14 | 19 | 1 | Yes |
| Week3_1.xlsm | 11 | 47 | 5 | 32 | Yes |
| WEEK-5_Sharable.xlsm | 8 | 22 | 11 | 2 | Yes |
| Superstore-Retail-Sep-2020_2_new.xlsm | 7 | 35 | 6 | 13 | Yes |

Interpretation:
- This confirms progression from foundational work to richer, multi-layer dashboard systems.
- The high counts of chart and pivot components indicate substantial visual and analytical build effort, not basic spreadsheet usage.

## Exact Outputs You Built Through This Learning Experience

1. Reusable macro-enabled workbook patterns
2. Region-wise analytical views and drilldown sheets
3. Dashboard pages for summary and comparison reporting
4. Period-over-period growth and distribution analysis
5. Business narrative reporting (not just technical output)
6. Sharable workbook versions suitable for presentation/demo

## How to Open and Review

1. Open any `.xlsm` or `.xltm` file in desktop Microsoft Excel.
2. Enable macros only for trusted local files.
3. Review sheets in this order: source data -> compute sheets -> dashboard.
4. Use slicers/filters/buttons (where available) to reproduce analysis outputs.

## Why This Repository Matters

This project demonstrates practical spreadsheet engineering and business analytics maturity:

- Strong fundamentals in Excel automation
- Clear transition from analysis to communication
- Progressive improvement over weekly milestones
- Real-world approach to reporting workflows

## Author

Tirth Shah
