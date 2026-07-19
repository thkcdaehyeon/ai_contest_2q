---
name: create-welfare-sales-ppt-report
description: Create a concise graph-driven Korean PowerPoint (.pptx) weekly welfare-equipment sales report from the linked Excel workbook produced by update-welfare-sales-workbook, using the retained Irowoom corporate template. Use when Codex needs Irowoom-branded meeting-ready slides with weekly KPIs, product-interest charts, stage and salesperson activity charts, follow-up priorities, and data-quality exceptions.
---

# Create Welfare Sales PPT Report

Turn the Excel source of truth into a short meeting deck. Favor readable charts and decisions over dense activity tables.

## Workflow

1. Read `references/workbook-schema.md`.
2. Confirm the workbook contains `영업일지`. Prefer a workbook already updated by `$update-welfare-sales-workbook`.
3. Run:

```bash
python scripts/create_weekly_ppt_report.py \
  --workbook updated-sales.xlsx \
  --output weekly-sales-report.pptx \
  --week-start 2026-07-13 \
  --week-end 2026-07-19
```

Omit the dates to use the Monday-to-Sunday week containing the latest activity.
The script uses `assets/irowoom-template.pptx` by default. Pass `--template` only when the user explicitly provides a replacement template.
The default cover is source page 1. Pass `--cover-style 2` to use source page 2 instead.

4. Inspect the generated deck when possible. Verify the chart labels and values against the workbook.
5. Report the output path, slide count, and the JSON summary printed by the script.

For large workbooks, run the script directly and validate aggregate counts. Do not load office or activity rows into the model context; the script reads and aggregates them locally.

## Deck structure

1. Title and reporting period
2. Weekly KPI overview
3. Product-interest chart
4. Sales-stage and salesperson activity charts
5. Follow-up priorities and review items

## Guardrails

- Keep every slide focused on one decision or comparison.
- Label product figures as 상담 관심도, not sales share.
- Show missing or ambiguous records as review items.
- Write to a new file unless replacement is explicitly requested.

## Resources

- `scripts/create_weekly_ppt_report.py`: generate the styled PPTX deck with native charts.
- `references/workbook-schema.md`: required workbook contract.
- `references/template-layouts.md`: intended use of each retained template layout.
- `assets/irowoom-template.pptx`: retained Irowoom master, theme, logo, backgrounds, and slide layouts.
