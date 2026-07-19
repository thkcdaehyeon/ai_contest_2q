---
name: create-welfare-sales-word-report
description: Create a polished Korean weekly welfare-equipment sales report in Word (.docx) from the linked Excel workbook produced by update-welfare-sales-workbook. Use when Codex needs a detailed manager-facing report covering weekly KPIs, activity records, product interest, follow-up actions, exceptions, and next-week priorities.
---

# Create Welfare Sales Word Report

Create a detailed, auditable Word report from the sales-management workbook. Use the workbook as the source of truth; do not recalculate numbers in prose.

## Workflow

1. Read `references/workbook-schema.md`.
2. Confirm the workbook contains `영업일지`. Prefer a workbook already updated by `$update-welfare-sales-workbook`.
3. Run:

```bash
python scripts/create_weekly_word_report.py \
  --workbook updated-sales.xlsx \
  --output weekly-sales-report.docx \
  --week-start 2026-07-13 \
  --week-end 2026-07-19
```

Omit the dates to use the Monday-to-Sunday week containing the latest activity.

4. Inspect the generated report when possible. Verify its date range, KPI values, activity count, and follow-up count against the workbook.
5. Report the output path and the JSON summary printed by the script.

For large workbooks, run the script directly and validate aggregate counts. Do not load office or activity rows into the model context; the script reads and aggregates them locally.

## Report structure

- Weekly headline and KPI table
- Product-interest and activity-stage summaries
- Major activity details by date
- Pending and overdue follow-up actions
- Review items requiring human confirmation

## Guardrails

- Describe product figures as 상담 건수 or 관심도, not confirmed sales.
- Do not invent executive commentary unsupported by the workbook.
- Preserve review items instead of smoothing over missing information.
- Write to a new file unless replacement is explicitly requested.

## Resources

- `scripts/create_weekly_word_report.py`: generate the styled DOCX report.
- `references/workbook-schema.md`: required workbook contract.
