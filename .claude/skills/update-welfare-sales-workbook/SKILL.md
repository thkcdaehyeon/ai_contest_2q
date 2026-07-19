---
name: update-welfare-sales-workbook
description: Convert batches of free-form or structured welfare-equipment sales logs into a validated multi-sheet Excel sales-management workbook. Use when Codex needs to create or update an .xlsx file containing office master data, sales activities, follow-up actions, product-interest analysis, weekly KPIs, and review items from daily or weekly sales notes.
---

# Update Welfare Sales Workbook

Turn many sales notes into one consistent workbook. Let the model interpret prose; let the bundled script perform deterministic matching, aggregation, formatting, and validation.

## Workflow

1. Read `references/input-schema.md` and `references/workbook-schema.md`.
2. Extract every sales activity into one JSON object. Never invent missing dates, office names, products, or next actions.
3. Save a JSON array using the input schema. Preserve the original note in `summary`.
4. Run:

```bash
python scripts/update_sales_workbook.py \
  --entries extracted-entries.json \
  --workbook existing-sales.xlsx \
  --output updated-sales.xlsx
```

Omit `--workbook` to create a new workbook. The script updates six linked sheets: `사업소_목록`, `영업일지`, `후속조치`, `품목_관심도`, `주간_현황`, and `확인필요`.

5. Read the JSON summary printed by the script. Report inserted, duplicate, and review counts.
6. Open or inspect the output workbook when possible. Confirm the expected sheets exist and the inserted count matches the accepted entries.

## Scale without wasting context

- Never paste or serialize the existing workbook into the model context. Let the Python script read all master and history rows locally.
- For more than 100 free-form activity notes or roughly 100,000 input characters, extract them in batches of 50–100 records and run the same deterministic update repeatedly.
- If activities are already structured as JSON, skip model interpretation and pass them directly to the script.
- Verify large runs from the printed counts and workbook sheet totals instead of reading every output row back into context.

## Guardrails

- Treat similar office names as review candidates; do not silently merge them.
- Skip exact duplicate activity IDs or duplicate composite records.
- Label product statistics as 상담 관심도, not sales share, unless actual order data is supplied.
- Keep real beneficiary, patient, resident, and personal contact information out of demo data.
- Write to a new output path unless the user explicitly asks to replace the original.

## Resources

- `scripts/update_sales_workbook.py`: create or update the linked workbook deterministically.
- `references/input-schema.md`: structured JSON contract and extraction rules.
- `references/workbook-schema.md`: workbook sheet contract shared with downstream report Skills.
