# Welfare sales workbook contract

The workbook uses these exact sheet names.

| Sheet | Purpose |
|---|---|
| `사업소_목록` | One row per office with current stage, latest activity, next action, and cumulative activity count |
| `영업일지` | Append-only normalized activity records; source of truth for reports |
| `후속조치` | Pending actions derived from activities, with overdue status |
| `품목_관심도` | Product mention counts and percentages; never imply confirmed sales |
| `주간_현황` | Latest-week KPIs and summary tables/charts |
| `확인필요` | Missing fields, possible office-name matches, and skipped duplicates |

Downstream report Skills must read headers by name, not fixed column positions. Dates are stored as Excel dates where possible. Product lists are separated with `, `.
