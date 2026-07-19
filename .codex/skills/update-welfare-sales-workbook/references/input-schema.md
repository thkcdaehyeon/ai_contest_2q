# Sales activity input schema

Pass a UTF-8 JSON array to `update_sales_workbook.py`.

```json
[
  {
    "entry_id": "2026-07-17-kim-001",
    "activity_date": "2026-07-17",
    "salesperson": "김영업",
    "office_name": "행복복지용구",
    "region": "서울 강남",
    "contact_type": "방문",
    "summary": "보행기와 안전손잡이 상담. 다음 주 견적 발송 예정.",
    "products": ["보행기", "안전손잡이"],
    "stage": "견적 예정",
    "next_action": "견적서 발송",
    "next_action_date": "2026-07-21"
  }
]
```

## Rules

- Required: `activity_date`, `salesperson`, `office_name`, `summary`.
- Use ISO dates (`YYYY-MM-DD`). Use an empty string when a date is not stated.
- Keep `products` as a JSON list. Use an empty list when none are stated.
- Use `contact_type` values such as 방문, 전화, 이메일, 온라인.
- Use short `stage` values such as 신규 접촉, 상담 중, 견적 예정, 검토 중, 계약 완료, 보류.
- Do not infer a missing office or next-action date from context that does not establish it.
- Make `entry_id` stable when source records provide their own IDs. Otherwise the script creates a deterministic ID.
