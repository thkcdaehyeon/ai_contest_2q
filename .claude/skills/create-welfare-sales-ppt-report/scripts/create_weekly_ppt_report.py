#!/usr/bin/env python3
"""Generate a graph-driven weekly sales PPTX from the linked workbook."""

from __future__ import annotations

import argparse
import json
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


NAVY = RGBColor(63, 63, 63)
BLUE = RGBColor(255, 139, 0)
CYAN = RGBColor(166, 207, 0)
TEAL = RGBColor(126, 179, 0)
ORANGE = RGBColor(255, 139, 0)
RED = RGBColor(185, 28, 28)
PALE = RGBColor(255, 242, 224)
LIGHT = RGBColor(252, 250, 246)
MID = RGBColor(230, 222, 210)
GRAY = RGBColor(112, 112, 112)
WHITE = RGBColor(255, 255, 255)
FONT = "Noto Sans KR Medium"
DEFAULT_TEMPLATE = Path(__file__).resolve().parents[1] / "assets" / "irowoom-template.pptx"
BRAND_SERIES = [ORANGE, CYAN, RGBColor(166, 166, 166), RGBColor(255, 190, 102), TEAL]


def clean(value):
    return str(value or "").strip()


def parse_date(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            pass
    return None


def rows_as_dicts(ws):
    headers = [clean(c.value) for c in ws[1]]
    return [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True) if any(v not in (None, "") for v in row)]


def split_products(value):
    return [x.strip() for x in clean(value).replace(";", ",").split(",") if x.strip()]


def add_text(slide, text, x, y, w, h, size=20, color=NAVY, bold=False, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = tf.margin_right = Inches(0.02)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = FONT
    r_pr = run._r.get_or_add_rPr()
    for old in list(r_pr):
        if old.tag.endswith("}ea"):
            r_pr.remove(old)
    ea = OxmlElement("a:ea")
    ea.set("typeface", FONT)
    r_pr.append(ea)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def add_rect(slide, x, y, w, h, fill, radius=True, line=None):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line or fill
    return shape


def set_placeholder_text(slide, idx, text):
    placeholder = next((ph for ph in slide.placeholders if ph.placeholder_format.idx == idx), None)
    if placeholder is None:
        return None
    placeholder.text = text
    for paragraph in placeholder.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = FONT
            r_pr = run._r.get_or_add_rPr()
            ea = OxmlElement("a:ea")
            ea.set("typeface", FONT)
            r_pr.append(ea)
    return placeholder


def add_header(slide, section, title, subtitle=None):
    set_placeholder_text(slide, 0, f"{section}      /  {title}")
    set_placeholder_text(slide, 1, "")
    if subtitle:
        add_text(slide, subtitle, 0.48, 1.08, 12.1, 0.34, 9.5, GRAY)


def add_footer(slide, number):
    add_text(slide, "영업관리 Excel 기반 자동 생성", 0.48, 7.12, 4.5, 0.18, 7.5, GRAY)
    add_text(slide, str(number), 12.2, 7.1, 0.5, 0.2, 8, GRAY, True, PP_ALIGN.RIGHT)


def add_kpi(slide, x, y, w, label, value, accent=BLUE):
    add_rect(slide, x, y, w, 1.25, LIGHT, True, MID)
    add_rect(slide, x, y, 0.08, 1.25, accent, False)
    add_text(slide, label, x + 0.22, y + 0.18, w - 0.35, 0.25, 10, GRAY, True)
    add_text(slide, value, x + 0.22, y + 0.53, w - 0.35, 0.48, 23, accent, True)


def style_chart(chart, legend=False):
    chart.has_title = False
    chart.has_legend = legend
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.name = FONT
        chart.legend.font.size = Pt(9)
    chart.chart_style = 10
    try:
        chart.category_axis.tick_labels.font.name = FONT
        chart.category_axis.tick_labels.font.size = Pt(9)
        chart.value_axis.tick_labels.font.name = FONT
        chart.value_axis.tick_labels.font.size = Pt(8)
        chart.value_axis.has_major_gridlines = True
    except ValueError:
        pass
    for r_pr in chart._chartSpace.xpath(".//a:defRPr | .//a:endParaRPr"):
        for old in list(r_pr):
            if old.tag.endswith("}ea"):
                r_pr.remove(old)
        ea = OxmlElement("a:ea")
        ea.set("typeface", FONT)
        r_pr.append(ea)


def remove_sample_slides(prs):
    slide_ids = prs.slides._sldIdLst
    for slide_id in list(slide_ids):
        prs.part.drop_rel(slide_id.rId)
        slide_ids.remove(slide_id)


def layout_by_name(prs, name):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    raise SystemExit(f"Template layout not found: {name}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--week-start")
    parser.add_argument("--week-end")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--cover-style", choices=("1", "2"), default="1")
    args = parser.parse_args()

    wb = load_workbook(args.workbook, data_only=True)
    if "영업일지" not in wb.sheetnames:
        raise SystemExit("Workbook must contain 영업일지")
    logs = rows_as_dicts(wb["영업일지"])
    dated = [parse_date(r.get("활동일")) for r in logs if parse_date(r.get("활동일"))]
    if not dated:
        raise SystemExit("영업일지 contains no valid activity dates")
    if args.week_start:
        week_start = parse_date(args.week_start)
        if not week_start:
            raise SystemExit("Invalid --week-start")
    else:
        latest = max(dated)
        week_start = latest - timedelta(days=latest.weekday())
    if args.week_end:
        week_end = parse_date(args.week_end)
        if not week_end:
            raise SystemExit("Invalid --week-end")
    else:
        week_end = week_start + timedelta(days=6)

    week_logs = [r for r in logs if parse_date(r.get("활동일")) and week_start <= parse_date(r.get("활동일")) <= week_end]
    ids = {clean(r.get("활동ID")) for r in week_logs}
    follows = rows_as_dicts(wb["후속조치"]) if "후속조치" in wb.sheetnames else []
    follows = [r for r in follows if clean(r.get("활동ID")) in ids and clean(r.get("상태")) != "완료"]
    reviews = rows_as_dicts(wb["확인필요"]) if "확인필요" in wb.sheetnames else []
    reviews = [r for r in reviews if not clean(r.get("활동ID")) or clean(r.get("활동ID")) in ids]

    product_counts = Counter(p for r in week_logs for p in split_products(r.get("관심품목")))
    stage_counts = Counter(clean(r.get("진행단계")) or "미분류" for r in week_logs)
    person_counts = Counter(clean(r.get("담당자")) or "미지정" for r in week_logs)
    daily_counts = Counter(parse_date(r.get("활동일")) for r in week_logs)
    unique_offices = len({clean(r.get("사업소ID")) or clean(r.get("사업소명")) for r in week_logs})
    overdue = sum(1 for r in follows if clean(r.get("상태")) == "지연")

    if not args.template.exists():
        raise SystemExit(f"Template not found: {args.template}")
    prs = Presentation(args.template)
    remove_sample_slides(prs)
    title_layout = layout_by_name(prs, f"제목 슬라이드_0{args.cover_style}")
    content_layout = layout_by_name(prs, "내용")

    # 1. Title
    slide = prs.slides.add_slide(title_layout)
    set_placeholder_text(slide, 0, "주간 영업 활동 보고")
    set_placeholder_text(
        slide, 1,
        f"{week_start.isoformat()}  —  {week_end.isoformat()}\n영업일지 {len(week_logs)}건 · 사업소 {unique_offices}곳"
    )

    # 2. KPI overview
    slide = prs.slides.add_slide(content_layout)
    add_header(slide, "01", "주간 영업 현황", "같은 영업일지를 Excel·Word·PPT에서 다시 입력하지 않습니다.")
    add_kpi(slide, 0.48, 1.55, 2.95, "영업활동", f"{len(week_logs)}건", BLUE)
    add_kpi(slide, 3.62, 1.55, 2.95, "접촉 사업소", f"{unique_offices}곳", CYAN)
    add_kpi(slide, 6.76, 1.55, 2.95, "후속조치", f"{len(follows)}건", ORANGE)
    add_kpi(slide, 9.9, 1.55, 2.95, "확인필요", f"{len(reviews)}건", RED if reviews else TEAL)

    chart_data = ChartData()
    days = [week_start + timedelta(days=i) for i in range(7)]
    chart_data.categories = [d.strftime("%m/%d") for d in days]
    chart_data.add_series("활동 건수", [daily_counts.get(d, 0) for d in days])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.52), Inches(3.15), Inches(7.15), Inches(3.35), chart_data).chart
    style_chart(chart)
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = BLUE
    add_text(slide, "일별 활동", 0.52, 2.9, 3, 0.3, 12, NAVY, True)

    add_rect(slide, 8.02, 3.15, 4.8, 3.35, LIGHT, True, MID)
    add_text(slide, "이번 주 포인트", 8.34, 3.46, 4.0, 0.3, 13, NAVY, True)
    top_product = product_counts.most_common(1)[0] if product_counts else ("기록 없음", 0)
    top_stage = stage_counts.most_common(1)[0] if stage_counts else ("미분류", 0)
    bullets = [
        f"관심 품목 1위  {top_product[0]} · {top_product[1]}건",
        f"주요 진행 단계  {top_stage[0]} · {top_stage[1]}건",
        f"기한 경과 후속 조치  {overdue}건",
    ]
    for i, item in enumerate(bullets):
        add_rect(slide, 8.34, 4.08 + i * 0.62, 0.13, 0.13, [BLUE, CYAN, ORANGE][i], False)
        add_text(slide, item, 8.62, 3.96 + i * 0.62, 3.65, 0.4, 10.5, NAVY, i == 0)
    add_footer(slide, 2)

    # 3. Product interest
    slide = prs.slides.add_slide(content_layout)
    add_header(slide, "02", "품목 관심도", "영업일지의 관심 품목 언급을 집계한 값이며 실제 판매 비율이 아닙니다.")
    products = product_counts.most_common(7) or [("기록 없음", 0)]
    chart_data = ChartData()
    chart_data.categories = [p for p, _ in reversed(products)]
    chart_data.add_series("상담 건수", [c for _, c in reversed(products)])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.5), Inches(1.52), Inches(8.25), Inches(5.15), chart_data).chart
    style_chart(chart)
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = CYAN
    total_mentions = sum(product_counts.values())
    add_rect(slide, 9.05, 1.65, 3.75, 4.85, LIGHT, True, MID)
    add_text(slide, "상담 관심 품목", 9.38, 1.97, 3.05, 0.35, 13, NAVY, True)
    for i, (product, count) in enumerate(products[:5]):
        pct = count / total_mentions * 100 if total_mentions else 0
        add_text(slide, f"{i + 1}", 9.38, 2.68 + i * 0.64, 0.35, 0.3, 10, CYAN, True)
        add_text(slide, product, 9.8, 2.65 + i * 0.64, 1.55, 0.35, 10.5, NAVY, i == 0)
        add_text(slide, f"{count}건 · {pct:.0f}%", 11.38, 2.65 + i * 0.64, 1.02, 0.35, 9.5, GRAY, False, PP_ALIGN.RIGHT)
    add_footer(slide, 3)

    # 4. Stage and salesperson charts
    slide = prs.slides.add_slide(content_layout)
    add_header(slide, "03", "활동 구성", "진행 단계와 담당자별 활동을 함께 봅니다.")
    stages = stage_counts.most_common() or [("미분류", 0)]
    chart_data = ChartData()
    chart_data.categories = [x for x, _ in stages]
    chart_data.add_series("활동", [x for _, x in stages])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(0.52), Inches(1.55), Inches(5.95), Inches(5.05), chart_data).chart
    style_chart(chart, True)
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = BLUE
    for point, color in zip(chart.series[0].points, BRAND_SERIES):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = color
    add_text(slide, "진행 단계", 0.52, 1.35, 2.5, 0.3, 12, NAVY, True)

    people = person_counts.most_common() or [("미지정", 0)]
    chart_data2 = ChartData()
    chart_data2.categories = [x for x, _ in people]
    chart_data2.add_series("활동 건수", [x for _, x in people])
    chart2 = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6.72), Inches(1.62), Inches(5.95), Inches(4.98), chart_data2).chart
    style_chart(chart2)
    chart2.series[0].format.fill.solid()
    chart2.series[0].format.fill.fore_color.rgb = ORANGE
    add_text(slide, "담당자별 활동", 6.72, 1.35, 3, 0.3, 12, NAVY, True)
    add_footer(slide, 4)

    # 5. Follow-up and review
    slide = prs.slides.add_slide(content_layout)
    add_header(slide, "04", "다음 행동", "기한이 지났거나 정보가 빠진 항목은 자동으로 숨기지 않고 표시합니다.")
    add_text(slide, "후속 조치", 0.52, 1.48, 5, 0.3, 13, NAVY, True)
    add_text(slide, "사람이 확인할 항목", 6.94, 1.48, 5, 0.3, 13, NAVY, True)
    add_rect(slide, 0.5, 1.88, 6.05, 4.75, LIGHT, True, MID)
    add_rect(slide, 6.92, 1.88, 5.9, 4.75, LIGHT, True, MID)

    sorted_follows = sorted(follows, key=lambda r: (clean(r.get("상태")) != "지연", parse_date(r.get("예정일")) or date.max))[:6]
    if not sorted_follows:
        add_text(slide, "미완료 후속 조치가 없습니다.", 1.05, 3.0, 5.1, 0.5, 13, GRAY)
    for i, item in enumerate(sorted_follows):
        y = 2.28 + i * 0.62
        status = clean(item.get("상태")) or "예정"
        color = RED if status == "지연" else ORANGE
        due = parse_date(item.get("예정일"))
        add_rect(slide, 0.82, y + 0.05, 0.48, 0.25, color, True)
        add_text(slide, status, 0.82, y + 0.07, 0.48, 0.18, 7, WHITE, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        add_text(slide, clean(item.get("사업소명")), 1.48, y, 1.55, 0.3, 9.5, NAVY, True)
        add_text(slide, clean(item.get("다음행동")), 3.05, y, 2.1, 0.3, 9.5, NAVY)
        add_text(slide, due.strftime("%m/%d") if due else "일정 확인", 5.27, y, 0.88, 0.3, 8.5, GRAY, False, PP_ALIGN.RIGHT)

    if not reviews:
        add_text(slide, "확인할 데이터 예외가 없습니다.", 7.28, 2.62, 4.85, 0.5, 13, GRAY)
    for i, item in enumerate(reviews[:6]):
        y = 2.27 + i * 0.64
        add_rect(slide, 7.28, y + 0.04, 0.12, 0.12, RED, False)
        add_text(slide, clean(item.get("유형")), 7.58, y - 0.02, 1.08, 0.28, 9.5, RED, True)
        add_text(slide, clean(item.get("사업소명")) or "공통", 8.64, y - 0.02, 1.45, 0.28, 9.5, NAVY, True)
        add_text(slide, clean(item.get("내용")), 10.08, y - 0.02, 2.22, 0.4, 8.5, NAVY)

    add_text(slide, f"영업일지 {len(week_logs)}건 → Excel 관리대장 · Word 상세보고 · PPT 주간보고", 0.72, 6.75, 11.7, 0.3, 11, BLUE, True, PP_ALIGN.CENTER)
    add_footer(slide, 5)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(args.output)
    print(json.dumps({
        "output": str(args.output.resolve()), "slides": len(prs.slides),
        "week": f"{week_start.isoformat()}~{week_end.isoformat()}", "activities": len(week_logs),
        "offices": unique_offices, "followups": len(follows), "review_items": len(reviews),
        "template": str(args.template.resolve()),
        "cover_style": args.cover_style,
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
