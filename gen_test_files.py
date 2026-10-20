#!/usr/bin/env python3
"""生成 5 个测试文件到 test_files/ 目录，供手动验证各功能模块使用。"""

from __future__ import annotations

import io
import os
from pathlib import Path

from faker import Faker
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image as RLImage,
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

OUT_DIR = Path(__file__).resolve().parent / "test_files"
OUT_DIR.mkdir(parents=True, exist_ok=True)

fake = Faker("zh_CN")
Faker.seed(42)

A4_W, A4_H = A4

# ── 中文字体注册 ──
_CJK_FONT_PATHS = [
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Songti.ttc",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
]

_CJK_FONT_NAME = None
_CJK_FONT_PATH = None

for _p in _CJK_FONT_PATHS:
    if os.path.exists(_p):
        try:
            pdfmetrics.registerFont(TTFont("CJK", _p, subfontIndex=0))
            _CJK_FONT_NAME = "CJK"
            _CJK_FONT_PATH = _p
            break
        except Exception:
            continue
# ═══════════════════════════════════════════════════════════════
#  1. 带复杂表格的 10 页 PDF
# ═══════════════════════════════════════════════════════════════
def gen_pdf_table_10p() -> None:
    path = OUT_DIR / "test_table_10p.pdf"
    story = []

    title_style = _cjk_style("title", fontSize=18, leading=24, spaceAfter=12)
    normal_style = _cjk_style("normal", fontSize=9, leading=14)

    for page in range(1, 11):
        story.append(_p(f"<b>第 {page} 页 — 销售数据报表</b>", title_style))
        story.append(Spacer(1, 6))

        header = ["序号", "产品名称", "规格型号", "单价(元)", "数量", "金额(元)", "备注"]
        data = [header]
        for i in range(1, 9):
            row_num = (page - 1) * 8 + i
            name = fake.catch_phrase()[:12]
            spec = fake.lexify(text="???-????", letters="ABCDEFGHIJKLMNOPQRSTUVWXYZ")
            price = round(fake.pyfloat(min_value=10, max_value=5000, right_digits=2), 2)
            qty = fake.random_int(1, 100)
            amount = round(price * qty, 2)
            note = fake.sentence()[:20] if fake.boolean(70) else ""
            data.append([str(row_num), name, spec, f"\u00a5{price:,.2f}", str(qty), f"\u00a5{amount:,.2f}", note])

        col_widths = [30, 90, 70, 60, 40, 65, 100]
        tbl = Table(data, colWidths=[w * 1.3 for w in col_widths], repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a3a5c")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), PDF_FONT_CJK),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTNAME", (0, 1), (-1, -1), PDF_FONT_CJK),
            ("FONTSIZE", (0, 1), (-1, -1), 8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#eef2f7")]),
            ("ALIGN", (3, 1), (5, -1), "RIGHT"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#b0bec5")),
            ("LINEBELOW", (0, 0), (-1, 0), 1.2, colors.HexColor("#1a3a5c")),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 8))
        story.append(_p(f"本页包含 {len(data) - 1} 条销售记录 | 生成日期：{fake.date()}", normal_style))
# ═══════════════════════════════════════════════════════════════
#  2. 中英文字体混合的 PDF
# ═══════════════════════════════════════════════════════════════
def gen_pdf_mixed_fonts() -> None:
    path = OUT_DIR / "test_mixed_fonts.pdf"
    story = []

    title_style = _cjk_style("title", fontSize=22, leading=30, spaceAfter=16, alignment=1)
    story.append(_p("中英文混合排版 Mixed-Language Layout", title_style))

    cn_style = _cjk_style("cn", fontSize=12, leading=20, spaceAfter=8)
    en_style = ParagraphStyle(
        "en", parent="Normal", fontName=PDF_FONT_WESTERN,
        fontSize=12, leading=20, spaceAfter=8,
    )
    mixed_style = _cjk_style("mixed", fontSize=11, leading=18, spaceAfter=6)

    sections = [
        ("一、中文段落",
         "在数字化时代，人工智能技术的飞速发展正在深刻改变我们的工作和生活方式。"
         "从智能语音助手到自动驾驶汽车，从医疗影像诊断到金融风险预测，AI 已经渗透到各行各业。"),
        ("二、English Paragraph",
         "Artificial intelligence is transforming the way we work and live. "
         "From voice assistants to self-driving cars, from medical imaging to financial forecasting, "
         "AI has permeated every industry."),
        ("三、中英混排 Mixed Content",
         "根据 McKinsey 2025 年报告，全球 AI 市场规模预计达到 $1.8 trillion。"
         "Natural Language Processing (NLP) 和 Computer Vision 是最活跃的两个领域。"
         "The report highlights that 67% of enterprises have adopted at least one AI solution, "
         "而中国的 adoption rate 达到了 72%，位居全球前列。"),
        ("四、特殊字符与标点",
         "中英文标点混用：引号「」『』\"\"''，破折号——省略号……书名号《》〈〉。"
         "English quotes: \"double\" 'single'. Dashes: em-dash\u2014en-dash\u2013. "
         "Symbols: \u00a9 \u00ae \u2122 \u20ac \u00a5 $ \u00a2 \u2030"),
    ]

    for section_title, content in sections:
        story.append(_p(f"<b>{section_title}</b>",
                        _cjk_style("h", fontSize=15, leading=22, spaceAfter=6)))
        story.append(_p(content, mixed_style))
        story.append(Spacer(1, 4))

    _build_pdf(path, story)
    _print_ok(path)
# ═══════════════════════════════════════════════════════════════
#  3. 带透明背景水印的 PDF
# ═══════════════════════════════════════════════════════════════
def gen_pdf_watermark() -> None:
    path = OUT_DIR / "test_watermark.pdf"

    # Pillow 生成半透明水印 PNG
    wm_img = Image.new("RGBA", (600, 400), (255, 255, 255, 0))
    draw = ImageDraw.Draw(wm_img)
    try:
        font_large = ImageFont.truetype(_CJK_FONT_PATH or "/System/Library/Fonts/PingFang.ttc", 48)
        font_small = ImageFont.truetype(_CJK_FONT_PATH or "/System/Library/Fonts/PingFang.ttc", 24)
    except Exception:
        font_large = ImageFont.load_default()
        font_small = ImageFont.load_default()

    draw.text((30, 150), "机密文件", fill=(200, 50, 80, 80), font=font_large)
    draw.text((30, 220), "CONFIDENTIAL", fill=(200, 50, 80, 80), font=font_small)
    for y in range(0, 400, 60):
        draw.text((100, y), "内部资料 · 请勿外传", fill=(150, 150, 150, 40), font=font_small)

    wm_bytes = io.BytesIO()
    wm_img.save(wm_bytes, format="PNG")

    # reportlab 构建 PDF，每页嵌入水印
    story = []
    body_style = _cjk_style("body", fontSize=12, leading=20, spaceAfter=10)

    for p in range(1, 4):
        story.append(_p(f"<b>第 {p} 页 · 年度工作报告</b>",
                        _cjk_style("h", fontSize=16, leading=22, spaceAfter=14)))
        story.append(_p(
            f"本报告属于公司内部机密文件，仅供授权人员查阅。"
            f"未经许可，不得以任何形式复制、传播或公开。本页包含{p * 3}条重要业务数据。",
            body_style,
        ))
        story.append(Spacer(1, 20))
        wm_bytes.seek(0)
        wm = RLImage(wm_bytes, width=400, height=260)
        wm.hAlign = "CENTER"
        story.append(wm)
        story.append(Spacer(1, 10))
        story.append(_p(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
            body_style,
        ))
        if p < 3:
            story.append(PageBreak())

    _build_pdf(path, story)
    _print_ok(path)


# ═══════════════════════════════════════════════════════════════
#  4. 包含 100 行数据的 Excel
# ═══════════════════════════════════════════════════════════════
def gen_excel_100_rows() -> None:
    path = OUT_DIR / "test_data_100rows.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "员工信息表"

    headers = ["序号", "姓名", "部门", "职位", "入职日期", "月薪(元)", "绩效评级", "邮箱", "手机号", "备注"]
    header_fill = PatternFill(start_color="1a3a5c", end_color="1a3a5c", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="b0bec5"),
        right=Side(style="thin", color="b0bec5"),
        top=Side(style="thin", color="b0bec5"),
        bottom=Side(style="thin", color="b0bec5"),
    )

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border

    departments = ["技术部", "产品部", "市场部", "销售部", "人事部", "财务部", "运营部", "设计部"]
    positions = {
        "技术部": ["高级工程师", "架构师", "前端开发", "后端开发", "测试工程师"],
        "产品部": ["产品总监", "产品经理", "产品助理"],
        "市场部": ["市场总监", "市场专员", "品牌经理"],
        "销售部": ["销售总监", "大客户经理", "销售代表"],
        "人事部": ["HR总监", "招聘专员", "培训经理"],
        "财务部": ["财务总监", "会计", "出纳"],
        "运营部": ["运营总监", "运营专员", "数据分析师"],
        "设计部": ["设计总监", "UI设计师", "UX设计师"],
    }
    ratings = ["S", "A", "B", "C", "D"]

    for i in range(1, 101):
        row = i + 1
        dept = departments[(i - 1) % len(departments)]
        pos_pool = positions[dept]
        pos = pos_pool[(i - 1) % len(pos_pool)]

        values = [
            i,
            fake.name(),
            dept,
            pos,
            fake.date_between(start_date="-10y", end_date="today").strftime("%Y-%m-%d"),
            round(fake.pyfloat(min_value=6000, max_value=45000, right_digits=2), 2),
            fake.random_element(ratings),
            fake.email(),
            fake.phone_number(),
            fake.sentence()[:15] if fake.boolean(40) else "",
        ]

        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
            if col in (1, 6):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == 7:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if val == "S":
                    cell.font = Font(bold=True, color="1a8a2e")
                elif val in ("A", "B"):
                    cell.font = Font(color="2563eb")
                elif val in ("C", "D"):
                    cell.font = Font(color="dc2626")
            if i % 2 == 0:
                cell.fill = PatternFill(start_color="f6f8fc", end_color="f6f8fc", fill_type="solid")

    widths = [6, 14, 10, 14, 13, 12, 10, 28, 16, 18]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = w

    ws.freeze_panes = "A2"

    # 汇总统计表
    ws2 = wb.create_sheet("汇总统计")
    for col, h in enumerate(["部门", "人数", "平均月薪(元)"], 1):
        c = ws2.cell(row=1, column=col, value=h)
        c.font = Font(bold=True)
    for di, dept in enumerate(departments, 2):
        ws2.cell(row=di, column=1, value=dept)
        ws2.cell(row=di, column=2, value=fake.random_int(8, 18))
        ws2.cell(row=di, column=3, value=round(fake.pyfloat(8000, 35000, 2), 2))

    wb.save(str(path))
    _print_ok(path)


# ═══════════════════════════════════════════════════════════════
#  5. 带图片的 Word 文档
# ═══════════════════════════════════════════════════════════════
def gen_docx_with_image() -> None:
    path = OUT_DIR / "test_with_image.docx"

    # 用 Pillow 生成测试柱状图
    img = Image.new("RGB", (600, 400), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    bars = [
        ("Q1", 120, "#2563eb"),
        ("Q2", 200, "#10b981"),
        ("Q3", 160, "#f59e0b"),
        ("Q4", 280, "#ef4444"),
    ]
    bar_w = 80
    chart_base_y = 350
    max_val = 300

    for bi, (label, val, color) in enumerate(bars):
        x = 60 + bi * 120
        bar_h = int(val / max_val * 250)
        draw.rectangle([x, chart_base_y - bar_h, x + bar_w, chart_base_y], fill=color)
        draw.text((x + bar_w // 2 - 20, chart_base_y + 10), label, fill="black")

    for v, y in [(0, 350), (100, 267), (200, 183), (300, 100)]:
        draw.text((20, y - 8), str(v), fill="#6b7280")

    draw.text((200, 20), "2025 季度销售对比", fill="#1a3a5c")

    img_bytes = io.BytesIO()
    img.save(img_bytes, format="PNG")
    img_bytes.seek(0)

    # 构建 Word
    doc = Document()

    title = doc.add_heading("2025 年度销售分析报告", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(
        "本报告基于2025年全年销售数据，对四个季度的业绩进行了对比分析。"
        "数据显示，Q4 销售额较 Q1 增长了 133%，全年呈现稳步上升趋势。",
        style="Intense Quote",
    )

    doc.add_picture(img_bytes, width=Inches(5.5))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("图1：2025年各季度销售额对比（单位：万元）").alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("详细分析", level=2)
    doc.add_paragraph(
        "第一季度作为传统淡季，销售额为 120 万元，处于全年最低水平。"
        "进入第二季度后，受益于新品发布和市场推广活动，销售额攀升至 200 万元，环比增长 66.7%。"
    )
    doc.add_paragraph(
        "第三季度受季节性因素影响，销售额小幅回落至 160 万元。"
        "但在第四季度，得益于双十一、年终促销等多重利好，销售额突破 280 万元，创下全年新高。"
    )

    doc.add_heading("改进建议", level=2)
    for s in [
        "加大 Q1 市场预热投入，提前锁定客户订单",
        "优化 Q3 产品组合策略，平滑季节性波动",
        "强化 Q4 库存管理，避免供货不足影响销售节奏",
        "建立客户分级体系，提升高价值客户的复购率",
    ]:
        doc.add_paragraph(s, style="List Bullet")

    doc.save(str(path))
    _print_ok(path)


# ═══════════════════════════════════════════════════════════════
def main() -> None:
    print("\n生成测试文件到 test_files/ 目录 ...\n")
    gen_pdf_table_10p()
    gen_pdf_mixed_fonts()
    gen_pdf_watermark()
    gen_excel_100_rows()
    gen_docx_with_image()
    print(f"\n全部完成，共 5 个文件 → {OUT_DIR.resolve()}\n")


if __name__ == "__main__":
    main()
    _build_pdf(path, story)
    _print_ok(path)

        if page < 10:
            story.append(PageBreak())

    _build_pdf(path, story)
    _print_ok(path)

PDF_FONT_WESTERN = "Helvetica"
PDF_FONT_CJK = _CJK_FONT_NAME or "Helvetica"

def _print_ok(path: Path) -> None:
    size_kb = path.stat().st_size / 1024
    print(f"  ✔ {path.name:.<40s} {size_kb:>8.1f} KB")

def _cjk_style(name: str, parent: str = "Normal", **kw) -> ParagraphStyle:
    return ParagraphStyle(
        name,
        parent=parent,
        fontName=PDF_FONT_CJK if PDF_FONT_CJK != "Helvetica" else PDF_FONT_WESTERN,
        **kw,
    )

def _p(text: str, style: ParagraphStyle | None = None) -> Paragraph:
    if style is None:
        style = _cjk_style("auto", fontSize=10, leading=16)
    return Paragraph(text, style)

def _build_pdf(path: Path, story: list) -> None:
    doc = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        leftMargin=25 * mm,
        rightMargin=25 * mm,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
    )
    doc.build(story)