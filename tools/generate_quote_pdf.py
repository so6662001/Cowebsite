"""
客户报价单 PDF 生成器
=====================

把 Excel 报价模型的输出转成"客户能看的"PDF 报价单。
- 中文字体：文泉驿微米黑（系统已带）
- 单页 A4，含品牌头、客户信息、3 档版本对比、付费模式选择、12 期分期模拟、有效期与签字栏

运行方式：
    python3 tools/generate_quote_pdf.py             # 生成内置样例
    python3 tools/generate_quote_pdf.py --json data.json  # 用 JSON 输入

输出：tools/sample_quote.pdf （样例）
"""

from __future__ import annotations
import argparse
import json
from datetime import datetime, timedelta
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)


CN_FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
pdfmetrics.registerFont(TTFont("WQY", CN_FONT_PATH))

PRIMARY = colors.HexColor("#1F4E79")
ACCENT = colors.HexColor("#2E75B6")
LIGHT = colors.HexColor("#DEEBF7")
HIGHLIGHT = colors.HexColor("#FFF2CC")
SUCCESS = colors.HexColor("#E2EFDA")
WARN = colors.HexColor("#FCE4D6")


def build_styles():
    s = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle("title", parent=s["Title"], fontName="WQY",
                                fontSize=22, leading=28, textColor=PRIMARY,
                                alignment=0, spaceAfter=4),
        "subtitle": ParagraphStyle("subtitle", parent=s["Normal"], fontName="WQY",
                                   fontSize=11, leading=16, textColor=colors.grey,
                                   alignment=0, spaceAfter=10),
        "h1": ParagraphStyle("h1", parent=s["Heading2"], fontName="WQY",
                             fontSize=14, leading=20, textColor=PRIMARY,
                             spaceBefore=8, spaceAfter=4),
        "h2": ParagraphStyle("h2", parent=s["Heading3"], fontName="WQY",
                             fontSize=12, leading=18, textColor=PRIMARY,
                             spaceBefore=4, spaceAfter=2),
        "body": ParagraphStyle("body", parent=s["Normal"], fontName="WQY",
                               fontSize=10, leading=15),
        "small": ParagraphStyle("small", parent=s["Normal"], fontName="WQY",
                                fontSize=8.5, leading=12, textColor=colors.grey),
        "money": ParagraphStyle("money", parent=s["Normal"], fontName="WQY",
                                fontSize=18, leading=22, textColor=PRIMARY,
                                alignment=1),
        "tag": ParagraphStyle("tag", parent=s["Normal"], fontName="WQY",
                              fontSize=10, leading=14, textColor=colors.white,
                              alignment=1),
    }
    return styles


def fmt_money(v: float) -> str:
    return f"¥{int(round(v)):,}"


def header_table(client: dict, styles):
    today = datetime.now().strftime("%Y 年 %m 月 %d 日")
    valid = (datetime.now() + timedelta(days=30)).strftime("%Y 年 %m 月 %d 日")
    rows = [
        ["客户名称", client.get("client_name", ""), "报价编号", client.get("quote_no", "")],
        ["联系人 / 职务", client.get("contact", ""), "出具日期", today],
        ["所属行业", client.get("industry", ""), "有效期至", valid],
        ["客户经理", client.get("sales", ""), "联系电话", client.get("sales_phone", "")],
    ]
    t = Table(rows, colWidths=[28*mm, 60*mm, 28*mm, 54*mm])
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "WQY", 10),
        ("BACKGROUND", (0, 0), (0, -1), LIGHT),
        ("BACKGROUND", (2, 0), (2, -1), LIGHT),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t


def versions_table(quote: dict, styles):
    """3 档版本对比"""
    seed, biz, flag = quote["seed"], quote["biz"], quote["flag"]
    header_row = ["", "种子版", "经营版（推荐）", "旗舰版"]
    rows = [
        header_row,
        ["一次性实施费",
         fmt_money(seed["impl"]), fmt_money(biz["impl"]), fmt_money(flag["impl"])],
        ["月订阅费 (起)",
         fmt_money(seed["monthly"]), fmt_money(biz["monthly"]), fmt_money(flag["monthly"])],
        ["12 期月供合计",
         fmt_money(seed["impl"]/12 + seed["monthly"]),
         fmt_money(biz["impl"]/12 + biz["monthly"]),
         fmt_money(flag["impl"]/12 + flag["monthly"])],
        ["首年总投入",
         fmt_money(seed["impl"] + seed["monthly"]*12),
         fmt_money(biz["impl"] + biz["monthly"]*12),
         fmt_money(flag["impl"] + flag["monthly"]*12)],
        ["3 年总投入",
         fmt_money(seed["impl"] + seed["monthly"]*36),
         fmt_money(biz["impl"] + biz["monthly"]*36),
         fmt_money(flag["impl"] + flag["monthly"]*36)],
        ["核心模块",
         Paragraph(seed["modules"], styles["body"]),
         Paragraph(biz["modules"], styles["body"]),
         Paragraph(flag["modules"], styles["body"])],
        ["适用规模",
         Paragraph(seed["scale"], styles["body"]),
         Paragraph(biz["scale"], styles["body"]),
         Paragraph(flag["scale"], styles["body"])],
    ]
    t = Table(rows, colWidths=[36*mm, 44*mm, 44*mm, 44*mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "WQY", 10),
        ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (2, 1), (2, -1), HIGHLIGHT),  # 经营版高亮
        ("BACKGROUND", (0, 1), (0, -1), LIGHT),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("FONT", (0, 1), (-1, 5), "WQY", 11),
    ]))
    return t


def payment_modes_table(quote: dict, styles):
    """付费模式对比"""
    biz = quote["biz"]
    impl = biz["impl"]
    monthly = biz["monthly"]
    rows = [
        ["付费模式", "首付", "月费", "首年支出", "适合"],
        ["A 经典", fmt_money(impl), fmt_money(monthly),
         fmt_money(impl + monthly*12), "现金流好"],
        ["B 0 首付月供", "¥0", fmt_money(monthly*1.3),
         fmt_money(monthly*1.3*12), "现金紧/抠门老板"],
        ["C 利润分成", "¥0",
         fmt_money(monthly*0.5) + " + 节省额 15-20%",
         "按效果", "高利润大客户"],
        ["A + 12 期分期", fmt_money(impl/12) + "/月", fmt_money(monthly),
         fmt_money(impl + monthly*12), "★ 最常见"],
    ]
    t = Table(rows, colWidths=[34*mm, 32*mm, 36*mm, 40*mm, 26*mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "WQY", 10),
        ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, -1), (-1, -1), SUCCESS),  # 推荐高亮
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return t


def installment_table(quote: dict, styles):
    """12 期分期模拟"""
    biz = quote["biz"]
    monthly_install = biz["impl"] / 12
    monthly_sub = biz["monthly"]

    rows = [["期次", "实施费分期", "月订阅费", "当月合计"]]
    total = 0
    for m in range(1, 13):
        sum_m = monthly_install + monthly_sub
        total += sum_m
        rows.append([f"第 {m} 月",
                     fmt_money(monthly_install),
                     fmt_money(monthly_sub),
                     fmt_money(sum_m)])
    rows.append(["12 期合计",
                 fmt_money(biz["impl"]),
                 fmt_money(monthly_sub * 12),
                 fmt_money(total)])

    t = Table(rows, colWidths=[26*mm, 36*mm, 36*mm, 38*mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "WQY", 9),
        ("BACKGROUND", (0, 0), (-1, 0), ACCENT),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, -1), (-1, -1), SUCCESS),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def commitment_table(styles):
    """风险逆转 / 承诺条款"""
    rows = [
        ["项", "承诺内容"],
        ["试用期",  "签约后前 30 个自然日免月租，效果不达标可终止合作"],
        ["KPI 退款", "上线 6 个月内 KPI 不达标 → 退月租 + 免后 3 个月 + 实施费 50% 兜底"],
        ["数据所有权", "客户数据所有权归客户；合同终止 30 日内全部交还+销毁副本"],
        ["分期支持", "实施费可申请 12 / 24 期免息分期（合作金融机构垫资）"],
        ["价格锁定", "合同期内月费单价不变；3 年合同享 8.5 折"],
    ]
    t = Table(rows, colWidths=[30*mm, 140*mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "WQY", 9.5),
        ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (0, -1), LIGHT),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (1, 0), (1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t


def signature_table(client, styles):
    rows = [
        ["客户方（盖章）", "服务方（盖章）"],
        ["", ""],
        [Paragraph("授权代表（签字）：________________", styles["body"]),
         Paragraph("授权代表（签字）：________________", styles["body"])],
        [Paragraph("日期：______ 年 ______ 月 ______ 日", styles["body"]),
         Paragraph("日期：______ 年 ______ 月 ______ 日", styles["body"])],
    ]
    t = Table(rows, colWidths=[85*mm, 85*mm])
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "WQY", 10),
        ("BACKGROUND", (0, 0), (-1, 0), LIGHT),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("MINROWHEIGHT", (0, 1), (-1, 1), 30),
    ]))
    return t


def build_pdf(data: dict, output_path: str):
    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            leftMargin=18*mm, rightMargin=18*mm,
                            topMargin=14*mm, bottomMargin=14*mm,
                            title="数字化系统报价单")
    styles = build_styles()
    story = []

    # 标题
    story.append(Paragraph(data.get("brand", "钢铁数字化系统"), styles["title"]))
    story.append(Paragraph("数字化系统专属报价单（v2.0）", styles["subtitle"]))

    # 客户信息
    story.append(header_table(data["client"], styles))
    story.append(Spacer(1, 6*mm))

    # 价格
    story.append(Paragraph("一、3 档版本菜单（推荐：经营版）", styles["h1"]))
    story.append(versions_table(data["quote"], styles))
    story.append(Spacer(1, 4*mm))

    # 付费模式
    story.append(Paragraph("二、付费模式选择（基于『经营版』）", styles["h1"]))
    story.append(payment_modes_table(data["quote"], styles))
    story.append(Spacer(1, 4*mm))

    # 12 期模拟
    story.append(Paragraph("三、12 期分期模拟（最常见方式）", styles["h1"]))
    story.append(installment_table(data["quote"], styles))

    story.append(PageBreak())

    # 第二页：承诺 + 签字
    story.append(Paragraph("四、风险承诺与服务保障", styles["h1"]))
    story.append(commitment_table(styles))
    story.append(Spacer(1, 5*mm))

    story.append(Paragraph("五、本次推荐方案（销售备注）", styles["h1"]))
    note = data.get("note", "推荐经营版 + A 经典模式 + 12 期分期。前 30 天免月租试用。"
                            "首年质量追溯 + 成本核算优先上线，3 周内见效。")
    story.append(Paragraph(note, styles["body"]))
    story.append(Spacer(1, 8*mm))

    story.append(Paragraph("六、签署确认", styles["h1"]))
    story.append(signature_table(data["client"], styles))
    story.append(Spacer(1, 6*mm))

    story.append(Paragraph(
        "本报价单基于客户方提供的初步信息制作，最终条款以双方正式签署的服务合同为准。"
        "本报价单自出具日起 30 日内有效。", styles["small"]))

    doc.build(story)


def sample_data():
    """内置样例：邯郸 XX 焊管厂"""
    return {
        "brand": "钢联数字化",
        "client": {
            "client_name": "邯郸 XX 钢管科技有限公司",
            "quote_no": "Q-20260418-001",
            "contact": "张 XX 董事长",
            "industry": "钢管生产（4 条焊管线）",
            "sales": "李 XX",
            "sales_phone": "138-XXXX-XXXX",
        },
        "quote": {
            "seed": {
                "impl": 19800, "monthly": 3500,
                "modules": "质量追溯（单产线）",
                "scale": "1-2 产线试点",
            },
            "biz": {
                "impl": 68000, "monthly": 9800,
                "modules": "质量追溯+计划排程+成本自动化（≤5 产线）",
                "scale": "80% 主流",
            },
            "flag": {
                "impl": 188000, "monthly": 25000,
                "modules": "全套+EAM+多基地+不限产线",
                "scale": "6+ 产线大厂",
            },
        },
        "note": ("综合贵厂 4 条焊管线规模与出口客户占 15% 的现状，"
                 "推荐【经营版 + A 经典模式 + 12 期分期】，"
                 "首年实施费分 12 期月供 ¥5,667，加月租 ¥9,800，"
                 "每月合计 ¥15,467（相当于贵厂一台二手叉车每月折旧）。"
                 "前 30 天免月租试用，先上『质量追溯』+『成本核算』两个模块，"
                 "3 周内可向您交付首批 4 条线吨成本对比报表。"),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", help="JSON 输入文件路径")
    parser.add_argument("--out", default="tools/sample_quote.pdf", help="PDF 输出路径")
    args = parser.parse_args()

    if args.json:
        data = json.loads(Path(args.json).read_text(encoding="utf-8"))
    else:
        data = sample_data()

    build_pdf(data, args.out)
    print(f"已生成 PDF：{args.out}")


if __name__ == "__main__":
    main()
