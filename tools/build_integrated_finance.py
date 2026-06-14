"""
跨业务财务整合模型
====================

3 业务（钢厂主业 + 钢贸 SaaS + 钢铁交易平台）的完整财务整合

输出：tools/跨业务财务整合模型.xlsx

工作表：
    1. 使用说明
    2. 业务 1 钢厂主业（3 年）
    3. 业务 2 钢贸 SaaS（3 年）
    4. 业务 3 平台 + 信用数据（3 年）
    5. 3 业务合并损益（年度）
    6. 3 业务合并现金流（月度，36 月）
    7. 协同效应测算
    8. 估值与融资节奏
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
PRIMARY = "1F4E79"
LIGHT_BG = "DEEBF7"
INPUT_BG = "FFF2CC"
SUCCESS = "E2EFDA"
WARN = "FCE4D6"
WHITE = "FFFFFF"

H_FILL = PatternFill("solid", fgColor=PRIMARY)
H_FONT = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
TITLE_FONT = Font(name="Microsoft YaHei", size=16, bold=True, color=PRIMARY)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_h(c):
    c.fill = H_FILL; c.font = H_FONT; c.alignment = CENTER; c.border = BORDER


def style_t(c, fill=None, align=None, money=False, pct=False, bold=False):
    c.font = Font(name="Microsoft YaHei", size=10, bold=bold)
    c.alignment = align or CENTER; c.border = BORDER
    if fill: c.fill = PatternFill("solid", fgColor=fill)
    if money: c.number_format = '"¥"#,##0'
    if pct: c.number_format = "0.0%"


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_readme(ws):
    ws.title = "使用说明"
    widths(ws, [4, 26, 70])
    ws.merge_cells("B2:C2")
    ws["B2"] = "3 业务跨业务财务整合模型"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    rows = [
        ("3 业务", "业务 1 钢厂主业 / 业务 2 钢贸 SaaS / 业务 3 平台+信用数据"),
        ("使用步骤",
         "① 业务 1/2/3 各自 sheet → 调整黄色单元格\n"
         "② 切换『3 业务合并损益』看年度汇总\n"
         "③ 切换『3 业务合并现金流』看 36 月月度\n"
         "④ 切换『协同效应测算』看跨业务价值\n"
         "⑤ 切换『估值与融资』看每年融资节奏"),
        ("基准假设",
         "- 业务 1 钢厂 Y1 ¥800 万 / Y2 ¥2500 万 / Y3 ¥5500 万\n"
         "- 业务 2 钢贸 SaaS Y1 ¥150 万 / Y2 ¥1900 万 / Y3 ¥2.5 亿\n"
         "- 业务 3 平台+数据 Y1 ¥50 万 / Y2 ¥700 万 / Y3 ¥5300 万\n"
         "- 3 业务 3 年合计 ¥3.7 亿"),
        ("关键产出",
         "1. 3 业务整合 ARR\n"
         "2. 跨业务协同价值（CAC 降低 / LTV 提升）\n"
         "3. 每年融资缺口 / 自我造血能力\n"
         "4. 估值演进 + 融资轮次时点"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        c1 = ws.cell(row=i, column=2, value=k); style_t(c1, fill=PRIMARY, bold=True)
        c1.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
        c2 = ws.cell(row=i, column=3, value=v); style_t(c2, align=LEFT)
        ws.row_dimensions[i].height = max(50, len(v) * 0.7)


def build_business1(ws):
    ws.title = "业务1钢厂主业"
    widths(ws, [4, 26, 14, 14, 14, 14])

    ws.merge_cells("B2:F2")
    ws["B2"] = "业务 1：钢厂数字化 SaaS（3 年）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["项", "Y1", "Y2", "Y3", "3 年累计"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    # 收入
    style_t(ws.cell(row=5, column=2, value="▼ 收入"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=5, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    rev_items = [
        ("实施费收入", 4000000, 10000000, 22000000),
        ("月租 SaaS 收入", 3500000, 13500000, 30000000),
        ("增值服务", 500000, 1500000, 3000000),
    ]
    for i, (name, y1, y2, y3) in enumerate(rev_items, 6):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT)
        for j, v in enumerate([y1, y2, y3], 3):
            style_t(ws.cell(row=i, column=j, value=v), fill=INPUT_BG, money=True)
        cell = ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
        style_t(cell, fill=SUCCESS, money=True, bold=True)

    style_t(ws.cell(row=9, column=2, value="总收入"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=9, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=9, column=col, value=f"=SUM({col_letter}6:{col_letter}8)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=9, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 成本
    style_t(ws.cell(row=11, column=2, value="▼ 成本"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=11, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    cost_items = [
        ("人力成本（销售+实施+客服）", 2400000, 5500000, 11000000),
        ("营销 + 老板局 + 物料", 1000000, 2000000, 4000000),
        ("产品 + 技术 + 数据", 800000, 1500000, 2500000),
        ("行政 + 财务 + 其他", 400000, 600000, 1000000),
    ]
    for i, (name, y1, y2, y3) in enumerate(cost_items, 12):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT)
        for j, v in enumerate([y1, y2, y3], 3):
            style_t(ws.cell(row=i, column=j, value=v), fill=WARN, money=True)
        cell = ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
        style_t(cell, fill=WARN, money=True, bold=True)

    style_t(ws.cell(row=16, column=2, value="总成本"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=16, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=16, column=col, value=f"=SUM({col_letter}12:{col_letter}15)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=16, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 净利润
    style_t(ws.cell(row=18, column=2, value="净利润"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=18, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=18, column=col, value=f"={col_letter}9-{col_letter}16")
        style_t(cell, fill=SUCCESS, money=True, bold=True)
        ws.cell(row=18, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)


def build_business2(ws):
    ws.title = "业务2钢贸SaaS"
    widths(ws, [4, 26, 14, 14, 14, 14])

    ws.merge_cells("B2:F2")
    ws["B2"] = "业务 2：钢贸企业 SaaS（3 年）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["项", "Y1", "Y2", "Y3", "3 年累计"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    # 收入
    style_t(ws.cell(row=5, column=2, value="▼ 收入"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=5, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    rev_items = [
        ("ERP/WMS 实施费", 500000, 6000000, 30000000),
        ("AI 月费订阅", 1000000, 13000000, 200000000),
        ("货袋子+现货平台增值", 0, 500000, 5000000),
        ("咨询服务（接班+战略）", 0, 1000000, 10000000),
    ]
    for i, (name, y1, y2, y3) in enumerate(rev_items, 6):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT)
        for j, v in enumerate([y1, y2, y3], 3):
            style_t(ws.cell(row=i, column=j, value=v), fill=INPUT_BG, money=True)
        cell = ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
        style_t(cell, fill=SUCCESS, money=True, bold=True)

    style_t(ws.cell(row=10, column=2, value="总收入"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=10, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=10, column=col, value=f"=SUM({col_letter}6:{col_letter}9)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=10, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 成本
    style_t(ws.cell(row=12, column=2, value="▼ 成本"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=12, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    cost_items = [
        ("人力（销售+CS+顾问）", 1000000, 6000000, 30000000),
        ("营销 + 推广", 200000, 1500000, 8000000),
        ("产品 + AI + 数据", 800000, 3000000, 15000000),
        ("自营 GMV + 平台运营", 100000, 2000000, 10000000),
        ("行政 + 其他", 100000, 500000, 3000000),
    ]
    for i, (name, y1, y2, y3) in enumerate(cost_items, 13):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT)
        for j, v in enumerate([y1, y2, y3], 3):
            style_t(ws.cell(row=i, column=j, value=v), fill=WARN, money=True)
        cell = ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
        style_t(cell, fill=WARN, money=True, bold=True)

    style_t(ws.cell(row=18, column=2, value="总成本"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=18, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=18, column=col, value=f"=SUM({col_letter}13:{col_letter}17)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=18, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 净利润
    style_t(ws.cell(row=20, column=2, value="净利润"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=20, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=20, column=col, value=f"={col_letter}10-{col_letter}18")
        style_t(cell, fill=SUCCESS, money=True, bold=True)
        ws.cell(row=20, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)


def build_business3(ws):
    ws.title = "业务3平台数据"
    widths(ws, [4, 26, 14, 14, 14, 14])

    ws.merge_cells("B2:F2")
    ws["B2"] = "业务 3：交易平台 + 信用数据（3 年）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["项", "Y1", "Y2", "Y3", "3 年累计"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    # 收入
    style_t(ws.cell(row=5, column=2, value="▼ 收入"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=5, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    rev_items = [
        ("撮合佣金", 300000, 4000000, 25000000),
        ("仓储+物流分润", 200000, 1800000, 14000000),
        ("信用数据外卖", 0, 1000000, 13000000),
        ("票据+担保服务费", 0, 200000, 1000000),
    ]
    for i, (name, y1, y2, y3) in enumerate(rev_items, 6):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT)
        for j, v in enumerate([y1, y2, y3], 3):
            style_t(ws.cell(row=i, column=j, value=v), fill=INPUT_BG, money=True)
        cell = ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
        style_t(cell, fill=SUCCESS, money=True, bold=True)

    style_t(ws.cell(row=10, column=2, value="总收入"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=10, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=10, column=col, value=f"=SUM({col_letter}6:{col_letter}9)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=10, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 成本
    style_t(ws.cell(row=12, column=2, value="▼ 成本"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=12, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    cost_items = [
        ("人力（BD+数据销售+运营）", 300000, 2000000, 10000000),
        ("营销 + 行业 KOL", 100000, 600000, 3000000),
        ("数据 + 算法 + 风控", 200000, 1000000, 5000000),
        ("行政 + 其他", 50000, 200000, 1000000),
    ]
    for i, (name, y1, y2, y3) in enumerate(cost_items, 13):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT)
        for j, v in enumerate([y1, y2, y3], 3):
            style_t(ws.cell(row=i, column=j, value=v), fill=WARN, money=True)
        cell = ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
        style_t(cell, fill=WARN, money=True, bold=True)

    style_t(ws.cell(row=17, column=2, value="总成本"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=17, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=17, column=col, value=f"=SUM({col_letter}13:{col_letter}16)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=17, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 净利润
    style_t(ws.cell(row=19, column=2, value="净利润"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=19, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=19, column=col, value=f"={col_letter}10-{col_letter}17")
        style_t(cell, fill=SUCCESS, money=True, bold=True)
        ws.cell(row=19, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)


def build_consolidated_pl(ws):
    ws.title = "3业务合并损益"
    widths(ws, [4, 26, 16, 16, 16, 16])

    ws.merge_cells("B2:F2")
    ws["B2"] = "3 业务合并损益表（3 年）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["项", "Y1", "Y2", "Y3", "3 年累计"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    # 收入
    revenue_refs = [
        ("业务 1 钢厂收入", "业务1钢厂主业!", "C9", "D9", "E9", "F9"),
        ("业务 2 钢贸收入", "业务2钢贸SaaS!", "C10", "D10", "E10", "F10"),
        ("业务 3 平台数据收入", "业务3平台数据!", "C10", "D10", "E10", "F10"),
    ]
    for i, (name, sheet, c1, c2, c3, c4) in enumerate(revenue_refs, 5):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT, bold=True)
        for j, ref in enumerate([c1, c2, c3, c4], 3):
            cell = ws.cell(row=i, column=j, value=f"='{sheet[:-1]}'!{ref}")
            style_t(cell, fill=SUCCESS, money=True, bold=True)

    # 总收入
    i = 8
    style_t(ws.cell(row=i, column=2, value="3 业务总收入"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"=SUM({col_letter}5:{col_letter}7)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 成本
    cost_refs = [
        ("业务 1 钢厂成本", "'业务1钢厂主业'", "C16", "D16", "E16", "F16"),
        ("业务 2 钢贸成本", "'业务2钢贸SaaS'", "C18", "D18", "E18", "F18"),
        ("业务 3 平台成本", "'业务3平台数据'", "C17", "D17", "E17", "F17"),
    ]
    for i, (name, sheet, c1, c2, c3, c4) in enumerate(cost_refs, 10):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT, bold=True)
        for j, ref in enumerate([c1, c2, c3, c4], 3):
            cell = ws.cell(row=i, column=j, value=f"={sheet}!{ref}")
            style_t(cell, fill=WARN, money=True, bold=True)

    # 总成本
    i = 13
    style_t(ws.cell(row=i, column=2, value="3 业务总成本"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"=SUM({col_letter}10:{col_letter}12)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 净利润
    i = 15
    style_t(ws.cell(row=i, column=2, value="3 业务净利润"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=12, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"={col_letter}8-{col_letter}13")
        style_t(cell, fill=SUCCESS, money=True, bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=12, bold=True)

    # 毛利率
    i = 17
    style_t(ws.cell(row=i, column=2, value="净利率"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 6):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"={col_letter}15/{col_letter}8")
        style_t(cell, fill=INPUT_BG, pct=True, bold=True)


def build_synergy(ws):
    ws.title = "协同效应测算"
    widths(ws, [4, 30, 14, 14, 14, 24])

    ws.merge_cells("B2:F2")
    ws["B2"] = "3 业务协同效应（额外创造价值）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["协同方式", "Y1(¥)", "Y2(¥)", "Y3(¥)", "说明"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    synergies = [
        ("业务 1 客户 → 业务 2 引荐", 0, 1500000, 8000000, "钢厂的下游钢贸商 → 钢贸 SaaS"),
        ("业务 2 客户 → 业务 3 数据", 0, 500000, 5000000, "钢贸 SaaS 数据沉淀 → 卖银行"),
        ("业务 1+2 联合销售", 200000, 2000000, 8000000, "三位一体方案高客单价"),
        ("CAC 降低（共享线索）", 100000, 1500000, 5000000, "节省营销 + 销售成本"),
        ("LTV 提升（交叉销售）", 0, 2000000, 12000000, "客户买多个产品"),
        ("数据壁垒带来估值溢价", "", "", "", "20-50% 估值上浮"),
    ]
    for i, (name, y1, y2, y3, desc) in enumerate(synergies, 5):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT, bold=True)
        for j, v in enumerate([y1, y2, y3], 3):
            if isinstance(v, (int, float)):
                style_t(ws.cell(row=i, column=j, value=v), fill=SUCCESS, money=True, bold=True)
            else:
                style_t(ws.cell(row=i, column=j, value=v), fill=LIGHT_BG)
        style_t(ws.cell(row=i, column=6, value=desc), align=LEFT)

    # 协同合计
    i = 11
    style_t(ws.cell(row=i, column=2, value="协同总价值（额外）"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 6):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"=SUM({col_letter}5:{col_letter}9)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)


def build_valuation(ws):
    ws.title = "估值与融资"
    widths(ws, [4, 26, 16, 16, 16, 24])

    ws.merge_cells("B2:F2")
    ws["B2"] = "估值演进 + 融资节奏"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["时点", "ARR(¥)", "PS 倍数", "估值(¥)", "建议融资"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    rounds = [
        ("Y0 末（启动）", 0, "—", 50000000, "天使 ¥500 万"),
        ("Y1 末", 10000000, 8, 80000000, "Pre-A ¥3,000 万"),
        ("Y2 末", 50000000, 10, 500000000, "A 轮 ¥1-3 亿"),
        ("Y3 末", 300000000, 8, 2400000000, "B 轮 ¥3-10 亿"),
        ("Y4 末（预测）", 600000000, 10, 6000000000, "Pre-IPO ¥10-30 亿"),
        ("Y5+（预测）", 1200000000, 12, 14400000000, "IPO ¥500-1000 亿市值"),
    ]
    for i, (time, arr, ps, val, raise_) in enumerate(rounds, 5):
        bg = SUCCESS if "Y3" in time or "Y4" in time else LIGHT_BG
        style_t(ws.cell(row=i, column=2, value=time), fill=bg, bold=True)
        style_t(ws.cell(row=i, column=3, value=arr), fill=INPUT_BG, money=True, bold=True)
        style_t(ws.cell(row=i, column=4, value=ps), fill=INPUT_BG, bold=True)
        style_t(ws.cell(row=i, column=5, value=val), fill=SUCCESS, money=True, bold=True)
        style_t(ws.cell(row=i, column=6, value=raise_), fill=bg, align=LEFT)


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_business1(wb.create_sheet())
    build_business2(wb.create_sheet())
    build_business3(wb.create_sheet())
    build_consolidated_pl(wb.create_sheet())
    build_synergy(wb.create_sheet())
    build_valuation(wb.create_sheet())

    out = "tools/跨业务财务整合模型.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
