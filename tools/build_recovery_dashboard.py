"""
失败客户挽回数据看板
====================

输出：tools/失败客户挽回数据看板.xlsx

工作表：
    1. 使用说明
    2. 挽回客户主表（30+ 客户）
    3. 挽回漏斗（7 类型）
    4. 销售个人挽回率
    5. 月度趋势
    6. ROI 测算
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
PRIMARY = "1F4E79"
LIGHT_BG = "DEEBF7"
INPUT_BG = "FFF2CC"
SUCCESS = "E2EFDA"
WARN = "FCE4D6"
CRITICAL = "FFB6C1"
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
    ws["B2"] = "失败客户挽回数据看板"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT
    rows = [
        ("适用对象", "销售总监 + 业务 VP + 老板"),
        ("使用步骤",
         "① 销售把过去 30+ 失败客户录入『挽回客户主表』\n"
         "② 标注 7 类型 + 启动挽回\n"
         "③ 切换『挽回漏斗』看 7 类型进展\n"
         "④ 切换『销售个人挽回率』看每个销售的挽回能力\n"
         "⑤ 切换『ROI 测算』看挽回投入产出"),
        ("挽回率参考",
         "A 时机型 40-60% | B 价格型 20-30% | C 价值型 25-35%\n"
         "D 关系型 30-40% | E 信任型 15-20% | F 竞品型 5-10% | G 内部型 15-25%"),
        ("挽回 ROI",
         "单客挽回成本 ¥9,000 vs 新客 ¥30-50K\n"
         "挽回 ROI = 3-5 倍 vs 新客获取"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        c1 = ws.cell(row=i, column=2, value=k); style_t(c1, fill=PRIMARY, bold=True)
        c1.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
        c2 = ws.cell(row=i, column=3, value=v); style_t(c2, align=LEFT)
        ws.row_dimensions[i].height = max(50, len(v) * 0.7)


def build_main(ws):
    ws.title = "挽回客户主表"
    widths(ws, [5, 22, 14, 14, 8, 14, 14, 14, 14, 30])

    ws.merge_cells("A1:J1")
    ws["A1"] = "失败客户挽回主表（黄色 = 输入）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    headers = ["#", "客户名", "负责销售", "失败日期",
               "失败类型", "挽回阶段", "启动挽回日期",
               "预计签约金额(¥)", "实际签约金额(¥)", "备注"]
    for j, h in enumerate(headers, 1):
        style_h(ws.cell(row=3, column=j, value=h))
    ws.row_dimensions[3].height = 32

    for i in range(4, 34):
        idx = i - 3
        style_t(ws.cell(row=i, column=1, value=idx), fill=LIGHT_BG, bold=True)
        for col in range(2, 11):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG, align=LEFT)
        # 金额列
        for col in [8, 9]:
            ws.cell(row=i, column=col).number_format = '"¥"#,##0'

    # 数据有效性
    dv_type = DataValidation(type="list",
        formula1='"A时机,B价格,C价值,D关系,E信任,F竞品,G内部"', allow_blank=True)
    ws.add_data_validation(dv_type); dv_type.add("E4:E33")

    dv_stage = DataValidation(type="list",
        formula1='"待启动,联系中,接触成功,POC中,签约成功,挽回失败,永久放弃"',
        allow_blank=True)
    ws.add_data_validation(dv_stage); dv_stage.add("F4:F33")

    dv_sales = DataValidation(type="list",
        formula1='"A1,A2,A3,B1,B2,B3,C1,C2,C3"', allow_blank=True)
    ws.add_data_validation(dv_sales); dv_sales.add("C4:C33")

    # 条件格式
    for stage, color in [("签约成功", "00B050"), ("挽回失败", "FF6B6B"),
                          ("POC中", "FFC000"), ("接触成功", "FFEB9C")]:
        ws.conditional_formatting.add("F4:F33",
            CellIsRule(operator="equal", formula=[f'"{stage}"'],
                       fill=PatternFill("solid", fgColor=color)))

    # 示例
    sample = ["邯郸 XX 钢贸", "B1", "2026-03-15", "B价格", "联系中", "2026-06-01",
              100000, "", "降价 + 同行案例"]
    for j, v in enumerate(sample, 2):
        ws.cell(row=4, column=j, value=v)

    ws.freeze_panes = "D4"


def build_funnel(ws):
    ws.title = "挽回漏斗"
    widths(ws, [4, 16, 12, 12, 12, 14, 14, 14])

    ws.merge_cells("B2:H2")
    ws["B2"] = "7 类型挽回漏斗（自动统计）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["类型", "总失败数", "已启动", "联系中", "POC 中",
               "签约成功", "挽回失败", "挽回率"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    types = ["A时机", "B价格", "C价值", "D关系", "E信任", "F竞品", "G内部"]
    target_rates = ["40-60%", "20-30%", "25-35%", "30-40%", "15-20%", "5-10%", "15-25%"]

    for i, (t, rate) in enumerate(zip(types, target_rates), 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=t), fill=bg, bold=True)

        # 总失败数
        cell = ws.cell(row=i, column=3, value=f'=COUNTIF(\'挽回客户主表\'!E4:E33,"{t}")')
        style_t(cell, fill=SUCCESS, bold=True)

        # 已启动 / 联系中 / POC / 签约 / 失败
        for j, stage in enumerate(["联系中", "联系中", "POC中", "签约成功", "挽回失败"], 4):
            if j == 4:  # 已启动 = 联系中 + POC中 + 签约 + 失败
                cell = ws.cell(row=i, column=4,
                    value=(f'=COUNTIFS(\'挽回客户主表\'!E4:E33,"{t}",\'挽回客户主表\'!F4:F33,"<>待启动")'
                           f'-COUNTIFS(\'挽回客户主表\'!E4:E33,"{t}",\'挽回客户主表\'!F4:F33,"")'))
            else:
                cell = ws.cell(row=i, column=j,
                    value=f'=COUNTIFS(\'挽回客户主表\'!E4:E33,"{t}",\'挽回客户主表\'!F4:F33,"{stage}")')
            style_t(cell, fill=bg, bold=True)

        # 挽回率
        cell = ws.cell(row=i, column=9, value=rate)
        style_t(cell, fill=SUCCESS, bold=True)

    # 合计
    i = 12
    style_t(ws.cell(row=i, column=2, value="合计"), fill=PRIMARY, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 9):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"=SUM({col_letter}5:{col_letter}11)")
        style_t(cell, fill=PRIMARY, bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)


def build_sales_analysis(ws):
    ws.title = "销售个人挽回率"
    widths(ws, [4, 10, 14, 14, 14, 14, 14, 14])

    ws.merge_cells("B2:H2")
    ws["B2"] = "9 销售个人挽回数据"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["销售", "失败客户数", "已启动挽回", "签约成功", "挽回失败",
               "挽回率", "签约金额(¥)"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    sales = ["A1", "A2", "A3", "B1", "B2", "B3", "C1", "C2", "C3"]
    for i, name in enumerate(sales, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=name), fill=bg, bold=True)

        # 失败数
        cell = ws.cell(row=i, column=3,
            value=f'=COUNTIF(\'挽回客户主表\'!C4:C33,"{name}")')
        style_t(cell, fill=bg)

        # 已启动
        cell = ws.cell(row=i, column=4,
            value=(f'=COUNTIFS(\'挽回客户主表\'!C4:C33,"{name}",'
                   f'\'挽回客户主表\'!F4:F33,"<>待启动")'
                   f'-COUNTIFS(\'挽回客户主表\'!C4:C33,"{name}",'
                   f'\'挽回客户主表\'!F4:F33,"")'))
        style_t(cell, fill=bg)

        # 签约成功
        cell = ws.cell(row=i, column=5,
            value=f'=COUNTIFS(\'挽回客户主表\'!C4:C33,"{name}",\'挽回客户主表\'!F4:F33,"签约成功")')
        style_t(cell, fill=SUCCESS, bold=True)

        # 挽回失败
        cell = ws.cell(row=i, column=6,
            value=f'=COUNTIFS(\'挽回客户主表\'!C4:C33,"{name}",\'挽回客户主表\'!F4:F33,"挽回失败")')
        style_t(cell, fill=WARN)

        # 挽回率
        cell = ws.cell(row=i, column=7,
            value=f"=IFERROR(E{i}/C{i},0)")
        style_t(cell, fill=SUCCESS, pct=True, bold=True)

        # 签约金额
        cell = ws.cell(row=i, column=8,
            value=f'=SUMIFS(\'挽回客户主表\'!I4:I33,\'挽回客户主表\'!C4:C33,"{name}")')
        style_t(cell, fill=SUCCESS, money=True, bold=True)

    # 合计
    i = 14
    style_t(ws.cell(row=i, column=2, value="合计"), fill=PRIMARY, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in [3, 4, 5, 6, 8]:
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"=SUM({col_letter}5:{col_letter}13)")
        style_t(cell, fill=PRIMARY, money=(col == 8), bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)


def build_monthly_trend(ws):
    ws.title = "月度趋势"
    widths(ws, [4, 22, 14, 14, 14, 14])

    ws.merge_cells("B2:F2")
    ws["B2"] = "月度挽回趋势"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["指标", "M1", "M2", "M3", "3 月累计"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    metrics = [
        "新增失败客户",
        "启动挽回数",
        "联系成功数",
        "POC 启动数",
        "签约成功数",
        "签约金额(¥)",
        "投入成本(¥)",
        "净收益(¥)",
    ]
    for i, m in enumerate(metrics, 5):
        style_t(ws.cell(row=i, column=2, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        for col in range(3, 6):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG)
            if "金额" in m or "成本" in m or "收益" in m:
                ws.cell(row=i, column=col).number_format = '"¥"#,##0'
        # 累计
        cell = ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
        style_t(cell, fill=SUCCESS, bold=True)
        if "金额" in m or "成本" in m or "收益" in m:
            ws.cell(row=i, column=6).number_format = '"¥"#,##0'


def build_roi(ws):
    ws.title = "ROI测算"
    widths(ws, [4, 26, 14, 14, 14, 24])

    ws.merge_cells("B2:F2")
    ws["B2"] = "挽回 ROI 测算"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["指标", "挽回", "新客户", "对比", "结论"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    rows = [
        ("CAC 单客成本(¥)", 9000, 40000, "=D5/C5", "挽回成本 = 1/4"),
        ("平均客单价(¥)", 100000, 100000, 1, "一致"),
        ("转化率", 0.2, 0.07, "=C7/D7", "挽回成功率 = 2.9x"),
        ("ROI", "=C6/C5", "=D6/D5", "=C8/D8", "挽回 ROI = 3-5x"),
    ]
    for i, (label, c, n, comp, conc) in enumerate(rows, 5):
        style_t(ws.cell(row=i, column=2, value=label), fill=LIGHT_BG, align=LEFT, bold=True)
        cell = ws.cell(row=i, column=3, value=c)
        style_t(cell, fill=SUCCESS, bold=True)
        if "成本" in label or "客单" in label:
            cell.number_format = '"¥"#,##0'
        elif "率" in label or "ROI" in label:
            if isinstance(c, (int, float)):
                cell.number_format = "0.0%" if "率" in label else "0.0x"

        cell = ws.cell(row=i, column=4, value=n)
        style_t(cell, fill=LIGHT_BG, bold=True)
        if "成本" in label or "客单" in label:
            cell.number_format = '"¥"#,##0'
        elif "率" in label or "ROI" in label:
            if isinstance(n, (int, float)):
                cell.number_format = "0.0%" if "率" in label else "0.0x"

        cell = ws.cell(row=i, column=5, value=comp)
        style_t(cell, fill=INPUT_BG, bold=True)
        cell.number_format = "0.0x"

        style_t(ws.cell(row=i, column=6, value=conc), align=LEFT, fill=LIGHT_BG)

    # 30 个失败客户的预期收益
    ws.merge_cells("B10:F10")
    ws["B10"] = "▼ 30 个失败客户的预期收益"
    style_t(ws["B10"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B10"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    items = [
        ("30 个失败客户中『启动挽回』数", 25, 0),
        ("预期签约客户（20% 挽回率）", 5, 0),
        ("平均签约金额(¥)", 100000, 0),
        ("预期总签约金额(¥)", "=C12*C13", 0),
        ("投入成本（¥9,000 × 25）", "=25*9000", 0),
        ("净收益(¥)", "=C14-C15", 0),
        ("ROI", "=C14/C15", 0),
    ]
    for i, (label, val, _) in enumerate(items, 11):
        style_t(ws.cell(row=i, column=2, value=label), fill=LIGHT_BG, align=LEFT, bold=True)
        cell = ws.cell(row=i, column=3, value=val)
        style_t(cell, fill=SUCCESS, bold=True)
        if "金额" in label or "成本" in label or "收益" in label:
            cell.number_format = '"¥"#,##0'
        elif "ROI" in label:
            cell.number_format = "0.0x"


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_main(wb.create_sheet())
    build_funnel(wb.create_sheet())
    build_sales_analysis(wb.create_sheet())
    build_monthly_trend(wb.create_sheet())
    build_roi(wb.create_sheet())

    out = "tools/失败客户挽回数据看板.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
