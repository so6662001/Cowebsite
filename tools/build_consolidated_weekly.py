"""
3 业务统一周战报合并模板
=========================

输出：tools/3业务统一周战报.xlsx

工作表：
    1. 使用说明
    2. 集团周战报（一页全览）
    3. 业务 1 钢厂 SaaS 战报
    4. 业务 2 钢贸 + 平台 战报
    5. 业务 3 信用数据 + 增值 战报
    6. 跨业务协同跟踪
    7. 13 周累计趋势
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
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
    ws["B2"] = "3 业务统一周战报合并模板"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    rows = [
        ("适用对象", "CEO + COO + 3 业务 VP + 销售总监 + CS 总监 + 财务"),
        ("使用步骤",
         "① 每周日晚 3 业务总监各自填业务战报\n"
         "② 业务运营汇总到『集团周战报』\n"
         "③ 周一 9:00 集团周例会 30 分钟过完\n"
         "④ COO 跟进跨业务协同事项"),
        ("会议结构",
         "9:00-9:10 集团一页全览（CEO 主持）\n"
         "9:10-9:25 3 业务 VP 各 5 分钟（关键数据 + 1 故事 + 1 决策）\n"
         "9:25-9:30 跨业务协同 + 决策"),
        ("关键铁律",
         "1. 业务间数据可比（KPI 口径统一）\n"
         "2. 协同事项必须有 RACI（详见 docs/49）\n"
         "3. 决策当场拍板，不留尾巴"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        c1 = ws.cell(row=i, column=2, value=k); style_t(c1, fill=PRIMARY, bold=True)
        c1.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
        c2 = ws.cell(row=i, column=3, value=v); style_t(c2, align=LEFT)
        ws.row_dimensions[i].height = max(50, len(v) * 0.7)


def build_group_dashboard(ws):
    ws.title = "集团周战报"
    widths(ws, [4, 18, 14, 14, 14, 14, 14])

    ws.merge_cells("B2:G2")
    ws["B2"] = "集团周战报（一页全览）第 ___ 期 / 日期 ___"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    # 3 业务汇总
    headers = ["业务", "本周收入(¥)", "周环比", "签约客户", "活跃客户", "状态"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    businesses = [
        ("业务 1 钢厂 SaaS", "='业务1钢厂'!H6"),
        ("业务 2 钢贸 + 平台", "='业务2钢贸'!H6"),
        ("业务 3 信用数据", "='业务3数据'!H6"),
    ]
    for i, (name, formula) in enumerate(businesses, 5):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT, bold=True)
        cell = ws.cell(row=i, column=3, value=formula)
        style_t(cell, fill=SUCCESS, money=True, bold=True)
        # 周环比 / 签约 / 活跃 / 状态
        for col in range(4, 8):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG)

    # 合计
    i = 8
    style_t(ws.cell(row=i, column=2, value="3 业务合计"), fill=PRIMARY, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=i, column=3, value=f"=SUM(C5:C7)")
    style_t(cell, fill=PRIMARY, money=True, bold=True)
    ws.cell(row=i, column=3).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 关键决策（本周）
    ws.merge_cells("B10:G10")
    style_t(ws.cell(row=10, column=2, value="▼ 本周关键决策（CEO 拍板）"),
            fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=10, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    for i in range(11, 14):
        style_t(ws.cell(row=i, column=2, value=f"决策 {i-10}："), fill=LIGHT_BG, bold=True)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=7)
        style_t(ws.cell(row=i, column=3), fill=INPUT_BG, align=LEFT)

    # 跨业务协同
    ws.merge_cells("B15:G15")
    style_t(ws.cell(row=15, column=2, value="▼ 跨业务协同（本周）"),
            fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=15, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    coop_headers = ["协同项", "发起方", "接收方", "状态", "金额(¥)", "备注"]
    for j, h in enumerate(coop_headers, 2):
        style_h(ws.cell(row=16, column=j, value=h))
    for i in range(17, 22):
        for col in range(2, 8):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG, align=LEFT)


def build_business_template(ws, name, prefix):
    ws.title = name
    widths(ws, [4, 22, 12, 12, 12, 12, 12, 14])

    ws.merge_cells("B2:H2")
    ws["B2"] = f"{name} - 周战报"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    # 期号 + 日期
    style_t(ws.cell(row=4, column=2, value="期号 / 日期"), fill=LIGHT_BG, bold=True)
    ws.merge_cells("C4:H4")
    style_t(ws.cell(row=4, column=3, value=""), fill=INPUT_BG)

    # 5 个核心数字
    headers = ["指标", "上周实际", "上周目标", "达成率", "累计本月", "月目标", "趋势"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=5, column=j, value=h))

    # 5 个 KPI（业务定制）
    if "钢厂" in name:
        metrics = ["新增 MES POC", "新增钢贸 ERP 签约", "新增 WMS 签约",
                   "新增 AI 客户", "本周回款(¥)"]
    elif "钢贸" in name:
        metrics = ["新增钢贸 ERP 签约", "新增 AI 客户", "平台月活买家",
                   "平台月活卖家", "本周回款(¥)"]
    else:  # 数据
        metrics = ["新增银行客户", "新增保理客户", "数据 API 调用量",
                   "撮合 GMV(¥)", "本周回款(¥)"]

    for i, m in enumerate(metrics, 6):
        style_t(ws.cell(row=i, column=2, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        for col in range(3, 9):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG)

    # 3 故事
    ws.merge_cells("B12:H12")
    style_t(ws.cell(row=12, column=2, value="▼ 3 故事（亮点 / 失败 / 学习）"),
            fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=12, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    stories = ["亮点客户", "失败客户", "学到的事"]
    for i, s in enumerate(stories, 13):
        style_t(ws.cell(row=i, column=2, value=s), fill=LIGHT_BG, align=LEFT, bold=True)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=8)
        style_t(ws.cell(row=i, column=3), fill=INPUT_BG, align=LEFT)
        ws.row_dimensions[i].height = 40

    # 决策 + 风险
    ws.merge_cells("B17:H17")
    style_t(ws.cell(row=17, column=2, value="▼ 本周决策 + 风险"),
            fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=17, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    for i in range(18, 21):
        style_t(ws.cell(row=i, column=2, value=f"项 {i-17}："), fill=LIGHT_BG, bold=True)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=8)
        style_t(ws.cell(row=i, column=3), fill=INPUT_BG, align=LEFT)

    # 跨业务协同请求
    ws.merge_cells("B22:H22")
    style_t(ws.cell(row=22, column=2, value="▼ 需要跨业务协同的事"),
            fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=22, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for i in range(23, 26):
        style_t(ws.cell(row=i, column=2, value=f"项 {i-22}："), fill=LIGHT_BG, bold=True)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=8)
        style_t(ws.cell(row=i, column=3), fill=INPUT_BG, align=LEFT)


def build_synergy(ws):
    ws.title = "跨业务协同跟踪"
    widths(ws, [4, 8, 22, 14, 14, 14, 14, 14, 22])

    ws.merge_cells("B2:I2")
    ws["B2"] = "跨业务协同跟踪表"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    headers = ["#", "协同项", "发起方", "接收方", "金额(¥)", "状态", "RACI 责任人", "备注"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    # 30 行
    for i in range(5, 35):
        style_t(ws.cell(row=i, column=2, value=i-4), fill=LIGHT_BG, bold=True)
        for col in range(3, 10):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG, align=LEFT)

    from openpyxl.worksheet.datavalidation import DataValidation
    dv_status = DataValidation(type="list",
        formula1='"待启动,进行中,已完成,搁置"', allow_blank=True)
    ws.add_data_validation(dv_status); dv_status.add("G5:G34")

    for status, color in [("已完成", "00B050"), ("进行中", "FFC000"),
                          ("搁置", "FF6B6B")]:
        ws.conditional_formatting.add("G5:G34",
            CellIsRule(operator="equal", formula=[f'"{status}"'],
                       fill=PatternFill("solid", fgColor=color)))


def build_trend(ws):
    ws.title = "13周累计趋势"
    widths(ws, [22] + [10]*13 + [12])

    ws.merge_cells("A1:O1")
    ws["A1"] = "13 周累计趋势（3 业务总收入）"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = LEFT

    style_h(ws.cell(row=3, column=1, value="业务"))
    for i in range(13):
        style_h(ws.cell(row=3, column=2+i, value=f"W{i+1}"))
    style_h(ws.cell(row=3, column=15, value="累计"))

    businesses = ["业务 1 钢厂", "业务 2 钢贸", "业务 3 数据", "3 业务合计"]
    for i, name in enumerate(businesses, 4):
        is_total = "合计" in name
        style_t(ws.cell(row=i, column=1, value=name),
                fill=PRIMARY if is_total else LIGHT_BG,
                align=LEFT, bold=True)
        if is_total:
            ws.cell(row=i, column=1).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

        for w in range(13):
            if is_total:
                col = get_column_letter(2+w)
                cell = ws.cell(row=i, column=2+w, value=f"=SUM({col}4:{col}{i-1})")
                style_t(cell, fill=SUCCESS, money=True, bold=True)
            else:
                style_t(ws.cell(row=i, column=2+w), fill=INPUT_BG, money=True)

        # 累计
        col_start = get_column_letter(2)
        col_end = get_column_letter(14)
        cell = ws.cell(row=i, column=15,
                       value=f"=SUM({col_start}{i}:{col_end}{i})")
        style_t(cell, fill=PRIMARY if is_total else SUCCESS, money=True, bold=True)
        if is_total:
            ws.cell(row=i, column=15).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_group_dashboard(wb.create_sheet())
    build_business_template(wb.create_sheet(), "业务1钢厂", "B1")
    build_business_template(wb.create_sheet(), "业务2钢贸", "B2")
    build_business_template(wb.create_sheet(), "业务3数据", "B3")
    build_synergy(wb.create_sheet())
    build_trend(wb.create_sheet())

    out = "tools/3业务统一周战报.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
