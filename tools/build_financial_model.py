"""
首年财务预测模型生成器
=========================

输出：tools/钢铁数字化系统_首年财务预测模型.xlsx

工作表：
    1. 使用说明
    2. 关键假设          —— 黄色区，老板/财务可调（销售线索/转化率/客单价/人力配置）
    3. 月度漏斗          —— 12 个月线索→签约漏斗自动计算
    4. 月度收入          —— 实施费 + 月租收入 + 续费 + 累计 MRR
    5. 月度成本          —— 人力 + 营销 + 老板局 + 工具 + 应急
    6. 月度现金流        —— 收入 - 成本，含累计现金净额
    7. 损益与回本        —— 首年损益 + 何时单月转正 + 何时累计回本
    8. 三档情景对比      —— 保守 / 基准 / 乐观
"""

from __future__ import annotations
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
SUB_FILL = PatternFill("solid", fgColor="DEEBF7")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
RESULT_FILL = PatternFill("solid", fgColor="E2EFDA")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
HEAD_FONT = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
SUB_FONT = Font(name="Microsoft YaHei", size=11, bold=True)
TEXT_FONT = Font(name="Microsoft YaHei", size=10)
TITLE_FONT = Font(name="Microsoft YaHei", size=16, bold=True, color="1F4E79")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_h(c):
    c.fill = HEAD_FILL; c.font = HEAD_FONT; c.alignment = CENTER; c.border = BORDER


def style_s(c):
    c.fill = SUB_FILL; c.font = SUB_FONT; c.alignment = CENTER; c.border = BORDER


def style_t(c, fill=None, money=False, pct=False, align=None):
    c.font = TEXT_FONT; c.alignment = align or CENTER; c.border = BORDER
    if fill: c.fill = fill
    if money: c.number_format = '"¥"#,##0'
    if pct: c.number_format = "0.0%"


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ----------------------------------------------------------------------
# Sheet 1：使用说明
# ----------------------------------------------------------------------
def build_readme(ws):
    ws.title = "使用说明"
    widths(ws, [4, 28, 70])
    ws.merge_cells("B2:C2")
    ws["B2"] = "钢铁数字化系统 · 首年财务预测模型 v1.0"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    rows = [
        ("适用对象", "老板 / 财务 / 营销负责人 / 投资人"),
        ("使用步骤",
         "① 打开『关键假设』，按黄色单元格调参数（线索量/转化率/客单价/人力等）\n"
         "② 切换『月度漏斗』『月度收入』『月度成本』『月度现金流』查看自动结果\n"
         "③ 『损益与回本』给你 3 个核心数：首年损益、单月转正月份、累计回本月份\n"
         "④ 『三档情景对比』给你保守/基准/乐观对比"),
        ("假设解释",
         "• 线索来源：地推/老板局/合伙人/灯塔厂转介绍/内容营销 5 类\n"
         "• 转化率：MQL→SQL→POC→签约 4 段漏斗\n"
         "• 客单价：种子/经营/旗舰 3 档比例 + 12 期分期模式（默认 70%）\n"
         "• 续费：月租按客户保留率累计 MRR\n"
         "• 成本：人力 + 营销 + 老板局 + 工具 + 应急"),
        ("修改边界", "只改黄色单元格；公式和表头不要动"),
        ("注意",
         "1. 数字仅基于行业基准 + 钢铁 SaaS 经验估算，**不是承诺**\n"
         "2. 实际偏差 ±30% 属正常；偏差 >50% 时及时复盘\n"
         "3. 建议每月用真实数据替换『假设月数据』，做 Forward Forecast"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        ws[f"B{i}"] = k; ws[f"C{i}"] = v
        style_s(ws[f"B{i}"]); style_t(ws[f"C{i}"], align=LEFT)
        ws.row_dimensions[i].height = max(40, len(v) * 0.6)


# ----------------------------------------------------------------------
# Sheet 2：关键假设
# ----------------------------------------------------------------------
def build_assumptions(ws):
    ws.title = "关键假设"
    widths(ws, [4, 30, 18, 18, 50])

    ws.merge_cells("B2:E2")
    ws["B2"] = "关键假设输入区（黄色 = 可调）"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    sections = [
        ("一、线索获取（每月）",
         [
            ("地推线索数（每月）",        80, "条/月", "1-2 个地推人员稳定产出"),
            ("老板局线索数（每月）",      15, "条/月", "1.5 月办 1 场，每场 12 人"),
            ("合伙人推荐线索（每月）",    25, "条/月", "100 个合伙人后稳定"),
            ("灯塔厂转介绍（每月）",      8,  "条/月", "灯塔厂上线后才有"),
            ("内容/线上线索（每月）",     20, "条/月", "公众号/抖音/视频号"),
         ]),
        ("二、漏斗转化率",
         [
            ("线索 → MQL（添加微信）", 0.30, "%", "整体均值"),
            ("MQL → SQL（约访成功）",  0.40, "%", ""),
            ("SQL → POC",              0.50, "%", "上门诊断后启动 POC"),
            ("POC → 签约",             0.55, "%", "POC 成功率"),
         ]),
        ("三、客单价分布",
         [
            ("种子版占比",      0.30, "%", ""),
            ("经营版占比",      0.55, "%", ""),
            ("旗舰版占比",      0.15, "%", ""),
            ("种子版均价(实施)", 30000, "元", "钢管/镀锌/开平/一体加权"),
            ("种子版均价(月租)", 5000,  "元", ""),
            ("经营版均价(实施)", 90000, "元", ""),
            ("经营版均价(月租)", 12000, "元", ""),
            ("旗舰版均价(实施)", 230000, "元", ""),
            ("旗舰版均价(月租)", 35000,  "元", ""),
            ("12 期分期客户占比", 0.70, "%", "选 12 期分期的客户比例"),
            ("年付折扣客户占比",   0.10, "%", "选年付 9 折的比例"),
         ]),
        ("四、续费与流失",
         [
            ("月度客户流失率", 0.015, "%", "1.5% / 月，相当于年流失 16%"),
         ]),
        ("五、人力成本（月薪含五险一金）",
         [
            ("CMO（前 6 月空缺，老板兼）", 0,    "元", "前 6 月不计；7 月起按 40000"),
            ("销售总监 × 2",               2 * 22000, "元", "每人 22000"),
            ("销售 × 2-4",                 4 * 12000, "元", "前 6 月 2 人、7 月起 4 人，简化按 4"),
            ("售前 × 2",                   2 * 14000, "元", ""),
            ("实施顾问 × 2",               2 * 16000, "元", ""),
            ("地推 × 2",                   2 * 7000,  "元", "底薪，提成在动作变量"),
            ("销售运营经理",                12000, "元", ""),
            ("合伙人运营专员",              10000, "元", ""),
            ("市场/内容",                   13000, "元", ""),
            ("行政/商务",                   8000,  "元", ""),
         ]),
        ("六、营销成本（每月）",
         [
            ("内容制作（公众号/视频）",  20000, "元/月", ""),
            ("地推物料 + 提成",          15000, "元/月", ""),
            ("线上投放（抖音/小红书）",  15000, "元/月", "可选"),
            ("白皮书/活动物料",          10000, "元/月", "摊销"),
            ("CRM/工具/服务器",          5000,  "元/月", ""),
         ]),
        ("七、老板局（每场）",
         [
            ("老板局每场预算", 50000, "元/场", "见 docs/05"),
            ("年办场次",        4,    "场/年",  ""),
         ]),
        ("八、灯塔厂补贴（一次性，分摊）",
         [
            ("灯塔厂数量",      4,     "家",     ""),
            ("每家补贴",        80000, "元/家",  "实施费打 5 折等损失"),
            ("摊销月数",        12,    "月",     ""),
         ]),
        ("九、提成（销售 8% 实施费 + 3% 月租 12 月）",
         [
            ("销售实施费提成",   0.08, "%", ""),
            ("销售月租提成",     0.03, "%", "前 12 个月"),
            ("售前实施费提成",   0.03, "%", ""),
            ("售前月租提成",     0.01, "%", "前 6 月"),
         ]),
    ]

    row = 4
    for title, items in sections:
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
        ws.cell(row=row, column=2, value=title).font = SUB_FONT
        ws.cell(row=row, column=2).fill = SUB_FILL
        ws.cell(row=row, column=2).alignment = LEFT
        ws.cell(row=row, column=2).border = BORDER
        row += 1
        for label, val, unit, note in items:
            ws.cell(row=row, column=2, value=label).font = TEXT_FONT
            ws.cell(row=row, column=2).alignment = LEFT
            ws.cell(row=row, column=2).border = BORDER
            c = ws.cell(row=row, column=3, value=val); style_t(c, fill=INPUT_FILL)
            if unit == "元" or unit.startswith("元"):
                c.number_format = '"¥"#,##0'
            elif unit == "%":
                c.number_format = "0.0%"
            elif unit.startswith("条"):
                c.number_format = "0"
            ws.cell(row=row, column=4, value=unit).font = TEXT_FONT
            ws.cell(row=row, column=4).alignment = CENTER
            ws.cell(row=row, column=4).border = BORDER
            ws.cell(row=row, column=5, value=note).font = TEXT_FONT
            ws.cell(row=row, column=5).alignment = LEFT
            ws.cell(row=row, column=5).border = BORDER
            row += 1
        row += 1


# ----------------------------------------------------------------------
# 命名引用辅助（直接用单元格地址）
# ----------------------------------------------------------------------
# 假设关键单元格位置（按 build_assumptions 顺序）：
A = "'关键假设'"
LEAD_COMPONENTS = {
    "ditui":  f"{A}!$C$5",
    "lbj":    f"{A}!$C$6",
    "hhr":    f"{A}!$C$7",
    "dz":     f"{A}!$C$8",
    "content":f"{A}!$C$9",
}
# 转化率
CONV = {
    "lead2mql": f"{A}!$C$12",
    "mql2sql":  f"{A}!$C$13",
    "sql2poc":  f"{A}!$C$14",
    "poc2deal": f"{A}!$C$15",
}
# 客单
PRICE = {
    "seed_pct": f"{A}!$C$18",
    "biz_pct":  f"{A}!$C$19",
    "flag_pct": f"{A}!$C$20",
    "seed_imp": f"{A}!$C$21",
    "seed_mon": f"{A}!$C$22",
    "biz_imp":  f"{A}!$C$23",
    "biz_mon":  f"{A}!$C$24",
    "flag_imp": f"{A}!$C$25",
    "flag_mon": f"{A}!$C$26",
    "instal_pct": f"{A}!$C$27",
    "annual_pct": f"{A}!$C$28",
}
CHURN = f"{A}!$C$31"

# 人力（前 6 月 / 7 月起两套）
HR_CMO_PRE = 0  # 前 6 月 0
HR_CMO_AFTER = 40000

# 简化：人力按"成本基础 + CMO 第 7 月起"计


# ----------------------------------------------------------------------
# Sheet 3：月度漏斗
# ----------------------------------------------------------------------
def build_funnel(ws):
    ws.title = "月度漏斗"
    widths(ws, [22] + [13]*12 + [14])

    ws.merge_cells("A1:N1")
    ws["A1"] = "月度漏斗（自动计算 12 个月）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    months = [f"M{i+1}" for i in range(12)]
    headers = ["指标"] + months + ["全年合计"]
    for j, h in enumerate(headers, 1):
        style_h(ws.cell(row=3, column=j, value=h))

    # 月度线索可用量"曲线"（前期低、后期高）
    # 用乘子来反映团队成长 / 灯塔厂上线 / 合伙人增长
    # M1=0.30, M2=0.45, M3=0.60, M4=0.75, M5=0.85, M6=1.0, M7+=1.0~1.5
    growth = [0.30, 0.45, 0.60, 0.75, 0.85, 1.0, 1.10, 1.20, 1.30, 1.35, 1.40, 1.45]

    # 灯塔厂转介绍只在 M5 之后才有
    dz_ramp = [0, 0, 0, 0, 0.3, 0.5, 0.8, 1.0, 1.0, 1.0, 1.0, 1.0]
    # 合伙人推荐 M3 之后慢慢起
    hhr_ramp = [0.1, 0.2, 0.4, 0.6, 0.7, 0.8, 0.9, 1.0, 1.0, 1.0, 1.0, 1.0]

    row = 4
    # 各类线索
    for label, key, ramp in [
        ("地推线索",     "ditui",   growth),
        ("老板局线索",   "lbj",     [0,0,0.5,0.5,1,1,1,1,1,1,1,1]),
        ("合伙人推荐",   "hhr",     hhr_ramp),
        ("灯塔厂转介绍", "dz",      dz_ramp),
        ("内容/线上",    "content", growth),
    ]:
        ws.cell(row=row, column=1, value=label).font = SUB_FONT
        ws.cell(row=row, column=1).alignment = LEFT
        ws.cell(row=row, column=1).border = BORDER
        for m in range(12):
            cell = ws.cell(row=row, column=2+m,
                           value=f"=ROUND({LEAD_COMPONENTS[key]}*{ramp[m]},0)")
            style_t(cell)
        # 全年合计
        cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
        style_t(cell, fill=RESULT_FILL)
        row += 1

    # 总线索
    ws.cell(row=row, column=1, value="月度总线索").font = SUB_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        cell = ws.cell(row=row, column=2+m, value=f"=SUM({col}4:{col}{row-1})")
        style_t(cell, fill=SUB_FILL)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL)
    total_lead_row = row
    row += 1

    # 漏斗
    funnel_rows = [
        ("MQL（添加微信）", CONV["lead2mql"]),
        ("SQL（约访成功）", CONV["mql2sql"]),
        ("POC 启动",        CONV["sql2poc"]),
        ("签约",            CONV["poc2deal"]),
    ]
    prev_row = total_lead_row
    for label, conv in funnel_rows:
        ws.cell(row=row, column=1, value=label).font = SUB_FONT
        ws.cell(row=row, column=1).alignment = LEFT
        ws.cell(row=row, column=1).border = BORDER
        for m in range(12):
            col = get_column_letter(2+m)
            cell = ws.cell(row=row, column=2+m,
                           value=f"=ROUND({col}{prev_row}*{conv},0)")
            style_t(cell)
        cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
        style_t(cell, fill=RESULT_FILL)
        prev_row = row
        row += 1

    # 标注签约行
    ws.cell(row=prev_row, column=1).fill = SUCCESS = PatternFill("solid", fgColor="E2EFDA")


# ----------------------------------------------------------------------
# Sheet 4：月度收入
# ----------------------------------------------------------------------
def build_revenue(ws):
    ws.title = "月度收入"
    widths(ws, [26] + [13]*12 + [14])

    ws.merge_cells("A1:N1")
    ws["A1"] = "月度收入（自动计算）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    months = [f"M{i+1}" for i in range(12)]
    headers = ["指标"] + months + ["全年合计"]
    for j, h in enumerate(headers, 1):
        style_h(ws.cell(row=3, column=j, value=h))

    # 引用月度签约数（在『月度漏斗』签约行 = 第 10 行 - 但不一定）
    # 我们重新计算：签约数 = 总线索 * 4 段漏斗乘积
    deal_formula = (
        f"='月度漏斗'!B$9*{CONV['lead2mql']}*{CONV['mql2sql']}*{CONV['sql2poc']}*{CONV['poc2deal']}"
    )

    # 取『月度漏斗』签约行（第 13 行：4 类线索 + 总线索 + 4 漏斗）
    # 总线索在第 9 行（4 类 + 1 = 5，从 row 4 起到 8 是各类，第 9 是合计），漏斗 4 行 → 第 13 行是签约
    # build_funnel 中 4 类线索 row 4-8，total 行 9，4 漏斗 10-13；签约在 13

    row = 4
    # 月度签约数
    ws.cell(row=row, column=1, value="月度签约数").font = SUB_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        cell = ws.cell(row=row, column=2+m, value=f"='月度漏斗'!{col}13")
        style_t(cell, fill=SUB_FILL)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL)
    deals_row = row
    row += 1

    # 三档分布
    for label, pct in [("种子版签约", PRICE["seed_pct"]),
                       ("经营版签约", PRICE["biz_pct"]),
                       ("旗舰版签约", PRICE["flag_pct"])]:
        ws.cell(row=row, column=1, value=label).font = TEXT_FONT
        ws.cell(row=row, column=1).alignment = LEFT
        ws.cell(row=row, column=1).border = BORDER
        for m in range(12):
            col = get_column_letter(2+m)
            cell = ws.cell(row=row, column=2+m, value=f"={col}{deals_row}*{pct}")
            style_t(cell)
        cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
        style_t(cell, fill=RESULT_FILL)
        row += 1

    seed_row, biz_row, flag_row = deals_row+1, deals_row+2, deals_row+3

    # 当月新增实施费收入
    # 实施费一次性 → 当月收入：30% 月内回收（首付）；其余在分期模式下逐月回款。
    # 简化：把 100% 实施费视为合同期内入账，但现金按"30% 当月 + 70% 分 12 期"。
    # 计算"当月新签合同的实施费总值（按契约入账）"
    ws.cell(row=row, column=1, value="新增实施费合同额").font = SUB_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        f = (f"={col}{seed_row}*{PRICE['seed_imp']}+"
             f"{col}{biz_row}*{PRICE['biz_imp']}+"
             f"{col}{flag_row}*{PRICE['flag_imp']}")
        cell = ws.cell(row=row, column=2+m, value=f); style_t(cell, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    impl_row = row
    row += 1

    # 当月新增 MRR
    ws.cell(row=row, column=1, value="当月新增 MRR").font = TEXT_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        f = (f"={col}{seed_row}*{PRICE['seed_mon']}+"
             f"{col}{biz_row}*{PRICE['biz_mon']}+"
             f"{col}{flag_row}*{PRICE['flag_mon']}")
        cell = ws.cell(row=row, column=2+m, value=f); style_t(cell, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    new_mrr_row = row
    row += 1

    # 累计 MRR（前期累计 * (1-流失) + 当月新增）
    ws.cell(row=row, column=1, value="累计 MRR（含流失）").font = SUB_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    # M1：等于当月新增
    ws.cell(row=row, column=2, value=f"=B{new_mrr_row}")
    style_t(ws.cell(row=row, column=2), money=True, fill=SUB_FILL)
    for m in range(1, 12):
        col = get_column_letter(2+m)
        prev = get_column_letter(1+m)
        f = f"={prev}{row}*(1-{CHURN})+{col}{new_mrr_row}"
        cell = ws.cell(row=row, column=2+m, value=f)
        style_t(cell, money=True, fill=SUB_FILL)
    # 全年合计无意义
    ws.cell(row=row, column=14, value="—").font = TEXT_FONT
    ws.cell(row=row, column=14).alignment = CENTER
    cum_mrr_row = row
    row += 1

    # 月度月租收入 = 当月累计 MRR
    ws.cell(row=row, column=1, value="月度月租收入").font = TEXT_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        cell = ws.cell(row=row, column=2+m, value=f"={col}{cum_mrr_row}")
        style_t(cell, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    monthly_rev_row = row
    row += 1

    # 月度实施费现金回收：30% 当月 + 70% 分 12 期
    # 简化：实施费 100% 当月+次月按比例进入。本模型按"30% 当月 + 50% 次月（上线款） + 20% 验收（按 +3 月）"
    ws.cell(row=row, column=1, value="月度实施费现金回收").font = TEXT_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        # 当月：30% 本月 + 50% 上一月（如果存在）+ 20% 三个月前（如果存在）
        parts = [f"{col}{impl_row}*0.3"]
        if m >= 1:
            parts.append(f"{get_column_letter(1+m)}{impl_row}*0.5")
        if m >= 3:
            parts.append(f"{get_column_letter(m-1)}{impl_row}*0.2")
        f = "=" + "+".join(parts)
        cell = ws.cell(row=row, column=2+m, value=f); style_t(cell, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    impl_cash_row = row
    row += 1

    # 月度总现金收入
    ws.cell(row=row, column=1, value="月度总现金收入").font = SUB_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        cell = ws.cell(row=row, column=2+m,
                       value=f"={col}{monthly_rev_row}+{col}{impl_cash_row}")
        style_t(cell, fill=RESULT_FILL, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)


SUCCESS = PatternFill("solid", fgColor="E2EFDA")


# ----------------------------------------------------------------------
# Sheet 5：月度成本
# ----------------------------------------------------------------------
def build_cost(ws):
    ws.title = "月度成本"
    widths(ws, [26] + [13]*12 + [14])

    ws.merge_cells("A1:N1")
    ws["A1"] = "月度成本（自动计算）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    months = [f"M{i+1}" for i in range(12)]
    headers = ["指标"] + months + ["全年合计"]
    for j, h in enumerate(headers, 1):
        style_h(ws.cell(row=3, column=j, value=h))

    row = 4
    # 人力成本（前 6 月 vs. 7 月起）
    ws.cell(row=row, column=1, value="人力成本").font = SUB_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER

    # 引用关键假设里 SUM(C32:C41)+CMO（C32 是 CMO=0，C32-C41 共 10 项；CMO 是 C32 已经是 0）
    # 简化：人力 = SUM(C32:C41) + IF(M>=7, 40000, 0)
    base_hr = f"SUM({A}!$C$35:$C$43)"  # CMO=C34（前6=0），其他 9 项 C35:C43
    for m in range(12):
        col = get_column_letter(2+m)
        if m >= 6:
            f = f"={base_hr} + 40000"  # 第 7 月起 CMO 上岗
        else:
            f = f"={base_hr}"
        cell = ws.cell(row=row, column=2+m, value=f); style_t(cell, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    hr_row = row
    row += 1

    # 营销成本：内容+地推+投放+活动物料+工具
    ws.cell(row=row, column=1, value="营销固定成本").font = TEXT_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    mk_base = f"SUM({A}!$C$46:$C$50)"
    for m in range(12):
        cell = ws.cell(row=row, column=2+m, value=f"={mk_base}")
        style_t(cell, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    row += 1

    # 老板局：每年 4 场分摊到 M3、M6、M9、M12
    ws.cell(row=row, column=1, value="老板局费用").font = TEXT_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        if m+1 in [3, 6, 9, 12]:
            f = f"={A}!$C$53"
        else:
            f = "=0"
        cell = ws.cell(row=row, column=2+m, value=f); style_t(cell, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    row += 1

    # 灯塔厂补贴（每月分摊）
    ws.cell(row=row, column=1, value="灯塔厂补贴(摊销)").font = TEXT_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        f = f"={A}!$C$57*{A}!$C$58/{A}!$C$59"
        cell = ws.cell(row=row, column=2+m, value=f); style_t(cell, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    row += 1

    # 销售提成（基于当月新增合同 + 累计 MRR）
    ws.cell(row=row, column=1, value="销售/售前提成").font = TEXT_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        # 实施费提成 + 月租提成（按累计 MRR）
        # 销售实施 8% + 月租 3%；售前 3% + 1% 简化合一
        f = (f"='月度收入'!{col}8*({A}!$C$62+{A}!$C$64) + "
             f"'月度收入'!{col}11*({A}!$C$63+{A}!$C$65)")
        cell = ws.cell(row=row, column=2+m, value=f); style_t(cell, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    row += 1

    # 应急 / 杂费（每月固定 5000）
    ws.cell(row=row, column=1, value="应急/杂费").font = TEXT_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        cell = ws.cell(row=row, column=2+m, value=5000); style_t(cell, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    row += 1

    # 总成本
    ws.cell(row=row, column=1, value="月度总成本").font = SUB_FONT
    ws.cell(row=row, column=1).fill = SUB_FILL
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        cell = ws.cell(row=row, column=2+m, value=f"=SUM({col}4:{col}{row-1})")
        style_t(cell, fill=RESULT_FILL, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)


# ----------------------------------------------------------------------
# Sheet 6：月度现金流
# ----------------------------------------------------------------------
def build_cashflow(ws):
    ws.title = "月度现金流"
    widths(ws, [26] + [13]*12 + [14])

    ws.merge_cells("A1:N1")
    ws["A1"] = "月度现金流（自动计算）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    months = [f"M{i+1}" for i in range(12)]
    headers = ["指标"] + months + ["全年合计"]
    for j, h in enumerate(headers, 1):
        style_h(ws.cell(row=3, column=j, value=h))

    # 月度收入：『月度收入』总现金行（最后一行 = 第 12 行）
    row = 4
    ws.cell(row=row, column=1, value="月度现金收入").font = SUB_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        cell = ws.cell(row=row, column=2+m, value=f"='月度收入'!{col}13")
        style_t(cell, fill=SUB_FILL, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    row += 1

    # 月度成本：『月度成本』最后一行 = 第 10 行
    ws.cell(row=row, column=1, value="月度总成本").font = SUB_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        cell = ws.cell(row=row, column=2+m, value=f"='月度成本'!{col}10")
        style_t(cell, fill=WARN_FILL, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    row += 1

    # 月度净现金
    ws.cell(row=row, column=1, value="月度净现金").font = SUB_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        cell = ws.cell(row=row, column=2+m, value=f"={col}4-{col}5")
        style_t(cell, fill=RESULT_FILL, money=True)
    cell = ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})")
    style_t(cell, fill=RESULT_FILL, money=True)
    net_row = row
    row += 1

    # 累计现金净额
    ws.cell(row=row, column=1, value="累计现金净额").font = SUB_FONT
    ws.cell(row=row, column=1).alignment = LEFT
    ws.cell(row=row, column=1).border = BORDER
    for m in range(12):
        col = get_column_letter(2+m)
        if m == 0:
            f = f"={col}{net_row}"
        else:
            prev = get_column_letter(1+m)
            f = f"={prev}{row}+{col}{net_row}"
        cell = ws.cell(row=row, column=2+m, value=f)
        style_t(cell, fill=RESULT_FILL, money=True)


# ----------------------------------------------------------------------
# Sheet 7：损益与回本
# ----------------------------------------------------------------------
def build_pnl(ws):
    ws.title = "损益与回本"
    widths(ws, [4, 32, 24, 60])

    ws.merge_cells("B2:D2")
    ws["B2"] = "首年损益与回本分析"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    rows = [
        ("全年现金收入合计", "='月度现金流'!N4", "实施费现金 + 月租收入"),
        ("全年现金成本合计", "='月度现金流'!N5", "人力 + 营销 + 老板局 + 提成 + 应急"),
        ("全年净现金（首年损益）", "='月度现金流'!N6", "正 = 自我造血；负 = 需要补贴"),
        ("第 12 月累计 MRR", "='月度收入'!M11", "月经常性收入，是公司估值的关键"),
        ("年化 MRR（×12）", "='月度收入'!M11*12", "ARR"),
        ("第 12 月累计签约客户数",
         "=SUM('月度漏斗'!B13:M13)",
         "首年签约总数"),
        ("单月净现金转正月份",
         '=IFERROR(MATCH(TRUE,(\'月度现金流\'!B6:M6>0),0),"未转正")',
         "首次单月现金为正的月份（数字 1-12）"),
        ("累计现金净额转正月份",
         '=IFERROR(MATCH(TRUE,(\'月度现金流\'!B7:M7>0),0),"未在首年回本")',
         "首次累计现金为正的月份"),
    ]
    for i, (lbl, formula, note) in enumerate(rows, start=4):
        ws.cell(row=i, column=2, value=lbl).font = SUB_FONT
        ws.cell(row=i, column=2).alignment = LEFT
        ws.cell(row=i, column=2).border = BORDER
        if "月份" in lbl:
            cell = ws.cell(row=i, column=3, value=f"{{{formula}}}")
            cell.value = formula
            style_t(cell, fill=RESULT_FILL)
        else:
            cell = ws.cell(row=i, column=3, value=formula)
            style_t(cell, fill=RESULT_FILL, money=True)
        ws.cell(row=i, column=4, value=note).font = TEXT_FONT
        ws.cell(row=i, column=4).alignment = LEFT
        ws.cell(row=i, column=4).border = BORDER

    # 备注
    ws.merge_cells("B14:D18")
    ws["B14"] = (
        "解读建议：\n"
        "• 首年现金净额为负属正常 —— B 端 SaaS 第 1 年通常烧钱\n"
        "• 重点看「累计 MRR」和「ARR」 —— 这两个数字决定了第 2 年起的爆发能力\n"
        "• 若单月净现金转正在 M9 之前 → 极优秀\n"
        "• 若 M12 仍未转正 → 检查转化率/客单价/人力规模是否合理\n"
        "• 累计现金回本通常发生在第 13-18 月，属于行业常态"
    )
    ws["B14"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["B14"].font = Font(name="Microsoft YaHei", size=10, color="595959")
    ws["B14"].border = BORDER


# ----------------------------------------------------------------------
# Sheet 8：三档情景对比
# ----------------------------------------------------------------------
def build_scenarios(ws):
    ws.title = "三档情景对比"
    widths(ws, [4, 28, 18, 18, 18, 40])

    ws.merge_cells("B2:F2")
    ws["B2"] = "保守 / 基准 / 乐观 三档情景对比（手工调假设页参数验证）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    # 这里只是文字说明 + 基准引用
    headers = ["情景", "线索量调整", "转化率调整", "首年净现金（基准）", "操作建议"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    rows = [
        ("保守", "× 0.7", "× 0.7", "（手动验证）",
         "把『关键假设』里地推/老板局/合伙人/内容线索数 ×0.7，再把 4 段转化率 ×0.7"),
        ("基准", "× 1.0", "× 1.0", "='损益与回本'!C6",
         "默认参数（建议起步用此情景）"),
        ("乐观", "× 1.3", "× 1.2", "（手动验证）",
         "把『关键假设』里所有线索数 ×1.3，转化率 ×1.2，目标超额完成"),
    ]
    for i, row in enumerate(rows, start=5):
        for j, v in enumerate(row, start=2):
            cell = ws.cell(row=i, column=j, value=v)
            if j in (3, 4) and isinstance(v, str) and v.startswith("×"):
                style_t(cell, fill=INPUT_FILL)
            elif isinstance(v, str) and v.startswith("="):
                style_t(cell, fill=RESULT_FILL, money=True)
            elif j == 6:
                style_t(cell, align=LEFT)
            else:
                style_t(cell)
        ws.cell(row=i, column=2).font = SUB_FONT

    ws.merge_cells("B10:F12")
    ws["B10"] = (
        "使用建议：\n"
        "1. 用基准跑一次完整结果，记录关键数字\n"
        "2. 切换到保守 / 乐观，看现金流压力区间\n"
        "3. 决策时按『保守 = 公司能扛』标准调整融资 / 现金储备"
    )
    ws["B10"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["B10"].font = Font(name="Microsoft YaHei", size=10, color="595959")
    ws["B10"].border = BORDER


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_assumptions(wb.create_sheet())
    build_funnel(wb.create_sheet())
    build_revenue(wb.create_sheet())
    build_cost(wb.create_sheet())
    build_cashflow(wb.create_sheet())
    build_pnl(wb.create_sheet())
    build_scenarios(wb.create_sheet())

    out = "tools/钢铁数字化系统_首年财务预测模型.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
