"""
首年财务预测模型 Plan A 压缩版
==============================

继承 build_financial_model.py 的全部结构，
只调整『关键假设』中的参数为 Plan A 节奏（更激进的线索/转化/人力）。

输出：tools/钢铁数字化系统_首年财务预测模型_PlanA.xlsx

调整对照（基准 vs. Plan A）：
- 地推线索: 80 → 100
- 老板局线索: 15 → 25
- 合伙人推荐: 25 → 40
- 灯塔厂转介绍: 8 → 15
- 内容/线上: 20 → 30
- MQL→SQL: 0.4 → 0.45
- 销售总监 × 2: 44000 → 50000
- 销售: 4×12000 → 5×13000 = 65000
- 内容制作: 20000 → 30000
- 老板局每场: 50000 → 60000
- 老板局年场次: 4 → 6
"""

import sys
sys.path.insert(0, "tools")

# 直接 import 原模块所有结构和函数
from build_financial_model import (
    Workbook, Font, PatternFill, Alignment, Border, Side, get_column_letter,
    DataValidation,
    THIN, BORDER, HEAD_FILL, SUB_FILL, INPUT_FILL, RESULT_FILL, WARN_FILL,
    HEAD_FONT, SUB_FONT, TEXT_FONT, TITLE_FONT, CENTER, LEFT,
    style_h, style_s, style_t, widths,
    build_readme as _orig_readme,
    build_funnel, build_revenue, build_cost, build_cashflow, build_pnl, build_scenarios,
    A,
)


def build_readme_plana(ws):
    ws.title = "使用说明"
    widths(ws, [4, 28, 70])
    ws.merge_cells("B2:C2")
    ws["B2"] = "钢铁数字化系统 · 首年财务预测模型 · Plan A 压缩版"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    rows = [
        ("版本说明",
         "Plan A = 21 个月压缩节奏（窗口紧张应对）\n"
         "对应 docs/14 时间窗口复盘文档；与基准版对比线索/转化/人力均上调"),
        ("适用对象", "老板 / 财务 / 营销负责人 / 投资人"),
        ("使用步骤",
         "① 打开『关键假设』，按黄色单元格调参数\n"
         "② 切换『月度漏斗』『月度收入』等 sheet 看自动结果\n"
         "③ 『损益与回本』给你 3 个核心数字\n"
         "④ 与基准版对比 → 看 Plan A 真实代价 + 真实回报"),
        ("Plan A 预期",
         "• 全年签约: 70-90 单（基准 54）\n"
         "• 12 月末 ARR: ¥1,000-1,400 万（基准 ¥813 万）\n"
         "• M3 单月转正（基准 M4）\n"
         "• M5-M6 累计回正（基准 M7）\n"
         "• 全年净现金: +¥450-600 万（基准 +¥310 万）"),
        ("Plan A 代价",
         "• 月度成本上涨 ¥3-5 万（人力 + 内容 + 老板局）\n"
         "• 老板自身投入 70%+ 时间\n"
         "• 6 个月生死线：见 docs/14"),
        ("注意",
         "数字基于行业基准 + 钢铁 SaaS 经验估算，不是承诺；"
         "实际偏差 ±30% 属正常；偏差 >50% 时及时复盘"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        ws[f"B{i}"] = k; ws[f"C{i}"] = v
        style_s(ws[f"B{i}"]); style_t(ws[f"C{i}"], align=LEFT)
        ws.row_dimensions[i].height = max(40, len(v) * 0.6)


def build_assumptions_plana(ws):
    """Plan A 调整版假设页：参数更激进"""
    ws.title = "关键假设"
    widths(ws, [4, 30, 18, 18, 50])

    ws.merge_cells("B2:E2")
    ws["B2"] = "关键假设输入区（Plan A 压缩版 · 黄色 = 可调）"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    sections = [
        ("一、线索获取（每月）·【Plan A 调整】",
         [
            ("地推线索数（每月）",        100, "条/月", "Plan A: 100（基准 80），加大地推 + 1 人"),
            ("老板局线索数（每月）",      25,  "条/月", "Plan A: 25（基准 15），2 月 1 场，每场 12 人"),
            ("合伙人推荐线索（每月）",    40,  "条/月", "Plan A: 40（基准 25），第 2 月立刻发动"),
            ("灯塔厂转介绍（每月）",      15,  "条/月", "Plan A: 15（基准 8），第 4 月起就有"),
            ("内容/线上线索（每月）",     30,  "条/月", "Plan A: 30（基准 20），内容投入加倍"),
         ]),
        ("二、漏斗转化率·【Plan A 调整】",
         [
            ("线索 → MQL（添加微信）", 0.30, "%", ""),
            ("MQL → SQL（约访成功）",  0.45, "%", "Plan A: 45%（基准 40%），老板局/灯塔厂助攻"),
            ("SQL → POC",              0.50, "%", ""),
            ("POC → 签约",             0.55, "%", ""),
         ]),
        ("三、客单价分布",
         [
            ("种子版占比",      0.30, "%", ""),
            ("经营版占比",      0.55, "%", ""),
            ("旗舰版占比",      0.15, "%", ""),
            ("种子版均价(实施)", 30000, "元", ""),
            ("种子版均价(月租)", 5000,  "元", ""),
            ("经营版均价(实施)", 90000, "元", ""),
            ("经营版均价(月租)", 12000, "元", ""),
            ("旗舰版均价(实施)", 230000, "元", ""),
            ("旗舰版均价(月租)", 35000,  "元", ""),
            ("12 期分期客户占比", 0.70, "%", ""),
            ("年付折扣客户占比",   0.10, "%", ""),
         ]),
        ("四、续费与流失",
         [
            ("月度客户流失率", 0.015, "%", "1.5%/月，年流失 16%"),
         ]),
        ("五、人力成本（月薪）·【Plan A 调整】",
         [
            ("CMO（前 6 月空缺，老板兼）", 0,    "元", ""),
            ("销售总监 × 2",               2 * 25000, "元", "Plan A: 25000/人（基准 22000），加薪锁定"),
            ("销售 × 5",                   5 * 13000, "元", "Plan A: 5 人 × 13000（基准 4×12000）"),
            ("售前 × 2",                   2 * 14000, "元", ""),
            ("实施顾问 × 2",               2 * 16000, "元", ""),
            ("地推 × 2",                   2 * 7000,  "元", ""),
            ("销售运营经理",                12000, "元", ""),
            ("合伙人运营专员",              10000, "元", ""),
            ("市场/内容",                   13000, "元", ""),
            ("行政/商务",                   8000,  "元", ""),
         ]),
        ("六、营销成本（每月）·【Plan A 调整】",
         [
            ("内容制作（公众号/视频）",  30000, "元/月", "Plan A: 30000（基准 20000），内容加倍"),
            ("地推物料 + 提成",          18000, "元/月", "Plan A: 18000（基准 15000）"),
            ("线上投放（抖音/小红书）",  20000, "元/月", "Plan A: 20000（基准 15000）"),
            ("白皮书/活动物料",          12000, "元/月", "Plan A: 12000（基准 10000）"),
            ("CRM/工具/服务器",          5000,  "元/月", ""),
         ]),
        ("七、老板局（每场）·【Plan A 调整】",
         [
            ("老板局每场预算", 60000, "元/场", "Plan A: 60000（基准 50000），高质量"),
            ("年办场次",        6,    "场/年",  "Plan A: 6 场（基准 4），每 2 月 1 场"),
         ]),
        ("八、灯塔厂补贴（一次性，分摊）",
         [
            ("灯塔厂数量",      4,     "家",     ""),
            ("每家补贴",        80000, "元/家",  ""),
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
        c = ws.cell(row=row, column=2, value=title)
        c.font = SUB_FONT; c.fill = SUB_FILL; c.alignment = LEFT; c.border = BORDER
        row += 1
        for label, val, unit, note in items:
            ws.cell(row=row, column=2, value=label).font = TEXT_FONT
            ws.cell(row=row, column=2).alignment = LEFT
            ws.cell(row=row, column=2).border = BORDER
            cc = ws.cell(row=row, column=3, value=val); style_t(cc, fill=INPUT_FILL)
            if unit.startswith("元"): cc.number_format = '"¥"#,##0'
            elif unit == "%": cc.number_format = "0.0%"
            elif unit.startswith("条"): cc.number_format = "0"
            ws.cell(row=row, column=4, value=unit).font = TEXT_FONT
            ws.cell(row=row, column=4).alignment = CENTER
            ws.cell(row=row, column=4).border = BORDER
            ws.cell(row=row, column=5, value=note).font = TEXT_FONT
            ws.cell(row=row, column=5).alignment = LEFT
            ws.cell(row=row, column=5).border = BORDER
            row += 1
        row += 1


def main():
    wb = Workbook()
    build_readme_plana(wb.active)
    build_assumptions_plana(wb.create_sheet())
    build_funnel(wb.create_sheet())
    build_revenue(wb.create_sheet())
    build_cost(wb.create_sheet())
    build_cashflow(wb.create_sheet())
    build_pnl(wb.create_sheet())
    build_scenarios(wb.create_sheet())

    out = "tools/钢铁数字化系统_首年财务预测模型_PlanA.xlsx"
    wb.save(out)
    print(f"已生成 Plan A 压缩版：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
