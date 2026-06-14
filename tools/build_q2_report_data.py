"""生成 2026 Q2 钢贸行业健康度报告配套数据集 Excel。

输出：tools/2026Q2钢贸健康度数据集.xlsx
工作表：
  0 报告概览 + 防伪
  1 总体分布（5 期对比）
  2 7 模块详解
  3 30 题完整数据
  4 区域对比
  5 规模对比
  6 业态对比
  7 6 个月预测（迁移矩阵）
  8 行业基准对照表
  9 政策环境
"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

THIN = Side(style="thin", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

NAVY = PatternFill("solid", fgColor="1F4E78")
RED = PatternFill("solid", fgColor="C00000")
ORANGE = PatternFill("solid", fgColor="ED7D31")
GREEN = PatternFill("solid", fgColor="00B050")
GOLD = PatternFill("solid", fgColor="FFC000")
LIGHT_GRAY = PatternFill("solid", fgColor="F2F2F2")
LIGHT_BLUE = PatternFill("solid", fgColor="D9E1F2")
LIGHT_YELLOW = PatternFill("solid", fgColor="FFF2CC")
LIGHT_GREEN = PatternFill("solid", fgColor="E2EFDA")
LIGHT_RED = PatternFill("solid", fgColor="FAE5E5")

WHITE_BOLD = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
BOLD = Font(name="Microsoft YaHei", size=11, bold=True)
NORMAL = Font(name="Microsoft YaHei", size=10)
SMALL = Font(name="Microsoft YaHei", size=9, color="555555")
TITLE = Font(name="Microsoft YaHei", size=16, bold=True, color="1F4E78")


def setw(ws: Worksheet, widths: list[float]) -> None:
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def title(ws: Worksheet, row: int, col: int, span: int, text: str, fill=NAVY, font=WHITE_BOLD, height=30) -> None:
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    c = ws.cell(row=row, column=col, value=text)
    c.fill = fill; c.font = font; c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[row].height = height


def row_vals(ws: Worksheet, row: int, values: list, fill=None, font=None, align=None) -> None:
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=i + 1, value=v)
        c.border = BORDER
        c.font = font or NORMAL
        if fill: c.fill = fill
        c.alignment = align or CENTER


def header_row(ws: Worksheet, row: int, headers: list, fill=NAVY, font=WHITE_BOLD) -> None:
    row_vals(ws, row, headers, fill=fill, font=font)


# -------------------------------------------------------------------
# Tab 0 报告概览 + 防伪
# -------------------------------------------------------------------
def build_overview(wb: Workbook) -> None:
    ws = wb.create_sheet("0 报告概览")
    setw(ws, [4, 22, 36, 22, 22])

    title(ws, 1, 1, 5, "2026 Q2 中国钢贸行业健康度研究报告 · 数据集")
    title(ws, 2, 1, 5, "钢联数字化研究院 · 基于 1,547 家样本", fill=LIGHT_BLUE, font=BOLD, height=22)

    r = 4
    title(ws, r, 1, 5, "报告基本信息", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    info = [
        ("报告编号", "2026-Q2-CN-1547"),
        ("发布日期", "2026-07-08"),
        ("数据周期", "2026-04-01 至 2026-06-30"),
        ("样本量", "1,547 家钢贸 / 钢厂企业"),
        ("覆盖", "14 省 / 22 城 / 5 业态"),
        ("样本质量", "自评 70% + 销售陪问 30% / 25% ERP 交叉验证"),
        ("置信度", "≥ 95%（整体）/ ≥ 90%（分组）"),
        ("发布机构", "钢联数字化研究院"),
        ("合作伙伴", "上海钢贸协会 / 中物联钢铁专委会"),
        ("防伪编号", "____ / 订阅人：____"),
    ]
    for k, v in info:
        ws.cell(row=r, column=2, value=k).font = BOLD
        ws.cell(row=r, column=2).fill = LIGHT_GRAY
        ws.cell(row=r, column=2).alignment = LEFT
        ws.cell(row=r, column=2).border = BORDER
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        c = ws.cell(row=r, column=3, value=v)
        c.font = NORMAL; c.alignment = LEFT; c.border = BORDER
        r += 1

    r += 1
    title(ws, r, 1, 5, "10 个工作表导航", fill=LIGHT_BLUE, font=BOLD, height=22)
    r += 1
    header_row(ws, r, ["#", "工作表", "内容", "用途", "—"])
    r += 1
    tabs = [
        ("0", "0 报告概览", "报告基本信息 + 防伪", "您正在看的页", "—"),
        ("1", "1 总体分布", "5 期分级分布 / 趋势 / 与其他行业对比 / 健康度与利润相关性", "看大盘", "—"),
        ("2", "2 7模块详解", "M1-M7 各模块得分 / 子指标 / 改善与恶化企业画像", "看模块", "—"),
        ("3", "3 30题完整数据", "30 题 × 5 期 × 7 模块完整数据 / 红区/绿区 TOP", "看细节", "—"),
        ("4", "4 区域对比", "6 大区域 7 模块 + 关键指标对比", "看地域", "—"),
        ("5", "5 规模对比", "4 大规模段 + 老板年龄段对比", "看规模", "—"),
        ("6", "6 业态对比", "5 大业态 + 业态差异化基准", "看业态", "—"),
        ("7", "7 6月预测", "12 月迁移矩阵 / 退出企业预测 / 行业洗牌指数 KSI", "看未来", "—"),
        ("8", "8 行业基准", "财务 / 运营 / 数字化 / 团队 / 业态 5 大基准对照表", "看标尺", "—"),
        ("9", "9 政策环境", "6 大政策 / 经济环境 / 钢价走势 / 国际市场", "看外部", "—"),
    ]
    for t in tabs:
        row_vals(ws, r, list(t))
        r += 1

    r += 1
    title(ws, r, 1, 5, "重要声明", fill=RED, font=WHITE_BOLD)
    r += 1
    statements = [
        "1. 本数据集仅供订阅人查阅，不得复制、分发、二次销售。",
        "2. 数据来自客户授权脱敏样本聚合，单家公司数据不可识别。",
        "3. 引用本数据须注明出处：'钢联数字化研究院 2026 Q2'。",
        "4. 防伪编号唯一对应订阅人，私自分发可追溯。",
        "5. 数据时效：发布之日有效，下期发布后自动失效。",
    ]
    for s in statements:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c = ws.cell(row=r, column=1, value=s)
        c.font = NORMAL; c.alignment = LEFT; c.fill = LIGHT_GRAY; c.border = BORDER
        ws.row_dimensions[r].height = 22
        r += 1


# -------------------------------------------------------------------
# Tab 1 总体分布 + 趋势
# -------------------------------------------------------------------
def build_distribution(wb: Workbook) -> None:
    ws = wb.create_sheet("1 总体分布")
    setw(ws, [4, 18, 12, 12, 12, 12, 12, 14])

    title(ws, 1, 1, 8, "1.1 等级分布（5 期对比）")
    title(ws, 2, 1, 8, "样本量 1,547 / 同口径 1,402（90.6%）", fill=LIGHT_BLUE, font=BOLD, height=22)

    r = 4
    header_row(ws, r, ["等级", "分数区间", "2025 Q3", "Q4", "2026 Q1", "Q2", "Q2 vs Q1", "Q3 预测"])
    r += 1
    data = [
        ("A 健康", "≥ 90", "4.5%", "7.4%", "7.8%", "8.5%", "+0.7", "9.2%"),
        ("B 良好", "80-89", "18.4%", "21.6%", "22.4%", "24.3%", "+1.9", "25.5%"),
        ("C 亚健康", "60-79", "52.5%", "51.2%", "50.6%", "48.9%", "-1.7", "47.8%"),
        ("D 高危", "< 60", "24.6%", "19.8%", "19.2%", "18.3%", "-0.9", "17.5%"),
    ]
    for d in data:
        cls = GREEN if d[0].startswith("A") else (LIGHT_GREEN if d[0].startswith("B") else (LIGHT_YELLOW if d[0].startswith("C") else LIGHT_RED))
        row_vals(ws, r, list(d), fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 8, "1.2 健康度总分趋势（2024 Q1 - 2026 Q2 + Q3 预测）", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["#", "季度", "健康度总分", "环比", "同比", "事件", "—", "—"])
    r += 1
    trend = [
        (1, "2024 Q1", 72.0, 0, 0, "高点（房地产未崩盘前）"),
        (2, "2024 Q2", 70.8, -1.2, "", ""),
        (3, "2024 Q3", 69.5, -1.3, "", ""),
        (4, "2024 Q4", 68.2, -1.3, "", ""),
        (5, "2025 Q1", 67.0, -1.2, -5.0, "房地产新开工 -32%"),
        (6, "2025 Q2", 66.4, -0.6, -4.4, "钢价单季 4 次大跌"),
        (7, "2025 Q3", 65.2, -1.2, -4.3, "★ 谷底"),
        (8, "2025 Q4", 66.5, +1.3, -1.7, "小幅反弹"),
        (9, "2026 Q1", 67.2, +0.7, +0.2, ""),
        (10, "2026 Q2", 68.5, +1.3, +2.1, "★ 当前期"),
        (11, "2026 Q3（预测）", 69.0, +0.5, +3.8, "AI 模型预测"),
        (12, "2026 Q4（预测）", 70.0, +1.0, +1.8, "重回 70 大关"),
    ]
    for t in trend:
        cls = GOLD if "★" in str(t[5]) else None
        row_vals(ws, r, list(t) + [""] * (8 - len(t)), fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 8, "1.3 与其他行业对比", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["排名", "行业", "平均健康度", "—", "—", "—", "—", "—"])
    r += 1
    ind = [
        (1, "互联网 / 软件", 84.5),
        (2, "医疗器械", 79.2),
        (3, "新能源", 78.6),
        (4, "高端制造", 76.4),
        (5, "汽车", 71.8),
        (6, "家电", 71.2),
        (7, "钢贸 / 钢厂", 68.5),
        (8, "房地产", 58.2),
        (9, "建筑", 56.4),
    ]
    for i in ind:
        cls = GOLD if i[1] == 68.5 else None
        row_vals(ws, r, list(i) + [""] * 5, fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 8, "1.4 健康度与利润相关性", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["等级", "毛利率", "净利率", "应收周转", "样本数", "占比", "—", "—"])
    r += 1
    lvl_p = [
        ("A 级 ≥90", "5.2%", "2.8%", "28 天", 131, "8.5%"),
        ("B 级 80-89", "3.8%", "2.0%", "35 天", 376, "24.3%"),
        ("C 级 60-79", "2.5%", "1.2%", "52 天", 757, "48.9%"),
        ("D 级 < 60", "1.2%", "0.1%", "78 天", 283, "18.3%"),
    ]
    for lp in lvl_p:
        cls = GREEN if "A" in lp[0] else (LIGHT_GREEN if "B" in lp[0] else (LIGHT_YELLOW if "C" in lp[0] else LIGHT_RED))
        row_vals(ws, r, list(lp) + [""] * 2, fill=cls)
        r += 1


# -------------------------------------------------------------------
# Tab 2 7 模块详解
# -------------------------------------------------------------------
def build_modules(wb: Workbook) -> None:
    ws = wb.create_sheet("2 7模块详解")
    setw(ws, [4, 18, 12, 12, 12, 12, 12, 28])

    title(ws, 1, 1, 8, "7 模块平均分（5 期 + 变化）")
    r = 4
    header_row(ws, r, ["#", "模块", "Q4 2025", "Q1 2026", "Q2 2026", "Q3 预测", "Q2 vs Q1", "解读"])
    r += 1
    modules = [
        ("M1", "财务健康度", 65.2, 65.2, 67.8, 69.0, "+2.6 ↗", "净利率 +0.3 个点"),
        ("M2", "客户管理", 62.5, 62.4, 64.3, 65.5, "+1.9 ↗", "TOP 5 依赖 -7"),
        ("M3", "数字化能力", 56.8, 58.7, 61.2, 64.0, "+2.5 ↗", "ERP 使用 47%"),
        ("M4", "风险控制", 72.0, 71.2, 70.5, 70.0, "-0.7 ↘", "★ 唯一下行 / 客户跑路 +28%"),
        ("M5", "增长能力", 62.5, 64.5, 67.2, 68.5, "+2.7 ↗", "服务化 30%"),
        ("M6", "团队能力", 68.5, 69.3, 70.1, 70.8, "+0.8 ↗", "二代接班 23%"),
        ("M7", "战略前瞻", 58.5, 60.1, 62.5, 64.0, "+2.4 ↗", "再做做看 -7"),
        ("综合", "", 66.5, 67.2, 68.5, 69.0, "+1.3 ↗", "—"),
    ]
    for m in modules:
        cls = GOLD if m[0] == "综合" else (LIGHT_RED if m[0] == "M4" else LIGHT_BLUE if m[0].startswith("M3") else None)
        row_vals(ws, r, list(m), fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 8, "M1 财务健康度 子指标", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["指标", "Q1 2026", "Q2 2026", "变化", "警示线", "A 级", "D 级", "解读"])
    r += 1
    m1 = [
        ("平均毛利率", "2.6%", "2.8%", "+0.2 ↗", "2.5%", "5.2%", "1.2%", "上 AI 定价"),
        ("平均净利率", "1.1%", "1.4%", "+0.3 ↗", "1.0%", "2.8%", "0.1%", "砍 D 客户"),
        ("90 天+ 应收", "22.5%", "24.8%", "+2.3 ↘", "25%", "12%", "48%", "★ 反向恶化"),
        ("库存周转(天)", "50", "48", "-2 ↗", "60", "28", "78", "AI 库存预警"),
        ("资产负债率", "70%", "68%", "-2 ↗", "75%", "58%", "76%", "—"),
    ]
    for x in m1:
        cls = LIGHT_RED if "★" in x[-1] else None
        row_vals(ws, r, list(x), fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 8, "M3 数字化分化（A 级 vs D 级）★ 历史峰值", fill=RED, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["指标", "A 级", "D 级", "差距", "—", "—", "—", "—"])
    r += 1
    m3 = [
        ("ERP 使用率", "100%", "8%", "92 个点", "", "", "", ""),
        ("WMS 使用率", "95%", "0%", "95 个点", "", "", "", ""),
        ("AI 工具数", "3.2", "0", "3.2 个", "", "", "", ""),
        ("数据日报自动化", "92%", "0%", "92 个点", "", "", "", ""),
        ("数字化总分", "93", "38", "★ 55 分", "", "", "", ""),
    ]
    for x in m3:
        cls = GOLD if "★" in str(x[3]) else None
        row_vals(ws, r, list(x), fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 8, "改善企业画像（185 家进入 A 级 / Q2）", fill=GREEN, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["共性", "比例", "—", "—", "—", "—", "—", "—"])
    r += 1
    improve = [
        ("已上 ERP + AI 经营包", "92%"),
        ("房地产客户 < 20%", "88%"),
        ("现金流储备 ≥ 3 个月", "85%"),
        ("单笔订单 ≥ ¥100 万", "80%"),
        ("二代实质参与", "62%"),
    ]
    for x in improve:
        row_vals(ws, r, list(x) + [""] * 6)
        r += 1

    r += 2
    title(ws, r, 1, 8, "恶化企业画像（127 家跌入 D 级 / Q2）", fill=RED, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["共性", "比例", "—", "—", "—", "—", "—", "—"])
    r += 1
    decline = [
        ("应收 > 35%", "95%"),
        ("房地产客户 > 50%", "78%"),
        ("仍用 Excel 管账", "72%"),
        ("老板 60+ 岁 + 无二代", "68%"),
        ("近 12 月无新客", "55%"),
    ]
    for x in decline:
        row_vals(ws, r, list(x) + [""] * 6, fill=LIGHT_RED)
        r += 1


# -------------------------------------------------------------------
# Tab 3 30 题完整数据
# -------------------------------------------------------------------
def build_30q(wb: Workbook) -> None:
    ws = wb.create_sheet("3 30题完整数据")
    setw(ws, [5, 8, 38, 10, 10, 10, 10, 10, 14])

    title(ws, 1, 1, 9, "30 题完整数据（按 7 模块分组）")
    r = 4
    header_row(ws, r, ["#", "题号", "题目", "Q4 2025", "Q1 2026", "Q2 2026", "Q3 预测", "红区%", "状态"])
    r += 1
    # 30 题数据
    questions = [
        # M1 财务 (Q1-Q6)
        (1, "Q1", "公司规模", 84, 85, 85, 86, "8%", "🟢 绿"),
        (2, "Q2", "经营年限", 78, 79, 79, 79, "12%", "🟢 绿"),
        (3, "Q3", "90 天+ 应收占比", 53, 52, 51, 50, "38%", "🔴 红"),
        (4, "Q4", "客户跑路次数", 76, 75, 75, 73, "15%", "🟢 绿"),
        (5, "Q5", "营收稳定性", 80, 81, 82, 82, "12%", "🟢 绿"),
        (6, "Q6", "现金流储备月数", 60, 60, 61, 62, "25%", "🟡 警"),
        # M2 客户 (Q7-Q13)
        (7, "Q7", "核心团队稳定性", 79, 79, 80, 80, "14%", "🟢 绿"),
        (8, "Q8", "主营产品集中度", 64, 65, 66, 67, "22%", "🟡 警"),
        (9, "Q9", "销售跑单率", 56, 55, 55, 54, "33%", "🔴 红"),
        (10, "Q10", "新客户开发能力", 62, 64, 66, 68, "26%", "🟡 警"),
        (11, "Q11", "房地产客户依赖", 58, 57, 56, 55, "32%", "🔴 红"),
        (12, "Q12", "银行抽贷预警", 65, 63, 62, 60, "21%", "🟡 警"),
        (13, "Q13", "TOP 5 客户合作年限", 71, 72, 72, 72, "16%", "🟢 绿"),
        # M3 数字化 (Q14-Q18)
        (14, "Q14", "老板个人投入度", 77, 78, 78, 78, "15%", "🟢 绿"),
        (15, "Q15", "数字化系统使用", 42, 45, 48, 53, "42%", "🔴 红"),
        (16, "Q16", "数据日报自动化", 58, 60, 62, 65, "28%", "🟡 警"),
        (17, "Q17", "数字化主线人员", 55, 57, 59, 62, "30%", "🟡 警"),
        (18, "Q18", "服务化转型程度", 53, 55, 58, 62, "30%", "🔴 红"),
        # M4 风控 (Q19-Q20)
        (19, "Q19", "风险评估流程", 72, 73, 74, 74, "18%", "🟢 绿"),
        (20, "Q20", "黑名单客户筛除", 68, 69, 69, 70, "20%", "🟡 警"),
        # M5 增长 (Q21-Q24)
        (21, "Q21", "新客户开发月均", 60, 62, 66, 68, "25%", "🟡 警"),
        (22, "Q22", "库存周转", 48, 50, 52, 55, "35%", "🔴 红"),
        (23, "Q23", "老客户流失率", 62, 60, 60, 60, "28%", "🟡 警"),
        (24, "Q24", "钢价波动应对工具", 55, 56, 57, 60, "30%", "🔴 红"),
        # M6 团队 (Q25-Q28)
        (25, "Q25", "财务核算规范度", 70, 72, 73, 74, "18%", "🟢 绿"),
        (26, "Q26", "销售团队战斗力", 73, 75, 76, 76, "17%", "🟢 绿"),
        (27, "Q27", "二代接班准备", 50, 52, 53, 55, "38%", "🔴 红"),
        (28, "Q28", "关键岗位备份", 65, 66, 67, 68, "23%", "🟡 警"),
        # M7 战略 (Q29-Q30)
        (29, "Q29", "服务化转型规划", 56, 58, 60, 63, "26%", "🟡 警"),
        (30, "Q30", "三年战略清晰度", 56, 58, 59, 62, "28%", "🔴 红"),
    ]
    for q in questions:
        cls = LIGHT_RED if "🔴" in q[-1] else (LIGHT_YELLOW if "🟡" in q[-1] else LIGHT_GREEN)
        row_vals(ws, r, list(q), fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 9, "红区 TOP 10 / 绿区 TOP 10 / 变化最大 TOP 5", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    # 改善最快 TOP 5
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    c = ws.cell(row=r, column=1, value="★ 改善最快 TOP 5（销售可重点宣传）")
    c.font = BOLD; c.fill = LIGHT_GREEN; c.alignment = LEFT; c.border = BORDER
    r += 1
    fast_up = [
        ("Q15", "数字化系统使用", "+5.2 ↗"),
        ("Q18", "服务化转型", "+4.8 ↗"),
        ("Q22", "库存周转", "+4.1 ↗"),
        ("Q21", "新客户开发", "+3.8 ↗"),
        ("Q30", "3 年战略清晰度", "+3.6 ↗"),
    ]
    header_row(ws, r, ["#", "题号", "题目", "变化", "—", "—", "—", "—", "—"])
    r += 1
    for i, q in enumerate(fast_up, 1):
        row_vals(ws, r, [i] + list(q) + [""] * 5, fill=LIGHT_GREEN)
        r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    c = ws.cell(row=r, column=1, value="★ 恶化最快 TOP 5（销售可警示）")
    c.font = BOLD; c.fill = LIGHT_RED; c.alignment = LEFT; c.border = BORDER
    r += 1
    fast_down = [
        ("Q3", "90 天+ 应收", "-3.2 ↘"),
        ("Q4", "客户跑路", "-2.8 ↘"),
        ("Q12", "银行抽贷", "-2.5 ↘"),
        ("Q11", "房地产客户依赖", "-1.8 ↘"),
        ("Q23", "老客户流失", "-1.6 ↘"),
    ]
    header_row(ws, r, ["#", "题号", "题目", "变化", "—", "—", "—", "—", "—"])
    r += 1
    for i, q in enumerate(fast_down, 1):
        row_vals(ws, r, [i] + list(q) + [""] * 5, fill=LIGHT_RED)
        r += 1


# -------------------------------------------------------------------
# Tab 4 区域对比
# -------------------------------------------------------------------
def build_region(wb: Workbook) -> None:
    ws = wb.create_sheet("4 区域对比")
    setw(ws, [4, 16, 8, 12, 10, 10, 10, 10, 10, 10, 10, 16])

    title(ws, 1, 1, 12, "6 大区域对比（含 7 模块）")
    r = 4
    header_row(ws, r, ["#", "区域", "样本", "平均分", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "关键特征"])
    r += 1
    regions = [
        (1, "长三角(江浙沪皖)", 612, 76.2, 72, 70, 83, 74, 75, 74, 70, "出口+数字化"),
        (2, "珠三角(粤桂闽)", 215, 71.5, 69, 68, 73, 72, 75, 69, 74, "供应链金融"),
        (3, "华中(鄂湘赣)", 165, 69.3, 73, 70, 65, 70, 71, 69, 66, "出口转型"),
        (4, "华南其他", 65, 67.8, 60, 65, 68, 70, 72, 71, 65, "区域整合"),
        (5, "西南(川渝云贵)", 102, 65.5, 65, 63, 60, 67, 65, 70, 62, "基建+内贸"),
        (6, "京津冀(京津冀豫鲁)", 388, 64.8, 60, 60, 55, 68, 62, 65, 60, "★ 应收最重"),
        ("—", "全国", 1547, 68.5, 67.8, 64.3, 61.2, 70.5, 67.2, 70.1, 62.5, "—"),
    ]
    for re in regions:
        cls = GOLD if "全国" in str(re[1]) else (LIGHT_RED if "★" in str(re[-1]) else (LIGHT_GREEN if re[3] >= 75 else None))
        row_vals(ws, r, list(re), fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 12, "长三角 vs 京津冀（差 11.4 分）", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["维度", "长三角", "京津冀", "差距", "解读", "—", "—", "—", "—", "—", "—", "—"])
    r += 1
    diff = [
        ("健康度总分", 76.2, 64.8, 11.4, "全国最大区域差"),
        ("M3 数字化", 83, 55, 28, "★ 单模块最大差距"),
        ("ERP 使用率", "63%", "38%", "25 个点", "数字化滞后"),
        ("90 天+ 应收", "18%", "32%", "14 个点", "★ 应收最严重"),
        ("房地产客户依赖", "20%", "35%", "15 个点", "建筑客户多"),
        ("60+ 老板比例", "22%", "38%", "16 个点", "★ 接班断层"),
    ]
    for d in diff:
        cls = LIGHT_RED if "★" in d[-1] else None
        row_vals(ws, r, list(d) + [""] * 7, fill=cls)
        r += 1


# -------------------------------------------------------------------
# Tab 5 规模对比
# -------------------------------------------------------------------
def build_size(wb: Workbook) -> None:
    ws = wb.create_sheet("5 规模对比")
    setw(ws, [4, 18, 10, 12, 12, 12, 12, 14])

    title(ws, 1, 1, 8, "5.1 按年流水规模分级")
    r = 4
    header_row(ws, r, ["#", "年流水", "样本", "平均分", "A+B 占比", "D 占比", "建议", "—"])
    r += 1
    sizes = [
        (1, "> ¥5 亿", 142, 82.5, "62%", "4%", "扩张/收购同行", ""),
        (2, "¥1-5 亿", 485, 71.3, "38%", "12%", "稳健+局部升级", ""),
        (3, "¥3,000 万-1 亿", 612, 65.4, "24%", "18%", "★ 必须升级", ""),
        (4, "< ¥3,000 万", 308, 56.8, "9%", "38%", "生死大考", ""),
        ("—", "全国平均", 1547, 68.5, "32.8%", "18.3%", "—", ""),
    ]
    for s in sizes:
        cls = LIGHT_RED if "★" in str(s[6]) else (GOLD if "全国" in str(s[1]) else None)
        row_vals(ws, r, list(s), fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 8, "5.2 按老板年龄段分级", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["#", "年龄段", "样本", "平均分", "二代接班程度", "5 年内传承断层风险", "建议", "—"])
    r += 1
    ages = [
        (1, "< 35 岁", 95, 74.8, "—（自己就是二代）", "0%", "扩张+资本化"),
        (2, "35-45 岁", 352, 72.6, "45%", "5%", "稳健"),
        (3, "45-55 岁", 685, 68.4, "18%", "20%", "推动二代接手"),
        (4, "55-65 岁", 312, 65.2, "8%", "45%", "★ 启动接班数字化"),
        (5, "> 65 岁", 103, 58.7, "2%", "60%", "★★ 传承迫在眉睫"),
    ]
    for a in ages:
        cls = LIGHT_RED if "★" in str(a[6]) else None
        row_vals(ws, r, list(a) + [""], fill=cls)
        r += 1


# -------------------------------------------------------------------
# Tab 6 业态对比
# -------------------------------------------------------------------
def build_biz(wb: Workbook) -> None:
    ws = wb.create_sheet("6 业态对比")
    setw(ws, [4, 18, 10, 12, 10, 10, 10, 12, 14])

    title(ws, 1, 1, 9, "5 大业态对比")
    r = 4
    header_row(ws, r, ["#", "业态", "样本", "平均分", "毛利率", "应收 90+", "周转(天)", "优势模块", "劣势模块"])
    r += 1
    biz = [
        (1, "综合型(贸+加工+服务)", 198, 75.8, "3.5%", "20%", "40", "全面均衡", "—"),
        (2, "加工型(开平/剪板)", 285, 73.2, "2.2%", "15%", "30", "M1 财务", "M5 拓客"),
        (3, "流通+仓储", 178, 70.1, "3.0%", "20%", "45", "M4 风控", "M5 增长"),
        (4, "钢厂客户", 364, 69.5, "4.5%", "18%", "55", "M1 财务", "M3 数字化"),
        (5, "纯贸易", 522, 65.4, "2.8%", "25%", "48", "M2 客户灵活", "★ 全面薄弱"),
    ]
    for b in biz:
        cls = LIGHT_RED if "★" in str(b[8]) else (LIGHT_GREEN if b[3] >= 73 else None)
        row_vals(ws, r, list(b), fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 9, "业态差异化基准（行业平均）", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["维度", "钢管厂", "镀锌厂", "开平厂", "综合钢贸", "钢厂", "—", "—", "—"])
    r += 1
    biz_metric = [
        ("毛利率", "3.8%", "4.2%", "2.2%", "2.8%", "4.5%"),
        ("90 天+ 应收", "22%", "18%", "15%", "25%", "18%"),
        ("库存周转(天)", 52, 55, 30, 48, 55),
        ("单笔均值", "¥45 万", "¥80 万", "¥8 万", "¥35 万", "¥120 万"),
        ("出口客户比例", "18%", "22%", "8%", "15%", "20%"),
    ]
    for m in biz_metric:
        row_vals(ws, r, list(m) + [""] * 3)
        r += 1


# -------------------------------------------------------------------
# Tab 7 6 月预测 + 迁移矩阵
# -------------------------------------------------------------------
def build_forecast(wb: Workbook) -> None:
    ws = wb.create_sheet("7 6月预测")
    setw(ws, [4, 18, 14, 14, 14, 14, 14, 16])

    title(ws, 1, 1, 8, "6 个月迁移矩阵（基于 12 月历史 + AI 预测）")
    r = 4
    header_row(ws, r, ["从→到", "→A", "→B", "→C", "→D", "→退出市场", "合计", "稳定率"])
    r += 1
    matrix = [
        ("A 级 (131)", "105", "11", "3", "2", "10", 131, "80.2%"),
        ("B 级 (376)", "26", "287", "31", "12", "20", 376, "76.3%"),
        ("C 级 (757)", "0", "87", "591", "76", "3", 757, "78.1%"),
        ("D 级 (283)", "0", "0", "5", "51", "227", 283, "18.0%"),
        ("合计", "131", "385", "630", "141", "260", 1547, "—"),
    ]
    for m in matrix:
        cls = GOLD if m[0] == "合计" else (LIGHT_RED if "D 级" in m[0] else None)
        row_vals(ws, r, list(m), fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 8, "关键预测结论", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    conclusions = [
        "★ 未来 6 个月约 260 家钢贸退出市场（占样本 16.8%）",
        "★ 其中 D 级流出 227 家（80% 来自 D 级）",
        "★ 但同时 87 家 C 级跃升 B 级 + 26 家 B 级跃升 A 级",
        "★ A + B 级合计占比 32.8% → 预测 36-38%（头部加速集中）",
        "★ 行业洗牌指数（KSI）：2.4 → 2.1（趋稳但仍剧烈）",
    ]
    for c in conclusions:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        cell = ws.cell(row=r, column=1, value=c)
        cell.font = NORMAL; cell.alignment = LEFT; cell.fill = LIGHT_YELLOW; cell.border = BORDER
        ws.row_dimensions[r].height = 22
        r += 1

    r += 2
    title(ws, r, 1, 8, "钢贸行业洗牌指数 KSI 历史", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["季度", "KSI", "状态", "—", "—", "—", "—", "—"])
    r += 1
    ksi = [
        ("2024 Q1", 1.0, "基准期"),
        ("2024 Q4", 1.8, "渐烈"),
        ("2025 Q3", 3.2, "★ 峰值"),
        ("2026 Q1", 2.7, "降"),
        ("2026 Q2", 2.4, "续降"),
        ("2026 Q3 预测", 2.1, "续降"),
        ("2027 Q1 预测", 1.5, "趋稳"),
    ]
    for k in ksi:
        cls = LIGHT_RED if "★" in k[2] else None
        row_vals(ws, r, list(k) + [""] * 5, fill=cls)
        r += 1


# -------------------------------------------------------------------
# Tab 8 行业基准对照
# -------------------------------------------------------------------
def build_benchmark(wb: Workbook) -> None:
    ws = wb.create_sheet("8 行业基准")
    setw(ws, [4, 20, 12, 12, 12, 12, 10, 10])

    title(ws, 1, 1, 8, "8.1 财务基准（钢贸全行业）")
    r = 4
    header_row(ws, r, ["#", "指标", "健康线", "警示线", "危险线", "行业平均", "A 级", "D 级"])
    r += 1
    fin = [
        (1, "毛利率", "≥ 4%", "2-4%", "< 2%", "2.8%", "5.2%", "1.2%"),
        (2, "净利率", "≥ 2%", "0.5-2%", "< 0.5%", "1.4%", "2.8%", "0.1%"),
        (3, "90 天+ 应收", "< 15%", "15-30%", "> 30%", "24.8%", "12%", "48%"),
        (4, "库存周转(天)", "≤ 30", "30-60", "> 60", "48", "28", "78"),
        (5, "资产负债率", "< 60%", "60-75%", "> 75%", "68%", "58%", "76%"),
        (6, "现金流储备(月)", "≥ 3", "1-3", "< 1", "1.8", "4.5", "0.4"),
    ]
    for f in fin:
        row_vals(ws, r, list(f))
        r += 1

    r += 2
    title(ws, r, 1, 8, "8.2 运营基准", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["#", "指标", "健康线", "警示线", "危险线", "行业平均", "—", "—"])
    r += 1
    ops = [
        (1, "客户流失率", "< 8%", "8-15%", "> 15%", "14%"),
        (2, "TOP 5 客户依赖", "< 30%", "30-50%", "> 50%", "38%"),
        (3, "销售跑单率", "< 8%", "8-15%", "> 15%", "12%"),
        (4, "客户回款周期", "< 30 天", "30-60 天", "> 60 天", "52 天"),
        (5, "出口客户比例", "≥ 20%", "5-20%", "< 5%", "15%"),
    ]
    for o in ops:
        row_vals(ws, r, list(o) + [""] * 2)
        r += 1

    r += 2
    title(ws, r, 1, 8, "8.3 数字化基准", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["#", "指标", "健康线", "警示线", "危险线", "行业平均", "A 级", "D 级"])
    r += 1
    digi = [
        (1, "ERP 使用率", "≥ 90%", "50-90%", "< 50%", "47%", "100%", "8%"),
        (2, "WMS 使用", "有", "—", "无", "25%", "95%", "0%"),
        (3, "数据日报", "自动", "半自动", "手工", "25%", "92%", "0%"),
        (4, "AI 工具数", "≥ 3", "1-2", "0", "0.5", "3.2", "0"),
    ]
    for d in digi:
        row_vals(ws, r, list(d))
        r += 1

    r += 2
    title(ws, r, 1, 8, "8.4 团队 + 战略基准", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["#", "指标", "健康线", "警示线", "危险线", "行业平均", "—", "—"])
    r += 1
    team = [
        (1, "销售流失率", "< 10%", "10-18%", "> 18%", "13%"),
        (2, "二代接班程度", "≥ 50%", "20-50%", "< 20%", "23%"),
        (3, "3 年战略清晰度", "≥ 80", "50-80", "< 50", 62.5),
        (4, "服务化转型", "≥ 30%", "10-30%", "< 10%", "30%"),
    ]
    for t in team:
        row_vals(ws, r, list(t) + [""] * 2)
        r += 1


# -------------------------------------------------------------------
# Tab 9 政策环境
# -------------------------------------------------------------------
def build_policy(wb: Workbook) -> None:
    ws = wb.create_sheet("9 政策环境")
    setw(ws, [4, 26, 14, 14, 38, 12, 12])

    title(ws, 1, 1, 7, "9.1 关键政策影响")
    r = 4
    header_row(ws, r, ["#", "政策", "发布", "生效", "影响", "利好方向", "利空方向"])
    r += 1
    policies = [
        (1, "增值税电子凭证全面强制", "2025-12", "2026-07", "倒逼数字化", "有 ERP 钢贸", "无 ERP / Excel"),
        (2, "钢铁出口退税调整(10→11%)", "2026-03", "2026 Q3", "利好出口型", "长三角/珠三角", "—"),
        (3, "银行授信新规(数字化加权)", "2026-04", "2026 H2", "倒逼数字化", "有 ERP 钢贸", "无 ERP 钢贸"),
        (4, "钢铁产能置换条例", "持续", "—", "倒逼上下游升级", "数字化钢贸", "传统钢贸"),
        (5, "雄安基建第 5 期", "2026-05", "2026 H2", "基建红利", "京津冀钢贸", "—"),
        (6, "反倾销关税(美国 +5%)", "2026-06", "2026 Q3", "出口结构调整", "东南亚转口", "美洲出口"),
    ]
    for p in policies:
        row_vals(ws, r, list(p))
        r += 1

    r += 2
    title(ws, r, 1, 7, "9.2 经济环境关键指标", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["#", "维度", "数据", "趋势", "影响", "—", "—"])
    r += 1
    econ = [
        (1, "房地产新开工同比", "-28%", "↘ 持续下行", "建筑钢材需求 -28%"),
        (2, "基建投资同比", "+6%", "↗ 缓升", "基建钢材需求 +6%"),
        (3, "出口订单同比", "+18%", "↗ 强劲", "出口钢贸利好"),
        (4, "PMI(钢铁)", "49.5", "→ 临界点", "中性"),
        (5, "LPR 1 年", "3.45%", "→ 稳定", "中性"),
        (6, "银行抽贷率", "21%", "↘ 警示", "现金流压力"),
    ]
    for e in econ:
        row_vals(ws, r, list(e) + [""] * 2)
        r += 1

    r += 2
    title(ws, r, 1, 7, "9.3 钢价 6 月走势", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["月份", "螺纹钢均价(元/吨)", "同比", "环比", "单日最大波动", "—", "—"])
    r += 1
    prices = [
        ("4 月", 3820, "-8%", "—", "-180"),
        ("5 月", 3650, "-12%", "-4.4%", "-240"),
        ("6 月", 3720, "-10%", "+1.9%", "-150"),
        ("Q3 预测", "3600-3800", "-7%", "—", "小幅震荡"),
    ]
    for p in prices:
        cls = GOLD if "预测" in str(p[0]) else None
        row_vals(ws, r, list(p) + [""] * 2, fill=cls)
        r += 1

    r += 2
    title(ws, r, 1, 7, "9.4 国际市场利好/利空", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    header_row(ws, r, ["#", "区域", "钢材进口需求", "利好程度", "—", "—", "—"])
    r += 1
    intl = [
        (1, "东南亚(越南/泰国/印尼)", "+25%", "强利好"),
        (2, "中东(沙特/阿联酋)", "+18%", "强利好"),
        (3, "非洲(埃及/尼日利亚)", "+15%", "利好"),
        (4, "美洲(巴西/墨西哥)", "+8%", "中性"),
        (5, "美国", "-10%", "利空"),
        (6, "欧洲", "-5%", "利空"),
    ]
    for i in intl:
        cls = LIGHT_GREEN if "利好" in i[3] else (LIGHT_RED if "利空" in i[3] else None)
        row_vals(ws, r, list(i) + [""] * 3, fill=cls)
        r += 1


# -------------------------------------------------------------------
def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)

    build_overview(wb)
    build_distribution(wb)
    build_modules(wb)
    build_30q(wb)
    build_region(wb)
    build_size(wb)
    build_biz(wb)
    build_forecast(wb)
    build_benchmark(wb)
    build_policy(wb)

    out = "tools/2026Q2钢贸健康度数据集.xlsx"
    wb.save(out)
    print(f"✓ 已生成：{out}")


if __name__ == "__main__":
    main()
