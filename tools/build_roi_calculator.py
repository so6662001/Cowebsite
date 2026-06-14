"""生成钢贸 7 大场景 ROI 计算器 Excel。

输出：tools/钢贸7大场景ROI计算器.xlsx
工作表：0 总览 + M1-M7 共 7 个场景 + 汇总
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
RIGHT = Alignment(horizontal="right", vertical="center")

NAVY = PatternFill("solid", fgColor="1F4E78")
RED = PatternFill("solid", fgColor="C00000")
ORANGE = PatternFill("solid", fgColor="ED7D31")
GREEN = PatternFill("solid", fgColor="00B050")
LIGHT_GRAY = PatternFill("solid", fgColor="F2F2F2")
LIGHT_BLUE = PatternFill("solid", fgColor="D9E1F2")
LIGHT_YELLOW = PatternFill("solid", fgColor="FFF2CC")
LIGHT_GREEN = PatternFill("solid", fgColor="E2EFDA")

WHITE_BOLD = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
BOLD = Font(name="Microsoft YaHei", size=11, bold=True)
NORMAL = Font(name="Microsoft YaHei", size=10)
TITLE = Font(name="Microsoft YaHei", size=16, bold=True, color="1F4E78")


def set_col_widths(ws: Worksheet, widths: list[float]) -> None:
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_title(ws: Worksheet, row: int, col: int, span: int, text: str, fill=NAVY, font=WHITE_BOLD) -> None:
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    c = ws.cell(row=row, column=col, value=text)
    c.fill = fill
    c.font = font
    c.alignment = CENTER
    c.border = BORDER
    ws.row_dimensions[row].height = 30


def write_row(ws: Worksheet, row: int, values: list, fills=None, fonts=None, borders=True, aligns=None) -> None:
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=i + 1, value=v)
        if fills:
            c.fill = fills[i] if isinstance(fills, list) else fills
        if fonts:
            c.font = fonts[i] if isinstance(fonts, list) else fonts
        else:
            c.font = NORMAL
        if borders:
            c.border = BORDER
        if aligns:
            c.alignment = aligns[i] if isinstance(aligns, list) else aligns
        else:
            c.alignment = CENTER


def build_overview(wb: Workbook) -> None:
    ws = wb.create_sheet("0 总览导航")
    set_col_widths(ws, [4, 18, 12, 28, 22, 14, 14, 14])

    write_title(ws, 1, 1, 8, "钢贸 7 大场景 ROI 计算器 · 总览")
    write_title(ws, 2, 1, 8, "30 题健康度评分 → 7 模块红区 → 对应场景客户故事 + ROI 计算", fill=LIGHT_BLUE, font=BOLD)

    write_title(ws, 4, 1, 8, "客户基础信息（请录入）", fill=ORANGE, font=WHITE_BOLD)
    info = [
        ("客户公司", "请填写"),
        ("老板姓名", "请填写"),
        ("年流水（万元）", 18000),
        ("员工人数", 50),
        ("成立年份", 2005),
        ("销售归属", "请填写"),
        ("拜访日期", "2026-06-15"),
        ("销售总监", "请填写"),
    ]
    r = 5
    for k, v in info:
        ws.cell(row=r, column=2, value=k).font = BOLD
        ws.cell(row=r, column=2).fill = LIGHT_GRAY
        ws.cell(row=r, column=2).alignment = LEFT
        ws.cell(row=r, column=2).border = BORDER
        c = ws.cell(row=r, column=3, value=v)
        c.fill = LIGHT_YELLOW
        c.font = NORMAL
        c.alignment = LEFT
        c.border = BORDER
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        r += 1

    r = 14
    write_title(ws, r, 1, 8, "7 大模块健康度评分（自评后填入）")
    r += 1
    header = ["#", "模块", "评分", "等级（自动）", "对应场景", "客户故事", "故事主角", "Tab 跳转"]
    write_row(ws, r, header, fills=LIGHT_BLUE, fonts=BOLD)
    r += 1

    modules = [
        ("M1", "财务健康度", 78, "利润 + 现金流", "山东 XX 钢管 王总", "王总", "M1 财务"),
        ("M2", "客户管理", 65, "客户结构 + 销售跑单", "广东 XX 钢贸 陈总", "陈总", "M2 客户"),
        ("M3", "数字化能力", 60, "系统升级 + 出口审厂", "江苏 XX 焊管 周总", "周总", "M3 数字化"),
        ("M4", "风险控制", 70, "应收 + 风控", "江苏 XX 钢贸 赵总", "赵总", "M4 风控"),
        ("M5", "增长能力", 72, "拓客 + 出口", "浙江 XX 钢贸 钱总", "钱总", "M5 增长"),
        ("M6", "团队能力", 85, "接班 + 团队", "江苏 XX 镀锌 周总", "周总", "M6 团队"),
        ("M7", "战略前瞻", 70, "服务化 + 转型", "上海 XX 钢贸 X 总", "X 总", "M7 战略"),
    ]
    for i, (code, name, score, scene, story, hero, tab) in enumerate(modules):
        row_data = [code, name, score, "", scene, story, hero, tab]
        for col_idx, val in enumerate(row_data, 1):
            c = ws.cell(row=r, column=col_idx, value=val)
            c.font = NORMAL
            c.border = BORDER
            c.alignment = CENTER
        ws.cell(row=r, column=4, value=f'=IF(C{r}<60,"D 红区",IF(C{r}<70,"C 警示",IF(C{r}<80,"B 良好","A 健康")))')
        score_cell = ws.cell(row=r, column=3)
        if score < 60:
            score_cell.fill = RED
            score_cell.font = WHITE_BOLD
        elif score < 70:
            score_cell.fill = ORANGE
            score_cell.font = WHITE_BOLD
        elif score < 80:
            score_cell.fill = LIGHT_YELLOW
        else:
            score_cell.fill = LIGHT_GREEN
        r += 1

    r += 1
    write_title(ws, r, 1, 8, "销售对号入座 SOP", fill=LIGHT_BLUE, font=BOLD)
    r += 1
    sop = [
        "1. 客户做完 30 题 → 7 模块自动出分",
        "2. 红区模块（<60 分）= 必讲故事 + 算 ROI",
        "3. 警示模块（60-69）= 备选讲（如客户继续询问）",
        "4. 健康模块（≥70）= 不主动讲（避免分散注意力）",
        "5. 销售铁律：每个红区都打开对应 Tab → 录入客户数据 → 当场算 ROI → 打印/截图给客户",
        "6. 最高 ROI = M4 风控（应收）→ 通常 >10 倍 / 1-3 个月回本",
        "7. 最大单 ROI = M6 接班 → 公司估值 +¥1 亿（适合年流水 ¥3 亿+ 客户）",
    ]
    for line in sop:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        c = ws.cell(row=r, column=1, value=line)
        c.font = NORMAL
        c.alignment = LEFT
        c.fill = LIGHT_GRAY
        c.border = BORDER
        ws.row_dimensions[r].height = 22
        r += 1


def build_scene(
    wb: Workbook,
    tab_name: str,
    module: str,
    scene_title: str,
    hero: str,
    pre_score: int,
    red_questions: list[str],
    inputs: list[tuple[str, str, float]],
    investments: list[tuple[str, float]],
    formula_lines: list[str],
    benefits: list[tuple[str, str]],
) -> None:
    """构建单个场景计算器 Tab.

    inputs: [(标签, 单位, 默认值)]
    investments: [(项目, 年费)]
    formula_lines: 描述
    benefits: [(标签, 公式)]
    """
    ws = wb.create_sheet(tab_name)
    set_col_widths(ws, [4, 28, 18, 14, 26, 16])

    write_title(ws, 1, 1, 6, f"{module} · {scene_title}")
    write_title(ws, 2, 1, 6, f"客户故事：{hero}（改造前评分 {pre_score} 分）", fill=LIGHT_BLUE, font=BOLD)

    r = 4
    write_title(ws, r, 1, 6, "1. 红区识别（30 题对应）", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    for q in red_questions:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        c = ws.cell(row=r, column=1, value=q)
        c.font = NORMAL
        c.alignment = LEFT
        c.fill = LIGHT_GRAY
        c.border = BORDER
        r += 1

    r += 1
    write_title(ws, r, 1, 6, "2. 客户数据输入（请客户提供 / 默认为行业平均）", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    headers = ["#", "字段", "单位", "默认值（行业平均）", "客户实际值（请填）", "备注"]
    write_row(ws, r, headers, fills=LIGHT_BLUE, fonts=BOLD)
    r += 1
    input_start_row = r
    for i, (label, unit, default) in enumerate(inputs, 1):
        ws.cell(row=r, column=1, value=i).border = BORDER
        ws.cell(row=r, column=1).alignment = CENTER
        ws.cell(row=r, column=1).font = NORMAL
        ws.cell(row=r, column=2, value=label).border = BORDER
        ws.cell(row=r, column=2).alignment = LEFT
        ws.cell(row=r, column=2).font = NORMAL
        ws.cell(row=r, column=3, value=unit).border = BORDER
        ws.cell(row=r, column=3).alignment = CENTER
        ws.cell(row=r, column=3).font = NORMAL
        ws.cell(row=r, column=4, value=default).border = BORDER
        ws.cell(row=r, column=4).alignment = CENTER
        ws.cell(row=r, column=4).font = NORMAL
        ws.cell(row=r, column=4).fill = LIGHT_GRAY
        c = ws.cell(row=r, column=5, value=default)
        c.fill = LIGHT_YELLOW
        c.border = BORDER
        c.font = BOLD
        c.alignment = CENTER
        ws.cell(row=r, column=6, value="可调").border = BORDER
        ws.cell(row=r, column=6).alignment = CENTER
        ws.cell(row=r, column=6).font = NORMAL
        r += 1
    input_end_row = r - 1

    r += 1
    write_title(ws, r, 1, 6, "3. 投入（默认值 / 可调）", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    headers = ["#", "投入项", "", "年费（万元）", "客户实际值", "备注"]
    write_row(ws, r, headers, fills=LIGHT_BLUE, fonts=BOLD)
    r += 1
    inv_start = r
    for i, (label, amount) in enumerate(investments, 1):
        ws.cell(row=r, column=1, value=i).border = BORDER
        ws.cell(row=r, column=1).font = NORMAL
        ws.cell(row=r, column=1).alignment = CENTER
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.cell(row=r, column=2, value=label).border = BORDER
        ws.cell(row=r, column=2).font = NORMAL
        ws.cell(row=r, column=2).alignment = LEFT
        ws.cell(row=r, column=4, value=amount).border = BORDER
        ws.cell(row=r, column=4).fill = LIGHT_GRAY
        ws.cell(row=r, column=4).font = NORMAL
        ws.cell(row=r, column=4).alignment = CENTER
        c = ws.cell(row=r, column=5, value=amount)
        c.fill = LIGHT_YELLOW
        c.border = BORDER
        c.font = BOLD
        c.alignment = CENTER
        ws.cell(row=r, column=6, value="可调").border = BORDER
        ws.cell(row=r, column=6).font = NORMAL
        ws.cell(row=r, column=6).alignment = CENTER
        r += 1
    inv_end = r - 1

    total_inv_row = r
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.cell(row=r, column=2, value="合计年投入（万元）").font = BOLD
    ws.cell(row=r, column=2).fill = RED
    ws.cell(row=r, column=2).font = WHITE_BOLD
    ws.cell(row=r, column=2).alignment = CENTER
    ws.cell(row=r, column=2).border = BORDER
    ws.cell(row=r, column=4, value=f"=SUM(D{inv_start}:D{inv_end})").border = BORDER
    ws.cell(row=r, column=4).fill = LIGHT_GRAY
    ws.cell(row=r, column=4).font = BOLD
    ws.cell(row=r, column=4).alignment = CENTER
    c = ws.cell(row=r, column=5, value=f"=SUM(E{inv_start}:E{inv_end})")
    c.fill = RED
    c.font = WHITE_BOLD
    c.border = BORDER
    c.alignment = CENTER
    ws.cell(row=r, column=6, value="自动").border = BORDER
    ws.cell(row=r, column=6).font = NORMAL
    ws.cell(row=r, column=6).alignment = CENTER
    r += 2

    write_title(ws, r, 1, 6, "4. 收益自动计算（基于客户实际值）", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    headers = ["#", "收益项", "", "公式说明", "金额（万元）", "类型"]
    write_row(ws, r, headers, fills=LIGHT_BLUE, fonts=BOLD)
    r += 1
    benefit_start = r
    for i, (label, formula) in enumerate(benefits, 1):
        ws.cell(row=r, column=1, value=i).border = BORDER
        ws.cell(row=r, column=1).font = NORMAL
        ws.cell(row=r, column=1).alignment = CENTER
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.cell(row=r, column=2, value=label).border = BORDER
        ws.cell(row=r, column=2).font = NORMAL
        ws.cell(row=r, column=2).alignment = LEFT
        ws.cell(row=r, column=4, value="见公式").border = BORDER
        ws.cell(row=r, column=4).font = NORMAL
        ws.cell(row=r, column=4).alignment = CENTER
        c = ws.cell(row=r, column=5, value=formula)
        c.fill = LIGHT_GREEN
        c.border = BORDER
        c.font = BOLD
        c.alignment = CENTER
        ws.cell(row=r, column=6, value="自动").border = BORDER
        ws.cell(row=r, column=6).font = NORMAL
        ws.cell(row=r, column=6).alignment = CENTER
        r += 1
    benefit_end = r - 1

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.cell(row=r, column=2, value="合计年收益（万元）").fill = GREEN
    ws.cell(row=r, column=2).font = WHITE_BOLD
    ws.cell(row=r, column=2).border = BORDER
    ws.cell(row=r, column=2).alignment = CENTER
    ws.cell(row=r, column=4, value="加总").border = BORDER
    ws.cell(row=r, column=4).font = NORMAL
    ws.cell(row=r, column=4).alignment = CENTER
    c = ws.cell(row=r, column=5, value=f"=SUM(E{benefit_start}:E{benefit_end})")
    c.fill = GREEN
    c.font = WHITE_BOLD
    c.border = BORDER
    c.alignment = CENTER
    total_benefit_row = r
    r += 2

    write_title(ws, r, 1, 6, "5. ROI 输出", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    roi_rows = [
        ("年投入（万元）", f"=E{total_inv_row}"),
        ("年收益（万元）", f"=E{total_benefit_row}"),
        ("年净收益（万元）", f"=E{total_benefit_row}-E{total_inv_row}"),
        ("3 年累计净收益（万元）", f"=(E{total_benefit_row}-E{total_inv_row})*3"),
        ("ROI（倍数）", f"=IFERROR(E{total_benefit_row}/E{total_inv_row},0)"),
        ("回本周期（月）", f"=IFERROR(E{total_inv_row}/E{total_benefit_row}*12,0)"),
    ]
    for label, formula in roi_rows:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c1 = ws.cell(row=r, column=2, value=label)
        c1.font = BOLD
        c1.alignment = LEFT
        c1.fill = LIGHT_BLUE
        c1.border = BORDER
        c = ws.cell(row=r, column=5, value=formula)
        c.fill = LIGHT_YELLOW
        c.font = BOLD
        c.border = BORDER
        c.alignment = CENTER
        ws.cell(row=r, column=6, value="自动").border = BORDER
        ws.cell(row=r, column=6).font = NORMAL
        ws.cell(row=r, column=6).alignment = CENTER
        r += 1

    r += 1
    write_title(ws, r, 1, 6, "6. 公式说明", fill=LIGHT_BLUE, font=BOLD)
    r += 1
    for line in formula_lines:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        c = ws.cell(row=r, column=1, value=line)
        c.font = NORMAL
        c.alignment = LEFT
        c.fill = LIGHT_GRAY
        c.border = BORDER
        ws.row_dimensions[r].height = 20
        r += 1

    return total_inv_row, total_benefit_row


def build_m1(wb: Workbook) -> None:
    inputs = [
        ("年流水（万元）", "万元", 18000),
        ("当前毛利率", "%", 3.2),
        ("当前净利率", "%", 0.8),
        ("90 天+ 应收占比", "%", 38),
        ("当前年利润", "万元", 144),
        ("改造后净利率（目标）", "%", 2.1),
        ("应收下降目标（百分点）", "百分点", 20),
    ]
    investments = [
        ("ERP + AI 经营包", 18),
        ("AI 风控 + 对账通", 12),
        ("战略咨询", 5),
    ]
    # E11=年流水 E12=当前毛利率 E13=当前净利率 E14=应收 E15=年利润 E16=目标净利率 E17=应收下降
    # input_start_row = 7, so:
    # E7=年流水 / E8=毛利 / E9=净利 / E10=应收 / E11=年利润 / E12=目标净利 / E13=应收下降
    benefits = [
        ("年增加利润（净利率提升）", "=E7*(E12-E9)/100"),
        ("释放现金流（应收下降）", "=E7*E13/100"),
        ("3 年累计追加利润", "=E7*(E12-E9)/100*3"),
    ]
    formula = [
        "年增加利润 = 年流水 × (目标净利率 - 当前净利率) / 100",
        "释放现金流 = 年流水 × 应收下降百分点 / 100（仅 1 次性，但用于年收益估算）",
        "案例：王总 ¥1.8 亿 / 净利 0.8%→2.1% / 应收 38%→18% → 年增 ¥234 万利润 + 释放 ¥3,600 万",
    ]
    red_q = [
        "Q3 应收账款占比（< 3 分）",
        "Q5 净利率（< 3 分）",
        "Q6 现金流储备（< 3 分）",
    ]
    build_scene(wb, "M1 财务", "M1", "财务健康度场景（利润 + 现金流）", "山东 XX 钢管 王总", 58, red_q, inputs, investments, formula, benefits)


def build_m2(wb: Workbook) -> None:
    inputs = [
        ("年流水（万元）", "万元", 30000),
        ("当前客户流失率", "%", 14),
        ("目标客户流失率", "%", 8),
        ("销售带走客户比例（前）", "%", 60),
        ("销售带走客户比例（后）", "%", 15),
        ("单客户平均年利润", "万元", 12),
        ("月新增客户数（改造后）", "个", 6),
        ("新客户成交率", "%", 30),
        ("平均毛利率", "%", 3),
    ]
    investments = [
        ("CRM + AI 客户管理", 15),
        ("AI 销售助理", 12),
        ("现货平台", 6),
        ("战略咨询", 5),
    ]
    benefits = [
        ("减少客户流失收益", "=E7*(E8-E9)/100*E15/100"),
        ("销售跑单避免", "=(E10-E11)/100*5*E12"),
        ("新客户增量利润", "=E13*E14/100*12*E12"),
    ]
    formula = [
        "减少流失收益 = 年流水 × (旧流失率 - 新流失率) / 100 × 平均毛利率",
        "销售跑单避免 = 销售离职带走客户差额 × 5 个销售 × 客均年利润",
        "新客户增量 = 月新增 × 成交率 × 12 月 × 客均年利润",
        "案例：广东陈总 客户留存 70→92% / 营收 +30%",
    ]
    red_q = [
        "Q9 销售跑单（< 3 分）",
        "Q11 客户结构 / 房地产依赖（< 3 分）",
        "Q12 TOP 5 客户依赖（< 3 分）",
    ]
    build_scene(wb, "M2 客户", "M2", "客户管理场景（客户结构 + 销售跑单）", "广东 XX 钢贸 陈总", 52, red_q, inputs, investments, formula, benefits)


def build_m3(wb: Workbook) -> None:
    inputs = [
        ("年流水（万元）", "万元", 12000),
        ("出口潜力订单（万元/年）", "万元", 1500),
        ("出口订单毛利率", "%", 8),
        ("当前账实差", "%", 12),
        ("目标账实差", "%", 1),
        ("月报出报现状（天）", "天", 18),
        ("月报目标（天）", "天", 4),
        ("财务人天成本（元）", "元", 800),
    ]
    investments = [
        ("钢贸 ERP", 18),
        ("WMS + 质量追溯", 12),
        ("数字化咨询 + 出口审厂", 8),
    ]
    benefits = [
        ("新增出口订单利润", "=E8*E9/100"),
        ("账实差减少节省（年）", "=E7*(E10-E11)/100*0.1"),
        ("月报效率节省（年）", "=(E12-E13)*E14/10000*12"),
    ]
    formula = [
        "新增出口收益 = 出口订单 × 毛利率",
        "节省内耗 = (当前账实差 - 目标账实差) × 年流水 × 10%（按沉淀资金成本估算）",
        "月报效率 = 节省 14 天 × 财务人天成本 × 12 月",
        "案例：江苏周总 数字化 → 通过审厂 → +¥1,500 万出口订单 / 6 月",
    ]
    red_q = [
        "Q15 ERP / WMS 使用（< 3 分）",
        "Q16 数据日报（< 3 分）",
        "Q17 数字化主线人员（< 3 分）",
    ]
    build_scene(wb, "M3 数字化", "M3", "数字化能力场景（系统升级 + 出口审厂）", "江苏 XX 焊管 周总", 38, red_q, inputs, investments, formula, benefits)


def build_m4(wb: Workbook) -> None:
    inputs = [
        ("年流水（万元）", "万元", 25000),
        ("当前 90 天+ 应收占比", "%", 38),
        ("目标 90 天+ 应收占比", "%", 13),
        ("资金占用利率", "%", 7.5),
        ("历史坏账率", "%", 1.5),
        ("目标坏账率", "%", 0.5),
        ("当前银行授信（万元）", "万元", 3000),
        ("授信提升预期", "%", 50),
    ]
    investments = [
        ("AI 风控", 8),
        ("对账通 + 应收预警", 6),
        ("法务咨询", 4),
    ]
    # E7=流水 E8=旧应收 E9=新应收 E10=利率 E11=旧坏账 E12=新坏账 E13=授信 E14=提升
    benefits = [
        ("释放现金流（应收下降）", "=E7*(E8-E9)/100"),
        ("节省财务成本（年）", "=E7*(E8-E9)/100*E10/100"),
        ("避免坏账（年）", "=E7*(E11-E12)/100"),
        ("授信提升带来低息融资节省（年）", "=E13*E14/100*0.02"),
    ]
    formula = [
        "释放现金流 = 年流水 × (旧应收 - 新应收) / 100（一次性大额）",
        "节省财务成本 = 释放现金流 × 资金利率",
        "避免坏账 = 年流水 × (旧坏账 - 新坏账) / 100",
        "案例：江苏赵总 应收 38%→13% / 9 月救回 ¥4,320 万 / ROI > 100 倍",
        "★ M4 是 7 大场景中 ROI 最高、回本最快的场景，销售优先讲 M4",
    ]
    red_q = [
        "Q3 应收账款占比（最关键题）",
        "Q4 客户跑路次数（< 3 分）",
        "Q19 风险评估流程（< 3 分）",
    ]
    build_scene(wb, "M4 风控", "M4", "风险控制场景（应收 + 风控）★ 销售首选", "江苏 XX 钢贸 赵总", 45, red_q, inputs, investments, formula, benefits)


def build_m5(wb: Workbook) -> None:
    inputs = [
        ("年流水（万元）", "万元", 20000),
        ("库存周转（前）", "天", 60),
        ("库存周转（后）", "天", 35),
        ("当前年库存损失", "万元", 540),
        ("库存损失下降比例", "%", 60),
        ("月新客户数（改造后）", "个", 6),
        ("客均年利润（万元）", "万元", 8),
        ("资金成本", "%", 7),
    ]
    investments = [
        ("AI 钢价预警 + AI 库存", 8),
        ("WMS", 6),
        ("现货平台 + 拓客", 4),
    ]
    benefits = [
        ("减少库存损失", "=E10*E11/100"),
        ("释放库存现金（年化收益）", "=E7*(E8-E9)/365*E14/100"),
        ("新增客户年利润", "=E12*6*E13"),
    ]
    formula = [
        "减少库存损失 = 当前损失 × 下降比例",
        "释放库存现金 = 年流水 × (旧周转 - 新周转) / 365 × 资金成本",
        "新增客户 = 月新客户 × 6 个月留存 × 客均利润",
        "案例：浙江钱总 库存周转 60→35 天 / 释放 ¥4,000 万 / 损失 -60%",
    ]
    red_q = [
        "Q21 新客户开发（< 3 分）",
        "Q22 出口业务比例（< 3 分）",
        "Q23 新业态拓展（< 3 分）",
    ]
    build_scene(wb, "M5 增长", "M5", "增长能力场景（拓客 + 出口）", "浙江 XX 钢贸 钱总", 50, red_q, inputs, investments, formula, benefits)


def build_m6(wb: Workbook) -> None:
    inputs = [
        ("年流水（万元）", "万元", 40000),
        ("当前公司估值（万元）", "万元", 15000),
        ("估值提升目标", "%", 60),
        ("营收提升目标", "%", 25),
        ("接班失败导致流失（万元/年）", "万元", 1000),
        ("接班成功概率提升", "%", 50),
        ("毛利率", "%", 3),
    ]
    investments = [
        ("数字化全套（ERP+WMS+AI）", 50),
        ("接班顾问陪跑（12 月）", 20),
        ("二代俱乐部入会", 5),
        ("股权激励设计", 5),
    ]
    # E7=流水 E8=估值 E9=估值% E10=营收% E11=失败流失 E12=成功率% E13=毛利
    benefits = [
        ("公司估值增加", "=E8*E9/100"),
        ("营收增加（年净利润）", "=E7*E10/100*E13/100"),
        ("避免接班失败损失", "=E11*E12/100"),
    ]
    formula = [
        "估值增加 = 当前估值 × 提升比例",
        "营收增加（年净利润）= 年流水 × 提升比例 × 毛利率",
        "避免接班失败 = 失败流失 × 成功概率提升",
        "案例：江苏周总 估值 ¥1.5 亿 → ¥2.5 亿（+¥1 亿）/ 营收 +25%",
        "★ M6 是 7 大场景中单 ROI 最大（估值层级）的场景，适合 ¥3 亿+ 客户",
    ]
    red_q = [
        "Q26 销售流失率（< 3 分）",
        "Q27 接班准备（< 3 分）",
        "Q28 关键岗位备份（< 3 分）",
    ]
    build_scene(wb, "M6 团队", "M6", "团队 + 接班场景 ★ 估值杠杆", "江苏 XX 镀锌 周总 + 二代", 55, red_q, inputs, investments, formula, benefits)


def build_m7(wb: Workbook) -> None:
    inputs = [
        ("年流水（万元）", "万元", 50000),
        ("当前毛利率", "%", 1.5),
        ("目标毛利率", "%", 3),
        ("服务收入占比目标", "%", 25),
        ("客户流失率（前）", "%", 18),
        ("客户流失率（后）", "%", 5),
        ("信用数据外卖年收入（万元）", "万元", 100),
    ]
    investments = [
        ("战略咨询 + 服务化设计", 30),
        ("数字化全套", 50),
        ("信用数据接入银行", 10),
    ]
    benefits = [
        ("毛利率提升带来年增利润", "=E7*(E9-E8)/100"),
        ("服务收入", "=E7*E10/100*0.2"),
        ("客户流失下降收益", "=E7*(E11-E12)/100*E8/100"),
        ("信用数据外卖收入", "=E13"),
    ]
    formula = [
        "毛利率提升带来年增利润 = 年流水 × (目标 - 当前) / 100",
        "服务收入 = 年流水 × 服务占比 × 20%（毛利率）",
        "客户流失下降 = 年流水 × 流失率差 × 当前毛利",
        "信用数据外卖 = 与银行合作的年订阅收入",
        "案例：上海 X 总 毛利 1.5%→3% / 服务收入 0→25% / 24 月转型",
    ]
    red_q = [
        "Q29 三年战略清晰度（< 3 分）",
        "Q30 服务化转型（< 3 分）",
    ]
    build_scene(wb, "M7 战略", "M7", "战略前瞻场景（服务化 + 转型）", "上海 XX 钢贸 X 总", 42, red_q, inputs, investments, formula, benefits)


def build_summary(wb: Workbook) -> None:
    ws = wb.create_sheet("8 汇总排序")
    set_col_widths(ws, [4, 18, 22, 22, 18, 18, 18, 16])

    write_title(ws, 1, 1, 8, "7 大场景 ROI 汇总（自动从各 Tab 取数）")
    write_title(ws, 2, 1, 8, "客户红区在哪，优先推荐哪个场景（按 ROI 排序）", fill=LIGHT_BLUE, font=BOLD)

    r = 4
    header = ["#", "模块", "场景", "故事主角", "年投入（万）", "年收益（万）", "ROI（倍）", "回本（月）"]
    write_row(ws, r, header, fills=NAVY, fonts=WHITE_BOLD)
    r += 1

    scenes = [
        ("M1", "财务健康度", "山东 XX 钢管 王总", "M1 财务"),
        ("M2", "客户管理", "广东 XX 钢贸 陈总", "M2 客户"),
        ("M3", "数字化能力", "江苏 XX 焊管 周总", "M3 数字化"),
        ("M4", "风险控制 ★", "江苏 XX 钢贸 赵总", "M4 风控"),
        ("M5", "增长能力", "浙江 XX 钢贸 钱总", "M5 增长"),
        ("M6", "团队 + 接班 ★", "江苏 XX 镀锌 周总", "M6 团队"),
        ("M7", "战略前瞻", "上海 XX 钢贸 X 总", "M7 战略"),
    ]

    # 由于各场景 Tab 中 total_inv 行号不同，这里给出查阅引用的提示
    # 实际中我们在每个 Tab 里把"合计"行放在固定位置（基本相同模式）。
    # 简化：写文字说明 + 让用户回到各 Tab 看。
    for i, (code, name, hero, tab) in enumerate(scenes, 1):
        row_data = [i, code, name, hero, f"见 [{tab}] 表", f"见 [{tab}] 表", f"见 [{tab}] 表", f"见 [{tab}] 表"]
        write_row(ws, r, row_data, fills=LIGHT_GRAY if i % 2 == 0 else None)
        r += 1

    r += 2
    write_title(ws, r, 1, 8, "销售推荐顺序（默认值预填后的 ROI 排序）", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    order = [
        "1. M4 风控 → ROI 通常 > 10 倍 / 1-3 月回本 ★ 销售首选",
        "2. M6 接班 → 估值 +¥1 亿（适合 ¥3 亿+ 客户）★ 大单首选",
        "3. M1 财务 → ROI 5-8 倍 / 5-8 月回本",
        "4. M2 客户 → ROI 4-6 倍 / 6-9 月回本",
        "5. M5 增长 → ROI 4-7 倍 / 5-8 月回本",
        "6. M3 数字化 → ROI 3-5 倍 / 8-12 月回本（适合出口客户）",
        "7. M7 战略 → 长期 / 战略级（适合 ¥5 亿+ 头部客户）",
    ]
    for line in order:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        c = ws.cell(row=r, column=1, value=line)
        c.font = NORMAL
        c.alignment = LEFT
        c.fill = LIGHT_GREEN
        c.border = BORDER
        ws.row_dimensions[r].height = 22
        r += 1


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)
    build_overview(wb)
    build_m1(wb)
    build_m2(wb)
    build_m3(wb)
    build_m4(wb)
    build_m5(wb)
    build_m6(wb)
    build_m7(wb)
    build_summary(wb)
    out = "tools/钢贸7大场景ROI计算器.xlsx"
    wb.save(out)
    print(f"✓ 已生成：{out}")


if __name__ == "__main__":
    main()
