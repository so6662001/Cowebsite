"""
6 个月生死线周战报跟踪表
========================

输出：tools/周战报_6个月生死线跟踪表.xlsx

工作表：
    1. 使用说明
    2. 当周战报模板        —— 单周一页式（可复制使用）
    3. 24 周数据看板       —— 5 KPI × 24 周 矩阵
    4. 生死线进度         —— 自动累计 + 风险评级
    5. 月度战报           —— 6 个月汇总
    6. W6 中期评估
    7. W12 生死线评估
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from datetime import date, timedelta


THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
PRIMARY = "1F4E79"
ACCENT = "2E75B6"
LIGHT_BG = "DEEBF7"
INPUT_BG = "FFF2CC"
SUCCESS = "E2EFDA"
WARN = "FCE4D6"
CRITICAL = "FFB6C1"
WHITE = "FFFFFF"

H_FILL = PatternFill("solid", fgColor=PRIMARY)
H_FONT = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
SUB_FONT = Font(name="Microsoft YaHei", size=11, bold=True)
TITLE_FONT = Font(name="Microsoft YaHei", size=16, bold=True, color=PRIMARY)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_h(c):
    c.fill = H_FILL; c.font = H_FONT; c.alignment = CENTER; c.border = BORDER


def style_t(c, fill=None, align=None, money=False, bold=False):
    c.font = Font(name="Microsoft YaHei", size=10, bold=bold)
    c.alignment = align or CENTER; c.border = BORDER
    if fill: c.fill = PatternFill("solid", fgColor=fill)
    if money: c.number_format = '"¥"#,##0'


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_readme(ws):
    ws.title = "使用说明"
    widths(ws, [4, 26, 70])
    ws.merge_cells("B2:C2")
    ws["B2"] = "6 个月生死线 · 周战报跟踪表（v1.0）"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    rows = [
        ("适用对象", "老板（主持）+ 销售运营经理（记录）+ 全员（出席）"),
        ("使用步骤",
         "① 每周日晚由销售运营经理整理上周数据，填入『24 周数据看板』黄色单元格\n"
         "② 周一 09:00 战报会上对照『当周战报模板』5+3+1 模型现场过\n"
         "③ 决策结果写入战报模板的『本周决策』栏\n"
         "④ 月底切换『月度战报』看 4 周累计趋势\n"
         "⑤ W6 / W12 切换『中期评估』『生死线评估』做关键决策"),
        ("5 个核心 KPI",
         "1. 新增线索数（周）\n"
         "2. 新增 POC 数（周）\n"
         "3. 新增签约数（周）\n"
         "4. 当周回款（周）\n"
         "5. 新增合伙人数（周）"),
        ("生死线 5 项",
         "灯塔厂签约 ≥ 3 / 灯塔厂上线 ≥ 1 / 老板局 ≥ 1 / 月签 ≥ 3 / 累计回款 ≥ ¥30 万"),
        ("会议铁律",
         "1. 每周一 09:00-09:30，30 分钟严格控时\n"
         "2. 老板亲自主持\n"
         "3. 5 数字 + 3 故事 + 1 决策\n"
         "4. 决策必须写下来 + 下周回看\n"
         "5. 失败案例不追责，只追教训"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        c1 = ws.cell(row=i, column=2, value=k); style_t(c1, fill=PRIMARY, bold=True)
        c1.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
        c2 = ws.cell(row=i, column=3, value=v); style_t(c2, align=LEFT)
        ws.row_dimensions[i].height = max(50, len(v) * 0.8)


def build_template(ws):
    ws.title = "当周战报模板"
    widths(ws, [4, 18, 14, 14, 14, 14, 14])
    ws.merge_cells("B2:G2")
    ws["B2"] = "当周战报模板（每周一战报会用，可复制 / 打印）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    # 期号 + 日期
    style_t(ws.cell(row=4, column=2, value="期号"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=4, column=3, value="第 ___ 期"), fill=INPUT_BG)
    style_t(ws.cell(row=4, column=4, value="日期"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=4, column=5, value="2026.__.__"), fill=INPUT_BG)
    style_t(ws.cell(row=4, column=6, value="主持"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=4, column=7, value="老板"), fill=INPUT_BG)

    # 一、5 个数字
    ws.merge_cells("B6:G6"); ws["B6"] = "▼ 一、5 个核心数字"
    style_t(ws["B6"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B6"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    headers = ["指标", "上周实际", "上周目标", "本周目标", "累计达成", "总目标"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=7, column=j, value=h))

    metrics = [
        ("新增线索数",     "", "", "", "", ""),
        ("新增 POC 数",   "", "", "", "", ""),
        ("新增签约数",     "", "", "", "", ""),
        ("当周回款",      "", "", "", "", ""),
        ("新增合伙人数",   "", "", "", "", ""),
    ]
    for i, m in enumerate(metrics, 8):
        style_t(ws.cell(row=i, column=2, value=m[0]), fill=LIGHT_BG, bold=True)
        for j in range(3, 8):
            style_t(ws.cell(row=i, column=j), fill=INPUT_BG)

    # 二、3 个故事
    ws.merge_cells("B14:G14"); ws["B14"] = "▼ 二、3 个故事"
    style_t(ws["B14"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B14"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    stories = [
        ("故事 1（亮点客户）", "公司 + 老板", "当周发生", "下周动作"),
        ("故事 2（失败客户）", "公司 + 老板", "失败原因", "教训"),
        ("故事 3（学到的事）", "讲述人", "学到什么", "沉淀到"),
    ]
    for i, (title, c1, c2, c3) in enumerate(stories, 15):
        style_t(ws.cell(row=i, column=2, value=title),
                fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=c1), fill=LIGHT_BG)
        ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=5)
        style_t(ws.cell(row=i, column=4), fill=INPUT_BG, align=LEFT)
        ws.merge_cells(start_row=i, start_column=6, end_row=i, end_column=7)
        style_t(ws.cell(row=i, column=6), fill=INPUT_BG, align=LEFT)
        ws.row_dimensions[i].height = 50

    # 三、生死线进度
    ws.merge_cells("B19:G19"); ws["B19"] = "▼ 三、6 月生死线进度（自动从『生死线进度』sheet 引用）"
    style_t(ws["B19"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B19"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    deadlines = [
        ("[1] 灯塔厂签约", "目标 ≥ 3"),
        ("[2] 灯塔厂上线", "目标 ≥ 1"),
        ("[3] 老板局", "目标 ≥ 1"),
        ("[4] 月签（非灯塔）", "目标 ≥ 3"),
        ("[5] 累计回款", "目标 ≥ ¥30 万"),
    ]
    style_h(ws.cell(row=20, column=2, value="指标"))
    style_h(ws.cell(row=20, column=3, value="目标"))
    style_h(ws.cell(row=20, column=4, value="累计"))
    style_h(ws.cell(row=20, column=5, value="风险"))
    ws.merge_cells("F20:G20"); style_h(ws.cell(row=20, column=6, value="本周新增"))
    for i, (m, target) in enumerate(deadlines, 21):
        style_t(ws.cell(row=i, column=2, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=target), fill=SUCCESS)
        style_t(ws.cell(row=i, column=4), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=5), fill=INPUT_BG)
        ws.merge_cells(start_row=i, start_column=6, end_row=i, end_column=7)
        style_t(ws.cell(row=i, column=6), fill=INPUT_BG)

    # 四、本周决策
    ws.merge_cells("B27:G27"); ws["B27"] = "▼ 四、本周老板决策"
    style_t(ws["B27"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B27"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    style_t(ws.cell(row=28, column=2, value="决策"), fill=LIGHT_BG, bold=True)
    ws.merge_cells("C28:G28")
    style_t(ws.cell(row=28, column=3), fill=INPUT_BG, align=LEFT)
    style_t(ws.cell(row=29, column=2, value="责任人"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=29, column=3), fill=INPUT_BG)
    style_t(ws.cell(row=29, column=4, value="截止日"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=29, column=5), fill=INPUT_BG)
    ws.merge_cells("F29:G29")
    style_t(ws.cell(row=29, column=6), fill=INPUT_BG)
    style_t(ws.cell(row=30, column=2, value="预期结果"), fill=LIGHT_BG, bold=True)
    ws.merge_cells("C30:G30")
    style_t(ws.cell(row=30, column=3), fill=INPUT_BG, align=LEFT)

    # 五、上周决策回顾
    ws.merge_cells("B32:G32"); ws["B32"] = "▼ 五、上周决策回顾"
    style_t(ws["B32"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B32"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    style_t(ws.cell(row=33, column=2, value="上周决策"), fill=LIGHT_BG, bold=True)
    ws.merge_cells("C33:E33"); style_t(ws.cell(row=33, column=3), fill=INPUT_BG, align=LEFT)
    style_t(ws.cell(row=33, column=6, value="完成情况"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=33, column=7), fill=INPUT_BG)


def build_dashboard(ws):
    ws.title = "24周数据看板"
    widths(ws, [22] + [9]*24 + [10])
    ws.merge_cells("A1:Z1")
    ws["A1"] = "24 周数据看板（黄色 = 输入，绿色 = 自动累计）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    # 周表头
    style_h(ws.cell(row=3, column=1, value="指标"))
    start = date(2026, 4, 20)
    for i in range(24):
        s = start + timedelta(days=7*i)
        style_h(ws.cell(row=3, column=2+i,
                        value=f"W{i+1}\n{s.month}/{s.day}"))
    style_h(ws.cell(row=3, column=26, value="累计"))
    ws.row_dimensions[3].height = 32

    # 5 KPI 行
    metrics = [
        "新增线索数",
        "新增 POC 数",
        "新增签约数",
        "当周回款（元）",
        "新增合伙人数",
    ]
    for i, m in enumerate(metrics, 4):
        style_t(ws.cell(row=i, column=1, value=m), fill=LIGHT_BG, bold=True, align=LEFT)
        for w in range(24):
            cell = ws.cell(row=i, column=2+w)
            style_t(cell, fill=INPUT_BG)
            if i == 7:
                cell.number_format = '"¥"#,##0'
        cell = ws.cell(row=i, column=26, value=f"=SUM(B{i}:Y{i})")
        style_t(cell, fill=SUCCESS, bold=True)
        if i == 7:
            cell.number_format = '"¥"#,##0'

    # 留空行
    # 月度小计行
    style_t(ws.cell(row=10, column=1, value="── 月度小计（4 周）──"),
            fill=PRIMARY, align=LEFT)
    ws.cell(row=10, column=1).font = Font(name="Microsoft YaHei",
                                           size=10, bold=True, color=WHITE)

    months = [
        ("4 月（W1-W2）", 0, 1),
        ("5 月（W3-W6）", 2, 5),
        ("6 月（W7-W11）", 6, 10),
        ("7 月（W12-W15）", 11, 14),
        ("8 月（W16-W19）", 15, 18),
        ("9 月（W20-W24）", 19, 23),
    ]
    for r, m_idx in enumerate(metrics, 4):
        for col_offset, (mname, ws_, we_) in enumerate(months):
            col = 11 + col_offset  # 这里只是占位，复杂月度逻辑放下表

    # 月度统计放在另一个区域
    row = 12
    style_h(ws.cell(row=row, column=1, value="月度统计"))
    for i, (mname, ws_, we_) in enumerate(months):
        style_h(ws.cell(row=row, column=2+i, value=mname))

    for r_idx, m in enumerate(metrics, 1):
        row_num = 12 + r_idx
        style_t(ws.cell(row=row_num, column=1, value=m),
                fill=LIGHT_BG, align=LEFT, bold=True)
        for c_idx, (mname, ws_, we_) in enumerate(months):
            start_col = get_column_letter(2 + ws_)
            end_col = get_column_letter(2 + we_)
            cell = ws.cell(row=row_num, column=2+c_idx,
                          value=f"=SUM({start_col}{r_idx+3}:{end_col}{r_idx+3})")
            style_t(cell, fill=SUCCESS, bold=True)
            if r_idx == 4:  # 回款
                cell.number_format = '"¥"#,##0'

    # 冻结
    ws.freeze_panes = "B4"


def build_deadline_progress(ws):
    ws.title = "生死线进度"
    widths(ws, [4, 22, 14, 14, 14, 14, 28])
    ws.merge_cells("B2:G2")
    ws["B2"] = "6 月生死线 5 项进度（自动累计 + 风险评级）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["指标", "目标值", "当前累计", "达成 %", "风险等级", "未达成对策"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    items = [
        ("[1] 灯塔厂签约（≥ 3）", 3, 0, "灯塔厂签约数手填", "调整价格 / 实施资源 / 灯塔筛选"),
        ("[2] 灯塔厂上线见效（≥ 1）", 1, 0, "公开案例数手填", "CTO 必须亲自上阵第 2 家"),
        ("[3] 老板局举办（≥ 1）", 1, 0, "已办场次手填", "立刻办，不再等"),
        ("[4] 月签客户（≥ 3）", 3, 0, "本月签约数手填", "复盘销售铁三角 + SOP 落地度"),
        ("[5] 累计回款（≥ ¥30 万）", 300000, 0, "回款累计手填", "现金流告急 → 启动 Plan C"),
    ]
    for i, (m, target, current, source, action) in enumerate(items, 5):
        style_t(ws.cell(row=i, column=2, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=target), fill=SUCCESS, bold=True)
        if i == 9:  # 回款
            ws.cell(row=i, column=3).number_format = '"¥"#,##0'
        style_t(ws.cell(row=i, column=4, value=current), fill=INPUT_BG)
        if i == 9:
            ws.cell(row=i, column=4).number_format = '"¥"#,##0'
        # 达成 %
        style_t(ws.cell(row=i, column=5,
                value=f'=IFERROR(D{i}/C{i},0)'), fill=SUCCESS, bold=True)
        ws.cell(row=i, column=5).number_format = "0%"
        # 风险等级（自动）
        style_t(ws.cell(row=i, column=6,
                value=f'=IF(E{i}>=1,"🟢 达成",IF(E{i}>=0.5,"🟡 警示","🔴 危险"))'),
                fill=INPUT_BG, bold=True)
        style_t(ws.cell(row=i, column=7, value=action), align=LEFT)

    # 条件格式
    risk_col = "F"
    ws.conditional_formatting.add(f"{risk_col}5:{risk_col}9",
        CellIsRule(operator="containsText", formula=['"达成"'],
                   fill=PatternFill("solid", fgColor="00B050")))
    ws.conditional_formatting.add(f"{risk_col}5:{risk_col}9",
        CellIsRule(operator="containsText", formula=['"警示"'],
                   fill=PatternFill("solid", fgColor="FFC000")))
    ws.conditional_formatting.add(f"{risk_col}5:{risk_col}9",
        CellIsRule(operator="containsText", formula=['"危险"'],
                   fill=PatternFill("solid", fgColor="FF6B6B")))


def build_w6_eval(ws):
    ws.title = "W6中期评估"
    widths(ws, [4, 24, 12, 12, 14, 28])
    ws.merge_cells("B2:F2")
    ws["B2"] = "W6 中期评估（5 月底，看是否需要提前干预）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["指标", "W6 应达成", "我的实际", "是否预警", "应对措施"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    items = [
        ("灯塔厂候选拜访",     "≥ 16 家", "调整销售总监工作分配"),
        ("灯塔厂签约 POC",    "≥ 1 家",  "老板亲自跑 + 价格让到 5 折"),
        ("灯塔厂部署完成",    "≥ 1 家",  "CTO 亲自上阵 + 实施加资源"),
        ("累计合伙人",        "≥ 8 个",  "销售总监个人 KPI 加合伙人项"),
        ("累计线索",          "≥ 200 条","加大地推 / 内容 / 老板局密度"),
    ]
    for i, (m, target, action) in enumerate(items, 5):
        style_t(ws.cell(row=i, column=2, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=target), fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=4), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=5), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=6, value=action), align=LEFT)

    # 决策矩阵
    ws.merge_cells("B12:F12")
    ws["B12"] = "★ W6 决策矩阵"
    style_t(ws["B12"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B12"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    decisions = [
        ("5/5 达成",  "优秀",  "加速节奏，考虑 Plan B", SUCCESS),
        ("4/5 达成",  "健康",  "Plan A 续打，强化最弱项", LIGHT_BG),
        ("3/5 达成",  "警示",  "立刻投入老板 80% 时间", WARN),
        ("≤ 2/5",    "危险",  "暂停扩张，重新对齐打法", CRITICAL),
    ]
    style_h(ws.cell(row=13, column=2, value="达成数"))
    style_h(ws.cell(row=13, column=3, value="评估"))
    ws.merge_cells("D13:F13"); style_h(ws.cell(row=13, column=4, value="行动"))
    for i, (lvl, eva, act, color) in enumerate(decisions, 14):
        style_t(ws.cell(row=i, column=2, value=lvl), fill=color, bold=True)
        style_t(ws.cell(row=i, column=3, value=eva), fill=color, bold=True)
        ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=6)
        style_t(ws.cell(row=i, column=4, value=act), fill=color, align=LEFT)


def build_w12_eval(ws):
    ws.title = "W12生死线评估"
    widths(ws, [4, 26, 14, 14, 14, 28])
    ws.merge_cells("B2:F2")
    ws["B2"] = "W12 生死线评估（2026.07.06 决策会）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    # 5 项评估
    headers = ["KPI", "目标", "实际", "达成", "等级"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    items = [
        ("灯塔厂签约",   3,   0,  "S/A/B"),
        ("灯塔厂上线",   1,   0,  "S/A/B"),
        ("老板局举办",   1,   0,  "S/A/B"),
        ("月签（非灯塔）", 3,   0,  "S/A/B"),
        ("累计回款（万）", 30,  0,  "S/A/B"),
    ]
    for i, (m, target, real, level) in enumerate(items, 5):
        style_t(ws.cell(row=i, column=2, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=target), fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=4, value=real), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=5,
                value=f'=IFERROR(D{i}/C{i},0)'), fill=SUCCESS, bold=True)
        ws.cell(row=i, column=5).number_format = "0%"
        style_t(ws.cell(row=i, column=6,
                value=f'=IF(E{i}>=0.8,"S 达成",IF(E{i}>=0.5,"A 警示","B 危险"))'),
                fill=INPUT_BG, bold=True)

    # 总体评估
    ws.merge_cells("B11:F11")
    ws["B11"] = "★ 总体决策（圈出 1 个）"
    style_t(ws["B11"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B11"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    decisions = [
        ("达成 ≥ 80%",   "Plan A 续打 / 加速到 Plan B",  SUCCESS),
        ("达成 50-80%", "Plan A 续打 + 调整短板",        WARN),
        ("达成 < 50%",  "立刻转 Plan C 聚焦",            CRITICAL),
        ("达成 < 30%",  "回头看产品/价格/客群定位（不是营销问题）", CRITICAL),
    ]
    for i, (lvl, dec, color) in enumerate(decisions, 12):
        style_t(ws.cell(row=i, column=2, value=lvl), fill=color, bold=True)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=6)
        style_t(ws.cell(row=i, column=3, value=dec), fill=color, align=LEFT, bold=True)

    # 红线条款
    ws.merge_cells("B17:F17")
    ws["B17"] = "★ 红线条款（触发即执行）"
    style_t(ws["B17"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B17"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    redlines = [
        "立刻换销售总监", "该总监 12 周 0 签约 + 0 灯塔厂",
        "立刻换 CMO（如已有）", "CMO 6 月 KPI < 60%",
        "立刻启动融资", "现金 < 3 月运营成本",
        "立刻关 1 个客群", "该客群 12 周 0 签约",
        "立刻关 1 个城市", "该城市无销售产出",
    ]
    for i in range(0, len(redlines), 2):
        row = 18 + i // 2
        style_t(ws.cell(row=row, column=2, value=redlines[i]),
                fill=CRITICAL, align=LEFT, bold=True)
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
        style_t(ws.cell(row=row, column=3, value=redlines[i+1]),
                fill=WARN, align=LEFT)


def build_monthly(ws):
    ws.title = "月度战报"
    widths(ws, [22] + [12]*6 + [12])
    ws.merge_cells("A1:H1")
    ws["A1"] = "月度战报（6 个月汇总）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    headers = ["指标", "4 月", "5 月", "6 月", "7 月", "8 月", "9 月", "6 月累计"]
    for j, h in enumerate(headers, 1):
        style_h(ws.cell(row=3, column=j, value=h))

    metrics = [
        "新增线索",
        "新增 POC",
        "新增签约",
        "当月回款",
        "新增合伙人",
        "灯塔厂签约（累计）",
        "灯塔厂上线（累计）",
        "老板局场次（累计）",
    ]
    for i, m in enumerate(metrics, 4):
        style_t(ws.cell(row=i, column=1, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        for c in range(2, 8):
            style_t(ws.cell(row=i, column=c), fill=INPUT_BG)
            if i == 7:  # 回款
                ws.cell(row=i, column=c).number_format = '"¥"#,##0'
        # 累计
        cell = ws.cell(row=i, column=8,
                      value=f"=SUM(B{i}:G{i})")
        style_t(cell, fill=SUCCESS, bold=True)
        if i == 7:
            cell.number_format = '"¥"#,##0'

    # 评分行
    ws.cell(row=13, column=1, value="月度自评（A/B/C）").font = SUB_FONT
    ws.cell(row=13, column=1).fill = PatternFill("solid", fgColor=LIGHT_BG)
    ws.cell(row=13, column=1).alignment = LEFT
    ws.cell(row=13, column=1).border = BORDER
    for c in range(2, 8):
        style_t(ws.cell(row=13, column=c), fill=INPUT_BG, bold=True)


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_template(wb.create_sheet())
    build_dashboard(wb.create_sheet())
    build_deadline_progress(wb.create_sheet())
    build_monthly(wb.create_sheet())
    build_w6_eval(wb.create_sheet())
    build_w12_eval(wb.create_sheet())

    out = "tools/周战报_6个月生死线跟踪表.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
