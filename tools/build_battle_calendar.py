"""
2026 年 4-6 月作战日历 Excel 甘特图生成器
==========================================

输出：tools/作战日历_2026年4到6月_甘特图.xlsx

工作表：
    1. 总览看板    —— 90 天关键里程碑 + 数字
    2. 周历甘特图  —— 12 周按工作流分类 × 每周状态
    3. 周任务清单  —— 每周具体动作 + 责任人 + 交付物
    4. 月度生死线  —— 6 月生死线 5 个指标的跟踪
    5. 90 天预算   —— 9 项费用预算
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, timedelta


THIN = Side(style="thin", color="BFBFBF")
THICK = Side(style="medium", color="1F4E79")
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
SUB_FILL = PatternFill("solid", fgColor=LIGHT_BG)
SUB_FONT = Font(name="Microsoft YaHei", size=11, bold=True)
TEXT_FONT = Font(name="Microsoft YaHei", size=10)
TITLE_FONT = Font(name="Microsoft YaHei", size=16, bold=True, color=PRIMARY)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_h(c):
    c.fill = H_FILL; c.font = H_FONT; c.alignment = CENTER; c.border = BORDER


def style_s(c, fill=LIGHT_BG):
    c.fill = PatternFill("solid", fgColor=fill); c.font = SUB_FONT
    c.alignment = CENTER; c.border = BORDER


def style_t(c, fill=None, align=None, money=False, bold=False):
    c.font = Font(name="Microsoft YaHei", size=10, bold=bold)
    c.alignment = align or CENTER; c.border = BORDER
    if fill: c.fill = PatternFill("solid", fgColor=fill)
    if money: c.number_format = '"¥"#,##0'


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# --------------------------------------------------------------------
# 12 周日期表
# --------------------------------------------------------------------
WEEKS = []
start = date(2026, 4, 20)
for i in range(12):
    s = start + timedelta(days=7*i)
    e = s + timedelta(days=6)
    WEEKS.append({"id": f"W{i+1}", "start": s, "end": e,
                  "label": f"W{i+1}\n{s.month}/{s.day}-{e.month}/{e.day}"})


# --------------------------------------------------------------------
# Sheet 1：总览看板
# --------------------------------------------------------------------
def build_dashboard(ws):
    ws.title = "总览看板"
    widths(ws, [4, 28, 22, 22, 22, 22])
    ws.merge_cells("B2:F2")
    ws["B2"] = "Plan A · 启动 90 天作战看板（2026.04.20 - 2026.07.12）"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    # KPI 卡片区
    ws.merge_cells("B4:F4")
    ws["B4"] = "★ 90 天 8 大必达 KPI（每周战报跟踪）"
    style_s(ws["B4"]); ws["B4"].alignment = LEFT

    kpis = [
        ("灯塔厂签约", "≥ 3 家", "必达"),
        ("灯塔厂上线见效", "≥ 1 家 + 视频", "必达"),
        ("老板局举办", "1 场，实到 ≥ 8", "必达"),
        ("6 月签约（不含灯塔）", "≥ 3 家", "必达"),
        ("合伙人招募", "≥ 15 个", "必达"),
        ("累计回款", "≥ ¥30 万", "必达"),
        ("销售铁三角组数", "≥ 1 组", "必达"),
        ("公众号 + 短视频", "8 篇 + 12 条", "努力"),
    ]
    headers = ["指标", "目标", "当前", "完成度", "等级"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=5, column=j, value=h))
    for i, (m, target, level) in enumerate(kpis, 6):
        style_t(ws.cell(row=i, column=2, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=target), fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=4, value=""), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=5, value=""), fill=INPUT_BG)
        fill = SUCCESS if level == "必达" else WARN
        style_t(ws.cell(row=i, column=6, value=level), fill=fill)

    # 阶段总结
    ws.merge_cells("B16:F16")
    ws["B16"] = "★ 3 阶段划分"
    style_s(ws["B16"]); ws["B16"].alignment = LEFT

    stages = [
        ("阶段 1（W1-W4，4.20-5.17）", "方案落地 + 灯塔厂启动", "签下第 1 家灯塔厂 + 第 1 场老板局邀约函发出"),
        ("阶段 2（W5-W8，5.18-6.14）", "灯塔厂见效 + 老板局倒计时", "灯塔厂 21 天 KPI 验收 + 老板局举办"),
        ("阶段 3（W9-W12，6.15-7.12）", "转化爆发 + 6 月收尾 + 生死线", "月签 ≥ 3 家 + 6 月生死线评估"),
    ]
    for i, (period, theme, milestone) in enumerate(stages, 17):
        style_t(ws.cell(row=i, column=2, value=period), fill=LIGHT_BG, align=LEFT, bold=True)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=4)
        style_t(ws.cell(row=i, column=3, value=theme), align=LEFT, fill=SUCCESS)
        ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=6)
        style_t(ws.cell(row=i, column=5, value=milestone), align=LEFT)
        ws.row_dimensions[i].height = 36

    # 老板时间分配
    ws.merge_cells("B22:F22")
    ws["B22"] = "★ 老板个人时间分配建议"
    style_s(ws["B22"]); ws["B22"].alignment = LEFT
    headers2 = ["阶段", "营销", "产品", "战略/财务", "其他"]
    for j, h in enumerate(headers2, 2):
        style_h(ws.cell(row=23, column=j, value=h))
    time_alloc = [
        ("4 月（W1-W2）", "70%", "15%", "10%", "5%"),
        ("5 月（W3-W7）", "70%", "15%", "10%", "5%"),
        ("6 月（W8-W11）", "60%", "20%", "15%", "5%"),
        ("7 月起（W12+）", "40%", "30%", "20%", "10%"),
    ]
    for i, row in enumerate(time_alloc, 24):
        for j, v in enumerate(row, 2):
            style_t(ws.cell(row=i, column=j, value=v),
                    fill=LIGHT_BG if j == 2 else SUCCESS, bold=(j == 2))

    # 备注
    ws.merge_cells("B30:F32")
    ws["B30"] = ("使用方法：每周一战报会上更新本看板的『当前/完成度』两列；\n"
                 "颜色：绿色 = 达标 / 黄色 = 警示 / 红色 = 危险；\n"
                 "决策：第 12 周（7.6-7.12）按生死线评估决定续打 Plan A / 转 Plan B / 转 Plan C")
    ws["B30"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["B30"].font = Font(name="Microsoft YaHei", size=10, color="595959")
    ws["B30"].border = BORDER


# --------------------------------------------------------------------
# Sheet 2：周历甘特图
# --------------------------------------------------------------------
def build_gantt(ws):
    ws.title = "周历甘特图"
    widths(ws, [22] + [12]*12)
    ws.merge_cells("A1:M1")
    ws["A1"] = "12 周作战甘特图"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    # 表头：周
    style_h(ws.cell(row=3, column=1, value="工作流"))
    for i, w in enumerate(WEEKS):
        style_h(ws.cell(row=3, column=2+i, value=w["label"]))
        ws.row_dimensions[3].height = 36

    # 工作流 × 周 矩阵
    # 1 = 启动 / 2 = 重点投入 / 3 = 收尾 / 0 = 不做
    streams = [
        ("方案对齐 + 法务",    [3, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
        ("灯塔厂招募调研",     [3, 3, 2, 1, 1, 1, 0, 0, 0, 0, 0, 0]),
        ("灯塔厂签约",         [0, 1, 3, 2, 2, 1, 1, 0, 1, 2, 0, 0]),
        ("灯塔厂部署/见效",    [0, 0, 0, 3, 3, 3, 3, 0, 1, 1, 0, 0]),
        ("内容产出（公众号/短视频）", [0, 0, 0, 3, 2, 3, 2, 0, 3, 2, 2, 1]),
        ("合伙人招募",         [0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]),
        ("老板局筹备",         [0, 0, 1, 2, 2, 3, 3, 0, 0, 0, 0, 0]),
        ("第 1 场老板局执行", [0, 0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0]),
        ("老板局后续转化",     [0, 0, 0, 0, 0, 0, 0, 0, 3, 3, 2, 1]),
        ("销售铁三角招聘",     [0, 1, 1, 1, 1, 1, 0, 0, 1, 1, 0, 0]),
        ("第 1 单 / 后续签约", [0, 0, 0, 0, 0, 0, 0, 0, 1, 2, 3, 1]),
        ("周战报 + 复盘",      [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 3]),
        ("月度复盘",           [0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 2, 0]),
        ("生死线评估",         [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3]),
    ]
    color_map = {
        0: WHITE, 1: LIGHT_BG, 2: ACCENT, 3: PRIMARY
    }
    label_map = {0: "", 1: "·", 2: "■", 3: "★"}
    for i, (stream, weights) in enumerate(streams, 4):
        style_t(ws.cell(row=i, column=1, value=stream),
                fill=LIGHT_BG, align=LEFT, bold=True)
        for w_idx, weight in enumerate(weights):
            cell = ws.cell(row=i, column=2+w_idx, value=label_map[weight])
            cell.font = Font(name="Microsoft YaHei", size=14, bold=True,
                             color=WHITE if weight >= 2 else "1F4E79")
            cell.alignment = CENTER
            cell.border = BORDER
            cell.fill = PatternFill("solid", fgColor=color_map[weight])
        ws.row_dimensions[i].height = 24

    # 图例
    legend_row = 4 + len(streams) + 2
    ws.cell(row=legend_row, column=1, value="图例").font = SUB_FONT
    ws.cell(row=legend_row, column=1).alignment = LEFT
    legends = [(WHITE, "0", "不做"), (LIGHT_BG, "·", "启动/收尾"),
               (ACCENT, "■", "重点投入"), (PRIMARY, "★", "里程碑")]
    for i, (color, sym, txt) in enumerate(legends):
        c = ws.cell(row=legend_row, column=2+i*2,
                    value=f"{sym} {txt}")
        c.fill = PatternFill("solid", fgColor=color)
        c.font = Font(name="Microsoft YaHei", size=10, bold=True,
                      color=WHITE if color in (ACCENT, PRIMARY) else "1F4E79")
        c.alignment = CENTER
        c.border = BORDER


# --------------------------------------------------------------------
# Sheet 3：周任务清单（详细）
# --------------------------------------------------------------------
def build_tasks(ws):
    ws.title = "周任务清单"
    widths(ws, [6, 16, 36, 16, 30])

    ws.merge_cells("A1:E1")
    ws["A1"] = "12 周精确任务清单（每周 5 大动作）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    headers = ["周", "日期", "动作", "责任人", "交付物"]
    for j, h in enumerate(headers, 1):
        style_h(ws.cell(row=3, column=j, value=h))

    # 每周任务（精简版，详见 docs/15）
    weekly_tasks = [
        # W1
        ("W1", "战略对齐会（v2.0 价格 + Plan A 节奏）", "老板", "战略决议签字版"),
        ("W1", "法务适配合同模板", "法务", "终版合同 v1.0"),
        ("W1", "财务核定价格 + 提成", "财务", "锁定报价 Excel"),
        ("W1", "2 销售总监列灯塔厂候选 30 家", "AB", "30 家清单"),
        ("W1", "首期周战报 + 接触银行 12 期分期", "全员", "周战报 + 1 家意向"),
        # W2
        ("W2", "AB 各走访 4 家灯塔厂候选（共 8 家）", "AB", "8 份拜访纪要"),
        ("W2", "老板亲自带队走访 1-2 家『梦中情人』", "老板", "顶级灯塔意向"),
        ("W2", "30 家清单按『成单概率 × 灯塔价值』排序", "老板+AB", "排序清单"),
        ("W2", "启动前 5 个合伙人（从现有满意客户）", "老板+运营", "5 个意向"),
        ("W2", "第 2 期周战报", "全员", "周战报"),
        # W3
        ("W3", "五一假期接待意向老板", "老板+AB", "拜访纪要"),
        ("W3", "第 1 家灯塔厂签约（POC 协议）", "销售总监 A", "签字 POC"),
        ("W3", "锁定第 1 场老板局时间地点（6 月第 2 周）", "老板", "时间地点"),
        ("W3", "第 1 家灯塔厂 CTO 跑现场 + 招聘第 1 销售", "CTO+HR", "实施方案 + 招聘启动"),
        ("W3", "第 3 期周战报", "全员", "周战报"),
        # W4
        ("W4", "第 1 家灯塔厂质量追溯部署（4 小时）", "CTO+实施", "扫码上线"),
        ("W4", "第 1 家灯塔厂操作工培训 3 天", "实施", "操作工掌握"),
        ("W4", "老板局邀约函发出（12 老板）", "老板+市场", "12 份邀约"),
        ("W4", "第 2 家灯塔厂 POC 启动 + 第 1 篇公众号", "B+市场", "POC 协议 + 公众号"),
        ("W4", "第 4 期周战报 + 4 月月度小结", "全员", "月报"),
        # W5
        ("W5", "第 1 家灯塔厂 7 天数据小结", "实施", "数据报告"),
        ("W5", "第 3 家灯塔厂签约推动", "AB", "POC 协议"),
        ("W5", "老板局 T-3：电话回访邀约确认", "运营", "确认表"),
        ("W5", "第 2 篇公众号 + 第 2-3 条短视频", "市场", "内容上线"),
        ("W5", "招聘第 2 销售 + 第 1 地推 + 第 5 期周战报", "HR+全员", "面试启动"),
        # W6
        ("W6", "第 1 家灯塔厂 14 天数据 + 拍视频", "实施+市场", "视频"),
        ("W6", "灯塔厂老板录 3 分钟推荐视频", "老板+市场", "视频成片"),
        ("W6", "老板局 T-2：菜单/包间/物料确认", "行政", "确认清单"),
        ("W6", "第 5 个合伙人 + 累计 10 个", "运营", "合伙人名单"),
        ("W6", "第 6 期周战报 + 5 月月度复盘", "全员", "月报"),
        # W7
        ("W7", "老板局 T-1：实到老板最终电话确认", "老板+AB", "最终名单"),
        ("W7", "第 1 家灯塔厂 21 天 KPI 验收", "实施+老板", "验收报告"),
        ("W7", "老板背 5 句金句 + 灯塔厂老板演练", "老板+灯塔", "演讲稿定稿"),
        ("W7", "现场踩点 + 物料装袋 + 全员预演", "全员", "预演完成"),
        ("W7", "第 7 期周战报 + Go/No-Go 决策", "老板", "决策"),
        # W8
        ("W8", "第 1 场老板局执行（17:30-23:00）", "全员", "老板局"),
        ("W8", "战报朋友圈 + 感谢短信", "老板+市场", "战报"),
        ("W8", "TOP 5 意向客户当周二次拜访", "老板+AB", "5 次拜访"),
        ("W8", "老板局 5 分钟纪录片剪辑", "市场", "成片"),
        ("W8", "第 8 期周战报 + T+1 复盘", "全员", "复盘报告"),
        # W9
        ("W9", "7 天内 ≥ 5 位老板二次拜访", "老板+AB", "拜访记录"),
        ("W9", "第 2 家灯塔厂正式实施", "CTO+实施", "部署完成"),
        ("W9", "老板局纪录片正式发布", "市场", "全网发布"),
        ("W9", "招聘第 3 销售 + 第 2 地推 + 售前 2", "HR", "新员工到岗"),
        ("W9", "第 1 单签约（来自老板局/合伙人）+ 第 9 期战报", "销售+全员", "签约"),
        # W10
        ("W10", "≥ 3 个客户 POC 协议同步推进", "AB+售前", "POC 协议"),
        ("W10", "第 3 家灯塔厂签约（如未签必签）", "销售总监", "灯塔合同"),
        ("W10", "第 4 篇公众号 + 第 4-5 条短视频", "市场", "内容上线"),
        ("W10", "第 2 单签约", "销售总监", "签约"),
        ("W10", "第 10 期周战报 + 月签预期复盘", "全员", "战报"),
        # W11
        ("W11", "6 月剩余签约/回款冲刺", "全员", "签约 + 回款"),
        ("W11", "第 3 单签约 + 第 4 单 POC 启动", "销售总监", "签约 + POC"),
        ("W11", "6 月月度复盘大会（半天）", "老板", "月度报告"),
        ("W11", "7 月作战日历输出（W13-W16）", "老板+AB", "下月日历"),
        ("W11", "第 11 期周战报 + 6 月战报全员发布", "全员", "战报"),
        # W12
        ("W12", "生死线评估会（达成度 ≥80% / 50-80% / <50%）", "老板", "评估结论"),
        ("W12", "评估结论 → 决定 Plan A 续打 / B 加速 / C 聚焦", "老板", "Plan 决策"),
        ("W12", "全员通气：未来 3 月调整方向", "老板", "战略宣讲"),
        ("W12", "7 月正式启动新节奏 + 第 12 期周战报", "全员", "战报"),
        ("W12", "(可选) 团队庆功会 + Q3 启动会", "全员", "凝聚力"),
    ]

    # 颜色按周分组
    week_colors = [LIGHT_BG, SUCCESS] * 6
    week_idx = -1
    last_week = ""
    for i, (week, action, owner, deliv) in enumerate(weekly_tasks, 4):
        if week != last_week:
            week_idx += 1
            last_week = week
        bg = week_colors[week_idx % 12]
        # 周列
        style_t(ws.cell(row=i, column=1, value=week), fill=bg, bold=True)
        # 日期列（W 编号查表）
        w_num = int(week[1:]) - 1
        date_str = f"{WEEKS[w_num]['start'].month}/{WEEKS[w_num]['start'].day}-{WEEKS[w_num]['end'].month}/{WEEKS[w_num]['end'].day}"
        style_t(ws.cell(row=i, column=2, value=date_str), fill=bg)
        style_t(ws.cell(row=i, column=3, value=action), fill=bg, align=LEFT)
        style_t(ws.cell(row=i, column=4, value=owner), fill=bg)
        style_t(ws.cell(row=i, column=5, value=deliv), fill=bg, align=LEFT)
        ws.row_dimensions[i].height = 26


# --------------------------------------------------------------------
# Sheet 4：6 月生死线
# --------------------------------------------------------------------
def build_deadline(ws):
    ws.title = "6月生死线"
    widths(ws, [4, 26, 18, 18, 18, 28])
    ws.merge_cells("B2:F2")
    ws["B2"] = "6 月生死线（W12 评估表）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["指标", "目标", "实际", "达成 %", "未达成对策"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    rows = [
        ("灯塔厂签约", "≥ 3 家", "", "", "调整价格 / 实施资源 / 灯塔筛选"),
        ("灯塔厂上线见效", "≥ 1 家公开案例", "", "", "CTO 必须亲自上阵第 2 家"),
        ("老板局举办", "≥ 1 场", "", "", "立刻办，不再等"),
        ("月签客户（不含灯塔）", "≥ 3 家", "", "", "复盘销售铁三角配置和 SOP 落地度"),
        ("合伙人招募", "≥ 15 个", "", "", "销售总监个人 KPI 加合伙人项"),
        ("累计回款", "≥ ¥30 万", "", "", "现金流告急 → 启动 Plan C"),
    ]
    for i, (m, t, real, pct, strategy) in enumerate(rows, 5):
        style_t(ws.cell(row=i, column=2, value=m), fill=LIGHT_BG, bold=True, align=LEFT)
        style_t(ws.cell(row=i, column=3, value=t), fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=4, value=real), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=5, value=pct), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=6, value=strategy), align=LEFT)

    # 评估决策
    ws.merge_cells("B12:F12")
    ws["B12"] = "评估决策（圈出 1 个）"
    style_s(ws["B12"]); ws["B12"].alignment = LEFT

    decisions = [
        ("达成 ≥ 80%", "Plan A 续打", "可考虑加速到 Plan B 节奏", SUCCESS),
        ("达成 50-80%", "Plan A 续打 + 调整短板", "重点强化最弱的 2 个指标", WARN),
        ("达成 < 50%", "立刻转 Plan C 聚焦", "只做 1-2 个客群、2 个产业带", CRITICAL),
        ("达成 < 30%", "商业模式根本问题", "回头看产品/价格/客群定位，不是营销问题", CRITICAL),
    ]
    headers2 = ["达成度", "决策", "动作", ""]
    for j, h in enumerate(headers2[:3], 2):
        style_h(ws.cell(row=13, column=j, value=h))
    for i, (level, dec, action, color) in enumerate(decisions, 14):
        style_t(ws.cell(row=i, column=2, value=level), fill=color, bold=True)
        style_t(ws.cell(row=i, column=3, value=dec), fill=color, bold=True)
        ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=6)
        style_t(ws.cell(row=i, column=4, value=action), fill=color, align=LEFT)
        ws.row_dimensions[i].height = 30


# --------------------------------------------------------------------
# Sheet 5：90 天预算
# --------------------------------------------------------------------
def build_budget(ws):
    ws.title = "90天预算"
    widths(ws, [4, 30, 18, 50])
    ws.merge_cells("B2:D2")
    ws["B2"] = "90 天总预算（含人力、灯塔补贴、老板局、内容、招聘等）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["项", "金额（元）", "备注"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    items = [
        ("团队薪酬 4-6 月（按 12 人）", 600000, "月均 ¥200K"),
        ("灯塔厂补贴（3 家 × ¥3 万）", 90000, "实施费打折部分"),
        ("第 1 场老板局费用", 60000, "佛山乐从 / 山东聊城 任选"),
        ("内容制作（视频 + 公众号 + 白皮书）", 30000, "3 个月"),
        ("物料 + 印刷（白皮书 + 礼品 + 合同夹）", 30000, ""),
        ("招聘费用（猎头 + 招聘网站）", 30000, ""),
        ("差旅（老板 + AB 全国跑灯塔厂）", 40000, ""),
        ("工具 + 系统（飞书 + 腾讯文档 + CRM）", 10000, ""),
        ("应急 / 杂费", 30000, ""),
    ]
    total = 0
    for i, (name, amount, note) in enumerate(items, 5):
        style_t(ws.cell(row=i, column=2, value=name), fill=LIGHT_BG, align=LEFT)
        style_t(ws.cell(row=i, column=3, value=amount), fill=SUCCESS, money=True)
        style_t(ws.cell(row=i, column=4, value=note), align=LEFT)
        total += amount

    # 合计
    i = 5 + len(items)
    style_t(ws.cell(row=i, column=2, value="90 天总投入"),
            fill=PRIMARY, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11,
                                          bold=True, color=WHITE)
    style_t(ws.cell(row=i, column=3, value=total),
            fill=PRIMARY, money=True, bold=True)
    ws.cell(row=i, column=3).font = Font(name="Microsoft YaHei", size=11,
                                          bold=True, color=WHITE)
    style_t(ws.cell(row=i, column=4, value="月均 ¥30 万"), align=LEFT)

    # 预期回款
    i += 2
    ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
    style_t(ws.cell(row=i, column=2,
                    value="预期 90 天回款：¥30-50 万 → 净投入 ¥40-60 万"),
            fill=SUCCESS, align=LEFT, bold=True)


def main():
    wb = Workbook()
    build_dashboard(wb.active)
    build_gantt(wb.create_sheet())
    build_tasks(wb.create_sheet())
    build_deadline(wb.create_sheet())
    build_budget(wb.create_sheet())

    out = "tools/作战日历_2026年4到6月_甘特图.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
