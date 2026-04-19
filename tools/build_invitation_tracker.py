"""
老板局邀约跟踪表
================

输出：tools/老板局邀约_跟踪表.xlsx

工作表：
    1. 使用说明
    2. 邀约总进度看板        —— 自动统计
    3. 12 位老板跟踪表        —— 主表
    4. 5 周时间轴            —— W3-W8 每周动作
    5. 现场座位安排           —— 视觉化布局
    6. T+7 后续跟进表
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule


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


def style_t(c, fill=None, align=None, bold=False):
    c.font = Font(name="Microsoft YaHei", size=10, bold=bold)
    c.alignment = align or CENTER; c.border = BORDER
    if fill: c.fill = PatternFill("solid", fgColor=fill)


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_readme(ws):
    ws.title = "使用说明"
    widths(ws, [4, 26, 70])
    ws.merge_cells("B2:C2")
    ws["B2"] = "第 1 场老板局邀约跟踪表（v1.0）"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    rows = [
        ("适用对象", "老板 + 销售总监 A/B + 销售运营经理"),
        ("使用步骤",
         "① 打开『12 位老板跟踪表』，填写 12 位老板基本信息（按必到/力争/加分分档）\n"
         "② 每周一战报会更新「最近联系」「确认状态」两列\n"
         "③ 切换『邀约总进度看板』看自动汇总\n"
         "④ T+7 后切换『后续跟进表』管理转化"),
        ("核心原则",
         "1. 邀 12 实到 8-10 是行业正常水平\n"
         "2. 必到 4 位由老板亲自邀\n"
         "3. T-3 周由灯塔厂老板亲自电话最关键\n"
         "4. T-1 周老板必须再亲自打 1 遍"),
        ("时间锚点",
         "T-5 周 = W3（5.4-5.10）锁定 + 圈定名单\n"
         "T-4 周 = W4（5.11-5.17）首次邀约\n"
         "T-3 周 = W5（5.18-5.24）跟进未确认\n"
         "T-2 周 = W6（5.25-5.31）实体邀请函\n"
         "T-1 周 = W7（6.1-6.7）最终确认 + 现场预演\n"
         "T  日 = W8（6.11 周四）老板局执行\n"
         "T+1-T+7 后续跟进"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        c1 = ws.cell(row=i, column=2, value=k); style_t(c1, fill=PRIMARY, bold=True)
        c1.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
        c2 = ws.cell(row=i, column=3, value=v); style_t(c2, align=LEFT)
        ws.row_dimensions[i].height = max(40, len(v) * 0.7)


def build_dashboard(ws):
    ws.title = "邀约总进度看板"
    widths(ws, [4, 22, 12, 12, 12, 28])
    ws.merge_cells("B2:F2")
    ws["B2"] = "邀约总进度看板（自动统计 12 位老板表）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["状态", "目标", "实际", "达成", "下一步动作"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    statuses = [
        ("名单已圈定",   12,  "T-5 周完成"),
        ("邀请函已发",   12,  "T-4 周完成"),
        ("已电话邀约",   12,  "T-4 周完成"),
        ("已确认实到",   8,   "T-1 周必达 ≥ 8"),
        ("待定中",       2,   "T-3 周必清"),
        ("已婉拒",       2,   "启动备选名单"),
        ("已回访（T-1）", 12, "T-1 周必达 = 12"),
    ]
    for i, (status, target, action) in enumerate(statuses, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=status), fill=bg, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=target), fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=4,
                value=f'=COUNTIF(\'12位老板跟踪表\'!K4:K15,"{status}")'),
                fill=bg, bold=True)
        # 达成度
        style_t(ws.cell(row=i, column=5,
                value=f'=IFERROR(D{i}/C{i},0)'), fill=bg)
        ws.cell(row=i, column=5).number_format = "0%"
        style_t(ws.cell(row=i, column=6, value=action), fill=bg, align=LEFT)

    # 关键 KPI
    ws.merge_cells("B14:F14")
    ws["B14"] = "★ 老板局 4 大成功 KPI"
    style_t(ws["B14"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B14"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    kpis = [
        ("实到老板数",       "≥ 8 位",    "如 < 6 位 → 临时启动备选"),
        ("现场添加微信",      "≥ 8 个",   "现场销售必须每个老板加微信"),
        ("T+7 二次拜访",      "≥ 5 位",   "立刻上门"),
        ("90 天内合同",      "≥ 3 单",    "POC → 签约的标准转化"),
    ]
    for i, (kpi, target, action) in enumerate(kpis, 15):
        style_t(ws.cell(row=i, column=2, value=kpi), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=target), fill=SUCCESS, bold=True)
        ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=5)
        style_t(ws.cell(row=i, column=4, value=""), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=6, value=action), align=LEFT)


def build_main(ws):
    ws.title = "12位老板跟踪表"
    widths(ws, [5, 8, 22, 12, 8, 14, 22, 14, 12, 14, 14, 22])
    ws.merge_cells("A1:L1")
    ws["A1"] = "12 位老板邀约跟踪表（黄色 = 输入，绿色 = 自动）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    headers = [
        "#", "档位", "公司", "老板姓名", "年龄",
        "城市", "客群", "邀约负责人", "首次邀约日",
        "最近联系", "确认状态", "备注 / 风险"
    ]
    for j, h in enumerate(headers, 1):
        style_h(ws.cell(row=3, column=j, value=h))
    ws.row_dimensions[3].height = 32

    # 12 行
    for i in range(4, 16):
        idx = i - 3
        # 自动序号
        style_t(ws.cell(row=i, column=1, value=idx), fill=LIGHT_BG, bold=True)
        # 输入字段
        for col in range(2, 13):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG, align=LEFT)

    # 数据有效性
    dv_arch = DataValidation(type="list",
        formula1='"必到,力争,加分,备选"', allow_blank=True)
    ws.add_data_validation(dv_arch); dv_arch.add("B4:B15")

    dv_kq = DataValidation(type="list",
        formula1='"钢管,镀锌冷卷,开平纵剪,一体厂,协会"', allow_blank=True)
    ws.add_data_validation(dv_kq); dv_kq.add("G4:G15")

    dv_owner = DataValidation(type="list",
        formula1='"老板,销售总监A,销售总监B,销售运营,灯塔厂老板"', allow_blank=True)
    ws.add_data_validation(dv_owner); dv_owner.add("H4:H15")

    dv_status = DataValidation(type="list",
        formula1='"名单已圈定,邀请函已发,已电话邀约,已确认实到,待定中,已婉拒,已回访（T-1）"',
        allow_blank=True)
    ws.add_data_validation(dv_status); dv_status.add("K4:K15")

    # 条件格式：状态颜色
    status_col = "K"
    ws.conditional_formatting.add(f"{status_col}4:{status_col}15",
        CellIsRule(operator="equal", formula=['"已确认实到"'],
                   fill=PatternFill("solid", fgColor="00B050")))
    ws.conditional_formatting.add(f"{status_col}4:{status_col}15",
        CellIsRule(operator="equal", formula=['"待定中"'],
                   fill=PatternFill("solid", fgColor="FFC000")))
    ws.conditional_formatting.add(f"{status_col}4:{status_col}15",
        CellIsRule(operator="equal", formula=['"已婉拒"'],
                   fill=PatternFill("solid", fgColor="FF6B6B")))
    ws.conditional_formatting.add(f"{status_col}4:{status_col}15",
        CellIsRule(operator="equal", formula=['"已回访（T-1）"'],
                   fill=PatternFill("solid", fgColor="00B0F0")))

    # 示例行（4 位必到）
    samples = [
        ("必到", "灯塔厂A", "张 XX", 55, "邯郸", "钢管", "老板", "5/11", "5/14", "已确认实到", "主讲嘉宾"),
        ("必到", "XX 协会", "李 XX", 60, "北京", "协会", "老板", "5/11", "5/15", "已确认实到", "特邀致辞"),
        ("必到", "XX 钢管科技", "王 XX", 52, "天津", "钢管", "老板", "5/11", "", "已电话邀约", ""),
        ("必到", "XX 镀锌厂", "赵 XX", 48, "佛山", "镀锌冷卷", "老板", "5/11", "", "已电话邀约", ""),
        ("力争", "XX 开平加工", "钱 XX", 45, "宁波", "开平纵剪", "销售总监B", "5/12", "", "邀请函已发", ""),
        ("力争", "XX 钢铁集团", "孙 XX", 58, "聊城", "一体厂", "销售总监A", "5/12", "", "邀请函已发", ""),
        ("力争", "XX 焊管厂", "周 XX", 50, "湖州", "钢管", "销售总监A", "5/12", "", "邀请函已发", ""),
        ("力争", "XX 冷卷厂", "吴 XX", 47, "邯郸", "镀锌冷卷", "销售总监A", "5/12", "", "邀请函已发", ""),
        ("力争", "XX 纵剪线", "郑 XX", 43, "苏州", "开平纵剪", "销售总监B", "5/13", "", "邀请函已发", ""),
        ("加分", "XX 钢管", "冯 XX", 41, "天津", "钢管", "销售运营", "5/13", "", "名单已圈定", ""),
        ("加分", "XX 镀锌", "陈 XX", 49, "广州", "镀锌冷卷", "销售运营", "5/13", "", "名单已圈定", ""),
        ("加分", "XX 集团", "褚 XX", 56, "苏州", "一体厂", "销售运营", "5/13", "", "名单已圈定", ""),
    ]
    for r, sample in enumerate(samples, 4):
        for c, v in enumerate(sample, 2):
            ws.cell(row=r, column=c, value=v)

    # 冻结
    ws.freeze_panes = "D4"


def build_timeline(ws):
    ws.title = "5周时间轴"
    widths(ws, [4, 12, 18, 36, 16, 20])
    ws.merge_cells("B2:F2")
    ws["B2"] = "5 周邀约时间轴（W3 → W8）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["阶段", "周次/日期", "动作", "责任人", "完成情况"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    timeline = [
        # T-5 周
        ("T-5", "W3 周一", "时间地点最终决定（佛山乐从）", "老板+总监+行政", ""),
        ("T-5", "W3 周二", "包间预定（高端粤菜）", "行政", ""),
        ("T-5", "W3 周三", "圈定 12 位老板初版名单", "老板+AB", ""),
        ("T-5", "W3 周四", "12 位老板分档（必到4/力争5/加分3）", "老板", ""),
        ("T-5", "W3 周五", "邀请函设计稿初版", "市场", ""),
        # T-4 周
        ("T-4", "W4 周一", "邀请函 PDF 终稿", "市场", ""),
        ("T-4", "W4 周一下午", "★ 老板亲自电话邀约 4 位必到", "老板", ""),
        ("T-4", "W4 周二", "销售总监 A 邀约 3 位力争（钢管类）", "A", ""),
        ("T-4", "W4 周二", "销售总监 B 邀约 2 位力争（开平类）", "B", ""),
        ("T-4", "W4 周三", "销售运营邀约 3 位加分", "运营", ""),
        ("T-4", "W4 周四", "邀请函电子版统一发送", "全员", ""),
        ("T-4", "W4 周五", "汇总：已确认/待定/婉拒", "运营", ""),
        # T-3 周
        ("T-3", "W5 周一", "未确认者电话跟进", "对应负责人", ""),
        ("T-3", "W5 周三", "★ 灯塔厂老板亲自致电 5 位老板", "灯塔厂老板", ""),
        ("T-3", "W5 周四", "如确认数 < 8 → 启动备选 3-5 位", "老板", ""),
        ("T-3", "W5 周五", "T-3 周状态汇总", "运营", ""),
        # T-2 周
        ("T-2", "W6 周一", "实体邀请函印刷完成", "市场+行政", ""),
        ("T-2", "W6 周二", "EMS 寄送至 12 位老板办公室", "行政", ""),
        ("T-2", "W6 周三", "短信 + 微信二次确认", "对应负责人", ""),
        ("T-2", "W6 周四", "菜单/酒水/包间/物料/礼品最终确认", "行政+市场", ""),
        ("T-2", "W6 周五", "灯塔厂老板录 3 分钟视频（备用）", "老板+市场", ""),
        # T-1 周
        ("T-1", "W7 周一", "★ 老板亲自电话最终确认", "老板", ""),
        ("T-1", "W7 周二", "灯塔厂老板演讲稿排练", "灯塔厂老板", ""),
        ("T-1", "W7 周三", "现场踩点（包间/走位/设备）", "市场+行政", ""),
        ("T-1", "W7 周四", "礼品装袋 + 合同夹准备", "行政", ""),
        ("T-1", "W7 周五", "全员预演 + Go/No-Go 决策", "全员", ""),
        # T 日
        ("T 日", "W8 周一-周三", "接待提前到达的老板", "老板+AB", ""),
        ("T 日", "★ W8 周四 17:30", "🔥 老板局正式开始", "全员", ""),
        ("T 日", "W8 周四 18:00-18:15", "主持人开场（避免讲我方公司）", "主持人", ""),
        ("T 日", "W8 周四 18:15-18:45", "灯塔厂老板分享", "灯塔厂老板", ""),
        ("T 日", "W8 周四 18:45-19:00", "★ 我方 CEO 15 分钟产品讲解", "老板", ""),
        ("T 日", "W8 周四 19:00-21:00", "边吃边聊 + 自由提问", "全员", ""),
        ("T 日", "W8 周四 21:00-22:00", "（可选）转场参观灯塔厂", "全员", ""),
        ("T 日", "W8 周四 22:00-23:00", "茶歇 + 个别老板深谈", "老板+AB", ""),
        # T+1 后续
        ("T+1", "W9 周五早晨", "战报朋友圈 + 个性化感谢短信", "老板+市场", ""),
        ("T+1", "W9 周五", "TOP 5 老板当天上门或视频", "老板+AB", ""),
        ("T+3", "W9 下周一-三", "≥ 5 位老板二次拜访", "老板+AB", ""),
        ("T+7", "W10 周二", "≥ 4 位 POC 协议谈判启动", "AB", ""),
    ]
    last_phase = ""
    for i, (phase, when, action, owner, _) in enumerate(timeline, 5):
        # 不同 phase 用不同颜色
        phase_colors = {
            "T-5": LIGHT_BG, "T-4": LIGHT_BG,
            "T-3": SUCCESS, "T-2": SUCCESS,
            "T-1": WARN, "T 日": CRITICAL,
            "T+1": "FFE699", "T+3": "FFE699", "T+7": "FFE699"
        }
        color = phase_colors.get(phase, WHITE)
        style_t(ws.cell(row=i, column=2, value=phase), fill=color, bold=True)
        style_t(ws.cell(row=i, column=3, value=when), fill=color)
        style_t(ws.cell(row=i, column=4, value=action), fill=color, align=LEFT)
        style_t(ws.cell(row=i, column=5, value=owner), fill=color)
        style_t(ws.cell(row=i, column=6, value=""), fill=INPUT_BG)
        ws.row_dimensions[i].height = 22


def build_seating(ws):
    ws.title = "现场座位安排"
    widths(ws, [4, 4, 18, 18, 18, 18, 4, 4])
    ws.merge_cells("B2:H2")
    ws["B2"] = "大圆桌座位安排示意图（10-12 人）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    # 圆桌示意（用 grid 模拟）
    ws.merge_cells("D4:F4"); ws["D4"] = "↑ 包间门口"
    style_t(ws["D4"], fill=LIGHT_BG, bold=True)

    seats = [
        # row, col, label
        (6, 4, "主位\n[老板/CEO]\n（你方）"),
        (7, 3, "副主位\n[协会秘书长]"),
        (7, 5, "主宾\n[灯塔厂老板]"),
        (8, 2, "客 1\n[必到老板]"),
        (8, 6, "客 2\n[必到老板]"),
        (9, 2, "客 3\n[力争老板]"),
        (9, 6, "客 4\n[力争老板]"),
        (10, 3, "客 5\n[力争老板]"),
        (10, 5, "客 6\n[力争老板]"),
        (11, 4, "末位\n[加分老板/二代]"),
    ]
    for r, c, label in seats:
        cell = ws.cell(row=r, column=c, value=label)
        style_t(cell, fill=PRIMARY)
        cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color=WHITE)
        ws.row_dimensions[r].height = 40

    # 注意事项
    ws.merge_cells("B13:H17")
    ws["B13"] = ("座位铁律：\n"
                 "1. 同业竞品绝不相邻（提前查公司业务核实）\n"
                 "2. 内向老板坐两个外向老板中间（自然破冰）\n"
                 "3. 二代接班人单独安排区（如有 ≥ 3 位，可后单独安排『二代联谊』）\n"
                 "4. 摄影师位置：主位斜对面（拍合影最佳）\n"
                 "5. 销售总监 A、B 不入主桌，在外间随时接应")
    ws["B13"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["B13"].font = Font(name="Microsoft YaHei", size=10, color="595959")
    ws["B13"].border = BORDER


def build_followup(ws):
    ws.title = "T+7后续跟进"
    widths(ws, [4, 22, 14, 14, 18, 14, 22])
    ws.merge_cells("B2:G2")
    ws["B2"] = "T+1 → T+30 后续跟进（实到老板转化跟踪）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["公司+老板", "T+1 感谢短信", "T+3 二次拜访", "T+7 POC 协议", "T+30 签约", "下一步"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    # 8 行（实到 8 位老板）
    for i in range(5, 13):
        for col in range(2, 8):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG, align=LEFT)

    # 数据有效性
    dv_done = DataValidation(type="list",
        formula1='"已完成,进行中,未完成,放弃"', allow_blank=True)
    ws.add_data_validation(dv_done)
    dv_done.add("C5:F12")

    # 条件格式
    for col in "CDEF":
        ws.conditional_formatting.add(f"{col}5:{col}12",
            CellIsRule(operator="equal", formula=['"已完成"'],
                       fill=PatternFill("solid", fgColor="00B050")))
        ws.conditional_formatting.add(f"{col}5:{col}12",
            CellIsRule(operator="equal", formula=['"未完成"'],
                       fill=PatternFill("solid", fgColor="FFC000")))
        ws.conditional_formatting.add(f"{col}5:{col}12",
            CellIsRule(operator="equal", formula=['"放弃"'],
                       fill=PatternFill("solid", fgColor="FF6B6B")))

    # 汇总
    ws.merge_cells("B14:G14")
    ws["B14"] = "★ 90 天 KPI：8 位实到老板 → ≥ 5 位 POC → ≥ 3 位签约"
    style_t(ws["B14"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B14"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_dashboard(wb.create_sheet())
    build_main(wb.create_sheet())
    build_timeline(wb.create_sheet())
    build_seating(wb.create_sheet())
    build_followup(wb.create_sheet())

    out = "tools/老板局邀约_跟踪表.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
