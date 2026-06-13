"""
钢贸版周战报 + 月复盘 跟踪表（定制版）
======================================

与 docs/18 + tools/周战报_6个月生死线跟踪表.xlsx 的差异：
- 针对钢贸场景，含 5 大产品 / 9 销售 / 1500 客户 / AI 累计 / 老板局 等专属指标
- 含每周战报模板 + 月度复盘 + Q1 90 天评估
- 含 9 销售个人战报 + 团队战报双视图

输出：tools/钢贸版周战报与月复盘.xlsx

工作表：
    1. 使用说明
    2. 周战报模板（钢贸版）
    3. 13 周数据看板（5 大产品维度）
    4. 9 销售个人战报
    5. 月度复盘模板
    6. 90 天评估表
    7. 客户跟进漏斗
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
    ws["B2"] = "钢贸版周战报 + 月复盘跟踪表"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    rows = [
        ("适用对象", "9 销售 + 销售总监 + 老板"),
        ("与 docs/18 区别",
         "本表针对钢贸 ¥500 万 / 月场景定制\n"
         "- 5 大产品维度（MES / ERP / WMS / 货袋子 / AI）\n"
         "- 9 销售个人战报视图\n"
         "- 1500 客户分级跟踪\n"
         "- AI 累计曲线监控"),
        ("使用步骤",
         "① 每周日晚销售运营整理上周数据，填入『13 周数据看板』黄色\n"
         "② 周一 9:00 战报会按『周战报模板』5+3+1 现场过\n"
         "③ 9 销售各自填『9 销售个人战报』\n"
         "④ 月底切『月度复盘』\n"
         "⑤ 90 天末切『90 天评估表』决定下一步"),
        ("会议铁律",
         "周一 9:00-9:30 30 分钟严格控时\n"
         "老板亲自主持\n"
         "5 数字（线索/POC/签约/回款/AI 客户）+ 3 故事 + 1 决策"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        c1 = ws.cell(row=i, column=2, value=k); style_t(c1, fill=PRIMARY, bold=True)
        c1.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
        c2 = ws.cell(row=i, column=3, value=v); style_t(c2, align=LEFT)
        ws.row_dimensions[i].height = max(50, len(v) * 0.7)


def build_template(ws):
    ws.title = "周战报模板"
    widths(ws, [4, 20, 16, 16, 16, 16, 16])

    ws.merge_cells("B2:G2")
    ws["B2"] = "钢贸版周战报模板（每周一会议用）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    # 期号 / 日期
    style_t(ws.cell(row=4, column=2, value="期号"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=4, column=3, value="第 ___ 期"), fill=INPUT_BG)
    style_t(ws.cell(row=4, column=4, value="日期"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=4, column=5, value="2026.__.__"), fill=INPUT_BG)
    style_t(ws.cell(row=4, column=6, value="主持"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=4, column=7, value="老板"), fill=INPUT_BG)

    # 5 个核心数字
    ws.merge_cells("B6:G6"); ws["B6"] = "▼ 一、5 个核心数字（钢贸版）"
    style_t(ws["B6"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B6"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    headers = ["指标", "上周实际", "上周目标", "本周目标", "累计 90 天", "状态"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=7, column=j, value=h))

    metrics = [
        "工业 MES 签约 / POC 数",
        "钢贸 ERP 签约数",
        "库准 WMS 签约数",
        "AI 新签客户数",
        "当周回款（¥）",
    ]
    for i, m in enumerate(metrics, 8):
        style_t(ws.cell(row=i, column=2, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        for col in range(3, 8):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG)

    # 3 故事
    ws.merge_cells("B14:G14"); ws["B14"] = "▼ 二、3 故事"
    style_t(ws["B14"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B14"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    stories = [
        ("亮点客户", "公司 + 老板", "成交细节", "下周动作"),
        ("失败客户", "公司 + 老板", "失败原因", "教训沉淀"),
        ("学到的事", "讲述人", "学了什么", "沉淀到哪"),
    ]
    for i, (title, c1, c2, c3) in enumerate(stories, 15):
        style_t(ws.cell(row=i, column=2, value=title), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=c1), fill=LIGHT_BG)
        ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=5)
        style_t(ws.cell(row=i, column=4), fill=INPUT_BG, align=LEFT)
        ws.merge_cells(start_row=i, start_column=6, end_row=i, end_column=7)
        style_t(ws.cell(row=i, column=6), fill=INPUT_BG, align=LEFT)
        ws.row_dimensions[i].height = 50

    # 1 决策
    ws.merge_cells("B19:G19"); ws["B19"] = "▼ 三、本周老板决策"
    style_t(ws["B19"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B19"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    style_t(ws.cell(row=20, column=2, value="决策"), fill=LIGHT_BG, bold=True)
    ws.merge_cells("C20:G20")
    style_t(ws.cell(row=20, column=3), fill=INPUT_BG, align=LEFT)
    style_t(ws.cell(row=21, column=2, value="责任人"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=21, column=3), fill=INPUT_BG)
    style_t(ws.cell(row=21, column=4, value="截止日"), fill=LIGHT_BG, bold=True)
    style_t(ws.cell(row=21, column=5), fill=INPUT_BG)
    ws.merge_cells("F21:G21")
    style_t(ws.cell(row=21, column=6), fill=INPUT_BG)


def build_dashboard(ws):
    ws.title = "13周数据看板"
    widths(ws, [22] + [9]*13 + [10])

    ws.merge_cells("A1:O1")
    ws["A1"] = "13 周数据看板（5 大产品维度，黄色 = 输入）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    style_h(ws.cell(row=3, column=1, value="指标"))
    for i in range(13):
        style_h(ws.cell(row=3, column=2+i, value=f"W{i+1}"))
    style_h(ws.cell(row=3, column=15, value="累计"))

    metrics = [
        "新增线索数",
        "工业 MES POC 数",
        "工业 MES 签约数",
        "钢贸 ERP 签约数",
        "库准 WMS 签约数",
        "AI 当周新增",
        "AI 累计客户",
        "当周回款（万）",
        "累计回款（万）",
        "9 销售活跃度（/9）",
    ]
    for i, m in enumerate(metrics, 4):
        style_t(ws.cell(row=i, column=1, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        for w in range(13):
            cell = ws.cell(row=i, column=2+w)
            style_t(cell, fill=INPUT_BG)
            if "万" in m:
                cell.number_format = '0.0"万"'
        cell = ws.cell(row=i, column=15, value=f"=SUM(B{i}:N{i})")
        style_t(cell, fill=SUCCESS, bold=True)
        if "万" in m:
            cell.number_format = '0.0"万"'

    ws.freeze_panes = "B4"


def build_individual(ws):
    ws.title = "9销售个人战报"
    widths(ws, [4, 12, 10, 12, 14, 14, 14, 16])

    ws.merge_cells("B2:H2")
    ws["B2"] = "9 销售个人周战报（每周一更新）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["销售", "角色", "本周拜访", "本周签约", "本周回款（万）", "当月累计（万）", "月目标达成度"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    sales = [
        ("A1", "大客户铁三角"),
        ("A2", "大客户铁三角"),
        ("A3", "大客户铁三角"),
        ("B1", "中端铁二角"),
        ("B2", "中端铁二角"),
        ("B3", "中端铁二角"),
        ("C1", "电销批量"),
        ("C2", "电销批量"),
        ("C3", "电销批量"),
    ]
    for i, (name, role) in enumerate(sales, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=name), fill=bg, bold=True)
        style_t(ws.cell(row=i, column=3, value=role), fill=bg)
        for col in range(4, 8):
            cell = ws.cell(row=i, column=col)
            style_t(cell, fill=INPUT_BG)
            if col in (6, 7):
                cell.number_format = '0.0"万"'
        cell = ws.cell(row=i, column=8)
        style_t(cell, fill=INPUT_BG)
        cell.number_format = "0%"

    # 合计
    i = 14
    style_t(ws.cell(row=i, column=2, value="9 销售合计"), fill=PRIMARY, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(4, 8):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"=SUM({col_letter}5:{col_letter}13)")
        style_t(cell, fill=PRIMARY, bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
        if col in (6, 7):
            ws.cell(row=i, column=col).number_format = '0.0"万"'


def build_monthly_review(ws):
    ws.title = "月度复盘"
    widths(ws, [4, 22, 16, 16, 16, 22])

    ws.merge_cells("B2:F2")
    ws["B2"] = "月度复盘模板（每月末用）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["指标", "M1", "M2", "M3", "90 天合计"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    metrics = [
        ("月营收（万）", "1.5", "3.5", "5.5"),
        ("工业 MES 签约", "0", "2", "3"),
        ("钢贸 ERP 签约", "9", "12", "15"),
        ("库准 WMS 签约", "1", "3", "6"),
        ("AI 累计", "100", "400", "800"),
        ("9 销售达成率(%)", "60%", "80%", "100%"),
        ("客户 NPS", "—", "—", "—"),
        ("流失率(%)", "—", "—", "—"),
    ]
    for i, (m, t1, t2, t3) in enumerate(metrics, 5):
        style_t(ws.cell(row=i, column=2, value=m), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=t1), fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=4, value=t2), fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=5, value=t3), fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=6), fill=INPUT_BG)

    # 月度评分
    ws.merge_cells("B14:F14")
    ws["B14"] = "▼ 月度自评（A/B/C）"
    style_t(ws["B14"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B14"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        style_t(ws.cell(row=15, column=col), fill=INPUT_BG, bold=True)

    # 月度三问
    ws.merge_cells("B17:F17"); ws["B17"] = "▼ 月度三问（每月填写）"
    style_t(ws["B17"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B17"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    questions = [
        "1. 本月做对的最重要 1 件事？",
        "2. 本月做错的最重要 1 件事？",
        "3. 下月最重要 1 件事？",
    ]
    for i, q in enumerate(questions, 18):
        style_t(ws.cell(row=i, column=2, value=q), fill=LIGHT_BG, align=LEFT, bold=True)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=6)
        style_t(ws.cell(row=i, column=3), fill=INPUT_BG, align=LEFT)
        ws.row_dimensions[i].height = 40


def build_90day_eval(ws):
    ws.title = "90天评估表"
    widths(ws, [4, 26, 16, 16, 16, 24])

    ws.merge_cells("B2:F2")
    ws["B2"] = "90 天总评估（2026.09 末用）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["核心 KPI", "目标", "实际", "达成度", "等级"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    kpis = [
        ("M3 月营收（万）", 500),
        ("工业 MES 签约数", 3),
        ("钢贸 ERP 签约数", 15),
        ("库准 WMS 签约数", 6),
        ("AI 累计客户", 800),
        ("9 销售达成率", "100%"),
        ("客户 NPS", "≥ 50"),
        ("现金流（万）", 660),
    ]
    for i, (k, t) in enumerate(kpis, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=k), fill=bg, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=t), fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=4), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=5), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=6), fill=INPUT_BG)

    # 决策
    ws.merge_cells("B14:F14")
    ws["B14"] = "▼ 90 天后决策（圈出 1 个）"
    style_t(ws["B14"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B14"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    decisions = [
        ("达成 ≥ 100%", "加速：目标提高到 ¥1,000 万 / 月", SUCCESS),
        ("达成 80-100%", "保持节奏 + 优化短板", LIGHT_BG),
        ("达成 50-80%", "暂停扩张 + 优化销售团队", WARN),
        ("达成 < 50%", "重新审视产品-市场匹配", CRITICAL),
    ]
    for i, (lvl, action, color) in enumerate(decisions, 15):
        style_t(ws.cell(row=i, column=2, value=lvl), fill=color, bold=True)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=6)
        style_t(ws.cell(row=i, column=3, value=action), fill=color, align=LEFT, bold=True)


def build_funnel(ws):
    ws.title = "客户跟进漏斗"
    widths(ws, [4, 22, 14, 14, 14, 14])

    ws.merge_cells("B2:F2")
    ws["B2"] = "1500 客户跟进漏斗（每周更新）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["阶段", "S 段", "A 段", "B 段", "C 段"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    stages = [
        ("客户总数", 50, 300, 450, 700),
        ("已联系", "", "", "", ""),
        ("有兴趣（MQL）", "", "", "", ""),
        ("约访成功（SQL）", "", "", "", ""),
        ("诊断完成", "", "", "", ""),
        ("POC 启动", "", "", "", ""),
        ("签约", "", "", "", ""),
        ("上线", "", "", "", ""),
        ("流失", "", "", "", ""),
    ]
    for i, row in enumerate(stages, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=row[0]), fill=bg, align=LEFT, bold=True)
        for j, v in enumerate(row[1:], 3):
            if i == 5:  # 第 1 行是总数
                style_t(ws.cell(row=i, column=j, value=v), fill=SUCCESS, bold=True)
            else:
                style_t(ws.cell(row=i, column=j, value=v), fill=INPUT_BG)


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_template(wb.create_sheet())
    build_dashboard(wb.create_sheet())
    build_individual(wb.create_sheet())
    build_monthly_review(wb.create_sheet())
    build_90day_eval(wb.create_sheet())
    build_funnel(wb.create_sheet())

    out = "tools/钢贸版周战报与月复盘.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
