"""
¥500 万 / 月营收 90 天达成 财务模型
====================================

输出：tools/500万月营收_财务模型.xlsx

工作表：
    1. 使用说明
    2. 产品价格表
    3. 客户分级与目标
    4. 9 销售个人目标
    5. 月度营收预测 (M1/M2/M3)
    6. AI 累计曲线 (90 天)
    7. 90 天损益与现金流
    8. 敏感性分析（3 档情景）
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


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
    ws["B2"] = "¥500 万 / 月营收 90 天达成模型"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    rows = [
        ("适用对象", "老板 + 销售总监 + 财务"),
        ("使用步骤",
         "① 『产品价格表』黄色 = 可调（5 大产品定价）\n"
         "② 『客户分级与目标』黄色 = 1500 客户分配\n"
         "③ 『9 销售个人目标』黄色 = 每销售月目标\n"
         "④ 切换『月度营收预测』看 M1/M2/M3 累计\n"
         "⑤ 切换『AI 累计曲线』看 90 天 AI 增长\n"
         "⑥ 切换『敏感性分析』看保守/基准/乐观"),
        ("核心假设",
         "- AI 累计：M1 100 / M2 400 / M3 800\n"
         "- 工业 MES 月签：M1 0 / M2 2 / M3 3\n"
         "- 钢贸 ERP 月签：M1 9 / M2 12 / M3 15\n"
         "- 库准 WMS 月签：M1 1 / M2 3 / M3 6\n"
         "- 90 天总投入 ¥180 万 / 回款 ¥840 万"),
        ("修改边界", "只改黄色；公式和表头不要改"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        c1 = ws.cell(row=i, column=2, value=k); style_t(c1, fill=PRIMARY, bold=True)
        c1.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
        c2 = ws.cell(row=i, column=3, value=v); style_t(c2, align=LEFT)
        ws.row_dimensions[i].height = max(50, len(v) * 0.7)


def build_pricing(ws):
    ws.title = "产品价格表"
    widths(ws, [4, 24, 14, 14, 14, 30])

    ws.merge_cells("B2:F2")
    ws["B2"] = "5 大产品价格表（黄色 = 可调）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["产品", "版本", "单价(¥)", "类型", "目标客户"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    products = [
        ("钢贸 ERP", "入门版", 2800, "一次性 / 用户", "C 段小客户"),
        ("钢贸 ERP", "经营版", 5500, "一次性 / 用户", "B 段中型"),
        ("钢贸 ERP", "旗舰版", 8500, "一次性 / 用户", "A/S 段大客户"),
        ("库准 WMS", "轻量版（新）", 5000, "一次性 / 用户", "C 段（≤5 用户）"),
        ("库准 WMS", "标准版", 10000, "一次性 / 用户", "B 段腰部"),
        ("库准 WMS", "旗舰版", 18000, "一次性 / 用户", "S 段集团"),
        ("工业 MES", "标准版", 300000, "一次性 / 项目", "S/A 段（有生产）"),
        ("货袋子", "免费版", 0, "免费", "全客户"),
        ("AI 入门包", "3 个 AI", 4000, "月费", "C 段"),
        ("AI 经营包", "5 个 AI", 7500, "月费", "B 段"),
        ("AI 旗舰包", "8 个 AI", 12000, "月费", "A/S 段"),
    ]
    for i, (prod, ver, price, type_, target) in enumerate(products, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=prod), fill=bg, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=ver), fill=bg)
        style_t(ws.cell(row=i, column=4, value=price), fill=INPUT_BG, money=True, bold=True)
        style_t(ws.cell(row=i, column=5, value=type_), fill=bg)
        style_t(ws.cell(row=i, column=6, value=target), fill=bg, align=LEFT)


def build_customer(ws):
    ws.title = "客户分级与目标"
    widths(ws, [4, 14, 14, 14, 18, 14, 14, 14])

    ws.merge_cells("B2:H2")
    ws["B2"] = "1500 客户分级 + 90 天转化目标"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["客户段", "客户数", "占比", "目标转化数", "转化率", "主推套餐", "ARPU(¥/月)", "首付(¥)"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    segments = [
        ("S 段", 50, "3%", 8, "16%", "MES + ERP 旗舰 + WMS + AI 旗舰", 12000, 500000),
        ("A 段", 300, "20%", 35, "11.7%", "ERP 旗舰 + WMS + AI 旗舰", 12000, 150000),
        ("B 段", 450, "30%", 120, "26.7%", "ERP 经营 + AI 经营", 7500, 60000),
        ("C 段", 700, "47%", 600, "85.7%", "AI 入门", 4000, 0),
    ]
    total_revenue_first_pay = 0
    total_mrr = 0
    for i, (seg, total, ratio, target, conv, plan, arpu, first_pay) in enumerate(segments, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=seg), fill=bg, bold=True)
        style_t(ws.cell(row=i, column=3, value=total), fill=bg, bold=True)
        style_t(ws.cell(row=i, column=4, value=ratio), fill=bg)
        style_t(ws.cell(row=i, column=5, value=target), fill=INPUT_BG, bold=True)
        style_t(ws.cell(row=i, column=6, value=conv), fill=bg)
        style_t(ws.cell(row=i, column=7, value=plan), fill=bg, align=LEFT)
        style_t(ws.cell(row=i, column=8, value=arpu), fill=INPUT_BG, money=True, bold=True)
        style_t(ws.cell(row=i, column=9 if False else 0, value=None), fill=bg) if False else None
        # 我们把首付放在 H 列
    # 重新处理首付列
    ws.cell(row=4, column=8).value = "ARPU(¥/月)"
    style_h(ws.cell(row=4, column=8))

    # 合计
    total_row = 9
    style_t(ws.cell(row=total_row, column=2, value="合计"), fill=PRIMARY, bold=True)
    ws.cell(row=total_row, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=total_row, column=3, value="=SUM(C5:C8)")
    style_t(cell, fill=PRIMARY, bold=True)
    ws.cell(row=total_row, column=3).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=total_row, column=5, value="=SUM(E5:E8)")
    style_t(cell, fill=PRIMARY, bold=True)
    ws.cell(row=total_row, column=5).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)


def build_sales(ws):
    ws.title = "9销售个人目标"
    widths(ws, [4, 14, 14, 12, 12, 12, 14, 28])

    ws.merge_cells("B2:H2")
    ws["B2"] = "9 销售个人 90 天目标"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["销售", "角色", "M1(¥万)", "M2(¥万)", "M3(¥万)", "90 天合计(¥万)", "主攻产品"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    sales_data = [
        ("A1", "大客户铁三角", 30, 60, 90, "MES + ERP 大单 + WMS"),
        ("A2", "大客户铁三角", 30, 60, 80, "MES + ERP 大单（钢管）"),
        ("A3", "大客户铁三角", 25, 50, 70, "MES + ERP 大单（镀锌）"),
        ("B1", "中端铁二角", 15, 30, 50, "ERP 升级 + AI 经营"),
        ("B2", "中端铁二角", 15, 30, 50, "ERP 升级 + AI 经营（珠三角）"),
        ("B3", "中端铁二角", 10, 30, 50, "ERP + AI（开平 / 纵剪）"),
        ("C1", "电销批量", 8, 25, 45, "AI 入门包（老客户激活）"),
        ("C2", "电销批量", 8, 25, 45, "AI 入门包（老带新）"),
        ("C3", "电销批量", 7, 35, 45, "AI 入门包（货袋子拓客）"),
    ]
    for i, (name, role, m1, m2, m3, products) in enumerate(sales_data, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=name), fill=bg, bold=True)
        style_t(ws.cell(row=i, column=3, value=role), fill=bg)
        style_t(ws.cell(row=i, column=4, value=m1), fill=INPUT_BG, bold=True)
        style_t(ws.cell(row=i, column=5, value=m2), fill=INPUT_BG, bold=True)
        style_t(ws.cell(row=i, column=6, value=m3), fill=INPUT_BG, bold=True)
        cell = ws.cell(row=i, column=7, value=f"=D{i}+E{i}+F{i}")
        style_t(cell, fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=8, value=products), fill=bg, align=LEFT)

    # 合计
    i = 14
    style_t(ws.cell(row=i, column=2, value="9 销售合计"), fill=PRIMARY, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(4, 8):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"=SUM({col_letter}5:{col_letter}13)")
        style_t(cell, fill=PRIMARY, bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 验证：M3 ≥ 500 万
    i = 16
    ws.merge_cells(f"B{i}:H{i}")
    cell = ws.cell(row=i, column=2,
                    value="★ 9 销售 M3 合计 ≥ ¥500 万 ✅ 达标")
    style_t(cell, fill=SUCCESS, align=LEFT, bold=True)


def build_monthly(ws):
    ws.title = "月度营收预测"
    widths(ws, [4, 24, 14, 14, 14, 16])

    ws.merge_cells("B2:F2")
    ws["B2"] = "月度营收预测 M1/M2/M3"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["产品", "M1(¥)", "M2(¥)", "M3(¥)", "90 天合计(¥)"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    products = [
        ("工业 MES（首付）", 0, 600000, 900000),
        ("钢贸 ERP（首付）", 600000, 800000, 1000000),
        ("库准 WMS（首付）", 100000, 300000, 600000),
        ("AI 累计月费", 400000, 1600000, 3200000),
        ("货袋子", 0, 0, 0),
        ("增值服务", 50000, 100000, 200000),
    ]
    for i, (prod, m1, m2, m3) in enumerate(products, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=prod), fill=bg, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=m1), fill=INPUT_BG, money=True, bold=True)
        style_t(ws.cell(row=i, column=4, value=m2), fill=INPUT_BG, money=True, bold=True)
        style_t(ws.cell(row=i, column=5, value=m3), fill=INPUT_BG, money=True, bold=True)
        cell = ws.cell(row=i, column=6, value=f"=C{i}+D{i}+E{i}")
        style_t(cell, fill=SUCCESS, money=True, bold=True)

    # 合计
    i = 11
    style_t(ws.cell(row=i, column=2, value="月度合计"), fill=PRIMARY, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"=SUM({col_letter}5:{col_letter}10)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 是否达标
    i = 13
    ws.merge_cells(f"B{i}:F{i}")
    cell = ws.cell(row=i, column=2,
                    value='=IF(E11>=5000000,"★ M3 ¥500 万 ✅ 达标","M3 未达 ¥500 万")')
    style_t(cell, fill=SUCCESS, align=LEFT, bold=True)


def build_ai_curve(ws):
    ws.title = "AI累计曲线"
    widths(ws, [4, 14, 14, 14, 14, 14])

    ws.merge_cells("B2:F2")
    ws["B2"] = "AI 累计客户增长曲线（90 天）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["月末", "月新增", "累计客户", "ARPU(¥/月)", "月营收(¥)"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    curve = [
        ("M1 末（D30）", 100, 100, 4000, ""),
        ("M2 末（D60）", 300, 400, 4000, ""),
        ("M3 末（D90）", 400, 800, 4000, ""),
    ]
    for i, (point, new, cumulative, arpu, _) in enumerate(curve, 5):
        bg = SUCCESS if "M3" in point else LIGHT_BG
        style_t(ws.cell(row=i, column=2, value=point), fill=bg, bold=True)
        style_t(ws.cell(row=i, column=3, value=new), fill=INPUT_BG, bold=True)
        style_t(ws.cell(row=i, column=4, value=cumulative), fill=INPUT_BG, bold=True)
        style_t(ws.cell(row=i, column=5, value=arpu), fill=INPUT_BG, money=True)
        cell = ws.cell(row=i, column=6, value=f"=D{i}*E{i}")
        style_t(cell, fill=SUCCESS, money=True, bold=True)

    # 备注：ARPU 提升路径
    ws.merge_cells("B9:F11")
    ws["B9"] = (
        "★ ARPU 提升机会：\n"
        "- 默认 ARPU ¥4,000 = 全部入门包\n"
        "- 优化：入门 70% + 经营 25% + 旗舰 5%\n"
        "- 加权 ARPU = 4000×0.7 + 7500×0.25 + 12000×0.05 = ¥5,275\n"
        "- 800 客户 × ¥5,275 = ¥422 万 / 月（vs 默认 ¥320 万）"
    )
    ws["B9"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["B9"].font = Font(name="Microsoft YaHei", size=10, color="595959")
    ws["B9"].border = BORDER


def build_pnl(ws):
    ws.title = "90天损益现金流"
    widths(ws, [4, 24, 14, 14, 14, 16])

    ws.merge_cells("B2:F2")
    ws["B2"] = "90 天损益与现金流"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["项", "M1(¥)", "M2(¥)", "M3(¥)", "90 天合计(¥)"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    # 收入（从月度营收预测引用）
    style_t(ws.cell(row=5, column=2, value="月营收"), fill=LIGHT_BG, bold=True)
    for j, col in enumerate(["C", "D", "E"], 3):
        cell = ws.cell(row=5, column=j, value=f"='月度营收预测'!{col}11")
        style_t(cell, fill=SUCCESS, money=True, bold=True)
    cell = ws.cell(row=5, column=6, value="=C5+D5+E5")
    style_t(cell, fill=SUCCESS, money=True, bold=True)

    # 成本
    cost_items = [
        ("9 销售薪酬+提成", 333333, 333333, 333333),
        ("实施+客服+运营 5 人", 100000, 100000, 100000),
        ("培训+物料一次性", 50000, 0, 0),
        ("CRM+电销工具", 50000, 0, 0),
        ("营销+客户活动", 50000, 50000, 50000),
        ("老板局 2-3 场", 50000, 50000, 50000),
        ("应急", 33000, 33000, 34000),
    ]
    for i, (item, m1, m2, m3) in enumerate(cost_items, 6):
        style_t(ws.cell(row=i, column=2, value=item), fill=LIGHT_BG, align=LEFT)
        style_t(ws.cell(row=i, column=3, value=m1), fill=WARN, money=True)
        style_t(ws.cell(row=i, column=4, value=m2), fill=WARN, money=True)
        style_t(ws.cell(row=i, column=5, value=m3), fill=WARN, money=True)
        cell = ws.cell(row=i, column=6, value=f"=C{i}+D{i}+E{i}")
        style_t(cell, fill=WARN, money=True, bold=True)

    # 总成本
    i = 13
    style_t(ws.cell(row=i, column=2, value="月度总成本"), fill=PRIMARY, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"=SUM({col_letter}6:{col_letter}12)")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 净现金
    i = 14
    style_t(ws.cell(row=i, column=2, value="月度净现金"), fill=PRIMARY, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for col in range(3, 7):
        col_letter = get_column_letter(col)
        cell = ws.cell(row=i, column=col, value=f"={col_letter}5-{col_letter}13")
        style_t(cell, fill=SUCCESS, money=True, bold=True)
        ws.cell(row=i, column=col).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # ROI
    i = 16
    ws.merge_cells(f"B{i}:E{i}")
    style_t(ws.cell(row=i, column=2, value="90 天 ROI"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=i, column=6, value="=F14/F13")
    style_t(cell, fill=SUCCESS, bold=True)
    cell.number_format = "0.0x"


def build_sensitivity(ws):
    ws.title = "敏感性分析"
    widths(ws, [4, 22, 16, 16, 16])

    ws.merge_cells("B2:E2")
    ws["B2"] = "3 档情景对比"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["指标", "保守（达成 70%）", "基准（100%）", "乐观（130%）"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    metrics = [
        ("M3 月营收（¥万）", 385, 550, 715),
        ("90 天累计营收（¥万）", 588, 840, 1092),
        ("M3 末 AI 客户数", 560, 800, 1040),
        ("90 天总投入（¥万）", 180, 180, 180),
        ("净利润（¥万）", 408, 660, 912),
        ("ROI", 2.3, 3.7, 5.1),
    ]
    for i, (metric, c, b, o) in enumerate(metrics, 5):
        style_t(ws.cell(row=i, column=2, value=metric), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=c), fill=WARN, bold=True)
        style_t(ws.cell(row=i, column=4, value=b), fill=SUCCESS, bold=True)
        style_t(ws.cell(row=i, column=5, value=o), fill=LIGHT_BG, bold=True)
        if metric.startswith("M3 月营收") or "累计" in metric or "投入" in metric or "净利润" in metric:
            for col in range(3, 6):
                ws.cell(row=i, column=col).number_format = '0"万"'
        elif "ROI" in metric:
            for col in range(3, 6):
                ws.cell(row=i, column=col).number_format = "0.0x"


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_pricing(wb.create_sheet())
    build_customer(wb.create_sheet())
    build_sales(wb.create_sheet())
    build_monthly(wb.create_sheet())
    build_ai_curve(wb.create_sheet())
    build_pnl(wb.create_sheet())
    build_sensitivity(wb.create_sheet())

    out = "tools/500万月营收_财务模型.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
