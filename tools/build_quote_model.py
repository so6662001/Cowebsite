"""
报价单 Excel 模型生成器  v2.0
==============================

v2.0 重大调整：
    1. 价格体系换成"3 档套餐价"（种子/经营/旗舰），不再按产线累加
    2. 增加付费模式选择器（A 经典 / B 0首付月供 / C 利润分成）
    3. 增加 12 期分期模拟
    4. 月租上浮系数（B 模式 130%）

运行方式：
    python3 tools/build_quote_model.py
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
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_header(c):
    c.fill = HEAD_FILL; c.font = HEAD_FONT; c.alignment = CENTER; c.border = BORDER


def style_sub(c):
    c.fill = SUB_FILL; c.font = SUB_FONT; c.alignment = CENTER; c.border = BORDER


def style_text(c, fill=None, align=None):
    c.font = TEXT_FONT; c.alignment = align or CENTER; c.border = BORDER
    if fill: c.fill = fill


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ----------------------------------------------------------------------
# 价格表 v2.0：客户类型 / 版本 / 实施费 / 月租 / 含模块
# ----------------------------------------------------------------------
PRICE_TABLE = [
    # 客户类型, 版本, 实施费, 月租, 含模块描述, 适用规模
    ("开平纵剪加工", "种子版",  9800,    1500, "成本核算（单产线）", "想先试试的小厂"),
    ("开平纵剪加工", "经营版", 36000,    4800, "成本核算+库存+卷号追溯（≤3 产线）", "80% 主流"),
    ("开平纵剪加工", "旗舰版", 98000,   12000, "全套(ERP+MES+WMS+财务)+不限产线", "4+ 产线大厂"),
    ("钢管生产企业", "种子版", 19800,    3500, "质量追溯（单产线）", "1-2 产线"),
    ("钢管生产企业", "经营版", 68000,    9800, "追溯+排程+成本（≤5 产线）", "80% 主流"),
    ("钢管生产企业", "旗舰版",188000,   25000, "全套+EAM+多基地+不限产线", "6+ 产线大厂"),
    ("镀锌冷卷企业", "种子版", 36000,    6800, "追溯+OEE（单机组）", "单机组试点"),
    ("镀锌冷卷企业", "经营版", 98000,   15000, "追溯+EAM+排程+成本（≤2 机组）", "80% 主流"),
    ("镀锌冷卷企业", "旗舰版",250000,   35000, "全套+APS接口+多基地", "3+ 机组大厂"),
    ("一体厂",      "种子版", 68000,   12000, "全厂质量追溯（单业务）", "试点先行"),
    ("一体厂",      "经营版",168000,   28000, "全厂追溯+排程+成本", "80% 主流"),
    ("一体厂",      "旗舰版",380000,   58000, "全套+集团合并报表+BI大屏", "集团化大厂"),
]


def build_readme(ws):
    ws.title = "使用说明"
    set_col_widths(ws, [4, 30, 70])
    ws.merge_cells("B2:C2")
    ws["B2"] = "钢铁数字化系统报价单 v2.0（降门槛版）"
    ws["B2"].font = Font(name="Microsoft YaHei", size=18, bold=True, color="1F4E79")
    ws["B2"].alignment = LEFT

    rows = [
        ("v2.0 主要变化",
         "① 实施费降幅 50-70%，月租降幅 20-30%\n"
         "② 简化为 3 档套餐价（种子/经营/旗舰），不再按产线累加\n"
         "③ 新增 3 种付费模式：A 经典 / B 0首付月供（+30%月租, 36 月）/ C 利润分成\n"
         "④ 新增 12 期分期模拟与年支出红线警示"),
        ("使用步骤",
         "① 打开『报价计算器』，黄色单元格填客户类型、版本、付费模式、客户年利润\n"
         "② 绿色单元格自动出报价、月供、年支出占年利润比\n"
         "③ 占比超过 1.5%→红色警示，建议调整版本或付费模式\n"
         "④ 切换『付费模式对比』看 3 种模式横向对比\n"
         "⑤ 切换『3 档版本一览』给客户看完整菜单（不要让客户看全价目表）"),
        ("销售铁律",
         "1. 永远只给客户看『3 档版本一览』中的 3 个版本，不要打开『价格基准表』\n"
         "2. 报价话术：先报『月供』，客户问总价再报实施费\n"
         "3. 实施费默认推 12 期分期，相当于『每月一杯酒钱』\n"
         "4. 现签现送：30 天月租免单（写进合同）"),
        ("权限",
         "销售：套餐价 ±5% 自批；销售总监 ±15%；C 模式必须总经理审批；种子版不打折"),
        ("注意事项",
         "1. 黄色 = 输入区；绿色 = 自动计算；红色 = 警示；蓝色 = 表头\n"
         "2. 修改『价格基准表』后保存即可，无需改公式\n"
         "3. C 模式（利润分成）只对灯塔厂候选客户开放，必须 CEO/总经理批"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        ws[f"B{i}"] = k; ws[f"C{i}"] = v
        style_sub(ws[f"B{i}"]); style_text(ws[f"C{i}"], align=LEFT)
        ws.row_dimensions[i].height = max(40, len(v) * 0.6)


def build_baseline(ws):
    ws.title = "价格基准表"
    headers = ["客户类型", "版本", "实施费(元)", "月租(元/月)", "含模块", "适用规模"]
    set_col_widths(ws, [16, 10, 14, 14, 50, 24])
    for j, h in enumerate(headers, start=1):
        style_header(ws.cell(row=1, column=j, value=h))
    for i, row in enumerate(PRICE_TABLE, start=2):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=v)
            if j in (3, 4):
                style_text(c); c.number_format = '"¥"#,##0'
            elif j in (5, 6):
                style_text(c, align=LEFT)
            else:
                style_text(c)


def build_versions(ws):
    """3 档版本一览：给客户看的"3 选 1"页"""
    ws.title = "3档版本一览"
    set_col_widths(ws, [4, 18, 18, 18, 18])
    ws.merge_cells("B2:E2")
    ws["B2"] = "3 档版本菜单（销售给客户看的页面）"
    ws["B2"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E79")

    # 4 个客户类型分块
    customer_types = ["开平纵剪加工", "钢管生产企业", "镀锌冷卷企业", "一体厂"]
    row = 4
    for ct in customer_types:
        ws.cell(row=row, column=2, value=f"▼ {ct}").font = Font(
            name="Microsoft YaHei", size=12, bold=True, color="1F4E79")
        row += 1
        # 表头
        for j, h in enumerate(["项目", "种子版", "经营版", "旗舰版"], start=2):
            style_header(ws.cell(row=row, column=j, value=h))
        row += 1
        # 取该类型 3 条
        rows_of_ct = [r for r in PRICE_TABLE if r[0] == ct]
        # 实施费行
        ws.cell(row=row, column=2, value="实施费(元)").font = SUB_FONT
        ws.cell(row=row, column=2).border = BORDER
        for j, r in enumerate(rows_of_ct, start=3):
            c = ws.cell(row=row, column=j, value=r[2]); style_text(c, fill=RESULT_FILL)
            c.number_format = '"¥"#,##0'
        row += 1
        # 月租
        ws.cell(row=row, column=2, value="月租起(元/月)").font = SUB_FONT
        ws.cell(row=row, column=2).border = BORDER
        for j, r in enumerate(rows_of_ct, start=3):
            c = ws.cell(row=row, column=j, value=r[3]); style_text(c, fill=RESULT_FILL)
            c.number_format = '"¥"#,##0'
        row += 1
        # 12 期月供
        ws.cell(row=row, column=2, value="12 期月供合计").font = SUB_FONT
        ws.cell(row=row, column=2).border = BORDER
        for j, r in enumerate(rows_of_ct, start=3):
            v = round(r[2] / 12 + r[3], -1)
            c = ws.cell(row=row, column=j, value=v); style_text(c, fill=RESULT_FILL)
            c.number_format = '"¥"#,##0'
        row += 1
        # 含模块
        ws.cell(row=row, column=2, value="含模块").font = SUB_FONT
        ws.cell(row=row, column=2).border = BORDER
        for j, r in enumerate(rows_of_ct, start=3):
            c = ws.cell(row=row, column=j, value=r[4]); style_text(c, align=LEFT)
        ws.row_dimensions[row].height = 50
        row += 1
        # 适用
        ws.cell(row=row, column=2, value="适用").font = SUB_FONT
        ws.cell(row=row, column=2).border = BORDER
        for j, r in enumerate(rows_of_ct, start=3):
            c = ws.cell(row=row, column=j, value=r[5]); style_text(c, align=LEFT)
        row += 2  # 分块空行


def build_calculator(ws):
    ws.title = "报价计算器"
    set_col_widths(ws, [4, 22, 22, 22, 22, 22])

    ws.merge_cells("B2:F2")
    ws["B2"] = "客户报价计算器（v2.0）"
    ws["B2"].font = Font(name="Microsoft YaHei", size=16, bold=True, color="1F4E79")
    ws["B2"].alignment = LEFT

    # 输入区
    ws.merge_cells("B4:F4"); ws["B4"] = "一、客户基础信息（黄色 = 必填）"; style_sub(ws["B4"])
    inputs = [
        ("客户全称", ""),
        ("联系人 / 电话", ""),
        ("客户类型", "钢管生产企业"),
        ("版本", "经营版"),
        ("付费模式", "A 经典"),
        ("销售折扣率(0.85~1.00)", 1.00),
        ("客户年利润(万元)", 1500),
    ]
    for i, (label, default) in enumerate(inputs, start=5):
        ws.cell(row=i, column=2, value=label).font = SUB_FONT
        ws.cell(row=i, column=2).alignment = LEFT
        ws.cell(row=i, column=2).border = BORDER
        c = ws.cell(row=i, column=3, value=default); style_text(c, fill=INPUT_FILL, align=LEFT)
    # 数据有效性
    dv_type = DataValidation(type="list", formula1='"开平纵剪加工,钢管生产企业,镀锌冷卷企业,一体厂"')
    ws.add_data_validation(dv_type); dv_type.add("C7")
    dv_ver = DataValidation(type="list", formula1='"种子版,经营版,旗舰版"')
    ws.add_data_validation(dv_ver); dv_ver.add("C8")
    dv_mode = DataValidation(type="list", formula1='"A 经典,B 0首付月供,C 利润分成"')
    ws.add_data_validation(dv_mode); dv_mode.add("C9")

    # 自动计算辅助
    ws.merge_cells("B13:F13"); ws["B13"] = "二、自动计算（绿色 = 输出，红色 = 警示）"; style_sub(ws["B13"])
    sheet_ref = "'价格基准表'"

    def lookup(col_letter):
        # 按客户类型 + 版本 双键查找
        return (
            f"=SUMPRODUCT(({sheet_ref}!$A$2:$A$13=$C$7)*({sheet_ref}!$B$2:$B$13=$C$8)*"
            f"{sheet_ref}!${col_letter}$2:${col_letter}$13)"
        )

    # H 列辅助
    ws.column_dimensions["H"].width = 18
    ws.column_dimensions["I"].width = 18
    ws.cell(row=5, column=8, value="—— 内部计算辅助 ——").font = Font(italic=True, color="808080")

    ws.cell(row=6, column=8, value="基准实施费").font = Font(color="808080")
    ws.cell(row=6, column=9, value=lookup("C")).number_format = '"¥"#,##0'
    ws.cell(row=7, column=8, value="基准月租").font = Font(color="808080")
    ws.cell(row=7, column=9, value=lookup("D")).number_format = '"¥"#,##0'

    # 模式系数
    # A: 实施费 100%，月租 100%
    # B: 实施费 0%，月租 130%（36 月）
    # C: 实施费 0%，月租 = 基础月租 50% + 后续按分成
    ws.cell(row=8, column=8, value="实施费系数").font = Font(color="808080")
    ws.cell(row=8, column=9,
            value='=IF($C$9="A 经典",1,IF($C$9="B 0首付月供",0,IF($C$9="C 利润分成",0,1)))').number_format = "0.00"
    ws.cell(row=9, column=8, value="月租系数").font = Font(color="808080")
    ws.cell(row=9, column=9,
            value='=IF($C$9="A 经典",1,IF($C$9="B 0首付月供",1.3,IF($C$9="C 利润分成",0.5,1)))').number_format = "0.00"

    # 折扣后实施费
    ws.cell(row=10, column=8, value="实付实施费").font = Font(color="808080")
    ws.cell(row=10, column=9, value="=ROUND(I6*I8*$C$10,-2)").number_format = '"¥"#,##0'

    # 折扣后月租
    ws.cell(row=11, column=8, value="实付月租").font = Font(color="808080")
    ws.cell(row=11, column=9, value="=ROUND(I7*I9*$C$10,-1)").number_format = '"¥"#,##0'

    # 12 期分期月供
    ws.cell(row=12, column=8, value="12 期月供合计").font = Font(color="808080")
    ws.cell(row=12, column=9, value="=ROUND(I10/12+I11,-1)").number_format = '"¥"#,##0'

    # 输出
    out = [
        ("一次性实施费",                "=I10"),
        ("月租（实际）",                 "=I11"),
        ("方案 1：一次付实施费 + 月租",   "=I10+I11*12"),
        ("方案 2：12 期分期月供",        "=I12"),
        ("3 年总投入",                  "=I10+I11*36"),
        ("销售提成",                    '=I10*IF($C$8="种子版",0.05,IF($C$8="经营版",0.08,0.10)) + I11*12*IF($C$8="种子版",0.02,IF($C$8="经营版",0.03,0.04))'),
    ]
    for i, (lbl, formula) in enumerate(out, start=15):
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
        ws.cell(row=i, column=2, value=lbl).font = SUB_FONT
        ws.cell(row=i, column=2).alignment = LEFT
        ws.cell(row=i, column=2).border = BORDER
        ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=6)
        c = ws.cell(row=i, column=5, value=formula); style_text(c, fill=RESULT_FILL)
        c.number_format = '"¥"#,##0'

    # 关键警示：年支出占年利润 %
    ws.merge_cells("B22:D22")
    ws["B22"] = "首年支出占客户年利润 %"
    ws["B22"].font = SUB_FONT; ws["B22"].alignment = LEFT; ws["B22"].border = BORDER
    ws.merge_cells("E22:F22")
    # 首年支出 = I10 + I11*12；除以年利润（万元转元）
    pct_formula = "=IFERROR((I10+I11*12)/(C11*10000),0)"
    pc = ws.cell(row=22, column=5, value=pct_formula)
    style_text(pc, fill=WARN_FILL); pc.number_format = "0.00%"

    ws.merge_cells("B23:F25")
    ws["B23"] = (
        "🚨 健康度参考：\n"
        "  ≤ 1%：非常健康，可推旗舰版\n"
        "  1-1.5%：合理，建议经营版+12 期分期\n"
        "  1.5-3%：偏高，强烈建议 B 模式（0首付月供）或降版本\n"
        "  > 3%：过高，必须降版本或推种子版试点\n\n"
        "🎯 销售三步话术：\n"
        "  ① 先报『方案 2 月供』数字\n"
        "  ② 客户问总价再报实施费\n"
        "  ③ 关单送『前 30 天月租免单』"
    )
    ws["B23"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["B23"].font = Font(name="Microsoft YaHei", size=9, color="595959")
    ws["B23"].border = BORDER


def build_mode_compare(ws):
    """3 种付费模式横向对比"""
    ws.title = "付费模式对比"
    set_col_widths(ws, [4, 24, 22, 22, 22])
    ws.merge_cells("B2:E2"); ws["B2"] = "3 种付费模式横向对比（基于报价计算器输入）"
    ws["B2"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E79")

    headers = ["项目", "A 经典", "B 0首付月供", "C 利润分成"]
    for j, h in enumerate(headers, start=2):
        style_header(ws.cell(row=4, column=j, value=h))

    rows = [
        ("首付实施费", "='报价计算器'!I6", 0, 0),
        ("月租（元/月）", "='报价计算器'!I7", "='报价计算器'!I7*1.3", "='报价计算器'!I7*0.5"),
        ("12 个月支出", "='报价计算器'!I6+'报价计算器'!I7*12",
         "='报价计算器'!I7*1.3*12", "='报价计算器'!I7*0.5*12 + 0"),
        ("36 个月支出", "='报价计算器'!I6+'报价计算器'!I7*36",
         "='报价计算器'!I7*1.3*36", "='报价计算器'!I7*0.5*36 + 0"),
        ("适合客户类型", "现金流好、不喜分期", "现金紧、抠门老板", "高利润大客户、灯塔候选"),
        ("销售难度", "★★★", "★★", "★（最易成交）"),
        ("公司利润", "稳健", "前低后高", "极高（封顶 1.5x）"),
    ]
    for i, row in enumerate(rows, start=5):
        for j, v in enumerate(row, start=2):
            c = ws.cell(row=i, column=j, value=v)
            if j == 2:
                style_text(c, align=LEFT)
            else:
                if i in (5, 6, 7, 8) and isinstance(v, str) and v.startswith("="):
                    style_text(c, fill=RESULT_FILL); c.number_format = '"¥"#,##0'
                else:
                    style_text(c)
        ws.cell(row=i, column=2).font = SUB_FONT


def build_installment(ws):
    """12 期分期模拟"""
    ws.title = "12期分期模拟"
    set_col_widths(ws, [4, 22, 16, 16, 16])
    ws.merge_cells("B2:E2"); ws["B2"] = "12 期分期月供模拟（基于报价计算器输入）"
    ws["B2"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E79")

    headers = ["期数", "实施费分期", "月租", "当月合计"]
    for j, h in enumerate(headers, start=2):
        style_header(ws.cell(row=4, column=j, value=h))
    for m in range(1, 13):
        i = 4 + m
        ws.cell(row=i, column=2, value=f"第 {m} 月").font = SUB_FONT
        ws.cell(row=i, column=2).alignment = CENTER; ws.cell(row=i, column=2).border = BORDER
        c1 = ws.cell(row=i, column=3, value="='报价计算器'!I10/12"); style_text(c1); c1.number_format = '"¥"#,##0'
        c2 = ws.cell(row=i, column=4, value="='报价计算器'!I11"); style_text(c2); c2.number_format = '"¥"#,##0'
        c3 = ws.cell(row=i, column=5, value=f"=C{i}+D{i}"); style_text(c3, fill=RESULT_FILL); c3.number_format = '"¥"#,##0'

    # 合计
    i = 17
    ws.cell(row=i, column=2, value="12 期合计").font = SUB_FONT
    ws.cell(row=i, column=2).alignment = CENTER; ws.cell(row=i, column=2).border = BORDER
    for col_letter in ("C", "D", "E"):
        col_idx = ord(col_letter) - ord("A") + 1
        c = ws.cell(row=i, column=col_idx, value=f"=SUM({col_letter}5:{col_letter}16)")
        style_text(c, fill=RESULT_FILL); c.number_format = '"¥"#,##0'

    # 销售话术
    ws.merge_cells("B19:E22")
    ws["B19"] = (
        "💬 现场销售话术：\n"
        '"X 总，您看这一栏（指 E 列）：每月总共付 ¥XX，相当于多请一个'
        '不要五险一金的数字化助理，3 年下来比您一台二手叉车还便宜。\n'
        "12 期分期免息，是我们和 XX 银行特别谈的合作政策，\n"
        '现在签还送『前 30 天月租免单』，您可以先试一个月再决定要不要继续。"'
    )
    ws["B19"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["B19"].font = Font(name="Microsoft YaHei", size=10, color="595959")
    ws["B19"].border = BORDER


def build_commission(ws):
    ws.title = "返佣测算"
    set_col_widths(ws, [4, 28, 16, 16, 16, 28])
    ws.merge_cells("B2:F2"); ws["B2"] = "推荐返佣测算（基于报价计算器输入）"
    ws["B2"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E79")

    headers = ["推荐人类型", "实施费返佣", "月租返佣", "返佣月数", "返佣金额合计", "结算说明"]
    for j, h in enumerate(headers, start=2):
        style_header(ws.cell(row=4, column=j, value=h))
    rows = [
        ("老客户（现金）",          0.10, 0.05, 12, "现金或对公"),
        ("老客户（抵自家月租 双倍）", 0.20, 0.10, 12, "抵扣本人月订阅费"),
        ("行业顾问 / KOL",         0.15, 0.08, 24, "对公月结"),
        ("区域代理（独家）",        0.38, 0.20, 36, "月结，需达业绩门槛"),
        ("设备厂商联合",           0.20, 0.10, 18, "季结"),
    ]
    for i, (name, im, mn, months, note) in enumerate(rows, start=5):
        ws.cell(row=i, column=2, value=name).font = SUB_FONT
        ws.cell(row=i, column=2).alignment = LEFT; ws.cell(row=i, column=2).border = BORDER

        c = ws.cell(row=i, column=3, value=im); style_text(c); c.number_format = "0%"
        c = ws.cell(row=i, column=4, value=mn); style_text(c); c.number_format = "0%"
        c = ws.cell(row=i, column=5, value=months); style_text(c)
        c = ws.cell(row=i, column=6,
                    value=f"='报价计算器'!I10*C{i} + '报价计算器'!I11*D{i}*E{i}")
        style_text(c, fill=RESULT_FILL); c.number_format = '"¥"#,##0'
        ws.cell(row=i, column=7, value=note).font = TEXT_FONT
        ws.cell(row=i, column=7).alignment = LEFT; ws.cell(row=i, column=7).border = BORDER


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_baseline(wb.create_sheet())
    build_versions(wb.create_sheet())
    build_calculator(wb.create_sheet())
    build_mode_compare(wb.create_sheet())
    build_installment(wb.create_sheet())
    build_commission(wb.create_sheet())

    out = "tools/钢铁数字化系统报价单_模型.xlsx"
    wb.save(out)
    print(f"已生成 v2.0：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
