"""
报价单 Excel 模型生成器
=======================

运行方式：
    python3 tools/build_quote_model.py

产出文件：
    tools/钢铁数字化系统报价单_模型.xlsx

工作表说明：
    1. 使用说明      —— 怎么用，谁能用，注意事项
    2. 价格基准表    —— 实施费/月租/模块的指导价（销售可在 ±15% 内调）
    3. 报价计算器    —— 销售在这里输入参数，自动出报价
    4. 模块加购菜单  —— 月租加购明细，被报价计算器引用
    5. 三年现金流    —— 一次性 + 月租折算 1/2/3 年总投入（含年付折扣）
    6. 返佣测算      —— 推荐人收益快速测算
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
TITLE_FONT = Font(name="Microsoft YaHei", size=14, bold=True, color="FFFFFF")
HEAD_FONT = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
SUB_FONT = Font(name="Microsoft YaHei", size=11, bold=True)
TEXT_FONT = Font(name="Microsoft YaHei", size=10)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_header(cell):
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = CENTER
    cell.border = BORDER


def style_sub(cell):
    cell.fill = SUB_FILL
    cell.font = SUB_FONT
    cell.alignment = CENTER
    cell.border = BORDER


def style_text(cell, fill=None, align=None):
    cell.font = TEXT_FONT
    cell.alignment = align or CENTER
    cell.border = BORDER
    if fill:
        cell.fill = fill


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ----------------------------------------------------------------------
# Sheet 1：使用说明
# ----------------------------------------------------------------------
def build_readme(ws):
    ws.title = "使用说明"
    set_col_widths(ws, [4, 60, 30])
    ws.merge_cells("B2:C2")
    ws["B2"] = "钢铁行业数字化系统报价单（v1.0）"
    ws["B2"].font = Font(name="Microsoft YaHei", size=18, bold=True, color="1F4E79")
    ws["B2"].alignment = LEFT

    rows = [
        ("适用对象", "公司销售、售前、渠道、合伙人。客户侧请用导出 PDF 版。"),
        ("使用步骤",
         "① 打开『报价计算器』，按黄色单元格填写客户类型、产线数、模块勾选、付款方式、折扣率\n"
         "② 绿色单元格自动计算报价\n"
         "③ 切换到『三年现金流』查看 1/2/3 年总成本（含年付折扣）\n"
         "④ 切换到『返佣测算』给推荐人算出预期收益"),
        ("折扣权限",
         "销售：±5% 自批；销售总监：±15% 自批；超出由总经理审批。\n"
         "一体厂自动叠加 88 折，已写在公式里，不要再二次打折。"),
        ("不达标退款承诺",
         "标准合同含 6 个月 KPI 退款条款，禁止销售口头承诺更激进的条件。"),
        ("数据维护",
         "『价格基准表』『模块加购菜单』可由产品/财务统一更新；不要在『报价计算器』里硬改公式。"),
        ("注意事项",
         "1. 黄色 = 输入区；绿色 = 自动计算；蓝色 = 表头；不要改其它单元格\n"
         "2. 任何低于成本线的报价必须出具《特价审批单》\n"
         "3. 报价单有效期默认 30 天，过期需重新生成"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        ws[f"B{i}"] = k
        ws[f"C{i}"] = v
        style_sub(ws[f"B{i}"])
        style_text(ws[f"C{i}"], align=LEFT)
        ws.row_dimensions[i].height = 60 if len(v) > 30 else 30
        # restore wide column for column C
        ws.column_dimensions["C"].width = 80


# ----------------------------------------------------------------------
# Sheet 2：价格基准表
# ----------------------------------------------------------------------
# 行：客户类型 + 档位；列：产线下限/上限/实施费下限/上限/基础月费/单产线月费
PRICE_ROWS = [
    # 客户类型 / 档位 / 产线下限 / 产线上限 / 实施费下限 / 实施费上限 / 基础月费 / 单产线月费
    ("开平纵剪加工", "入门",   1, 1,  60000,  90000, 3500, 2500),
    ("开平纵剪加工", "标准",   2, 3, 100000, 150000, 5500, 2200),
    ("开平纵剪加工", "旗舰",   4, 99, 180000, 250000, 8500, 1900),
    ("钢管生产企业", "入门",   1, 2, 120000, 180000, 7500, 3500),
    ("钢管生产企业", "标准",   3, 5, 220000, 300000, 11000, 3000),
    ("钢管生产企业", "旗舰",   6, 99, 350000, 500000, 17000, 2500),
    ("镀锌冷卷企业", "单机组", 1, 1, 180000, 250000, 9500, 6000),
    ("镀锌冷卷企业", "双机组", 2, 2, 280000, 380000, 14000, 5500),
    ("镀锌冷卷企业", "多机组", 3, 99, 450000, 650000, 21000, 5000),
    ("一体厂",      "入门",   1, 99, 500000, 700000, 18000, 4500),
    ("一体厂",      "标准",   3, 99, 700000, 900000, 24000, 4200),
    ("一体厂",      "旗舰",   5, 99, 900000, 1200000, 32000, 4000),
]


def build_baseline(ws):
    ws.title = "价格基准表"
    headers = ["客户类型", "档位", "产线数下限", "产线数上限",
               "实施费下限(元)", "实施费上限(元)", "基础月费(元/月)", "单产线月费(元/月)"]
    set_col_widths(ws, [16, 10, 12, 12, 16, 16, 16, 18])
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j, value=h)
        style_header(c)
    for i, row in enumerate(PRICE_ROWS, start=2):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=v)
            style_text(c)
            if j >= 5:
                c.number_format = '"¥"#,##0'

    # 一体厂综合折扣说明
    ws.cell(row=len(PRICE_ROWS) + 4, column=1, value="说明：")
    ws.cell(row=len(PRICE_ROWS) + 4, column=2,
            value="一体厂报价 = 钢管 + 镀锌叠加后再打 88 折；销售如直接选『一体厂』档位则使用上表数值。")
    ws.merge_cells(start_row=len(PRICE_ROWS) + 4, start_column=2, end_row=len(PRICE_ROWS) + 4, end_column=8)
    ws.cell(row=len(PRICE_ROWS) + 4, column=1).font = SUB_FONT
    ws.cell(row=len(PRICE_ROWS) + 4, column=2).font = TEXT_FONT


# ----------------------------------------------------------------------
# Sheet 3：模块加购菜单
# ----------------------------------------------------------------------
MODULE_ROWS = [
    # 模块 / 默认是否含 / 月租加价(元) / 备注
    ("ERP 基础",         "含",   0, "标配"),
    ("MES 核心(计划/排程/质检)", "含",   0, "标配"),
    ("质量追溯(卷号/炉批号)",  "可选", 2000, "A/C 类首选首单"),
    ("设备管理 EAM",      "可选", 2200, "C/D 类必推"),
    ("WMS 仓储",         "可选", 2200, "多仓再加 ¥1,000"),
    ("财务系统",          "可选", 2000, "可对接金蝶/用友"),
    ("成本自动化核算",     "可选", 3000, "A 类首选首单"),
    ("MES 旗舰版升级",    "可选", 2000, "高级排程/OEE/SPC"),
    ("多仓位扩展",        "可选", 1000, "WMS 启用后可加"),
    ("移动 BI 看板",      "可选", 1500, "老板手机端"),
]


def build_modules(ws):
    ws.title = "模块加购菜单"
    headers = ["模块", "默认", "月租加价(元/月)", "备注"]
    set_col_widths(ws, [28, 10, 18, 36])
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j, value=h)
        style_header(c)
    for i, row in enumerate(MODULE_ROWS, start=2):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=v)
            style_text(c)
            if j == 3:
                c.number_format = '"¥"#,##0'


# ----------------------------------------------------------------------
# Sheet 4：报价计算器
# ----------------------------------------------------------------------
def build_calculator(ws):
    ws.title = "报价计算器"
    set_col_widths(ws, [4, 22, 22, 22, 22, 22])

    # 标题
    ws.merge_cells("B2:F2")
    ws["B2"] = "客户报价计算器"
    ws["B2"].font = Font(name="Microsoft YaHei", size=16, bold=True, color="1F4E79")
    ws["B2"].alignment = LEFT

    # 输入区
    ws.merge_cells("B4:C4")
    ws["B4"] = "一、客户基础信息（黄色 = 必填）"
    style_sub(ws["B4"])

    labels_inputs = [
        ("客户全称", ""),
        ("联系人 / 职务 / 电话", ""),
        ("客户类型", "钢管生产企业"),
        ("档位", "标准"),
        ("产线 / 机组数", 3),
        ("销售折扣率(0.85~1.00)", 1.00),
        ("付款方式", "月付"),
    ]
    for i, (label, default) in enumerate(labels_inputs, start=5):
        ws.cell(row=i, column=2, value=label).font = SUB_FONT
        ws.cell(row=i, column=2).alignment = LEFT
        ws.cell(row=i, column=2).border = BORDER
        c = ws.cell(row=i, column=3, value=default)
        style_text(c, fill=INPUT_FILL, align=LEFT)
    # 档位/类型/付款方式 数据有效性
    dv_type = DataValidation(type="list", formula1='"开平纵剪加工,钢管生产企业,镀锌冷卷企业,一体厂"', allow_blank=False)
    ws.add_data_validation(dv_type); dv_type.add("C7")
    dv_tier = DataValidation(type="list", formula1='"入门,标准,旗舰,单机组,双机组,多机组"', allow_blank=False)
    ws.add_data_validation(dv_tier); dv_tier.add("C8")
    dv_pay = DataValidation(type="list", formula1='"月付,季付,年付"', allow_blank=False)
    ws.add_data_validation(dv_pay); dv_pay.add("C11")
    dv_disc = DataValidation(type="decimal", operator="between", formula1=0.85, formula2=1.0,
                              error="销售自批折扣范围 0.85-1.00", errorTitle="折扣超限")
    ws.add_data_validation(dv_disc); dv_disc.add("C10")

    # 模块勾选
    ws.merge_cells("E4:F4")
    ws["E4"] = "二、模块勾选（√ = 启用）"
    style_sub(ws["E4"])
    module_pick_rows = [
        ("质量追溯(卷号/炉批号)", "√"),
        ("设备管理 EAM", ""),
        ("WMS 仓储", ""),
        ("财务系统", ""),
        ("成本自动化核算", "√"),
        ("MES 旗舰版升级", ""),
        ("多仓位扩展", ""),
        ("移动 BI 看板", ""),
    ]
    for i, (m, p) in enumerate(module_pick_rows, start=5):
        c1 = ws.cell(row=i, column=5, value=m)
        c1.font = TEXT_FONT; c1.alignment = LEFT; c1.border = BORDER
        c2 = ws.cell(row=i, column=6, value=p)
        style_text(c2, fill=INPUT_FILL)
    # √ 数据有效性
    dv_chk = DataValidation(type="list", formula1='"√,"', allow_blank=True)
    ws.add_data_validation(dv_chk); dv_chk.add("F5:F12")

    # ------------- 计算区 -------------
    ws.merge_cells("B14:F14")
    ws["B14"] = "三、自动计算（绿色 = 输出）"
    style_sub(ws["B14"])

    # 基准查找：使用 INDEX/MATCH（按"客户类型"+"档位"组合）
    # 在『价格基准表』B2:H13 区域 (行 2..13)
    # MATCH 复合键：客户类型 & 档位
    # 我们在 calculator 内构建一个 helper：H 列存复合键

    # 先在 H 列做隐藏 helper（仍可见，但说明为内部用）
    ws.cell(row=5, column=8, value="—— 内部计算辅助列 ——").font = Font(italic=True, color="808080")
    ws.column_dimensions["H"].width = 18
    ws.column_dimensions["I"].width = 18
    ws.column_dimensions["J"].width = 18

    # 查找基准
    # 基准表行 2-13；列 1=类型 2=档位 5=实施下 6=实施上 7=基础月费 8=单产线月费
    sheet_ref = "'价格基准表'"
    # 用 SUMPRODUCT 做组合匹配
    def lookup(col_letter):
        return (
            f"=SUMPRODUCT(({sheet_ref}!$A$2:$A$13=$C$7)*({sheet_ref}!$B$2:$B$13=$C$8)*{sheet_ref}!${col_letter}$2:${col_letter}$13)"
        )

    ws.cell(row=6, column=8, value="实施费下限").font = Font(color="808080")
    ws.cell(row=6, column=9, value=lookup("E")).number_format = '"¥"#,##0'
    ws.cell(row=7, column=8, value="实施费上限").font = Font(color="808080")
    ws.cell(row=7, column=9, value=lookup("F")).number_format = '"¥"#,##0'
    ws.cell(row=8, column=8, value="基础月费").font = Font(color="808080")
    ws.cell(row=8, column=9, value=lookup("G")).number_format = '"¥"#,##0'
    ws.cell(row=9, column=8, value="单产线月费").font = Font(color="808080")
    ws.cell(row=9, column=9, value=lookup("H")).number_format = '"¥"#,##0'

    # 模块加价合计：按 F5:F12 与 模块菜单 C2:C9 对应（顺序一致）
    # 用条件求和
    mod_sheet = "'模块加购菜单'"
    # 模块菜单中：可选模块在 C4:C11（行 4-11，对应顺序 = 质量追溯..移动BI）
    # 在 F5:F12 中"√"代表选中
    ws.cell(row=10, column=8, value="模块加价合计").font = Font(color="808080")
    ws.cell(row=10, column=9, value=(
        f"=SUMPRODUCT(($F$5:$F$12=\"√\")*{mod_sheet}!$C$4:$C$11)"
    )).number_format = '"¥"#,##0'

    # 实施费推荐取上下限平均 * 折扣
    ws.cell(row=11, column=8, value="实施费推荐").font = Font(color="808080")
    ws.cell(row=11, column=9, value="=ROUND(((I6+I7)/2)*$C$10,-2)").number_format = '"¥"#,##0'

    # 月订阅费 = 基础 + 单产线*产线数 + 模块加价；统一乘以折扣
    ws.cell(row=12, column=8, value="月订阅费(月付)").font = Font(color="808080")
    ws.cell(row=12, column=9, value="=ROUND((I8 + I9*$C$9 + I10)*$C$10,-1)").number_format = '"¥"#,##0'

    # 季付/年付折扣
    ws.cell(row=13, column=8, value="支付方式系数").font = Font(color="808080")
    ws.cell(row=13, column=9, value='=IF($C$11="年付",0.9,IF($C$11="季付",0.97,1))').number_format = "0.00"

    ws.cell(row=14, column=8, value="月订阅费(实付)").font = Font(color="808080")
    ws.cell(row=14, column=9, value="=ROUND(I12*I13,-1)").number_format = '"¥"#,##0'

    # ----- 输出（绿色） -----
    out = [
        ("一次性实施费（建议成交价）", "=I11"),
        ("月订阅费（实际收取）",     "=I14"),
        ("一年总投入（实施费 + 月费*12）", "=I11 + I14*12"),
        ("三年总投入（实施费 + 月费*36）", "=I11 + I14*36"),
        ("销售提成（实施费 8% + 12 个月月租 3%）", "=I11*0.08 + I14*12*0.03"),
    ]
    for i, (lbl, formula) in enumerate(out, start=16):
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
        ws.cell(row=i, column=2, value=lbl).font = SUB_FONT
        ws.cell(row=i, column=2).alignment = LEFT
        ws.cell(row=i, column=2).border = BORDER
        ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=6)
        c = ws.cell(row=i, column=5, value=formula)
        style_text(c, fill=RESULT_FILL)
        c.number_format = '"¥"#,##0'

    # 备注
    ws.merge_cells("B22:F25")
    ws["B22"] = (
        "备注：\n"
        "1. 一次性实施费默认 = (基准下限+上限)/2 × 折扣率，可按客户复杂度上下调整\n"
        "2. 月订阅费 = (基础月费 + 单产线月费 × 产线数 + 模块加价) × 折扣率 × 支付方式系数\n"
        "3. 一体厂建议直接选『一体厂』客户类型，已隐含 88 折逻辑（基准表数值已综合）\n"
        "4. 销售自批折扣 0.95-1.00；销售总监 0.85-0.95；低于 0.85 须总经理审批"
    )
    ws["B22"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["B22"].font = Font(name="Microsoft YaHei", size=9, color="595959")
    ws["B22"].border = BORDER


# ----------------------------------------------------------------------
# Sheet 5：三年现金流
# ----------------------------------------------------------------------
def build_cashflow(ws):
    ws.title = "三年现金流"
    set_col_widths(ws, [4, 24, 18, 18, 18])
    ws.merge_cells("B2:E2")
    ws["B2"] = "三年总投入对比（自动引用『报价计算器』）"
    ws["B2"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E79")

    headers = ["项目", "第 1 年", "第 2 年", "第 3 年"]
    for j, h in enumerate(headers, start=2):
        c = ws.cell(row=4, column=j, value=h)
        style_header(c)
    rows = [
        ("一次性实施费", "='报价计算器'!I11", 0, 0),
        ("月订阅费 × 12", "='报价计算器'!I14*12", "='报价计算器'!I14*12", "='报价计算器'!I14*12"),
        ("年付折扣返还(年付时已含)", 0, 0, 0),
        ("年度合计", "=C5+C6+C7", "=D5+D6+D7", "=E5+E6+E7"),
        ("累计合计", "=C8", "=C9+D8", "=D9+E8"),
    ]
    for i, row in enumerate(rows, start=5):
        for j, v in enumerate(row, start=2):
            c = ws.cell(row=i, column=j, value=v)
            if j == 2:
                style_text(c, align=LEFT)
            else:
                style_text(c, fill=RESULT_FILL if i in (8, 9) else None)
                c.number_format = '"¥"#,##0'
        ws.cell(row=i, column=2).font = SUB_FONT


# ----------------------------------------------------------------------
# Sheet 6：返佣测算
# ----------------------------------------------------------------------
def build_commission(ws):
    ws.title = "返佣测算"
    set_col_widths(ws, [4, 28, 18, 18, 18, 28])
    ws.merge_cells("B2:F2")
    ws["B2"] = "推荐返佣快速测算（自动引用『报价计算器』）"
    ws["B2"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E79")

    ws.merge_cells("B4:F4")
    ws["B4"] = "假设以本次报价的客户作为被推荐客户来测算推荐人收益"
    style_sub(ws["B4"])

    headers = ["推荐人类型", "实施费返佣比例", "月租返佣比例", "返佣月数", "返佣金额合计", "结算说明"]
    for j, h in enumerate(headers, start=2):
        c = ws.cell(row=6, column=j, value=h)
        style_header(c)

    rows = [
        ("老客户(现金)",        0.10, 0.05, 12, "现金或对公"),
        ("老客户(抵自家月租 双倍)", 0.20, 0.10, 12, "抵扣本人月订阅费"),
        ("行业顾问/KOL",       0.15, 0.08, 24, "对公月结"),
        ("区域代理(独家)",      0.38, 0.20, 36, "月结，需达业绩门槛"),
        ("设备厂商联合",       0.20, 0.10, 18, "季结+联合活动费"),
    ]
    for i, (name, im, mn, months, note) in enumerate(rows, start=7):
        ws.cell(row=i, column=2, value=name).font = SUB_FONT
        ws.cell(row=i, column=2).alignment = LEFT
        ws.cell(row=i, column=2).border = BORDER

        c = ws.cell(row=i, column=3, value=im); style_text(c); c.number_format = "0%"
        c = ws.cell(row=i, column=4, value=mn); style_text(c); c.number_format = "0%"
        c = ws.cell(row=i, column=5, value=months); style_text(c)
        # 返佣金额 = 实施费*比例 + 月费*月租比例*月数
        formula = f"='报价计算器'!I11*C{i} + '报价计算器'!I14*D{i}*E{i}"
        c = ws.cell(row=i, column=6, value=formula); style_text(c, fill=RESULT_FILL); c.number_format = '"¥"#,##0'
        ws.cell(row=i, column=7, value=note).font = TEXT_FONT
        ws.cell(row=i, column=7).alignment = LEFT
        ws.cell(row=i, column=7).border = BORDER


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_baseline(wb.create_sheet())
    build_modules(wb.create_sheet())
    build_calculator(wb.create_sheet())
    build_cashflow(wb.create_sheet())
    build_commission(wb.create_sheet())

    out = "tools/钢铁数字化系统报价单_模型.xlsx"
    wb.save(out)
    print(f"已生成：{out}")


if __name__ == "__main__":
    main()
