"""
钢贸定价与 3 年测算模型
========================

输出：tools/钢贸定价与3年测算模型.xlsx

工作表：
    1. 使用说明
    2. 6 大产品价格表           —— ERP/WMS + A/B/C 闭环全套
    3. 5 层客户分层             —— 小微/小型/中型/大型/集团
    4. ARPU 配置器              —— 选客户层 + 选模块 自动算 ARPU + 是否超红线
    5. Y3 客户金字塔            —— 输入总数 + 自动分配 5 层
    6. 3 年营收预测              —— Y1/Y2/Y3 月度
    7. 信用数据服务收入          —— 银行/保理/大买家
    8. 估值与融资                —— PS 倍数 / 估值 / 融资轮次
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
    if pct: c.number_format = "0.00%"


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_readme(ws):
    ws.title = "使用说明"
    widths(ws, [4, 26, 70])
    ws.merge_cells("B2:C2")
    ws["B2"] = "钢贸定价与 3 年测算模型 v2.0"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    rows = [
        ("适用对象", "老板 + 产品 + 销售 + 财务"),
        ("使用步骤",
         "① 打开『6 大产品价格表』看你的全产品价格\n"
         "② 切换『5 层客户分层』看不同客户的付费上限\n"
         "③ 用『ARPU 配置器』给特定客户配套餐 → 自动出 ARPU + 是否超红线\n"
         "④ 切换『Y3 客户金字塔』输入总客户数 → 自动算各层数量 + 总营收\n"
         "⑤ 切换『3 年营收预测』看完整 36 月预测\n"
         "⑥ 信用数据 + 估值 / 融资页给老板做决策"),
        ("核心规则",
         "1. 钢贸商年付费红线 = 流水的 0.05-0.1%\n"
         "2. 中型客户（流水 ¥1-3 亿）是主力，目标 ARPU ¥9,100/月\n"
         "3. 80% 客户应落在「金卡」档位\n"
         "4. 信用数据对外销售（银行/保理/大买家）才是高利润核心"),
        ("修改边界", "只改黄色单元格；公式和表头不要动"),
        ("注意",
         "1. 数字基于行业基准 + 你 6 产品矩阵估算，不是承诺\n"
         "2. 实际偏差 ±30-50% 属正常\n"
         "3. 必须每月用真实数据替换"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        c1 = ws.cell(row=i, column=2, value=k); style_t(c1, fill=PRIMARY, bold=True)
        c1.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
        c2 = ws.cell(row=i, column=3, value=v); style_t(c2, align=LEFT)
        ws.row_dimensions[i].height = max(50, len(v) * 0.6)


def build_pricing(ws):
    ws.title = "6大产品价格表"
    widths(ws, [4, 22, 14, 14, 14, 30])

    ws.merge_cells("B2:F2")
    ws["B2"] = "6 大产品价格表（黄色 = 可调）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["产品", "档位", "实施费(¥)", "月费(¥)", "权益说明"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    products = [
        # (产品, 档位, 实施费, 月费, 权益, 是否主推)
        ("ERP/WMS 底座", "标准版", 4800, 1800, "CRM+基础库存+财务，≤5 用户", False),
        ("ERP/WMS 底座", "经营版★", 9800, 3000, "+多仓 WMS+应收风控+多门店，≤15 用户", True),
        ("ERP/WMS 底座", "旗舰版", 18000, 6800, "+BI+API+集团合并，不限用户", False),
        ("现货大厅", "普通会员", 0, 0, "基础位+5 SKU", False),
        ("现货大厅", "银卡会员", 0, 500, "50 SKU+前 50% 排序", False),
        ("现货大厅", "金卡会员★", 0, 1500, "不限 SKU+前 30% 排序+优先推送+数据看板", True),
        ("现货大厅", "钻卡会员", 0, 4800, "+首页 banner+城市置顶+专属客服", False),
        ("AI 询价", "基础（含底座）", 0, 0, "≤50 次/月", False),
        ("AI 询价", "升级版", 0, 200, "/用户/月，不限+AI 分类+派单+对比", False),
        ("报价工作台", "基础（含底座）", 0, 0, "手动+1 模板", False),
        ("报价工作台", "专业版★", 0, 300, "/用户/月，自动算价+多模板+一键发", True),
        ("报价工作台", "团队版", 0, 800, "/用户/月，+协作+审批+防跑单", False),
        ("提货通", "入门包", 0, 500, "100 单/月，超量 ¥3/单", False),
        ("提货通", "标准包★", 0, 1500, "500 单/月，超量 ¥2/单", True),
        ("提货通", "大客户包", 0, 4800, "2000 单/月，超量 ¥1.5/单", False),
        ("提货通", "不限包", 0, 9800, "不限单量", False),
        ("对账通", "入门版", 0, 300, "/用户/月，≤5 个买家", False),
        ("对账通", "专业版★", 0, 800, "/用户/月，不限买家+应收预警+催收", True),
        ("对账通", "旗舰版", 0, 2000, "/用户/月，+银行系统对接", False),
        ("信用评分", "全等级", 0, 0, "免费！但等级影响流量倾斜", False),
        ("智能分发", "银卡流量", 0, 1000, "/月，+500 询价分发", False),
        ("智能分发", "金卡流量★", 0, 3000, "/月，+1500 询价+城市定向+智能推荐", True),
        ("智能分发", "钻卡流量", 0, 8800, "/月，+5000 询价+全国曝光", False),
    ]

    for i, (prod, level, impl, monthly, desc, is_main) in enumerate(products, 5):
        bg = SUCCESS if is_main else (LIGHT_BG if i % 2 == 0 else WHITE)
        style_t(ws.cell(row=i, column=2, value=prod), fill=bg, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=level), fill=bg, bold=is_main)
        style_t(ws.cell(row=i, column=4, value=impl), fill=INPUT_BG, money=True)
        style_t(ws.cell(row=i, column=5, value=monthly), fill=INPUT_BG, money=True)
        style_t(ws.cell(row=i, column=6, value=desc), fill=bg, align=LEFT)

    # 信用数据服务（外卖）
    last_row = 4 + len(products)
    ws.merge_cells(start_row=last_row+2, start_column=2, end_row=last_row+2, end_column=6)
    ws.cell(row=last_row+2, column=2, value="▼ 信用数据服务（对外卖给银行/保理/大买家，毛利 80%+）").font = Font(
        name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
    ws.cell(row=last_row+2, column=2).fill = PatternFill("solid", fgColor=PRIMARY)
    ws.cell(row=last_row+2, column=2).alignment = LEFT

    headers2 = ["客户类型", "年单价（万）", "目标客户数 Y3", "Y3 年收入（万）"]
    for j, h in enumerate(headers2, 3):
        style_h(ws.cell(row=last_row+3, column=j, value=h))

    data_services = [
        ("银行/信贷", 50, 8, ""),
        ("保理/担保", 30, 12, ""),
        ("大型买家", 10, 25, ""),
        ("行业研究", 20, 8, ""),
        ("政府/协会", 50, 4, ""),
    ]
    for i, (ctype, price, count, _) in enumerate(data_services, last_row+4):
        style_t(ws.cell(row=i, column=2, value=""), fill=WHITE)
        style_t(ws.cell(row=i, column=3, value=ctype), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=4, value=price), fill=INPUT_BG, bold=True)
        style_t(ws.cell(row=i, column=5, value=count), fill=INPUT_BG)
        cell = ws.cell(row=i, column=6, value=f"=D{i}*E{i}")
        style_t(cell, fill=SUCCESS, money=True, bold=True)

    # 合计
    total_row = last_row + 4 + len(data_services)
    style_t(ws.cell(row=total_row, column=3, value="信用数据 Y3 总年收入"),
            fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=total_row, column=3).font = Font(name="Microsoft YaHei", size=10, bold=True, color=WHITE)
    cell = ws.cell(row=total_row, column=6,
                  value=f"=SUM(F{last_row+4}:F{total_row-1})")
    style_t(cell, fill=PRIMARY, money=True, bold=True)
    cell.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)


def build_segments(ws):
    ws.title = "5层客户分层"
    widths(ws, [4, 14, 18, 12, 16, 16, 16, 18])

    ws.merge_cells("B2:H2")
    ws["B2"] = "5 层钢贸商客户分层 + 推荐套餐"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["层", "流水规模(亿)", "占比", "年付费红线(万)",
               "推荐月费(¥)", "推荐年费(万)", "占流水‰"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    segments = [
        ("小微", "≤ 0.3", "40%", 1.8, 2300, 2.76, "0.92‰"),
        ("小型", "0.3-1", "30%", 3.6, 4500, 5.4, "0.6-1.8‰"),
        ("中型★", "1-3", "20%", 7.0, 9100, 10.92, "0.4-1.1‰"),
        ("中型+流量", "1-3", "(主推)", 14.0, 12100, 14.5, "0.5-1.5‰"),
        ("大型", "3-10", "8%", 18.0, 27900, 33.5, "0.3-1.1‰"),
        ("集团", "> 10", "2%", 30.0, 80000, 96.0, "0.3-0.96‰"),
    ]
    for i, row in enumerate(segments, 5):
        bg = SUCCESS if "★" in row[0] else (LIGHT_BG if i % 2 == 0 else WHITE)
        for j, v in enumerate(row, 2):
            cell = ws.cell(row=i, column=j, value=v)
            if isinstance(v, (int, float)) and j in (5, 6, 7):
                if j == 5:
                    style_t(cell, fill=bg, money=True, bold=True)
                elif j == 6:
                    style_t(cell, fill=bg, money=True, bold=True)
                else:
                    style_t(cell, fill=bg, bold=True)
            else:
                style_t(cell, fill=bg, align=LEFT if j == 2 else CENTER, bold=(j == 2))


def build_arpu_calculator(ws):
    ws.title = "ARPU配置器"
    widths(ws, [4, 24, 18, 18, 30])

    ws.merge_cells("B2:E2")
    ws["B2"] = "ARPU 配置器（黄色 = 选你给客户的套餐）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    # 客户基本信息
    style_t(ws.cell(row=4, column=2, value="客户名称"), fill=LIGHT_BG, align=LEFT, bold=True)
    style_t(ws.cell(row=4, column=3, value="某钢贸商"), fill=INPUT_BG, align=LEFT)
    style_t(ws.cell(row=4, column=4, value="年流水(万元)"), fill=LIGHT_BG, align=LEFT, bold=True)
    style_t(ws.cell(row=4, column=5, value=15000), fill=INPUT_BG, align=LEFT)

    # 各产品选择
    ws.merge_cells("B6:E6")
    style_t(ws.cell(row=6, column=2, value="▼ 选你要给客户配的套餐"),
            fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=6, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    headers = ["产品", "档位选择", "用户/单量", "月费(¥)"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=7, column=j, value=h))

    # 产品行
    config_rows = [
        ("ERP/WMS", "经营版", 1, '=IF(C8="标准版",1800,IF(C8="经营版",3000,IF(C8="旗舰版",6800,0)))'),
        ("现货大厅", "金卡", 1, '=IF(C9="普通",0,IF(C9="银卡",500,IF(C9="金卡",1500,IF(C9="钻卡",4800,0))))'),
        ("AI 询价", "升级版", 3, '=IF(C10="升级版",200*D10,0)'),
        ("报价工作台", "专业版", 5, '=IF(C11="专业版",300*D11,IF(C11="团队版",800*D11,0))'),
        ("提货通", "标准包", 1, '=IF(C12="入门",500,IF(C12="标准包",1500,IF(C12="大客户包",4800,IF(C12="不限",9800,0))))'),
        ("对账通", "专业版", 2, '=IF(C13="入门",300*D13,IF(C13="专业版",800*D13,IF(C13="旗舰版",2000*D13,0)))'),
        ("智能分发流量", "金卡", 1, '=IF(C14="银卡",1000,IF(C14="金卡",3000,IF(C14="钻卡",8800,0)))'),
    ]
    for i, (prod, default_lvl, users, formula) in enumerate(config_rows, 8):
        style_t(ws.cell(row=i, column=2, value=prod), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=default_lvl), fill=INPUT_BG, bold=True)
        style_t(ws.cell(row=i, column=4, value=users), fill=INPUT_BG)
        style_t(ws.cell(row=i, column=5, value=formula), fill=SUCCESS, money=True, bold=True)

    # 数据有效性
    dvs = [
        ("C8", '"标准版,经营版,旗舰版"'),
        ("C9", '"普通,银卡,金卡,钻卡"'),
        ("C10", '"基础,升级版"'),
        ("C11", '"基础,专业版,团队版"'),
        ("C12", '"入门,标准包,大客户包,不限"'),
        ("C13", '"入门,专业版,旗舰版"'),
        ("C14", '"无,银卡,金卡,钻卡"'),
    ]
    for cell_ref, formula in dvs:
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        ws.add_data_validation(dv); dv.add(cell_ref)

    # 输出
    out_row = 16
    style_t(ws.cell(row=out_row, column=2, value="月总费用"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=out_row, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=out_row, column=5, value="=SUM(E8:E14)")
    style_t(cell, fill=SUCCESS, money=True, bold=True)
    cell.font = Font(name="Microsoft YaHei", size=12, bold=True)

    style_t(ws.cell(row=out_row+1, column=2, value="年总费用"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=out_row+1, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=out_row+1, column=5, value=f"=E{out_row}*12")
    style_t(cell, fill=SUCCESS, money=True, bold=True)
    cell.font = Font(name="Microsoft YaHei", size=12, bold=True)

    # 占流水比
    style_t(ws.cell(row=out_row+2, column=2, value="占客户年流水"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=out_row+2, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=out_row+2, column=5, value=f"=E{out_row+1}/(E4*10000)")
    style_t(cell, fill=INPUT_BG, pct=True, bold=True)

    # 红线判断
    style_t(ws.cell(row=out_row+3, column=2, value="红线判断"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=out_row+3, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=out_row+3, column=5,
                  value=f'=IF(E{out_row+2}<=0.001,"🟢 健康",IF(E{out_row+2}<=0.0015,"🟡 略偏高","🔴 超线，请降配"))')
    style_t(cell, fill=INPUT_BG, bold=True)

    # 说明
    ws.merge_cells(f"B{out_row+5}:E{out_row+8}")
    ws[f"B{out_row+5}"] = (
        "★ 红线规则：\n"
        "  ≤ 1‰ → 🟢 健康\n"
        "  1-1.5‰ → 🟡 略偏高，建议降配 1 项\n"
        "  > 1.5‰ → 🔴 超线，钢贸商付费意愿急剧下降"
    )
    ws[f"B{out_row+5}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws[f"B{out_row+5}"].font = Font(name="Microsoft YaHei", size=10, color="595959")
    ws[f"B{out_row+5}"].border = BORDER


def build_pyramid(ws):
    ws.title = "Y3客户金字塔"
    widths(ws, [4, 14, 14, 16, 18, 18, 18])

    ws.merge_cells("B2:G2")
    ws["B2"] = "Y3 末（2029.03）客户金字塔 + 营收测算（自动）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    # 输入：Y3 末总客户数
    style_t(ws.cell(row=4, column=2, value="Y3 末总客户数（黄色可调）"),
            fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=4, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    style_t(ws.cell(row=4, column=3, value=2000), fill=INPUT_BG, bold=True)

    headers = ["层", "占比", "客户数", "ARPU(¥/月)", "MRR(¥)", "ARR(¥)"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=6, column=j, value=h))

    layers = [
        ("小微",   0.40, 2300),
        ("小型",   0.30, 4500),
        ("中型",   0.20, 9100),
        ("大型",   0.08, 27900),
        ("集团",   0.02, 80000),
    ]
    for i, (layer, ratio, arpu) in enumerate(layers, 7):
        bg = SUCCESS if layer == "中型" else (LIGHT_BG if i % 2 == 0 else WHITE)
        style_t(ws.cell(row=i, column=2, value=layer), fill=bg, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=ratio), fill=INPUT_BG, pct=True, bold=True)
        cell = ws.cell(row=i, column=4, value=f"=ROUND($C$4*C{i},0)")
        style_t(cell, fill=bg, bold=True)
        style_t(ws.cell(row=i, column=5, value=arpu), fill=INPUT_BG, money=True, bold=True)
        cell = ws.cell(row=i, column=6, value=f"=D{i}*E{i}")
        style_t(cell, fill=SUCCESS, money=True, bold=True)
        cell = ws.cell(row=i, column=7, value=f"=F{i}*12")
        style_t(cell, fill=SUCCESS, money=True, bold=True)

    # 总计
    i = 12
    style_t(ws.cell(row=i, column=2, value="合计"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=i, column=4, value="=SUM(D7:D11)")
    style_t(cell, fill=PRIMARY, bold=True)
    cell.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=i, column=6, value="=SUM(F7:F11)")
    style_t(cell, fill=PRIMARY, money=True, bold=True)
    cell.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=i, column=7, value="=SUM(G7:G11)")
    style_t(cell, fill=PRIMARY, money=True, bold=True)
    cell.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 比率
    ws.merge_cells("B14:G14")
    ws["B14"] = "★ 占比合计应等于 100%（如不等，调整黄色比例）"
    style_t(ws["B14"], fill=WARN, align=LEFT, bold=True)
    cell = ws.cell(row=15, column=2, value="占比合计")
    style_t(cell, fill=LIGHT_BG, align=LEFT, bold=True)
    cell = ws.cell(row=15, column=3, value="=SUM(C7:C11)")
    style_t(cell, fill=INPUT_BG, pct=True, bold=True)


def build_3yr_revenue(ws):
    ws.title = "3年营收预测"
    widths(ws, [4, 26, 14, 14, 14, 14])

    ws.merge_cells("B2:F2")
    ws["B2"] = "钢贸业务 3 年营收预测（基准情景）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["收入来源", "Y1", "Y2", "Y3", "3 年累计"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    revenues = [
        ("钢贸 SaaS 订阅（5 层客户）", 100, 1200, 19000, ""),
        ("信用数据服务（外卖）", 0, 100, 1300, ""),
        ("平台撮合佣金", 30, 400, 2500, ""),
        ("仓储托盘分润", 10, 100, 800, ""),
        ("物流调度分润", 10, 80, 600, ""),
        ("票据/担保服务费", 0, 20, 1000, ""),
    ]
    for i, (src, y1, y2, y3, _) in enumerate(revenues, 5):
        style_t(ws.cell(row=i, column=2, value=src), fill=LIGHT_BG, align=LEFT, bold=True)
        for j, v in enumerate([y1, y2, y3], 3):
            cell = ws.cell(row=i, column=j, value=v*10000)
            style_t(cell, fill=INPUT_BG, money=True)
        # 累计
        cell = ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
        style_t(cell, fill=SUCCESS, money=True, bold=True)

    # 合计
    i = 5 + len(revenues)
    style_t(ws.cell(row=i, column=2, value="钢贸业务合计"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    for j in range(3, 7):
        col = get_column_letter(j)
        cell = ws.cell(row=i, column=j, value=f"=SUM({col}5:{col}{i-1})")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        cell.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    # 加上钢厂主业
    i += 2
    style_t(ws.cell(row=i, column=2, value="+ 钢厂主业（Plan A）"),
            fill=LIGHT_BG, align=LEFT, bold=True)
    for j, v in enumerate([800, 2500, 5500], 3):
        cell = ws.cell(row=i, column=j, value=v*10000)
        style_t(cell, fill=LIGHT_BG, money=True)
    cell = ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
    style_t(cell, fill=SUCCESS, money=True, bold=True)

    # 3 业务总营收
    i += 2
    style_t(ws.cell(row=i, column=2, value="★ 3 业务总营收"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=12, bold=True, color=WHITE)
    for j in range(3, 7):
        col = get_column_letter(j)
        cell = ws.cell(row=i, column=j, value=f"={col}{i-4}+{col}{i-2}")
        style_t(cell, fill=PRIMARY, money=True, bold=True)
        cell.font = Font(name="Microsoft YaHei", size=12, bold=True, color=WHITE)


def build_credit_data(ws):
    ws.title = "信用数据服务"
    widths(ws, [4, 22, 14, 16, 16])

    ws.merge_cells("B2:E2")
    ws["B2"] = "C 闭环：信用数据服务（外卖给银行/保理/大买家）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    ws.merge_cells("B4:E4")
    ws["B4"] = "★ 这是 3 业务里毛利最高的部分，毛利 80%+，必须早布局"
    style_t(ws["B4"], fill=SUCCESS, align=LEFT, bold=True)

    headers = ["客户类型", "年单价(万)", "Y3 客户数", "Y3 年收入(万)"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=6, column=j, value=h))

    services = [
        ("银行 / 信贷",     50, 8),
        ("保理 / 担保",     30, 12),
        ("大型买家",        10, 25),
        ("行业研究 / 评级", 20, 8),
        ("政府 / 协会",     50, 4),
    ]
    for i, (ctype, price, count) in enumerate(services, 7):
        style_t(ws.cell(row=i, column=2, value=ctype), fill=LIGHT_BG, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=price), fill=INPUT_BG, money=True)
        ws.cell(row=i, column=3).number_format = '0.0"万"'
        style_t(ws.cell(row=i, column=4, value=count), fill=INPUT_BG)
        cell = ws.cell(row=i, column=5, value=f"=C{i}*D{i}")
        style_t(cell, fill=SUCCESS, bold=True)
        cell.number_format = '0"万"'

    # 合计
    i = 12
    style_t(ws.cell(row=i, column=2, value="Y3 信用数据总收入"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    cell = ws.cell(row=i, column=5, value="=SUM(E7:E11)")
    style_t(cell, fill=PRIMARY, bold=True)
    cell.font = Font(name="Microsoft YaHei", size=12, bold=True, color=WHITE)
    cell.number_format = '0"万"'

    # 增长曲线
    ws.merge_cells("B14:E14")
    ws["B14"] = "★ 信用数据 3 年增长曲线"
    style_t(ws["B14"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B14"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    headers2 = ["年", "客户数", "年收入(万)", "毛利率"]
    for j, h in enumerate(headers2, 2):
        style_h(ws.cell(row=15, column=j, value=h))
    growth = [
        ("Y1", 0, 0, "—"),
        ("Y2", 5, 100, "60%"),
        ("Y3", 57, 1300, "80%"),
        ("Y4 预期", 100, 2500, "85%"),
        ("Y5 预期", 200, 5000, "85%"),
    ]
    for i, row in enumerate(growth, 16):
        bg = SUCCESS if row[0] in ("Y3", "Y4 预期", "Y5 预期") else LIGHT_BG
        for j, v in enumerate(row, 2):
            cell = ws.cell(row=i, column=j, value=v)
            if j == 2:
                style_t(cell, fill=bg, bold=True)
            elif j == 4 and isinstance(v, (int, float)):
                style_t(cell, fill=bg, bold=True)
                cell.number_format = '0"万"'
            else:
                style_t(cell, fill=bg)


def build_valuation(ws):
    ws.title = "估值与融资"
    widths(ws, [4, 24, 18, 18, 30])

    ws.merge_cells("B2:E2")
    ws["B2"] = "估值与融资测算（按 SaaS PS 倍数）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    # 输入
    style_t(ws.cell(row=4, column=2, value="Y3 末年化 ARR (¥)"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=4, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    style_t(ws.cell(row=4, column=3, value=300000000), fill=INPUT_BG, money=True, bold=True)

    style_t(ws.cell(row=5, column=2, value="增长率 (年)"), fill=PRIMARY, align=LEFT, bold=True)
    ws.cell(row=5, column=2).font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
    style_t(ws.cell(row=5, column=3, value=0.7), fill=INPUT_BG, pct=True, bold=True)

    # PS 倍数表
    ws.merge_cells("B7:E7")
    ws["B7"] = "PS 倍数 → 估值"
    style_t(ws["B7"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B7"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    headers = ["情景", "PS 倍数", "估值 (¥)", "说明"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=8, column=j, value=h))

    scenarios = [
        ("保守", 5, "增长率不及预期"),
        ("基准", 8, "市场平均"),
        ("乐观", 12, "高增长 + 数据壁垒"),
        ("极端乐观", 15, "上市前估值"),
    ]
    for i, (sce, ps, note) in enumerate(scenarios, 9):
        bg = SUCCESS if sce == "基准" else (LIGHT_BG if i % 2 == 0 else WHITE)
        style_t(ws.cell(row=i, column=2, value=sce), fill=bg, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=ps), fill=INPUT_BG, bold=True)
        cell = ws.cell(row=i, column=4, value=f"=$C$4*C{i}")
        style_t(cell, fill=SUCCESS, money=True, bold=True)
        style_t(ws.cell(row=i, column=5, value=note), fill=bg, align=LEFT)

    # 融资轮次建议
    ws.merge_cells("B14:E14")
    ws["B14"] = "★ 融资轮次建议（基于估值）"
    style_t(ws["B14"], fill=PRIMARY, align=LEFT, bold=True)
    ws["B14"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)

    headers2 = ["阶段", "时点", "融资规模(¥)", "稀释 %"]
    for j, h in enumerate(headers2, 2):
        style_h(ws.cell(row=15, column=j, value=h))

    rounds = [
        ("天使轮", "Y1 末 (2027.03)", 5000000, "5-10%"),
        ("Pre-A 轮", "Y2 末 (2028.03)", 30000000, "10-15%"),
        ("A 轮", "Y3 末 (2029.03)", 100000000, "10-15%"),
        ("B 轮", "Y4 末 (2030.03)", 300000000, "8-12%"),
        ("Pre-IPO", "Y5 (2031)", 800000000, "5-10%"),
    ]
    for i, (rnd, when, amount, dilute) in enumerate(rounds, 16):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=rnd), fill=bg, align=LEFT, bold=True)
        style_t(ws.cell(row=i, column=3, value=when), fill=bg)
        style_t(ws.cell(row=i, column=4, value=amount), fill=INPUT_BG, money=True, bold=True)
        style_t(ws.cell(row=i, column=5, value=dilute), fill=bg)


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_pricing(wb.create_sheet())
    build_segments(wb.create_sheet())
    build_arpu_calculator(wb.create_sheet())
    build_pyramid(wb.create_sheet())
    build_3yr_revenue(wb.create_sheet())
    build_credit_data(wb.create_sheet())
    build_valuation(wb.create_sheet())

    out = "tools/钢贸定价与3年测算模型.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
