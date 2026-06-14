"""生成 CRM 分级评分计算器 Excel。

输出：tools/CRM分级评分计算器.xlsx
工作表：
  0 客户档案（基础信息 + 业态选择）
  1 价值维度（V1-V4 录入 + 自动打分）
  2 利润维度（P1-P3）
  3 付款维度（R1-R4）
  4 粘性维度（L1-L3）
  5 潜力维度（G1-G4）
  6 综合评分（自动加权 + 等级 + 强制规则）
  7 5 级 SOP（自动对应行动清单）
  8 批量录入（50 客户批量评级）
  9 维度说明
"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

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

WHITE_BOLD = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
BOLD = Font(name="Microsoft YaHei", size=11, bold=True)
NORMAL = Font(name="Microsoft YaHei", size=10)
SMALL = Font(name="Microsoft YaHei", size=9, color="555555")
TITLE = Font(name="Microsoft YaHei", size=16, bold=True, color="1F4E78")


def set_col_widths(ws: Worksheet, widths: list[float]) -> None:
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_title(ws: Worksheet, row: int, col: int, span: int, text: str, fill=NAVY, font=WHITE_BOLD, height=30) -> None:
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    c = ws.cell(row=row, column=col, value=text)
    c.fill = fill
    c.font = font
    c.alignment = CENTER
    c.border = BORDER
    ws.row_dimensions[row].height = height


def write_row(ws: Worksheet, row: int, values: list, fills=None, fonts=None, aligns=None) -> None:
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=i + 1, value=v)
        c.border = BORDER
        c.font = fonts[i] if isinstance(fonts, list) else (fonts or NORMAL)
        fill_val = fills[i] if isinstance(fills, list) else fills
        if isinstance(fill_val, PatternFill):
            c.fill = fill_val
        c.alignment = aligns[i] if isinstance(aligns, list) else (aligns or CENTER)


# -----------------------------------------------------------------------------
# 0 客户档案
# -----------------------------------------------------------------------------
def build_profile(wb: Workbook) -> None:
    ws = wb.create_sheet("0 客户档案")
    set_col_widths(ws, [4, 22, 28, 8, 22, 22])

    write_title(ws, 1, 1, 6, "CRM 分级评分计算器 · 客户档案")
    write_title(ws, 2, 1, 6, "录入客户基础信息（业态会影响阈值调整）", fill=LIGHT_BLUE, font=BOLD, height=22)

    fields = [
        ("客户公司名", "山东 XX 钢管科技"),
        ("老板姓名", "王 X X"),
        ("客户业态（钢贸/钢厂/加工/镀锌）", "钢贸"),
        ("所在地区（长三角/珠三角/京津冀/华中/西南/其他）", "京津冀"),
        ("年流水（万元）", 18000),
        ("员工数", 50),
        ("成立年份", 2005),
        ("销售归属", "李 X X"),
        ("评分日期", "2026-06-15"),
        ("数据完整度（%）", 95),
    ]
    r = 4
    for k, v in fields:
        ws.cell(row=r, column=2, value=k).font = BOLD
        ws.cell(row=r, column=2).fill = LIGHT_GRAY
        ws.cell(row=r, column=2).alignment = LEFT
        ws.cell(row=r, column=2).border = BORDER
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        c = ws.cell(row=r, column=3, value=v)
        c.fill = LIGHT_YELLOW
        c.font = BOLD
        c.alignment = LEFT
        c.border = BORDER
        r += 1

    # 业态选择数据验证
    dv = DataValidation(type="list", formula1='"钢贸,钢厂,加工,镀锌,综合"', allow_blank=True)
    dv.add("C6")
    ws.add_data_validation(dv)
    dv2 = DataValidation(type="list", formula1='"长三角,珠三角,京津冀,华中,西南,其他"', allow_blank=True)
    dv2.add("C7")
    ws.add_data_validation(dv2)

    r += 1
    write_title(ws, r, 1, 6, "业态阈值调整说明", fill=ORANGE, font=WHITE_BOLD, height=22)
    r += 1
    tips = [
        ("钢贸（默认）", "基准阈值"),
        ("钢厂", "V1 ×3 / P1 ×1.5 / L1 权重 ↑"),
        ("加工", "V1 ×0.5 / V2 权重 ↑ / P1 ×0.7"),
        ("镀锌", "V1 ×2 / P1 ×1.2 / 加入锌价对冲信号"),
    ]
    for k, v in tips:
        ws.cell(row=r, column=2, value=k).font = BOLD
        ws.cell(row=r, column=2).border = BORDER
        ws.cell(row=r, column=2).alignment = LEFT
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        c = ws.cell(row=r, column=3, value=v)
        c.font = NORMAL
        c.alignment = LEFT
        c.border = BORDER
        r += 1


# 通用维度构建器
def build_dimension(
    wb: Workbook,
    tab_name: str,
    title: str,
    sub_indicators: list[tuple[str, str, str, str, list[tuple], int]],
    overall_formula_desc: str,
    weight_in_total: str,
) -> tuple[str, int]:
    """构建一个维度的 Tab.

    sub_indicators: [(代码, 名称, 数据来源, 单位, 等级阈值表, 默认值)]
        等级阈值表: [(阈值描述, 单项分)]
    返回：(tab_name, 维度综合分所在行号) — 综合评分 Tab 引用用
    """
    ws = wb.create_sheet(tab_name)
    set_col_widths(ws, [4, 24, 22, 14, 14, 30, 12])

    write_title(ws, 1, 1, 7, f"{title}（权重 {weight_in_total} 占综合）")
    write_title(ws, 2, 1, 7, "在【实测值】填入您的客户数据，单项分自动计算", fill=LIGHT_BLUE, font=BOLD, height=22)

    r = 4
    write_title(ws, r, 1, 7, "子指标录入", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    headers = ["#", "子指标名称", "数据来源", "单位", "实测值", "等级阈值表", "单项分"]
    write_row(ws, r, headers, fills=LIGHT_BLUE, fonts=BOLD)
    r += 1

    score_rows = []
    weights = []

    for idx, (code, name, source, unit, thresholds, default) in enumerate(sub_indicators, 1):
        # 主行
        ws.cell(row=r, column=1, value=idx).border = BORDER
        ws.cell(row=r, column=1).alignment = CENTER
        ws.cell(row=r, column=1).font = NORMAL

        ws.cell(row=r, column=2, value=f"{code} {name}").font = BOLD
        ws.cell(row=r, column=2).border = BORDER
        ws.cell(row=r, column=2).alignment = LEFT

        ws.cell(row=r, column=3, value=source).font = SMALL
        ws.cell(row=r, column=3).border = BORDER
        ws.cell(row=r, column=3).alignment = LEFT

        ws.cell(row=r, column=4, value=unit).font = NORMAL
        ws.cell(row=r, column=4).border = BORDER
        ws.cell(row=r, column=4).alignment = CENTER

        c = ws.cell(row=r, column=5, value=default)
        c.fill = LIGHT_YELLOW
        c.font = BOLD
        c.border = BORDER
        c.alignment = CENTER

        # 阈值描述（多行）
        thresh_str = "\n".join(f"{desc} → {pts}" for desc, pts, _ in thresholds)
        ws.cell(row=r, column=6, value=thresh_str).font = SMALL
        ws.cell(row=r, column=6).border = BORDER
        ws.cell(row=r, column=6).alignment = LEFT
        ws.row_dimensions[r].height = max(60, 14 * len(thresholds))

        # 单项分公式（nested IF）
        score_formula = build_score_formula(f"E{r}", thresholds)
        c = ws.cell(row=r, column=7, value=score_formula)
        c.fill = LIGHT_GREEN
        c.font = BOLD
        c.border = BORDER
        c.alignment = CENTER

        score_rows.append((code, r, thresholds))
        r += 1

    # 权重设置区
    r += 1
    write_title(ws, r, 1, 7, "权重 / 综合分计算", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    write_row(ws, r, ["#", "子指标", "权重（%）", "单项分", "加权分", "", ""], fills=LIGHT_BLUE, fonts=BOLD)
    r += 1

    # 默认权重（应来自第 1 部分的细则）
    weight_lookup = {
        # 价值
        "V1": 40, "V2": 20, "V3": 20, "V4": 20,
        # 利润
        "P1": 60, "P2": 28, "P3": 12,
        # 付款
        "R1": 35, "R2": 30, "R3": 20, "R4": 15,
        # 粘性
        "L1": 30, "L2": 40, "L3": 30,
        # 潜力
        "G1": 30, "G2": 27, "G3": 27, "G4": 16,
    }

    weighted_rows = []
    for idx, (code, score_row, _) in enumerate(score_rows, 1):
        w = weight_lookup.get(code, 25)
        ws.cell(row=r, column=1, value=idx).border = BORDER
        ws.cell(row=r, column=1).font = NORMAL
        ws.cell(row=r, column=1).alignment = CENTER

        ws.cell(row=r, column=2, value=code).border = BORDER
        ws.cell(row=r, column=2).font = NORMAL
        ws.cell(row=r, column=2).alignment = CENTER

        wc = ws.cell(row=r, column=3, value=w)
        wc.fill = LIGHT_YELLOW
        wc.font = BOLD
        wc.border = BORDER
        wc.alignment = CENTER

        sc = ws.cell(row=r, column=4, value=f"=G{score_row}")
        sc.fill = LIGHT_GRAY
        sc.font = NORMAL
        sc.border = BORDER
        sc.alignment = CENTER

        wsc = ws.cell(row=r, column=5, value=f"=C{r}*D{r}/100")
        wsc.fill = LIGHT_GREEN
        wsc.font = BOLD
        wsc.border = BORDER
        wsc.alignment = CENTER

        weighted_rows.append(r)
        r += 1

    # 维度综合分
    write_row(ws, r, ["", "维度综合分", f"=SUM(C{weighted_rows[0]}:C{weighted_rows[-1]})", "—", f"=SUM(E{weighted_rows[0]}:E{weighted_rows[-1]})", "", ""],
              fills=[None, GOLD, GOLD, GOLD, GOLD, None, None],
              fonts=[NORMAL, WHITE_BOLD, WHITE_BOLD, WHITE_BOLD, WHITE_BOLD, NORMAL, NORMAL])
    dim_total_row = r
    r += 2

    write_title(ws, r, 1, 7, "维度公式说明", fill=LIGHT_BLUE, font=BOLD, height=22)
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    c = ws.cell(row=r, column=1, value=overall_formula_desc)
    c.font = NORMAL
    c.alignment = LEFT
    c.fill = LIGHT_GRAY
    c.border = BORDER
    ws.row_dimensions[r].height = 22

    # 返回维度综合分单元格引用
    return (f"'{tab_name}'!E{dim_total_row}", dim_total_row)


def build_score_formula(value_cell: str, thresholds: list[tuple[str, int, tuple]]) -> str:
    """构建 nested IF 公式做单项打分.

    thresholds: [(描述, 分数, (operator, value))]
        operator: ">=" / "<=" / ">" / "<" / "between"
    """
    # 从高到低嵌套 IF
    parts = []
    for desc, score, cond in thresholds:
        if len(cond) == 0:
            continue
        op = cond[0]
        if op == ">=":
            parts.append((f"{value_cell}>={cond[1]}", score))
        elif op == ">":
            parts.append((f"{value_cell}>{cond[1]}", score))
        elif op == "<=":
            parts.append((f"{value_cell}<={cond[1]}", score))
        elif op == "<":
            parts.append((f"{value_cell}<{cond[1]}", score))
        elif op == "between":
            parts.append((f"AND({value_cell}>={cond[1]},{value_cell}<{cond[2]})", score))

    if not parts:
        return "=0"
    # 嵌套 IF
    expr = str(parts[-1][1])  # 最后默认值
    for cond_str, score in reversed(parts[:-1]):
        expr = f"IF({cond_str},{score},{expr})"
    # 最外层
    cond_str, score = parts[0]
    expr = f"IF({cond_str},{score},{expr})"
    return "=" + expr


# -----------------------------------------------------------------------------
# 5 维度构建
# -----------------------------------------------------------------------------
def build_value(wb: Workbook):
    indicators = [
        ("V1", "年采购额", "订单表：SUM(订单金额) 近12月", "万元",
         [(">= ¥3,000 万", 100, (">=", 3000)),
          ("¥1,000-3,000 万", 85, ("between", 1000, 3000)),
          ("¥500-1,000 万", 70, ("between", 500, 1000)),
          ("¥100-500 万", 50, ("between", 100, 500)),
          ("¥30-100 万", 30, ("between", 30, 100)),
          ("< ¥30 万", 10, ("<", 30))],
         4200),
        ("V2", "月均订单数", "订单表：COUNT(订单)/12", "单/月",
         [(">= 8", 100, (">=", 8)),
          ("4-8", 80, ("between", 4, 8)),
          ("2-4", 60, ("between", 2, 4)),
          ("1-2", 40, ("between", 1, 2)),
          ("< 1", 20, ("<", 1))],
         6.5),
        ("V3", "单笔订单均值", "订单：SUM/COUNT", "万元",
         [(">= ¥100 万", 100, (">=", 100)),
          ("¥50-100 万", 85, ("between", 50, 100)),
          ("¥20-50 万", 70, ("between", 20, 50)),
          ("¥5-20 万", 50, ("between", 5, 20)),
          ("¥1-5 万", 30, ("between", 1, 5)),
          ("< ¥1 万", 10, ("<", 1))],
         54),
        ("V4", "SKU 多样度", "订单明细：COUNT(DISTINCT SKU)", "种",
         [(">= 30", 100, (">=", 30)),
          ("15-30", 80, ("between", 15, 30)),
          ("8-15", 60, ("between", 8, 15)),
          ("3-8", 40, ("between", 3, 8)),
          ("< 3", 20, ("<", 3))],
         18),
    ]
    return build_dimension(wb, "1 价值维度", "价值维度（V1-V4）", indicators,
        "维度综合分 = V1×40% + V2×20% + V3×20% + V4×20%（钢厂业态 V1 阈值 ×3 / 加工业态 ×0.5）",
        "25%")


def build_profit(wb: Workbook):
    indicators = [
        ("P1", "综合毛利率", "(销售-成本-运费-加工)/销售", "%",
         [(">= 6%", 100, (">=", 6)),
          ("4-6%", 85, ("between", 4, 6)),
          ("2-4%", 65, ("between", 2, 4)),
          ("1-2%", 40, ("between", 1, 2)),
          ("0-1%", 15, ("between", 0, 1)),
          ("< 0%", 0, ("<", 0))],
         4.5),
        ("P2", "年贡献毛利占公司比", "客户毛利/公司总毛利", "%",
         [(">= 5%", 100, (">=", 5)),
          ("2-5%", 80, ("between", 2, 5)),
          ("0.5-2%", 60, ("between", 0.5, 2)),
          ("0.1-0.5%", 40, ("between", 0.1, 0.5)),
          ("< 0.1%", 20, ("<", 0.1))],
         1.8),
        ("P3", "利润同比变化", "(今年-去年)/去年", "%",
         [("> +20%", 100, (">", 20)),
          ("+5% ~ +20%", 80, ("between", 5, 20)),
          ("-5% ~ +5%", 60, ("between", -5, 5)),
          ("-20% ~ -5%", 40, ("between", -20, -5)),
          ("< -20%", 20, ("<", -20))],
         18),
    ]
    return build_dimension(wb, "2 利润维度", "利润维度（P1-P3）", indicators,
        "维度综合分 = P1×60% + P2×28% + P3×12%（钢厂 P1 阈值 ×1.5 / 加工 ×0.7）★ P1<0% 触发强制不超 C 级",
        "25%")


def build_payment(wb: Workbook):
    indicators = [
        ("R1", "应收周转天数", "平均应收/销售额×365", "天",
         [("<= 30", 100, ("<=", 30)),
          ("30-60", 80, ("between", 30, 60)),
          ("60-90", 60, ("between", 60, 90)),
          ("90-120", 35, ("between", 90, 120)),
          ("> 120", 10, (">", 120))],
         42),
        ("R2", "90 天+ 应收占比", "90天+/总应收", "%",
         [("0%", 100, ("<=", 0)),
          ("0-10%", 85, ("between", 0, 10)),
          ("10-25%", 60, ("between", 10, 25)),
          ("25-50%", 30, ("between", 25, 50)),
          ("> 50%", 5, (">", 50))],
         8),
        ("R3", "付款准时率", "按时付款笔数/总笔数", "%",
         [(">= 95%", 100, (">=", 95)),
          ("85-95%", 80, ("between", 85, 95)),
          ("70-85%", 55, ("between", 70, 85)),
          ("50-70%", 30, ("between", 50, 70)),
          ("< 50%", 10, ("<", 50))],
         92),
        ("R4", "历史坏账次数", "坏账登记表", "次",
         [("0 次", 100, ("<=", 0)),
          ("1 次 < ¥10 万（折合 1-1.5）", 70, ("between", 0.5, 1.5)),
          ("1 次 ¥10-100 万（折合 1.5-2.5）", 40, ("between", 1.5, 2.5)),
          ("1 次 > ¥100 万（折合 2.5-3.5）", 15, ("between", 2.5, 3.5)),
          (">= 2 次（强制黑名单）", 0, (">=", 4))],
         0),
    ]
    return build_dimension(wb, "3 付款维度", "付款维度（R1-R4）", indicators,
        "维度综合分 = R1×35% + R2×30% + R3×20% + R4×15%（★ R4>=2 次 / R3<50% / R2>50% 6 月+ 均触发强制 D）",
        "20%")


def build_loyalty(wb: Workbook):
    indicators = [
        ("L1", "合作年限", "首次合作至今", "年",
         [(">= 5", 100, (">=", 5)),
          ("3-5", 85, ("between", 3, 5)),
          ("1-3", 65, ("between", 1, 3)),
          ("0.5-1", 40, ("between", 0.5, 1)),
          ("< 0.5（新客）", 20, ("<", 0.5))],
         7),
        ("L2", "近 12 月活跃月数", "下单月+主动询价月", "月",
         [(">= 10", 100, (">=", 10)),
          ("7-9", 80, ("between", 7, 10)),
          ("4-6", 60, ("between", 4, 7)),
          ("1-3", 30, ("between", 1, 4)),
          ("0（沉睡）", 0, ("<=", 0))],
         12),
        ("L3", "对接深度评分", "对接人数+EDI+多部门", "分(0-100)",
         [(">= 80（多部门+EDI）", 100, (">=", 80)),
          ("60-80（2+ 对接人含老板）", 80, ("between", 60, 80)),
          ("40-60（1 对接人 含财务）", 50, ("between", 40, 60)),
          ("20-40（1 对接人 业务）", 30, ("between", 20, 40)),
          ("< 20（绑销售）", 15, ("<", 20))],
         80),
    ]
    return build_dimension(wb, "4 粘性维度", "粘性维度（L1-L3）", indicators,
        "维度综合分 = L1×30% + L2×40% + L3×30%",
        "15%")


def build_potential(wb: Workbook):
    indicators = [
        ("G1", "客户行业景气", "行业研究 + 公开数据", "分",
         [(">= 85（出口/新能源/汽车）", 100, (">=", 85)),
          ("65-85（家电/集装箱/五金）", 80, ("between", 65, 85)),
          ("50-65（基建/机械）", 60, ("between", 50, 65)),
          ("30-50（钢结构/工业建筑）", 40, ("between", 30, 50)),
          ("< 30（房地产/普通建筑）", 20, ("<", 30))],
         85),
        ("G2", "客户自身增长", "客户营收同比", "%",
         [("> +30%", 100, (">", 30)),
          ("+10% ~ +30%", 80, ("between", 10, 30)),
          ("-10% ~ +10%", 60, ("between", -10, 10)),
          ("-30% ~ -10%", 35, ("between", -30, -10)),
          ("< -30%", 10, ("<", -30))],
         22),
        ("G3", "业态拓展空间", "客户已买/我能卖 比例", "%",
         [("< 30%（巨大空间）", 100, ("<", 30)),
          ("30-60%", 70, ("between", 30, 60)),
          ("60-80%", 50, ("between", 60, 80)),
          ("80-95%", 30, ("between", 80, 95)),
          (">= 95%（饱和）", 10, (">=", 95))],
         40),
        ("G4", "区域 / 政策红利", "政策研究", "分",
         [(">= 80（长三角/珠三角）", 100, (">=", 80)),
          ("60-80（京津冀雄安）", 70, ("between", 60, 80)),
          ("40-60（华中出口转型）", 55, ("between", 40, 60)),
          ("20-40（西南基建）", 40, ("between", 20, 40)),
          ("< 20", 25, ("<", 20))],
         90),
    ]
    return build_dimension(wb, "5 潜力维度", "潜力维度（G1-G4）", indicators,
        "维度综合分 = G1×30% + G2×27% + G3×27% + G4×16%",
        "15%")


# -----------------------------------------------------------------------------
# 6 综合评分
# -----------------------------------------------------------------------------
def build_total(wb: Workbook, dim_refs: dict[str, str], force_rule_refs: dict[str, str]) -> None:
    ws = wb.create_sheet("6 综合评分")
    set_col_widths(ws, [4, 22, 14, 14, 14, 18, 22])

    write_title(ws, 1, 1, 7, "综合评分 + 等级判定 + 强制规则")
    write_title(ws, 2, 1, 7, "自动汇总 5 维度 + 加权计算 + 强制规则触发", fill=LIGHT_BLUE, font=BOLD, height=22)

    r = 4
    write_title(ws, r, 1, 7, "5 维度综合分", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    write_row(ws, r, ["#", "维度", "权重(%)", "维度分", "加权分", "等级判定（独立）", "备注"],
              fills=LIGHT_BLUE, fonts=BOLD)
    r += 1

    dims = [
        ("①", "价值维度", 25, dim_refs["value"], "—"),
        ("②", "利润维度", 25, dim_refs["profit"], "P1<0% 强制不超 C"),
        ("③", "付款维度", 20, dim_refs["payment"], "R4>=2 强制 D / R3<50% 强制 D"),
        ("④", "粘性维度", 15, dim_refs["loyalty"], "—"),
        ("⑤", "潜力维度", 15, dim_refs["potential"], "—"),
    ]
    dim_rows = []
    for idx, name, w, ref, note in dims:
        ws.cell(row=r, column=1, value=idx).border = BORDER
        ws.cell(row=r, column=1).font = NORMAL
        ws.cell(row=r, column=1).alignment = CENTER

        ws.cell(row=r, column=2, value=name).border = BORDER
        ws.cell(row=r, column=2).font = BOLD
        ws.cell(row=r, column=2).alignment = LEFT

        ws.cell(row=r, column=3, value=w).border = BORDER
        ws.cell(row=r, column=3).font = NORMAL
        ws.cell(row=r, column=3).alignment = CENTER

        sc = ws.cell(row=r, column=4, value=f"={ref}")
        sc.fill = LIGHT_GRAY
        sc.font = BOLD
        sc.border = BORDER
        sc.alignment = CENTER

        wsc = ws.cell(row=r, column=5, value=f"=C{r}*D{r}/100")
        wsc.fill = LIGHT_GREEN
        wsc.font = BOLD
        wsc.border = BORDER
        wsc.alignment = CENTER

        lc = ws.cell(row=r, column=6, value=f'=IF(D{r}>=85,"S",IF(D{r}>=70,"A",IF(D{r}>=55,"B",IF(D{r}>=40,"C","D"))))')
        lc.border = BORDER
        lc.font = NORMAL
        lc.alignment = CENTER

        ws.cell(row=r, column=7, value=note).font = SMALL
        ws.cell(row=r, column=7).border = BORDER
        ws.cell(row=r, column=7).alignment = LEFT

        dim_rows.append(r)
        r += 1

    # 综合总分
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    c = ws.cell(row=r, column=1, value="综合分（基础）= 5 维度加权和")
    c.font = WHITE_BOLD
    c.fill = NAVY
    c.alignment = CENTER
    c.border = BORDER
    total_row = r
    tc = ws.cell(row=r, column=5, value=f"=ROUND(SUM(E{dim_rows[0]}:E{dim_rows[-1]}),1)")
    tc.fill = NAVY
    tc.font = WHITE_BOLD
    tc.alignment = CENTER
    tc.border = BORDER
    ws.cell(row=r, column=6, value=f'=IF(E{r}>=85,"S 钻石",IF(E{r}>=70,"A 黄金",IF(E{r}>=55,"B 白银",IF(E{r}>=40,"C 青铜","D 警示"))))').font = BOLD
    ws.cell(row=r, column=6).border = BORDER
    ws.cell(row=r, column=6).fill = GOLD
    ws.cell(row=r, column=6).alignment = CENTER
    ws.cell(row=r, column=7, value="基础等级").font = NORMAL
    ws.cell(row=r, column=7).border = BORDER
    ws.cell(row=r, column=7).alignment = CENTER
    r += 2

    # 强制规则
    write_title(ws, r, 1, 7, "强制规则（自动检测）", fill=RED, font=WHITE_BOLD)
    r += 1
    write_row(ws, r, ["#", "规则", "触发条件", "—", "—", "是否触发", "效果"],
              fills=LIGHT_BLUE, fonts=BOLD)
    r += 1

    # 引用付款维度的 R4 / R3 / R2 实测值（这里假设 R4/R3/R2 在 '3 付款维度' E 列特定行）
    # 简化：基于实测值 5 个子指标在 E6/E7/E8/E9 的位置（V1=E6, V2=E7, ...; 各 Tab 略不同）
    # 由于 build_dimension 中数据从 row=6 起，每个子指标 1 行（headers row=5）
    # 付款维度：R1=E6, R2=E7, R3=E8, R4=E9
    # 利润维度：P1=E6
    rules = [
        ("①", "R4 >= 2 次坏账", "付款 R4 实测 >= 4", f"='3 付款维度'!E9>=4", "强制 D + 黑名单"),
        ("②", "R3 准时率 < 50%", "付款 R3 实测 < 50", f"='3 付款维度'!E8<50", "强制 D"),
        ("③", "R2 90 天+ > 50%", "付款 R2 实测 > 50", f"='3 付款维度'!E7>50", "强制 D + 冻结订单"),
        ("④", "P1 毛利率 < 0%", "利润 P1 实测 < 0", f"='2 利润维度'!E6<0", "强制不超 C 级"),
    ]
    rule_rows = []
    for idx, name, cond_desc, formula, effect in rules:
        ws.cell(row=r, column=1, value=idx).border = BORDER
        ws.cell(row=r, column=1).font = NORMAL
        ws.cell(row=r, column=1).alignment = CENTER

        ws.cell(row=r, column=2, value=name).border = BORDER
        ws.cell(row=r, column=2).font = BOLD
        ws.cell(row=r, column=2).alignment = LEFT

        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.cell(row=r, column=3, value=cond_desc).border = BORDER
        ws.cell(row=r, column=3).font = SMALL
        ws.cell(row=r, column=3).alignment = LEFT

        c = ws.cell(row=r, column=6, value=f'=IF({formula},"🔴 触发","✅ 未触发")')
        c.font = BOLD
        c.border = BORDER
        c.alignment = CENTER

        ws.cell(row=r, column=7, value=effect).font = NORMAL
        ws.cell(row=r, column=7).border = BORDER
        ws.cell(row=r, column=7).alignment = LEFT
        rule_rows.append((r, formula))
        r += 1

    # 最终等级（基础等级 + 强制规则覆盖）
    r += 1
    write_title(ws, r, 1, 7, "最终等级（基础 + 强制规则覆盖）", fill=GREEN, font=WHITE_BOLD)
    r += 1
    base_grade_cell = f"F{total_row}"
    # 复合规则
    final_formula = (
        f'=IF({rule_rows[0][1]},"D 钻石 → 强制 D + 黑名单",'
        f'IF({rule_rows[1][1]},"D → 强制 D",'
        f'IF({rule_rows[2][1]},"D → 强制 D + 冻结订单",'
        f'IF(AND({rule_rows[3][1]},OR({base_grade_cell}="S 钻石",{base_grade_cell}="A 黄金",{base_grade_cell}="B 白银")),"C 青铜 → P1<0 不超 C",'
        f'{base_grade_cell}))))'
    )
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    c = ws.cell(row=r, column=1, value="客户最终等级")
    c.font = WHITE_BOLD
    c.fill = GREEN
    c.alignment = CENTER
    c.border = BORDER
    fc = ws.cell(row=r, column=5, value=final_formula)
    fc.font = WHITE_BOLD
    fc.fill = GREEN
    fc.alignment = CENTER
    fc.border = BORDER
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
    final_grade_row = r
    ws.row_dimensions[r].height = 26
    r += 2

    # 特殊标签
    write_title(ws, r, 1, 7, "特殊标签（独立于等级）", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    tags = [
        ("🔴 风险预警", f'=IF(OR(\'3 付款维度\'!E6>90,\'3 付款维度\'!E7>30),"是","否")'),
        ("🟡 沉睡客户", f'=IF(\'4 粘性维度\'!E7<=0,"是","否")'),
        ("🟢 战略客户", "（老板手动标记）"),
        ("⚫ 黑名单", f'=IF(\'3 付款维度\'!E9>=4,"是","否")'),
    ]
    for name, formula in tags:
        ws.cell(row=r, column=2, value=name).border = BORDER
        ws.cell(row=r, column=2).font = BOLD
        ws.cell(row=r, column=2).alignment = LEFT

        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        c = ws.cell(row=r, column=3, value=formula)
        c.font = NORMAL
        c.border = BORDER
        c.alignment = CENTER
        c.fill = LIGHT_YELLOW
        r += 1

    # 把 final_grade_row 暴露给 SOP Tab
    return final_grade_row


# -----------------------------------------------------------------------------
# 7 5 级 SOP
# -----------------------------------------------------------------------------
def build_sop(wb: Workbook, final_grade_ref: str) -> None:
    ws = wb.create_sheet("7 五级 SOP")
    set_col_widths(ws, [4, 18, 22, 22, 22, 18])

    write_title(ws, 1, 1, 6, "客户等级 → 销售 SOP（自动匹配）")
    write_title(ws, 2, 1, 6, f"当前客户等级：从'6 综合评分'读取（{final_grade_ref}）", fill=LIGHT_BLUE, font=BOLD, height=22)

    r = 4
    # 当前等级显示
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws.cell(row=r, column=2, value="当前客户等级：")
    c.font = WHITE_BOLD
    c.fill = NAVY
    c.alignment = CENTER
    c.border = BORDER
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    cc = ws.cell(row=r, column=4, value=f"={final_grade_ref}")
    cc.font = WHITE_BOLD
    cc.fill = GREEN
    cc.alignment = CENTER
    cc.border = BORDER
    ws.row_dimensions[r].height = 28
    r += 2

    # SOP 总表
    write_title(ws, r, 1, 6, "5 级 SOP 总表（销售必读）", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    write_row(ws, r, ["等级", "主跟人", "维护频率", "资源配额(年/客户)", "总数控制", "续单率KPI"],
              fills=LIGHT_BLUE, fonts=BOLD)
    r += 1

    sops = [
        ("S 钻石 ≥85", "销售老板/总监", "月度上门+周电话+日微信", "¥5-10 万", "≤ 5%（75 个）", "≥ 95%"),
        ("A 黄金 70-84", "销售总监/资深", "双周上门+周电话", "¥1-3 万", "15-20%（250 个）", "≥ 90%"),
        ("B 白银 55-69", "销售员", "月度电话+季上门", "¥3,000-1 万", "40-50%（700 个）", "≥ 75%"),
        ("C 青铜 40-54", "自动化为主", "季度 SMS+半年电话", "≤ ¥1,000", "25-30%（400 个）", "升 B≥15%/年"),
        ("D 警示 <40", "法务/风控/应收", "催收/黑名单评估", "—", "≤ 10%（150 个）", "坏账≤1%"),
    ]
    fills_map = {"S": GOLD, "A": LIGHT_GREEN, "B": LIGHT_BLUE, "C": LIGHT_YELLOW, "D": PatternFill("solid", fgColor="FAE5E5")}
    for sop_row in sops:
        first_char = sop_row[0][0]
        write_row(ws, r, list(sop_row), fills=fills_map.get(first_char, None))
        r += 1

    r += 1
    # 当前等级对应的详细 SOP
    write_title(ws, r, 1, 6, "★ 当前客户的详细 SOP（自动匹配）", fill=GREEN, font=WHITE_BOLD)
    r += 1

    # 用 IF 公式根据等级返回不同 SOP
    sop_text_map = {
        "S": "S 钻石：① 老板亲自月度电话\n② 销售总监双周上门\n③ 月度对账+应收清账+下月预订单\n④ 季度邀请进老板局\n⑤ 半年二代俱乐部\n⑥ 年度战略对话+送年货",
        "A": "A 黄金：① 销售总监双周上门\n② 月度对账+应收\n③ 季度业务回顾\n④ 老板局年度 1-2 次\n⑤ 升 S 提名（连续 2 季 ≥85 分）",
        "B": "B 白银：① 销售员月度电话\n② 季度上门 1 次\n③ 半年健康度自评邀请\n④ 升级路径明确化\n⑤ 跨季评估升 A",
        "C": "C 青铜：① 自动化群发月度行情\n② 季度群发新品/优惠\n③ 半年电话调研\n④ 升 B / 降 D 评估\n⑤ 自动化覆盖 ≥ 90%",
        "D": "D 警示：① 法务催收（30/60/90 天三级）\n② 持续不付款=法律诉讼\n③ 进黑名单\n④ 客户主动表态触发反向挽回 SOP\n⑤ 6 月观察期重新评估",
    }
    # 简化：用 LEFT 提取等级
    detail_formula = (
        f'=IF(LEFT(D{r-3 - len(sops)},1)="S","{sop_text_map["S"]}",'
        f'IF(LEFT(D{r-3 - len(sops)},1)="A","{sop_text_map["A"]}",'
        f'IF(LEFT(D{r-3 - len(sops)},1)="B","{sop_text_map["B"]}",'
        f'IF(LEFT(D{r-3 - len(sops)},1)="C","{sop_text_map["C"]}",'
        f'"{sop_text_map["D"]}"))))'
    )
    # 偏移取顶部"当前等级"显示的单元格（D4 是上面合并的客户等级值）
    detail_formula = (
        f'=IF(ISNUMBER(SEARCH("S",D4)),"{sop_text_map["S"]}",'
        f'IF(ISNUMBER(SEARCH("A",D4)),"{sop_text_map["A"]}",'
        f'IF(ISNUMBER(SEARCH("B",D4)),"{sop_text_map["B"]}",'
        f'IF(ISNUMBER(SEARCH("C",D4)),"{sop_text_map["C"]}",'
        f'"{sop_text_map["D"]}"))))'
    )
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 6, end_column=6)
    c = ws.cell(row=r, column=1, value=detail_formula)
    c.font = BOLD
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    c.fill = LIGHT_GREEN
    c.border = BORDER
    ws.row_dimensions[r].height = 30
    for off in range(1, 7):
        ws.row_dimensions[r + off].height = 26
    r += 8

    write_title(ws, r, 1, 6, "提成挂钩规则", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    write_row(ws, r, ["等级", "提成系数", "—", "—", "—", "—"], fills=LIGHT_BLUE, fonts=BOLD)
    r += 1
    coefs = [("S", "×1.5"), ("A", "×1.2"), ("B", "×1.0"), ("C", "×0.6"), ("D", "×0 + 倒罚")]
    for k, v in coefs:
        write_row(ws, r, [k, v, "", "", "", ""])
        r += 1


# -----------------------------------------------------------------------------
# 8 批量录入（50 客户批量评级）
# -----------------------------------------------------------------------------
def build_batch(wb: Workbook) -> None:
    ws = wb.create_sheet("8 批量录入")
    set_col_widths(ws, [4, 16, 10, 8, 8, 9, 8, 7, 8, 8, 8, 8, 8, 8, 9, 10, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 10])

    write_title(ws, 1, 1, 28, "50 客户批量录入与评级（每行 1 个客户）")
    write_title(ws, 2, 1, 28, "黄色列录入实测值 / 绿色列自动计算 / 末列自动等级", fill=LIGHT_BLUE, font=BOLD, height=22)

    headers = [
        "#", "客户名", "业态", "V1", "V2", "V3", "V4", "P1", "P2", "P3",
        "R1", "R2", "R3", "R4", "L1", "L2", "L3",
        "G1", "G2", "G3", "G4",
        "价值分", "利润分", "付款分", "粘性分", "潜力分",
        "综合分", "等级",
    ]
    write_row(ws, 4, headers, fills=NAVY, fonts=WHITE_BOLD)

    # 示例数据 + 公式
    sample_data = [
        ("山东 XX 钢管", "钢贸", 4200, 6.5, 54, 18, 4.5, 1.8, 18, 42, 8, 92, 0, 7, 12, 80, 85, 22, 40, 90),
        ("江苏 XX 焊管", "钢贸", 1200, 4.0, 30, 12, 3.8, 0.6, 8, 55, 18, 78, 0, 5, 10, 60, 75, 18, 50, 80),
        ("浙江 XX 冷卷", "镀锌", 6000, 8.0, 75, 20, 4.2, 2.0, 20, 38, 5, 95, 0, 8, 12, 90, 92, 25, 35, 85),
        ("广东 XX 钢贸", "钢贸", 3000, 7.0, 43, 15, 3.2, 1.5, 10, 48, 12, 85, 0, 6, 11, 70, 70, 20, 45, 75),
        ("邯郸 XX 钢管", "钢贸", 3000, 5.5, 55, 18, 4.0, 1.2, -5, 65, 25, 75, 1, 6, 12, 75, 60, -10, 30, 60),
        ("天津 XX 钢贸", "钢贸", 800, 3.5, 23, 10, 2.0, 0.3, -15, 95, 38, 60, 2, 4, 8, 50, 45, -20, 55, 60),
    ]
    # 各维度权重
    val_w = [40, 20, 20, 20]   # V1-V4
    pro_w = [60, 28, 12]       # P1-P3
    pay_w = [35, 30, 20, 15]   # R1-R4
    loy_w = [30, 40, 30]       # L1-L3
    pot_w = [30, 27, 27, 16]   # G1-G4

    def v_score(v):
        # V1
        return None
    # 简化：批量录入用 SUMPRODUCT 模式，给出公式
    for idx, row_data in enumerate(sample_data, 1):
        r = 4 + idx
        # 客户名+业态
        ws.cell(row=r, column=1, value=idx).border = BORDER
        ws.cell(row=r, column=1).font = NORMAL
        ws.cell(row=r, column=1).alignment = CENTER

        ws.cell(row=r, column=2, value=row_data[0]).border = BORDER
        ws.cell(row=r, column=2).font = NORMAL
        ws.cell(row=r, column=2).alignment = LEFT

        ws.cell(row=r, column=3, value=row_data[1]).border = BORDER
        ws.cell(row=r, column=3).font = NORMAL
        ws.cell(row=r, column=3).alignment = CENTER

        # V1-V4 / P1-P3 / R1-R4 / L1-L3 / G1-G4 共 18 个数据
        for col_idx, val in enumerate(row_data[2:], 4):
            c = ws.cell(row=r, column=col_idx, value=val)
            c.fill = LIGHT_YELLOW
            c.font = NORMAL
            c.border = BORDER
            c.alignment = CENTER

        # 价值分（V1-V4 转换为分数后加权）
        # 简化：用阶梯函数估算
        def stepfn_v1(cell):
            return f'(IF({cell}>=3000,100,IF({cell}>=1000,85,IF({cell}>=500,70,IF({cell}>=100,50,IF({cell}>=30,30,10))))))'
        def stepfn_v2(cell):
            return f'(IF({cell}>=8,100,IF({cell}>=4,80,IF({cell}>=2,60,IF({cell}>=1,40,20)))))'
        def stepfn_v3(cell):
            return f'(IF({cell}>=100,100,IF({cell}>=50,85,IF({cell}>=20,70,IF({cell}>=5,50,IF({cell}>=1,30,10))))))'
        def stepfn_v4(cell):
            return f'(IF({cell}>=30,100,IF({cell}>=15,80,IF({cell}>=8,60,IF({cell}>=3,40,20)))))'
        def stepfn_p1(cell):
            return f'(IF({cell}>=6,100,IF({cell}>=4,85,IF({cell}>=2,65,IF({cell}>=1,40,IF({cell}>=0,15,0))))))'
        def stepfn_p2(cell):
            return f'(IF({cell}>=5,100,IF({cell}>=2,80,IF({cell}>=0.5,60,IF({cell}>=0.1,40,20)))))'
        def stepfn_p3(cell):
            return f'(IF({cell}>20,100,IF({cell}>=5,80,IF({cell}>=-5,60,IF({cell}>=-20,40,20)))))'
        def stepfn_r1(cell):
            return f'(IF({cell}<=30,100,IF({cell}<=60,80,IF({cell}<=90,60,IF({cell}<=120,35,10)))))'
        def stepfn_r2(cell):
            return f'(IF({cell}<=0,100,IF({cell}<=10,85,IF({cell}<=25,60,IF({cell}<=50,30,5)))))'
        def stepfn_r3(cell):
            return f'(IF({cell}>=95,100,IF({cell}>=85,80,IF({cell}>=70,55,IF({cell}>=50,30,10)))))'
        def stepfn_r4(cell):
            return f'(IF({cell}<=0,100,IF({cell}<=1,70,IF({cell}<=2,40,IF({cell}<=3,15,0)))))'
        def stepfn_l1(cell):
            return f'(IF({cell}>=5,100,IF({cell}>=3,85,IF({cell}>=1,65,IF({cell}>=0.5,40,20)))))'
        def stepfn_l2(cell):
            return f'(IF({cell}>=10,100,IF({cell}>=7,80,IF({cell}>=4,60,IF({cell}>=1,30,0)))))'
        def stepfn_l3(cell):
            return f'(IF({cell}>=80,100,IF({cell}>=60,80,IF({cell}>=40,50,IF({cell}>=20,30,15)))))'
        def stepfn_g1(cell):
            return f'(IF({cell}>=85,100,IF({cell}>=65,80,IF({cell}>=50,60,IF({cell}>=30,40,20)))))'
        def stepfn_g2(cell):
            return f'(IF({cell}>30,100,IF({cell}>=10,80,IF({cell}>=-10,60,IF({cell}>=-30,35,10)))))'
        def stepfn_g3(cell):
            return f'(IF({cell}<30,100,IF({cell}<60,70,IF({cell}<80,50,IF({cell}<95,30,10)))))'
        def stepfn_g4(cell):
            return f'(IF({cell}>=80,100,IF({cell}>=60,70,IF({cell}>=40,55,IF({cell}>=20,40,25)))))'

        # 价值分 (V1=D, V2=E, V3=F, V4=G)
        val_formula = f"=ROUND({stepfn_v1(f'D{r}')}*0.4+{stepfn_v2(f'E{r}')}*0.2+{stepfn_v3(f'F{r}')}*0.2+{stepfn_v4(f'G{r}')}*0.2,1)"
        c = ws.cell(row=r, column=22, value=val_formula)
        c.fill = LIGHT_GREEN
        c.font = BOLD
        c.border = BORDER
        c.alignment = CENTER

        # 利润分 (P1=H, P2=I, P3=J)
        pro_formula = f"=ROUND({stepfn_p1(f'H{r}')}*0.6+{stepfn_p2(f'I{r}')}*0.28+{stepfn_p3(f'J{r}')}*0.12,1)"
        c = ws.cell(row=r, column=23, value=pro_formula)
        c.fill = LIGHT_GREEN
        c.font = BOLD
        c.border = BORDER
        c.alignment = CENTER

        # 付款分 (R1=K, R2=L, R3=M, R4=N)
        pay_formula = f"=ROUND({stepfn_r1(f'K{r}')}*0.35+{stepfn_r2(f'L{r}')}*0.3+{stepfn_r3(f'M{r}')}*0.2+{stepfn_r4(f'N{r}')}*0.15,1)"
        c = ws.cell(row=r, column=24, value=pay_formula)
        c.fill = LIGHT_GREEN
        c.font = BOLD
        c.border = BORDER
        c.alignment = CENTER

        # 粘性分 (L1=O, L2=P, L3=Q)
        loy_formula = f"=ROUND({stepfn_l1(f'O{r}')}*0.3+{stepfn_l2(f'P{r}')}*0.4+{stepfn_l3(f'Q{r}')}*0.3,1)"
        c = ws.cell(row=r, column=25, value=loy_formula)
        c.fill = LIGHT_GREEN
        c.font = BOLD
        c.border = BORDER
        c.alignment = CENTER

        # 潜力分 (G1=R, G2=S, G3=T, G4=U)
        pot_formula = f"=ROUND({stepfn_g1(f'R{r}')}*0.3+{stepfn_g2(f'S{r}')}*0.27+{stepfn_g3(f'T{r}')}*0.27+{stepfn_g4(f'U{r}')}*0.16,1)"
        c = ws.cell(row=r, column=26, value=pot_formula)
        c.fill = LIGHT_GREEN
        c.font = BOLD
        c.border = BORDER
        c.alignment = CENTER

        # 综合分（V=22, W=23, X=24, Y=25, Z=26 → 列字母分别 V,W,X,Y,Z）
        total_formula = f"=ROUND(V{r}*0.25+W{r}*0.25+X{r}*0.2+Y{r}*0.15+Z{r}*0.15,1)"
        c = ws.cell(row=r, column=27, value=total_formula)
        c.fill = GOLD
        c.font = WHITE_BOLD
        c.border = BORDER
        c.alignment = CENTER

        # 等级（基础 + 强制规则）
        # 强制 D 触发条件：R4>=2（N列）/ R3<50（M列）/ R2>50（L列）
        grade_formula = (
            f'=IF(N{r}>=2,"D 黑名单",'
            f'IF(M{r}<50,"D 强制",'
            f'IF(L{r}>50,"D 冻结",'
            f'IF(H{r}<0,IF(AA{r}>=55,"C P1<0",IF(AA{r}>=40,"C",IF(AA{r}>=30,"C","D"))),'
            f'IF(AA{r}>=85,"S 钻石",IF(AA{r}>=70,"A 黄金",IF(AA{r}>=55,"B 白银",IF(AA{r}>=40,"C 青铜","D 警示"))))))))'
        )
        c = ws.cell(row=r, column=28, value=grade_formula)
        c.fill = GREEN
        c.font = WHITE_BOLD
        c.border = BORDER
        c.alignment = CENTER

    # 给空白行 50 - 6 = 44 行也加公式（让用户填实测值即可）
    for idx in range(len(sample_data) + 1, 51):
        r = 4 + idx
        ws.cell(row=r, column=1, value=idx).border = BORDER
        ws.cell(row=r, column=1).alignment = CENTER
        ws.cell(row=r, column=1).font = NORMAL
        for col_idx in range(2, 22):
            c = ws.cell(row=r, column=col_idx, value=None)
            c.border = BORDER
            if 4 <= col_idx <= 21:
                c.fill = LIGHT_YELLOW

    # 表格说明
    r = 56
    write_title(ws, r, 1, 28, "使用说明", fill=LIGHT_BLUE, font=BOLD, height=22)
    r += 1
    tips = [
        "1. 黄色单元格 = 录入区（V1-V4 / P1-P3 / R1-R4 / L1-L3 / G1-G4 共 18 个子指标）",
        "2. 绿色单元格 = 自动计算（5 维度分 + 综合分）",
        "3. 金色单元格 = 综合分",
        "4. 深绿色单元格 = 最终等级（含强制规则）",
        "5. 强制规则：R4≥2=D黑名单 / R3<50=D强制 / R2>50=D冻结 / P1<0=不超 C",
        "6. 可拖拽公式至 50 行批量评级",
    ]
    for t in tips:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=28)
        c = ws.cell(row=r, column=1, value=t)
        c.font = NORMAL
        c.alignment = LEFT
        c.fill = LIGHT_GRAY
        c.border = BORDER
        r += 1


def build_legend(wb: Workbook) -> None:
    ws = wb.create_sheet("9 维度说明")
    set_col_widths(ws, [4, 16, 18, 16, 18, 28])

    write_title(ws, 1, 1, 6, "5 维度 × 18 子指标速查")
    write_title(ws, 2, 1, 6, "完整规则见 docs/69-CRM分级落地细则与评估细节.md", fill=LIGHT_BLUE, font=BOLD, height=22)

    r = 4
    write_row(ws, r, ["#", "维度", "权重", "子指标", "权重", "说明"], fills=NAVY, fonts=WHITE_BOLD)
    r += 1

    items = [
        ("①", "价值", "25%", "V1 年采购额", "40%", "近 12 月订单总额（钢厂×3 / 加工×0.5）"),
        ("", "", "", "V2 月均订单数", "20%", "近 12 月订单数 / 12"),
        ("", "", "", "V3 单笔均值", "20%", "订单总额 / 订单数"),
        ("", "", "", "V4 SKU 多样度", "20%", "近 12 月不重复 SKU 数"),
        ("②", "利润", "25%", "P1 综合毛利率", "60%", "★ <0% 强制不超 C"),
        ("", "", "", "P2 年贡献毛利占公司", "28%", "客户毛利 / 公司总毛利"),
        ("", "", "", "P3 利润同比", "12%", "今年毛利 vs 去年"),
        ("③", "付款", "20%", "R1 应收周转天数", "35%", "平均应收/销售×365"),
        ("", "", "", "R2 90 天+ 应收占比", "30%", "★ >50% 强制 D + 冻结"),
        ("", "", "", "R3 付款准时率", "20%", "★ <50% 强制 D"),
        ("", "", "", "R4 历史坏账次数", "15%", "★ ≥2 次强制 D + 黑名单"),
        ("④", "粘性", "15%", "L1 合作年限", "30%", "首次合作至今"),
        ("", "", "", "L2 近 12 月活跃度", "40%", "下单/询价月数"),
        ("", "", "", "L3 对接深度", "30%", "对接人数 + EDI + 多部门"),
        ("⑤", "潜力", "15%", "G1 客户行业景气", "30%", "汽车/家电/出口 高"),
        ("", "", "", "G2 客户自身增长", "27%", "客户营收同比"),
        ("", "", "", "G3 业态拓展空间", "27%", "客户已买/我能卖 比例"),
        ("", "", "", "G4 区域/政策红利", "16%", "长三角/雄安/出口"),
    ]
    for row_data in items:
        write_row(ws, r, list(row_data))
        r += 1

    r += 2
    write_title(ws, r, 1, 6, "5 级评分边界 + 强制规则", fill=ORANGE, font=WHITE_BOLD)
    r += 1
    write_row(ws, r, ["等级", "分数区间", "—", "占比目标", "—", "强制触发"], fills=LIGHT_BLUE, fonts=BOLD)
    r += 1
    grades = [
        ("S 钻石", "≥ 85", "", "≤ 5%", "", "—"),
        ("A 黄金", "70-84", "", "15-20%", "", "—"),
        ("B 白银", "55-69", "", "40-50%", "", "—"),
        ("C 青铜", "40-54", "", "25-30%", "", "P1<0% 不超 C"),
        ("D 警示", "< 40", "", "≤ 10%", "", "R4≥2 / R3<50% / R2>50%"),
    ]
    for g in grades:
        write_row(ws, r, list(g))
        r += 1


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)

    build_profile(wb)
    val_ref, _ = build_value(wb)
    pro_ref, _ = build_profit(wb)
    pay_ref, _ = build_payment(wb)
    loy_ref, _ = build_loyalty(wb)
    pot_ref, _ = build_potential(wb)

    dim_refs = {
        "value": val_ref,
        "profit": pro_ref,
        "payment": pay_ref,
        "loyalty": loy_ref,
        "potential": pot_ref,
    }
    final_grade_row = build_total(wb, dim_refs, {})

    build_sop(wb, f"'6 综合评分'!E{final_grade_row}")
    build_batch(wb)
    build_legend(wb)

    out = "tools/CRM分级评分计算器.xlsx"
    wb.save(out)
    print(f"✓ 已生成：{out}")


if __name__ == "__main__":
    main()
