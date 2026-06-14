"""
灯塔厂候选 30 家评分表
======================

输出：tools/灯塔厂候选30家_评分表.xlsx

工作表：
    1. 使用说明
    2. 评分维度参考       —— 10 个维度的详细评分参考
    3. 30 家候选评分      —— 主表，30 行 × 26 列，自动算总分 + 等级
    4. 等级分布           —— 自动统计 S/A/B/C 各几家
    5. 推进进度看板        —— 按推进阶段统计
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule


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
TEXT_FONT = Font(name="Microsoft YaHei", size=10)
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
    widths(ws, [4, 28, 70])
    ws.merge_cells("B2:C2")
    ws["B2"] = "灯塔厂候选 30 家评分表（v1.0）"
    ws["B2"].font = TITLE_FONT
    ws["B2"].alignment = LEFT

    rows = [
        ("适用对象", "老板 + 2 销售总监 + 销售运营经理"),
        ("使用步骤",
         "① 打开『30 家候选评分』sheet，黄色单元格填写客户基本信息\n"
         "② 黄色 10 列评分（每维度 1-5 分）由老板、A 总监、B 总监 3 人独立打分再平均\n"
         "③ 绿色单元格自动计算总分 + 等级（S/A/B/C）\n"
         "④ 切换『等级分布』『推进进度看板』看自动统计"),
        ("评分铁律",
         "1. 每家候选必须满足『3 不要』后才能进表\n"
         "2. 评分由 3 人独立打分平均，避免单人主观\n"
         "3. S 级（≥80）老板亲自跑；A 级（65-79）销售总监主跟\n"
         "4. 30 家里争取 ≥6 家 S/A 级"),
        ("时间节点",
         "W1 周一-周四：列初版 30 家草稿\n"
         "W1 周五：评分定稿\n"
         "W2-W3：S 级老板亲自跑、A 级销售总监跑\n"
         "W3-W4：第 1 家灯塔厂签约 POC"),
        ("注意事项",
         "1. 黄色 = 输入；绿色 = 自动；蓝色 = 表头\n"
         "2. 修改评分立即看到等级变化\n"
         "3. 每周一战报会更新『当前进度』+『下次跟进日期』\n"
         "4. 每月一次重排序"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        ws[f"B{i}"] = k; ws[f"C{i}"] = v
        c1 = ws[f"B{i}"]; c1.font = SUB_FONT; c1.fill = H_FILL
        c1.font = Font(name="Microsoft YaHei", size=11, bold=True, color=WHITE)
        c1.alignment = CENTER; c1.border = BORDER
        c2 = ws[f"C{i}"]; c2.font = TEXT_FONT; c2.alignment = LEFT; c2.border = BORDER
        ws.row_dimensions[i].height = max(40, len(v) * 0.7)


def build_dimensions(ws):
    ws.title = "评分维度参考"
    widths(ws, [4, 22, 8, 14, 24, 22, 22, 22, 22, 22])
    ws.merge_cells("B2:J2")
    ws["B2"] = "10 个评分维度参考（评分时对照本页）"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["维度", "权重", "5 分", "4 分", "3 分", "2 分", "1 分"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    dims = [
        ("① 老板年龄段", 2, "40-55 岁", "35-39", "56-60", "≥61", "≥65（不选）"),
        ("② 老板出镜意愿", 2, "已有抖音/视频号", "愿出镜需包装", "勉强同意", "极不情愿", "明确拒绝"),
        ("③ 厂区可看度", 1, "示范工厂", "整洁先进", "标准车间", "偏脏乱", "完全不能拍"),
        ("④ 产业带地位", 2, "前 3 协会副会长级", "前 10", "中等知名", "本地圈知道", "默默无闻"),
        ("⑤ 年产值规模", 2, "5-30 亿", "3-5 或 30-50 亿", "1-3 或 50-100 亿", "<1 或 >100 亿", "极端"),
        ("⑥ 客户结构", 2, "出口/汽车/家电≥30%", "10-30%", "本地工程为主", "小客户为主", "不清晰"),
        ("⑦ 数字化基础", 1, "金蝶+仓库 Excel", "仅金蝶", "手工+Excel", "纯手工", "失败 ERP 包袱"),
        ("⑧ 决策链", 3, "老板一人拍板", "老板+1 副总", "老板+多副总", "经理人/二代分歧", "股东多方博弈"),
        ("⑨ 与我方关系", 3, "老板已熟", "强引荐路径", "弱引荐", "知道公司无人脉", "纯陌生"),
        ("⑩ 见效速度", 2, "已主动咨询有预算", "今年必做(审厂)", "1-2 年计划", "在观望", "无意识"),
    ]
    for i, row in enumerate(dims, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        for j, v in enumerate(row, 2):
            cell = ws.cell(row=i, column=j, value=v)
            if j == 2:
                style_t(cell, fill=bg, align=LEFT, bold=True)
            elif j == 3:
                style_t(cell, fill=ACCENT, bold=True)
                cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color=WHITE)
            else:
                style_t(cell, fill=bg, align=LEFT)
        ws.row_dimensions[i].height = 36

    # 总分公式说明
    ws.merge_cells("B17:J17")
    ws["B17"] = ("总分 = ①×2 + ②×2 + ③×1 + ④×2 + ⑤×2 + ⑥×2 + ⑦×1 + ⑧×3 + ⑨×3 + ⑩×2 = 满分 100\n"
                 "等级：S（≥80） / A（65-79） / B（50-64） / C（<50，淘汰）")
    ws["B17"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["B17"].font = Font(name="Microsoft YaHei", size=11, bold=True, color=PRIMARY)
    ws["B17"].fill = PatternFill("solid", fgColor=SUCCESS)
    ws["B17"].border = BORDER
    ws.row_dimensions[17].height = 36


def build_main(ws):
    ws.title = "30家候选评分"
    # 26 列 + 序号
    col_widths = [5, 8, 22, 12, 8, 14, 12, 8, 22, 14, 8] + [10]*10 + [8, 8, 12, 14, 22, 12]
    widths(ws, col_widths)

    ws.merge_cells("A1:AA1")
    ws["A1"] = "灯塔厂候选 30 家评分表（黄色 = 必填，绿色 = 自动）"
    ws["A1"].font = TITLE_FONT; ws["A1"].alignment = LEFT

    headers = [
        "#", "客群", "公司全称", "老板姓名", "年龄", "城市",
        "年产值(亿)", "产线", "主要客户", "介绍人", "关系强度",
        "①老板年龄", "②出镜", "③厂区", "④产业带", "⑤产值",
        "⑥客户", "⑦数字化", "⑧决策链", "⑨关系", "⑩见效",
        "总分", "等级", "负责人", "当前进度", "下一步动作", "跟进日"
    ]
    for j, h in enumerate(headers, 1):
        style_h(ws.cell(row=3, column=j, value=h))
    ws.row_dimensions[3].height = 32

    # 30 行
    weights = [2, 2, 1, 2, 2, 2, 1, 3, 3, 2]  # 10 维度权重
    for i in range(4, 34):  # 30 行
        idx = i - 3
        # 序号
        style_t(ws.cell(row=i, column=1, value=idx), fill=LIGHT_BG, bold=True)

        # 输入字段（黄色）
        for col in range(2, 12):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG, align=LEFT)

        # 评分字段（黄色）col 12-21
        for col in range(12, 22):
            style_t(ws.cell(row=i, column=col), fill=INPUT_BG)

        # 总分（绿色，公式）
        # =L*2 + M*2 + N*1 + O*2 + P*2 + Q*2 + R*1 + S*3 + T*3 + U*2
        score_formula = (f"=IF(COUNTA(L{i}:U{i})<10,\"\","
                         f"L{i}*2+M{i}*2+N{i}*1+O{i}*2+P{i}*2+Q{i}*2+R{i}*1+S{i}*3+T{i}*3+U{i}*2)")
        cell = ws.cell(row=i, column=22, value=score_formula)
        style_t(cell, fill=SUCCESS, bold=True)
        cell.number_format = "0"

        # 等级（绿色，公式）
        grade_formula = (f"=IF(V{i}=\"\",\"\","
                         f"IF(V{i}>=80,\"S\","
                         f"IF(V{i}>=65,\"A\","
                         f"IF(V{i}>=50,\"B\",\"C\"))))")
        cell = ws.cell(row=i, column=23, value=grade_formula)
        style_t(cell, fill=SUCCESS, bold=True)

        # 负责人（黄色 + 下拉）
        style_t(ws.cell(row=i, column=24), fill=INPUT_BG)
        # 当前进度（黄色 + 下拉）
        style_t(ws.cell(row=i, column=25), fill=INPUT_BG)
        # 下一步动作（黄色）
        style_t(ws.cell(row=i, column=26), fill=INPUT_BG, align=LEFT)
        # 跟进日（黄色）
        style_t(ws.cell(row=i, column=27), fill=INPUT_BG)

    # 数据有效性
    dv_kq = DataValidation(type="list",
        formula1='"钢管,镀锌冷卷,开平纵剪,一体厂"', allow_blank=True)
    ws.add_data_validation(dv_kq)
    dv_kq.add(f"B4:B33")

    dv_rel = DataValidation(type="list",
        formula1='"强,中,弱,无"', allow_blank=True)
    ws.add_data_validation(dv_rel)
    dv_rel.add("K4:K33")

    dv_score = DataValidation(type="list",
        formula1='"1,2,3,4,5"', allow_blank=True)
    ws.add_data_validation(dv_score)
    dv_score.add("L4:U33")

    dv_owner = DataValidation(type="list",
        formula1='"老板,销售总监A,销售总监B,销售1,销售2"', allow_blank=True)
    ws.add_data_validation(dv_owner)
    dv_owner.add("X4:X33")

    dv_status = DataValidation(type="list",
        formula1='"候选,已接触,首次拜访,工作坊,POC,签约,上线,放弃"',
        allow_blank=True)
    ws.add_data_validation(dv_status)
    dv_status.add("Y4:Y33")

    # 条件格式：等级
    grade_col = "W"
    # S 绿色
    ws.conditional_formatting.add(f"{grade_col}4:{grade_col}33",
        CellIsRule(operator="equal", formula=['"S"'],
                   fill=PatternFill("solid", fgColor="00B050")))
    ws.conditional_formatting.add(f"{grade_col}4:{grade_col}33",
        CellIsRule(operator="equal", formula=['"A"'],
                   fill=PatternFill("solid", fgColor="92D050")))
    ws.conditional_formatting.add(f"{grade_col}4:{grade_col}33",
        CellIsRule(operator="equal", formula=['"B"'],
                   fill=PatternFill("solid", fgColor="FFC000")))
    ws.conditional_formatting.add(f"{grade_col}4:{grade_col}33",
        CellIsRule(operator="equal", formula=['"C"'],
                   fill=PatternFill("solid", fgColor="FF6B6B")))

    # 示例行（第 4 行预填 1 个示例）
    sample = ["钢管", "邯郸 XX 钢管科技", "张 XX", 55, "邯郸",
              1.8, 4, "建筑/家具/出口", "王 XX(轧机厂)", "中",
              4, 3, 4, 4, 5, 5, 4, 4, 3, 4]
    for j, v in enumerate(sample, 2):
        ws.cell(row=4, column=j, value=v)

    # 冻结窗格
    ws.freeze_panes = "D4"


def build_distribution(ws):
    ws.title = "等级分布"
    widths(ws, [4, 14, 10, 14, 30])
    ws.merge_cells("B2:E2")
    ws["B2"] = "30 家清单等级自动统计"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["等级", "家数", "占比", "建议动作"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    grades = [
        ("S（≥80）", "S", "老板亲自跑 + 5 折/利润分成", SUCCESS),
        ("A（65-79）", "A", "销售总监主跟 + 6-7 折", LIGHT_BG),
        ("B（50-64）", "B", "排队跟进", WARN),
        ("C（<50）", "C", "出 30 家清单", CRITICAL),
    ]
    for i, (label, code, action, color) in enumerate(grades, 5):
        style_t(ws.cell(row=i, column=2, value=label), fill=color, bold=True)
        style_t(ws.cell(row=i, column=3,
                value=f'=COUNTIF(\'30家候选评分\'!W4:W33,"{code}")'),
                fill=color, bold=True)
        # 占比
        style_t(ws.cell(row=i, column=4,
                value=f'=IFERROR(C{i}/COUNTA(\'30家候选评分\'!W4:W33),0)'),
                fill=color)
        ws.cell(row=i, column=4).number_format = "0%"
        style_t(ws.cell(row=i, column=5, value=action), fill=color, align=LEFT)

    # 总计
    i = 9
    style_t(ws.cell(row=i, column=2, value="累计已评分"), fill=PRIMARY)
    ws.cell(row=i, column=2).font = Font(name="Microsoft YaHei", size=10,
                                          bold=True, color=WHITE)
    style_t(ws.cell(row=i, column=3,
            value="=COUNTA('30家候选评分'!W4:W33)"),
            fill=PRIMARY)
    ws.cell(row=i, column=3).font = Font(name="Microsoft YaHei", size=10,
                                          bold=True, color=WHITE)


def build_progress(ws):
    ws.title = "推进进度看板"
    widths(ws, [4, 16, 10, 14, 30])
    ws.merge_cells("B2:E2")
    ws["B2"] = "30 家清单推进进度自动统计"
    ws["B2"].font = TITLE_FONT; ws["B2"].alignment = LEFT

    headers = ["阶段", "家数", "占比", "目标值"]
    for j, h in enumerate(headers, 2):
        style_h(ws.cell(row=4, column=j, value=h))

    stages = [
        ("候选",        "≥ 30 家"),
        ("已接触",      "≥ 25 家"),
        ("首次拜访",    "≥ 20 家"),
        ("工作坊",      "≥ 8 家"),
        ("POC",         "≥ 6 家"),
        ("签约",        "≥ 4 家（含灯塔）"),
        ("上线",        "≥ 2 家"),
        ("放弃",        "≤ 5 家"),
    ]
    for i, (stage, target) in enumerate(stages, 5):
        bg = LIGHT_BG if i % 2 == 0 else WHITE
        style_t(ws.cell(row=i, column=2, value=stage), fill=bg, bold=True)
        style_t(ws.cell(row=i, column=3,
                value=f'=COUNTIF(\'30家候选评分\'!Y4:Y33,"{stage}")'),
                fill=bg, bold=True)
        style_t(ws.cell(row=i, column=4,
                value=f'=IFERROR(C{i}/COUNTA(\'30家候选评分\'!Y4:Y33),0)'),
                fill=bg)
        ws.cell(row=i, column=4).number_format = "0%"
        style_t(ws.cell(row=i, column=5, value=target), fill=bg, align=LEFT)


def main():
    wb = Workbook()
    build_readme(wb.active)
    build_dimensions(wb.create_sheet())
    build_main(wb.create_sheet())
    build_distribution(wb.create_sheet())
    build_progress(wb.create_sheet())

    out = "tools/灯塔厂候选30家_评分表.xlsx"
    wb.save(out)
    print(f"已生成：{out}")
    print(f"工作表：{wb.sheetnames}")


if __name__ == "__main__":
    main()
