# Cowebsite —— 钢铁数字化系统营销作战包

本仓库交付一整套面向钢铁加工行业（钢管 / 镀锌冷卷 / 一体厂 / 开平纵剪）民营企业的
"**实施费 + 按产线月租 + 模块解锁**"商业模式落地材料，含合同、报价工具、销售 SOP、
合伙人小程序 PRD、灯塔厂与老板局执行手册。

## 快速入口

### 总览
| 文件 | 说明 |
|---|---|
| [`docs/00-总览-营销作战包.md`](docs/00-总览-营销作战包.md) | **先看这个**，全局总览 + 落地路线 |

### v1.0 基础包（5 份）
| 文件 | 说明 |
|---|---|
| [`docs/01-合同模板.md`](docs/01-合同模板.md) | 标准合同（含三段付款、6 个月 KPI 退款、返佣条款） |
| [`tools/钢铁数字化系统报价单_模型.xlsx`](tools/钢铁数字化系统报价单_模型.xlsx) | 销售自助报价工具，**v2.0 7 个 sheet**（已更新为新版价格） |
| [`tools/build_quote_model.py`](tools/build_quote_model.py) | 重新生成报价 Excel 的脚本 |
| [`docs/03-销售SOP手册与异议话术库.md`](docs/03-销售SOP手册与异议话术库.md) | 五步成交 SOP + 10 个高频异议话术 |
| [`docs/04-合伙人小程序PRD.md`](docs/04-合伙人小程序PRD.md) | 合伙人小程序产品需求 |
| [`docs/05-灯塔厂招募与老板局执行手册.md`](docs/05-灯塔厂招募与老板局执行手册.md) | 灯塔厂招募 + 第一场老板局 T-14 清单 |

### v2.0 增量包（5 份，回应一线落地反馈）
| 文件 | 说明 |
|---|---|
| [`docs/06-收费方案修订版-降门槛.md`](docs/06-收费方案修订版-降门槛.md) | **价格腰斩**、3 档套餐、3 种付费模式（含 0 首付/利润分成） |
| [`docs/07-健康度自评表与诊断初稿样例.md`](docs/07-健康度自评表与诊断初稿样例.md) | 30 题自评表 + 1 页诊断初稿手写样例 |
| [`docs/08-流程现状图与数字化路径图样例.md`](docs/08-流程现状图与数字化路径图样例.md) | As-Is + To-Be 完整样例（以邯郸某焊管厂为对象） |
| [`docs/09-钢管行业数字化白皮书样例.md`](docs/09-钢管行业数字化白皮书样例.md) | 60-80 页行业白皮书完整结构 + 4 个标杆案例 |
| [`docs/10-里程碑实际时间与团队搭建.md`](docs/10-里程碑实际时间与团队搭建.md) | **12-15 个月时间盘** + 团队结构 + 负责人画像 + 老板自己怎么排时间 |

### v3.0 增量包（回应"价格落地难+团队搭建+5 份扩展"）
| 文件 | 说明 |
|---|---|
| [`docs/11-四个落地问题集中回答.md`](docs/11-四个落地问题集中回答.md) | 试用版无ERP/MES怎么跑、无CMO怎么办、谁招合伙人、地推有没必要 |
| [`docs/12-三份姐妹行业白皮书结构.md`](docs/12-三份姐妹行业白皮书结构.md) | 镀锌冷卷 / 开平纵剪 / 一体厂 三份姐妹白皮书差异化结构 |
| [`docs/13-CMO面试题库与试用期KPI.md`](docs/13-CMO面试题库与试用期KPI.md) | CMO 招聘画像 + 5 道核心题 + 6 个月试用期 KPI 模板 |
| [`tools/generate_quote_pdf.py`](tools/generate_quote_pdf.py) | 客户用 PDF 报价单生成器（中文 ReportLab） |
| [`tools/sample_quote.pdf`](tools/sample_quote.pdf) | PDF 报价单样例（邯郸 XX 焊管厂） |
| [`tools/health_assessment.html`](tools/health_assessment.html) | 30 题在线健康度自评单页（手机/桌面响应式，可分享、可打印 PDF） |
| [`tools/build_financial_model.py`](tools/build_financial_model.py) | 首年财务预测模型生成脚本 |
| [`tools/钢铁数字化系统_首年财务预测模型.xlsx`](tools/钢铁数字化系统_首年财务预测模型.xlsx) | **首年财务模型 8 sheet**：假设/漏斗/收入/成本/现金流/损益/情景对比 |

### v4.0 增量包（**战略时机复盘**：行业窗口紧迫，节奏必须压缩）
| 文件 | 说明 |
|---|---|
| [`docs/14-时间窗口复盘与压缩节奏方案.md`](docs/14-时间窗口复盘与压缩节奏方案.md) | **回答"窗口期紧 / 还做不做"**：32 个月真实窗口 / 21 个月压缩节奏 / Plan A B C 三档 / 6 月生死线 |
| [`tools/build_financial_model_planA.py`](tools/build_financial_model_planA.py) | Plan A 压缩版财务模型生成器 |
| [`tools/钢铁数字化系统_首年财务预测模型_PlanA.xlsx`](tools/钢铁数字化系统_首年财务预测模型_PlanA.xlsx) | **Plan A 压缩版财务模型**（与基准版对比验证） |

### v5.0 增量包（**Plan A 启动 90 天工具箱**）
| 文件 | 说明 |
|---|---|
| [`docs/15-作战日历2026年4到6月.md`](docs/15-作战日历2026年4到6月.md) | **2026.04-06 三个月作战日历**：12 周精确到天 + 90 天预算 |
| [`tools/作战日历_2026年4到6月_甘特图.xlsx`](tools/作战日历_2026年4到6月_甘特图.xlsx) | 作战日历 Excel：总览看板 + 12 周甘特图 + 周任务清单 + 6 月生死线 + 90 天预算 |
| [`docs/16-灯塔厂候选名单调研与评分模板.md`](docs/16-灯塔厂候选名单调研与评分模板.md) | **灯塔厂调研**：3 不要 + 5 必要 + 10 维度评分模型 + 调研 4 板斧 + 1 页档案模板 |
| [`tools/灯塔厂候选30家_评分表.xlsx`](tools/灯塔厂候选30家_评分表.xlsx) | 30 家候选评分 Excel：自动算总分 + 等级（S/A/B/C）+ 分布 + 推进进度 |
| [`docs/17-老板局倒推5周邀约脚本.md`](docs/17-老板局倒推5周邀约脚本.md) | **第 1 场老板局邀约脚本**：T-5 → T+7 完整话术 + 5 周时间轴 + 7 个细节 |
| [`tools/老板局邀约_跟踪表.xlsx`](tools/老板局邀约_跟踪表.xlsx) | 邀约跟踪 Excel：12 老板跟踪 + 5 周时间轴 + 现场座位安排 + T+7 跟进 |
| [`docs/18-6个月生死线周战报模板.md`](docs/18-6个月生死线周战报模板.md) | **周战报模板**：5 数字+3 故事+1 决策 + 月度战报 + W6/W12 评估 |
| [`tools/周战报_6个月生死线跟踪表.xlsx`](tools/周战报_6个月生死线跟踪表.xlsx) | 周战报 Excel：单周模板 + 24 周看板 + 生死线自动评级 + 月度汇总 + W6/W12 评估 |

## 重新生成工具

```bash
pip install openpyxl reportlab

# 报价 Excel
python3 tools/build_quote_model.py
# 输出: tools/钢铁数字化系统报价单_模型.xlsx

# 客户 PDF 报价单
python3 tools/generate_quote_pdf.py
# 输出: tools/sample_quote.pdf

# 首年财务预测模型（基准版）
python3 tools/build_financial_model.py
# 输出: tools/钢铁数字化系统_首年财务预测模型.xlsx

# 首年财务预测模型（Plan A 压缩版，21 个月节奏）
python3 tools/build_financial_model_planA.py
# 输出: tools/钢铁数字化系统_首年财务预测模型_PlanA.xlsx

# Plan A 启动 90 天 4 个 Excel 工具
python3 tools/build_battle_calendar.py        # 作战日历甘特图
python3 tools/build_lighthouse_scoring.py     # 灯塔厂 30 家评分表
python3 tools/build_invitation_tracker.py     # 老板局邀约跟踪表
python3 tools/build_weekly_report.py          # 周战报跟踪表

# 在线健康度自评（直接用浏览器打开即可）
# tools/health_assessment.html
```
