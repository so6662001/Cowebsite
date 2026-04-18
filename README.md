# Cowebsite —— 钢铁数字化系统营销作战包

本仓库交付一整套面向钢铁加工行业（钢管 / 镀锌冷卷 / 一体厂 / 开平纵剪）民营企业的
"**实施费 + 按产线月租 + 模块解锁**"商业模式落地材料，含合同、报价工具、销售 SOP、
合伙人小程序 PRD、灯塔厂与老板局执行手册。

## 快速入口

| 文件 | 说明 |
|---|---|
| [`docs/00-总览-营销作战包.md`](docs/00-总览-营销作战包.md) | **先看这个**，全局总览 + 落地路线 |
| [`docs/01-合同模板.md`](docs/01-合同模板.md) | 标准合同（含三段付款、6 个月 KPI 退款、返佣条款） |
| [`tools/钢铁数字化系统报价单_模型.xlsx`](tools/钢铁数字化系统报价单_模型.xlsx) | 销售自助报价工具，6 个 sheet |
| [`tools/build_quote_model.py`](tools/build_quote_model.py) | 重新生成报价 Excel 的脚本 |
| [`docs/03-销售SOP手册与异议话术库.md`](docs/03-销售SOP手册与异议话术库.md) | 五步成交 SOP + 10 个高频异议话术 |
| [`docs/04-合伙人小程序PRD.md`](docs/04-合伙人小程序PRD.md) | 合伙人小程序产品需求 |
| [`docs/05-灯塔厂招募与老板局执行手册.md`](docs/05-灯塔厂招募与老板局执行手册.md) | 灯塔厂招募 + 第一场老板局 T-14 清单 |

## 重新生成报价模型

```bash
pip install openpyxl
python3 tools/build_quote_model.py
```

输出：`tools/钢铁数字化系统报价单_模型.xlsx`
