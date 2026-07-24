# Week6 上市前融资分析报告

## 1. 研究目标

对七家公司招股说明书中的股本演变、增资、股权转让及PE/VC进入路径进行工程化提取，
并建立可复现的Auto、Manual Gold、Final和验证链路。

## 2. 公司范围

- 001282 三联锻造
- 301581 黄山谷捷
- 603418 友升股份
- 688758 赛分科技
- 688775 影石创新
- 920100 三协电机
- 920116 星图测控

## 3. 工程化流程

分页文本 → 章节定位 → 候选事件Auto → 候选冻结 → 结构化事件Auto →
数值校验 → PE/VC路径Auto → Manual Gold → Cross-check → Final。

## 4. 主要结果

- 公司：7
- 事件：26
- 交易：27
- 数值校验失败：0
- PE/VC候选主体：2
- PE/VC投资路径：2
- 人工复核决定：19
- 开放复核：0

最终PE/VC结论为友升股份2020年9月增资中的金浦临港基金和金浦科创基金，
均为发行人层面的直接增资进入。

## 5. Auto、Gold与Final

- `auto_output/`：按阶段合并的未经人工修改Auto结果。
- `manual_gold/`：人工确认后的Gold数据和前后值台账。
- `final/`：按数据类型合并的JSONL与合并Excel。

本轮人工复核没有补写未披露事实；接受当前结果的记录前后值相同，
重复或错误候选的after_value为null。

## 6. 证据与数值

事件保留PDF页码、正文页码和证据ID。
原文披露值标记为DISCLOSED；公式计算值标记为CALCULATED。

## 7. 已知边界

股东名称与持股比例尚未形成稳定逐行对应，因此不强行生成确定股权结构快照。
招股书仅聚合披露但未逐项展开的事件按披露限制保留，不拆分、不编造。

## 8. 复现

```powershell
python -m pipeline.cli run-all `
  --input-dir "data" `
  --repo-root "." `
  --workspace-dir "workspace" `
  --offline-replay
```

人工介入位置为 `review/` 和 `manual_gold/`。


## 9. 可复现性验证

统一命令已在隔离副本中从 `data/`中的`<证券代码>_page_text.jsonl` 完整执行。
章节定位、候选事件、候选冻结、结构化事件、数值校验、PE/VC识别、
Manual Gold合并和Final均成功运行。

- 单元测试：102项通过；
- Replay Cross-check：PASSED；
- 事件26、交易27；
- 数值失败0；
- PE/VC主体2、路径2；
- 人工Gold开放复核0。


## 10. 精简提交说明

同阶段、同语义文件已经合并，所有合并记录保留源文件名、
记录类型、运行ID或证券代码。精简不改变最终数据结论。
