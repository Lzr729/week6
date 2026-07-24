# Week6 七家公司IPO股本与PE/VC工程化提取

## 完成状态

- 公司7家；
- 事件26条；
- 交易27条；
- 数值校验失败0；
- PE/VC主体2个；
- PE/VC直接增资路径2条；
- 人工复核决定19条；
- 开放复核0；
- Final验证PASSED。

## 一条统一运行命令

```powershell
python -m pipeline.cli run-all `
  --input-dir "data" `
  --repo-root "." `
  --workspace-dir "workspace" `
  --offline-replay
```

完整多层运行中间文件写入`workspace/runtime_repo/`。
提交目录中的精简Auto、Final、Validation、Review和Logs保持扁平。

## 输入

- `data/<证券代码>_page_text.jsonl`
- `data/pdf_manifest.csv`
- `data/company_registry.json`

## 结果目录

- `auto_output/`：五个阶段合并Auto JSONL；
- `manual_gold/`：人工Gold和前后值台账；
- `final/`：按数据类型合并的Final JSONL及合并Excel；
- `validation/`：合并Schema、摘要、记录、复核队列和Excel；
- `review/`：章节Patch、人工决定、摘要和Excel；
- `logs/`：运行、执行、交互、重放、数量和哈希；
- `report/week6_report.md`：研究报告。

## 精简与追溯

同阶段、同语义的数据合并；每条记录保留`source_file`、
`record_type`、运行ID或证券代码。候选证据TXT逐条嵌入JSONL。

代码、配置、测试、提示词和七家公司分页文本保持独立，
避免形成难以维护的超大文件。

## 人工介入位置

- `review/chapter_location_review_patch.jsonl`
- `review/pagination_review_patch.jsonl`
- `review/review_decisions.jsonl`
- `manual_gold/review_decisions_gold.jsonl`
- `manual_gold/manual_gold.xlsx`

人工决定保留before_value、after_value、decision、reason、
evidence_ids、PDF页码和正文页码。

## Prompt与交互材料

阶段指令位于`prompts/`；交互索引和结构化输出合并至
`logs/model_interactions.jsonl`。未记录字段标记为
`UNKNOWN_NOT_RECORDED`，不补造。

## 运行环境

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m unittest discover -s pipeline/tests
```

## 已知边界

- 股东与持股比例未形成稳定逐行对应，不生成确定股权结构快照；
- 聚合披露但未逐项展开的事件按披露限制保留；
- 未披露或无法可靠识别的字段保持为空。
