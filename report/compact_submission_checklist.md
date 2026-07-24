# Week6精简提交验收

## 文件数量

- 精简前：281
- 精简后：104
- 减少：177
- 减少比例：63.0%

## 目录文件数

- `data/`：13
- `auto_output/`：7
- `final/`：10
- `validation/`：9
- `review/`：6
- `logs/`：6
- `manual_gold/`：7
- `prompts/`：8
- `pipeline/`：31
- `report/`：2

## 验收结果

- 单元测试：PASSED；
- 完整离线重放：PASSED；
- Replay Cross-check：PASSED；
- 公司7家；
- 事件26条；
- 交易27条；
- 数值校验记录81条；
- 数值校验失败0；
- PE/VC主体2个；
- PE/VC路径2条；
- 人工复核决定19条；
- 开放复核0；
- 重放后结果目录仍无子文件夹；
- 旧Final/Review/Logs路径引用0；
- 两份新增汇总Excel公式错误0。

## 精简原则

- 同阶段JSONL合并，保留`source_file`与`record_type`；
- TXT证据片段嵌入JSONL；
- CSV与JSONL重复时优先JSONL；
- 七份单公司Excel删除，保留合并Excel；
- 逐公司Final JSON按数据类型合并；
- 代码、配置、测试、提示词和分页文本保持独立。
