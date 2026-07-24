# Auto输出（精简版）

同一阶段的Auto结果合并为一个JSONL：

- `chapter_location_auto.jsonl`
- `candidate_events_auto.jsonl`
- `structured_events_auto.jsonl`
- `numeric_validation_auto.jsonl`
- `pevc_paths_auto.jsonl`

每条记录保留源文件名、记录类型和运行ID。
候选证据TXT已逐条嵌入JSONL，没有删除原文。
