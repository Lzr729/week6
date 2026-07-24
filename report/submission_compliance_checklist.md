# GitHub提交合规检查

| 要求 | 状态 | 位置 |
|---|---|---|
| 一条统一运行命令 | 通过 | `README.md` |
| 分页文本输入 | 通过 | `data/` |
| Auto结果 | 通过 | `auto_output/` |
| Manual Gold | 通过 | `manual_gold/` |
| Final结果 | 通过 | `final/` |
| Schema与Cross-check | 通过 | `validation/` |
| 人工前后值与证据 | 通过 | `review/review_decisions.jsonl` |
| Prompt与交互记录 | 通过 | `prompts/`, `logs/model_interactions.jsonl` |
| 运行日志与文件哈希 | 通过 | `logs/` |
| 周报 | 通过 | `report/week6_report.md` |
| 单元测试 | 通过，102项 | `validation/unit_test_summary.json` |
| 离线重放 | PASSED | `logs/offline_replay.json` |
| 开放人工复核 | 0 | `review/review_summary.json` |

精简包没有删除事实或证据，只合并重复和同类文件。
