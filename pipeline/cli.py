from __future__ import annotations

import argparse
from pathlib import Path

from pipeline.batch_check import run_batch_check
from pipeline.candidate_events import run_candidate_event_generation
from pipeline.chapter_location import run_chapter_location
from pipeline.offline_replay import run_all


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="week6-pipeline",
        description="IPO招股说明书批处理公共流水线",
    )
    subparsers = parser.add_subparsers(
        dest="command",
        required=True,
    )

    run_all_parser = subparsers.add_parser(
        "run-all",
        help=(
            "从分页文本开始重建Auto、验证、"
            "Manual Gold合并和Final"
        ),
    )
    run_all_parser.add_argument(
        "--input-dir",
        type=Path,
        required=True,
    )
    run_all_parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path("."),
    )
    run_all_parser.add_argument(
        "--workspace-dir",
        type=Path,
        required=True,
    )
    run_all_parser.add_argument(
        "--offline-replay",
        action="store_true",
    )
    run_all_parser.add_argument(
        "--expected-count",
        type=int,
        default=7,
    )

    batch_check = subparsers.add_parser(
        "batch-check",
        help="批量执行PDF基础检查",
    )
    batch_check.add_argument(
        "--input-dir",
        type=Path,
        required=True,
    )
    batch_check.add_argument(
        "--repo-root",
        type=Path,
        default=Path("."),
    )
    batch_check.add_argument(
        "--expected-count",
        type=int,
        default=None,
    )

    locate = subparsers.add_parser(
        "locate-chapters",
        help="定位发行人基本情况及必需章节",
    )
    locate.add_argument(
        "--input-dir",
        type=Path,
        required=True,
    )
    locate.add_argument(
        "--repo-root",
        type=Path,
        default=Path("."),
    )
    locate.add_argument(
        "--workspace-dir",
        type=Path,
        required=True,
    )
    locate.add_argument(
        "--rules-file",
        type=Path,
        default=Path(
            "pipeline/configs/"
            "chapter_locator_rules.json"
        ),
    )
    locate.add_argument(
        "--expected-count",
        type=int,
        default=None,
    )
    locate.add_argument(
        "--reuse-run-id",
        default=None,
    )

    candidates = subparsers.add_parser(
        "generate-candidate-events",
        help="在冻结章节范围内生成候选事件",
    )
    candidates.add_argument(
        "--input-dir",
        type=Path,
        required=True,
    )
    candidates.add_argument(
        "--repo-root",
        type=Path,
        default=Path("."),
    )
    candidates.add_argument(
        "--workspace-dir",
        type=Path,
        required=True,
    )
    candidates.add_argument(
        "--chapter-patch-file",
        type=Path,
        default=Path(
            "review/chapter_location/"
            "chapter_location_review_patch.jsonl"
        ),
    )
    candidates.add_argument(
        "--pagination-patch-file",
        type=Path,
        default=Path(
            "review/chapter_location/"
            "pagination_review_patch.jsonl"
        ),
    )
    candidates.add_argument(
        "--rules-file",
        type=Path,
        default=Path(
            "pipeline/configs/"
            "candidate_event_rules.json"
        ),
    )
    candidates.add_argument(
        "--page-text-run-id",
        required=True,
    )
    candidates.add_argument(
        "--expected-count",
        type=int,
        default=None,
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.command == "run-all":
        return run_all(
            input_dir=args.input_dir,
            repo_root=args.repo_root,
            workspace_dir=args.workspace_dir,
            offline_replay=args.offline_replay,
            expected_count=args.expected_count,
        )

    if args.command == "batch-check":
        return run_batch_check(
            input_dir=args.input_dir,
            repo_root=args.repo_root,
            expected_count=args.expected_count,
        )

    if args.command == "locate-chapters":
        return run_chapter_location(
            input_dir=args.input_dir,
            repo_root=args.repo_root,
            workspace_dir=args.workspace_dir,
            rules_file=args.rules_file,
            expected_count=args.expected_count,
            reuse_run_id=args.reuse_run_id,
        )

    if args.command == "generate-candidate-events":
        return run_candidate_event_generation(
            input_dir=args.input_dir,
            repo_root=args.repo_root,
            workspace_dir=args.workspace_dir,
            chapter_patch_file=(
                args.chapter_patch_file
            ),
            pagination_patch_file=(
                args.pagination_patch_file
            ),
            rules_file=args.rules_file,
            page_text_run_id=(
                args.page_text_run_id
            ),
            expected_count=args.expected_count,
        )

    parser.error(
        f"不支持的命令：{args.command}"
    )
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
