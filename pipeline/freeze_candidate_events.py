from __future__ import annotations

import argparse
import csv
import hashlib
import json
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


FREEZE_VERSION = "candidate_event_freeze_patch_v1"

EXPECTED_COUNTS = {
    "001282": 3,
    "301581": 6,
    "603418": 5,
    "688758": 7,
    "688775": 3,
    "920100": 1,
    "920116": 1,
}


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def make_freeze_id() -> str:
    return datetime.now().astimezone().strftime(
        "CANDIDATEFREEZE_V1_%Y%m%d_%H%M%S"
    )


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8-sig") as handle:
        for line_number, line in enumerate(handle, start=1):
            if not line.strip():
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError as exc:
                raise ValueError(
                    f"JSONL解析失败：{path} 第{line_number}行：{exc}"
                ) from exc
    return rows


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def write_jsonl(path: Path, rows: Iterable[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        for row in rows:
            handle.write(json.dumps(row, ensure_ascii=False) + "\n")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def normalized_period(candidate: dict[str, Any]) -> str:
    for key in (
        "event_period",
        "event_date_text",
    ):
        value = candidate.get(key)
        if value:
            return str(value).replace("年", "-").replace("月", "-").strip("-")
    values = candidate.get("event_date_candidates") or []
    return str(values[0]) if values else ""


def title_text(candidate: dict[str, Any]) -> str:
    return str(
        candidate.get("event_title")
        or candidate.get("title")
        or ""
    )


def candidate_signals(candidate: dict[str, Any]) -> list[str]:
    values = candidate.get("matched_signals") or candidate.get("signals") or []
    return [str(value) for value in values]


def is_timeline_recovery(candidate: dict[str, Any]) -> bool:
    text = " ".join(candidate_signals(candidate)).lower()
    source_kind = str(candidate.get("source_kind") or "").lower()
    return (
        "page_timeline_recovery" in text
        or "page_timeline_recovery" in source_kind
    )


def is_limited_establishment(candidate: dict[str, Any]) -> bool:
    primary = str(candidate.get("event_type_candidate") or "")
    values = candidate.get("event_type_candidates") or []
    return (
        primary == "LIMITED_COMPANY_ESTABLISHMENT"
        or "LIMITED_COMPANY_ESTABLISHMENT" in values
    )


def rank_establishment_candidate(
    company_id: str,
    candidate: dict[str, Any],
) -> tuple[int, int, str]:
    period = normalized_period(candidate)
    title = title_text(candidate)
    recovery = is_timeline_recovery(candidate)

    if company_id == "001282":
        if period.startswith("2004-06-18"):
            return (0, 0, title)
        if "有限公司设立情况" in title and not recovery:
            return (1, 0, title)
        if period.startswith("1999") or recovery:
            return (9, 0, title)

    if company_id == "603418":
        if period.startswith("1992-12-04"):
            return (0, 0, title)
        if "有限责任公司" in title and not recovery:
            return (1, 0, title)
        if period.startswith("1992-10-22") or recovery:
            return (9, 0, title)

    if company_id == "301581":
        if period.startswith("2012-06"):
            return (0, 0, title)
        if recovery and "发行人前身" in title:
            return (1, 0, title)
        if "昆山谷捷" in title or period.startswith("2009"):
            return (9, 0, title)

    return (
        5,
        1 if recovery else 0,
        title,
    )


def select_establishment_candidate(
    company_id: str,
    candidates: list[dict[str, Any]],
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    if not candidates:
        raise ValueError(
            f"{company_id}没有有限公司设立候选，无法冻结"
        )
    ranked = sorted(
        candidates,
        key=lambda item: rank_establishment_candidate(
            company_id,
            item,
        ),
    )
    best = ranked[0]
    best_rank = rank_establishment_candidate(
        company_id,
        best,
    )[0]
    if company_id in {"001282", "301581", "603418"} and best_rank > 1:
        raise ValueError(
            f"{company_id}未找到预期的设立候选："
            f"{[(normalized_period(x), title_text(x)) for x in ranked]}"
        )
    return best, ranked[1:]


def build_patch_record(
    *,
    freeze_id: str,
    source_run_id: str,
    candidate: dict[str, Any],
    decision: str,
    reason: str,
    replacement_id: str | None,
) -> dict[str, Any]:
    return {
        "review_patch_id": (
            f"CEPATCH-{candidate['company_id']}-"
            f"{candidate['candidate_event_id']}"
        ),
        "freeze_id": freeze_id,
        "source_run_id": source_run_id,
        "candidate_event_id": candidate["candidate_event_id"],
        "company_id": candidate["company_id"],
        "event_type_candidate": candidate.get(
            "event_type_candidate"
        ),
        "event_period": candidate.get("event_period"),
        "event_title": title_text(candidate),
        "decision": decision,
        "replacement_candidate_event_id": replacement_id,
        "reason": reason,
        "before_value": candidate,
        "reviewer_type": "AI_ASSISTED_EVIDENCE_REVIEW",
        "human_signoff_status": "NOT_CLAIMED",
        "created_at": now_iso(),
    }


def apply_review_patch(
    candidates: list[dict[str, Any]],
    source_run_id: str,
    freeze_id: str,
) -> tuple[
    list[dict[str, Any]],
    list[dict[str, Any]],
]:
    by_company: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for candidate in candidates:
        by_company[str(candidate["company_id"])].append(candidate)

    keep_ids: set[str] = {
        str(candidate["candidate_event_id"])
        for candidate in candidates
    }
    patches: list[dict[str, Any]] = []

    for company_id in ("001282", "301581", "603418"):
        establishment = [
            candidate
            for candidate in by_company.get(company_id, [])
            if is_limited_establishment(candidate)
        ]
        selected, rejected = select_establishment_candidate(
            company_id,
            establishment,
        )

        if company_id == "001282":
            selected_reason = (
                "保留PDF详细披露的2004年6月18日有限公司设立；"
                "时间线恢复的1999年背景日期不构成独立设立事件。"
            )
        elif company_id == "603418":
            selected_reason = (
                "保留PDF详细披露的1992年12月4日有限公司设立；"
                "1992年10月22日为政府批复日期，不是工商设立日期。"
            )
        else:
            selected_reason = (
                "保留2012年6月发行人前身设立时间线候选；"
                "昆山谷捷为被吸收方，其2009年设立不属于发行人自身历史。"
            )

        patches.append(build_patch_record(
            freeze_id=freeze_id,
            source_run_id=source_run_id,
            candidate=selected,
            decision="ACCEPTED",
            reason=selected_reason,
            replacement_id=None,
        ))

        for candidate in rejected:
            candidate_id = str(
                candidate["candidate_event_id"]
            )
            keep_ids.discard(candidate_id)
            patches.append(build_patch_record(
                freeze_id=freeze_id,
                source_run_id=source_run_id,
                candidate=candidate,
                decision="REJECTED_DUPLICATE_OR_WRONG_ENTITY",
                reason=(
                    "与已接受的发行人设立事件重复，"
                    "或属于被吸收方/程序性日期；不进入冻结候选。"
                ),
                replacement_id=str(
                    selected["candidate_event_id"]
                ),
            ))

    frozen = [
        candidate
        for candidate in candidates
        if str(candidate["candidate_event_id"]) in keep_ids
    ]
    frozen.sort(
        key=lambda item: (
            str(item["company_id"]),
            int(item.get("pdf_page_start") or 0),
            str(item["candidate_event_id"]),
        )
    )
    return frozen, patches


def validate_frozen(
    frozen: list[dict[str, Any]],
    evidence: list[dict[str, Any]],
) -> dict[str, Any]:
    errors: list[str] = []
    counts = Counter(
        str(item["company_id"])
        for item in frozen
    )
    actual_counts = {
        company_id: counts.get(company_id, 0)
        for company_id in EXPECTED_COUNTS
    }
    if actual_counts != EXPECTED_COUNTS:
        errors.append(
            "冻结后公司候选数量不符合预期："
            f"actual={actual_counts}, expected={EXPECTED_COUNTS}"
        )

    ids = [
        str(item["candidate_event_id"])
        for item in frozen
    ]
    if len(ids) != len(set(ids)):
        errors.append("冻结候选ID重复")

    evidence_ids = {
        str(item["evidence_id"])
        for item in evidence
    }
    for candidate in frozen:
        primary = candidate.get("primary_evidence_id")
        if primary and str(primary) not in evidence_ids:
            errors.append(
                f"{candidate['candidate_event_id']}缺少主证据{primary}"
            )
        for supporting in (
            candidate.get("supporting_evidence_ids") or []
        ):
            if str(supporting) not in evidence_ids:
                errors.append(
                    f"{candidate['candidate_event_id']}缺少支持证据"
                    f"{supporting}"
                )

    mapped = sum(
        item.get("printed_page_start") is not None
        and item.get("printed_page_end") is not None
        for item in frozen
    )
    coverage = round(
        mapped / max(len(frozen), 1),
        4,
    )
    if coverage != 1.0:
        errors.append(
            f"正文页码覆盖率不是100%：{coverage}"
        )

    return {
        "validation_status": (
            "FAILED" if errors else "PASSED"
        ),
        "errors": errors,
        "company_candidate_counts": actual_counts,
        "candidate_event_count": len(frozen),
        "printed_page_mapped_count": mapped,
        "printed_page_coverage_rate": coverage,
    }


def write_csv(
    path: Path,
    candidates: list[dict[str, Any]],
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = [
        "company_id",
        "candidate_event_id",
        "event_type_candidate",
        "event_period",
        "event_title",
        "pdf_page_start",
        "pdf_page_end",
        "printed_page_start",
        "printed_page_end",
        "printed_page_value_type",
        "disclosure_scope",
        "candidate_confidence",
        "primary_evidence_id",
    ]
    with path.open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=headers,
            extrasaction="ignore",
        )
        writer.writeheader()
        for candidate in candidates:
            row = dict(candidate)
            row["event_title"] = title_text(candidate)
            writer.writerow(row)


def run_freeze(
    repo_root: Path,
    run_id: str,
) -> int:
    repo_root = repo_root.expanduser().resolve()
    auto_dir = (
        repo_root
        / "auto_output"
        / "candidate_events"
        / "runs"
        / run_id
    )
    validation_dir = (
        repo_root
        / "validation"
        / "candidate_events"
        / "runs"
        / run_id
    )

    input_paths = {
        "candidates": (
            auto_dir / "candidate_events_auto.jsonl"
        ),
        "evidence": (
            auto_dir / "candidate_evidence_auto.jsonl"
        ),
        "shareholder": (
            auto_dir / "shareholder_evidence_auto.jsonl"
        ),
        "coverage": (
            auto_dir / "coverage_gaps_auto.jsonl"
        ),
        "metrics": (
            validation_dir / "candidate_event_metrics.json"
        ),
    }
    missing = [
        str(path)
        for path in input_paths.values()
        if not path.is_file()
    ]
    if missing:
        raise FileNotFoundError(
            "缺少v0.11输入文件：\n" + "\n".join(missing)
        )

    source_metrics = read_json(input_paths["metrics"])
    if source_metrics.get("company_failure_count") != 0:
        raise ValueError("源运行存在公司失败，禁止冻结")
    if source_metrics.get("validation_error_count") != 0:
        raise ValueError("源运行存在验证错误，禁止冻结")

    candidates = read_jsonl(input_paths["candidates"])
    evidence = read_jsonl(input_paths["evidence"])
    shareholder = read_jsonl(input_paths["shareholder"])
    coverage = read_jsonl(input_paths["coverage"])

    freeze_id = make_freeze_id()
    frozen, patches = apply_review_patch(
        candidates,
        run_id,
        freeze_id,
    )

    kept_candidate_ids = {
        str(item["candidate_event_id"])
        for item in frozen
    }
    frozen_evidence = [
        item
        for item in evidence
        if str(item["candidate_event_id"])
        in kept_candidate_ids
    ]

    validation = validate_frozen(
        frozen,
        frozen_evidence,
    )

    output_dir = (
        repo_root
        / "review"
        / "candidate_events"
        / "frozen"
        / freeze_id
    )
    output_dir.mkdir(parents=True, exist_ok=False)

    write_jsonl(
        output_dir / "candidate_event_review_patch.jsonl",
        patches,
    )
    write_jsonl(
        output_dir / "candidate_events_frozen.jsonl",
        frozen,
    )
    write_jsonl(
        output_dir / "candidate_evidence_frozen.jsonl",
        frozen_evidence,
    )
    write_jsonl(
        output_dir / "shareholder_evidence_frozen.jsonl",
        shareholder,
    )
    write_jsonl(
        output_dir / "coverage_gaps_frozen.jsonl",
        coverage,
    )
    write_jsonl(
        output_dir / "review_history.jsonl",
        patches,
    )
    write_csv(
        output_dir / "candidate_events_frozen.csv",
        frozen,
    )

    freeze_metrics = {
        "freeze_version": FREEZE_VERSION,
        "freeze_id": freeze_id,
        "source_run_id": run_id,
        "created_at": now_iso(),
        "freeze_status": (
            "FROZEN"
            if validation["validation_status"] == "PASSED"
            else "FAILED"
        ),
        "source_candidate_count": len(candidates),
        "frozen_candidate_count": len(frozen),
        "removed_candidate_count": (
            len(candidates) - len(frozen)
        ),
        "review_patch_count": len(patches),
        "coverage_gap_count": len(coverage),
        "coverage_gap_status": (
            "ACCEPTED_DISCLOSURE_LIMITATION"
        ),
        "company_candidate_counts": (
            validation["company_candidate_counts"]
        ),
        "printed_page_coverage_rate": (
            validation["printed_page_coverage_rate"]
        ),
        "validation_error_count": len(
            validation["errors"]
        ),
        "source_metrics_sha256": sha256_file(
            input_paths["metrics"]
        ),
        "source_candidates_sha256": sha256_file(
            input_paths["candidates"]
        ),
        "human_signoff_status": "NOT_CLAIMED",
        "next_stage": (
            "STRUCTURED_EVENT_FIELD_EXTRACTION"
            if validation["validation_status"] == "PASSED"
            else "STOP_AND_REVIEW"
        ),
    }
    write_json(
        output_dir / "freeze_metrics.json",
        freeze_metrics,
    )
    write_json(
        output_dir / "freeze_validation.json",
        validation,
    )
    write_json(
        (
            repo_root
            / "review"
            / "candidate_events"
            / "latest_frozen.json"
        ),
        {
            "freeze_id": freeze_id,
            "source_run_id": run_id,
            "freeze_status": freeze_metrics[
                "freeze_status"
            ],
            "output_dir": str(output_dir),
            "created_at": now_iso(),
        },
    )

    readme = f"""# 候选事件冻结结果

- Freeze ID: `{freeze_id}`
- Source Run ID: `{run_id}`
- 公共代码版本：`candidate_event_generation_v0.11`
- 冻结候选数：`{len(frozen)}`
- 覆盖缺口：`{len(coverage)}`
- 正文页码覆盖率：`{validation['printed_page_coverage_rate']:.2%}`
- 状态：`{freeze_metrics['freeze_status']}`

## 冻结规则

1. 三联锻造保留2004年6月18日详细设立事件，删除时间线恢复重复项。
2. 黄山谷捷保留2012年6月发行人前身设立，排除被吸收方昆山谷捷2009年设立。
3. 友升股份保留1992年12月4日详细设立事件，删除1992年10月22日批复日期恢复项。
4. 三联锻造和影石创新的覆盖缺口按“招股书未逐项披露”保留，不编造单事件字段。
5. 当前结果为AI辅助证据复核，不声称人工Gold或人工签字。

## 下一阶段

以 `candidate_events_frozen.jsonl` 和
`candidate_evidence_frozen.jsonl` 作为结构化事件字段抽取的唯一业务输入。
"""
    (output_dir / "README.md").write_text(
        readme,
        encoding="utf-8",
    )

    print()
    print("候选事件Review Patch冻结完成")
    print(f"Freeze ID：{freeze_id}")
    print(f"源候选：{len(candidates)}")
    print(f"冻结候选：{len(frozen)}")
    print(
        f"删除重复/错误候选："
        f"{len(candidates) - len(frozen)}"
    )
    print(f"覆盖缺口：{len(coverage)}")
    print(
        "正文页码覆盖率："
        f"{validation['printed_page_coverage_rate']:.2%}"
    )
    print(
        f"验证错误：{len(validation['errors'])}"
    )
    print(
        f"冻结状态：{freeze_metrics['freeze_status']}"
    )
    print(f"输出目录：{output_dir}")

    return (
        0
        if validation["validation_status"] == "PASSED"
        else 2
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "对candidate_event_generation_v0.11结果"
            "应用Review Patch并冻结26条候选事件"
        )
    )
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path("."),
    )
    parser.add_argument(
        "--run-id",
        required=True,
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    return run_freeze(
        repo_root=args.repo_root,
        run_id=args.run_id,
    )


if __name__ == "__main__":
    raise SystemExit(main())
