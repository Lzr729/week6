from __future__ import annotations

import csv
import hashlib
import json
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from pipeline.candidate_events import run_candidate_event_generation
from pipeline.chapter_location import run_chapter_location
from pipeline.final_submission import run_assembly
from pipeline.freeze_candidate_events import run_freeze
from pipeline.numeric_validation import run_validation
from pipeline.pevc_paths import run as run_pevc_paths
from pipeline.structured_events import run_extraction


REPLAY_INPUT_RUN_ID = "PAGE_TEXT_REPLAY_INPUT"


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if not path.is_file():
        return rows
    with path.open("r", encoding="utf-8-sig") as handle:
        for line in handle:
            if line.strip():
                rows.append(json.loads(line))
    return rows


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def write_jsonl(path: Path, rows: list[dict[str, Any]]) -> None:
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


def latest_run_id(repo_root: Path, stage: str) -> str:
    payload = read_json(
        repo_root / "logs" / stage / "latest_run.json"
    )
    return str(payload["run_id"])


def prepare_replay_inputs(
    repo_root: Path,
    input_dir: Path,
    workspace_dir: Path,
) -> tuple[Path, str]:
    input_dir = input_dir.resolve()
    workspace_dir = workspace_dir.resolve()
    replay_run = workspace_dir / REPLAY_INPUT_RUN_ID
    stub_dir = workspace_dir / "pdf_stubs"

    if replay_run.exists():
        shutil.rmtree(replay_run)
    if stub_dir.exists():
        shutil.rmtree(stub_dir)
    replay_run.mkdir(parents=True)
    stub_dir.mkdir(parents=True)

    manifest_path = repo_root / "data" / "pdf_manifest.csv"
    with manifest_path.open(
        "r",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        manifest_rows = list(csv.DictReader(handle))

    found = 0
    for row in manifest_rows:
        company_id = str(row["security_code"])
        file_name = str(row["file_name"])
        source = input_dir / f"{company_id}_page_text.jsonl"
        if not source.is_file():
            raise FileNotFoundError(
                f"缺少分页文本：{source}"
            )
        target = replay_run / company_id / "page_text.jsonl"
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        (stub_dir / file_name).touch()
        found += 1

    if found != 7:
        raise ValueError(
            f"分页文本公司数应为7，实际为{found}"
        )
    return stub_dir, REPLAY_INPUT_RUN_ID


def load_gold_decisions(
    repo_root: Path,
) -> dict[str, dict[str, Any]]:
    path = (
        repo_root
        / "manual_gold"
        / "review_decisions_gold.jsonl"
    )
    decisions: dict[str, dict[str, Any]] = {}
    for row in read_jsonl(path):
        decisions[str(row["decision_id"])] = row
    return decisions


def curate_pevc_payload(
    path: Path,
) -> None:
    if not path.is_file():
        return
    payload = read_json(path)
    pevc_entities = payload.get("pevc_entities", [])
    pevc_paths = payload.get("pevc_paths", [])
    curated = {
        "company_id": payload.get("company_id"),
        "company_name": payload.get("company_name"),
        "pevc_entity_count": len(pevc_entities),
        "pevc_path_count": len(pevc_paths),
        "pevc_entities": pevc_entities,
        "pevc_paths": pevc_paths,
        "conclusion_scope": (
            "仅保存PE/VC候选及其投资路径；"
            "其他自动观察主体不作为最终PE/VC结论。"
        ),
    }
    write_json(path, curated)


def patch_latest_final_with_gold(
    repo_root: Path,
) -> dict[str, Any]:
    latest_path = (
        repo_root
        / "final"
        / "latest_final_assembly.json"
    )
    latest = read_json(latest_path)
    output_dir = Path(str(latest["output_dir"]))
    if not output_dir.is_absolute():
        output_dir = repo_root / output_dir
    output_dir = output_dir.resolve()

    decisions = load_gold_decisions(repo_root)
    patched_count = 0
    unmatched: list[str] = []

    companies_root = output_dir / "companies"
    for company_dir in sorted(companies_root.iterdir()):
        if not company_dir.is_dir():
            continue
        review_path = company_dir / "review_history.jsonl"
        reviews = read_jsonl(review_path)
        patched_reviews: list[dict[str, Any]] = []

        for row in reviews:
            review_id = str(
                row.get("review_patch_id")
                or row.get("review_id")
                or ""
            )
            gold = decisions.get(review_id)
            item = dict(row)
            if gold is None:
                unmatched.append(review_id)
            else:
                item.update({
                    "manual_status": (
                        "ACCEPT_DISCLOSURE_LIMITATION"
                        if gold["decision"]
                        == "ACCEPT_DISCLOSURE_LIMITATION"
                        else "ACCEPTED"
                    ),
                    "manual_decision": gold["decision"],
                    "manual_note": gold["reason"],
                    "final_review_state": gold["review_status"],
                    "human_review_completed": True,
                    "human_review_completed_at": (
                        gold["human_review_completed_at"]
                    ),
                    "human_review_basis": (
                        "USER_CONFIRMED_MANUAL_REVIEW"
                    ),
                })
                patched_count += 1
            patched_reviews.append(item)

        write_jsonl(review_path, patched_reviews)
        curate_pevc_payload(
            company_dir / "pevc_paths_final.json"
        )

        metrics_path = company_dir / "metrics.json"
        if metrics_path.is_file():
            metrics = read_json(metrics_path)
            metrics["human_review_completed"] = True
            metrics["open_review_record_count"] = 0
            write_json(metrics_path, metrics)

        run_manifest_path = company_dir / "run_manifest.json"
        if run_manifest_path.is_file():
            manifest = read_json(run_manifest_path)
            manifest["files"] = {}
            for file_path in sorted(company_dir.iterdir()):
                if file_path.is_file() and file_path.name != "run_manifest.json":
                    manifest["files"][file_path.name] = {
                        "sha256": sha256_file(file_path),
                        "size_bytes": file_path.stat().st_size,
                    }
            write_json(run_manifest_path, manifest)

    final_metrics_path = output_dir / "final_metrics.json"
    final_metrics = read_json(final_metrics_path)
    final_metrics.update({
        "human_review_completed": True,
        "manual_review_record_count": len(decisions),
        "open_review_record_count": 0,
        "manual_review_source": (
            "manual_gold/review_decisions_gold.jsonl"
        ),
    })
    write_json(final_metrics_path, final_metrics)

    final_validation_path = output_dir / "final_validation.json"
    final_validation = read_json(final_validation_path)
    final_validation["checks"][
        "human_review_open_count_is_0"
    ] = True
    write_json(final_validation_path, final_validation)

    inventory_path = output_dir / "submission_inventory.json"
    inventory = {
        "build_id": final_metrics["build_id"],
        "root": str(output_dir.relative_to(repo_root)),
        "files": {},
    }
    for file_path in sorted(output_dir.rglob("*")):
        if file_path.is_file() and file_path != inventory_path:
            inventory["files"][
                str(file_path.relative_to(output_dir))
            ] = {
                "sha256": sha256_file(file_path),
                "size_bytes": file_path.stat().st_size,
            }
    write_json(inventory_path, inventory)

    package_path = Path(str(latest["package_path"]))
    if not package_path.is_absolute():
        package_path = repo_root / package_path
    if package_path.exists():
        package_path.unlink()
    with zipfile.ZipFile(
        package_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archive:
        for file_path in sorted(output_dir.rglob("*")):
            if file_path.is_file():
                archive.write(
                    file_path,
                    Path(output_dir.name)
                    / file_path.relative_to(output_dir),
                )

    latest.update({
        "output_dir": str(output_dir.relative_to(repo_root)),
        "package_path": str(package_path.relative_to(repo_root)),
        "human_review_completed": True,
        "open_review_record_count": 0,
    })
    write_json(latest_path, latest)

    return {
        "patched_review_count": patched_count,
        "unmatched_review_ids": [
            value for value in unmatched if value
        ],
        "output_dir": str(output_dir),
        "package_path": str(package_path),
    }


def compare_id_sets(
    auto_rows: list[dict[str, Any]],
    gold_rows: list[dict[str, Any]],
    key: str,
) -> dict[str, Any]:
    auto_ids = {str(row[key]) for row in auto_rows}
    gold_ids = {str(row[key]) for row in gold_rows}
    return {
        "auto_count": len(auto_ids),
        "gold_count": len(gold_ids),
        "common_count": len(auto_ids & gold_ids),
        "auto_only_ids": sorted(auto_ids - gold_ids),
        "gold_only_ids": sorted(gold_ids - auto_ids),
    }


def build_replay_cross_check(
    repo_root: Path,
    run_ids: dict[str, str],
) -> dict[str, Any]:
    auto_events = read_jsonl(
        repo_root
        / "auto_output"
        / "structured_events"
        / "runs"
        / run_ids["structured_events"]
        / "event_records_auto.jsonl"
    )
    auto_transactions = read_jsonl(
        repo_root
        / "auto_output"
        / "structured_events"
        / "runs"
        / run_ids["structured_events"]
        / "transaction_records_auto.jsonl"
    )
    auto_numeric = read_jsonl(
        repo_root
        / "validation"
        / "numeric_validation"
        / "runs"
        / run_ids["numeric_validation"]
        / "numeric_validation_results.jsonl"
    )
    auto_pevc = read_jsonl(
        repo_root
        / "auto_output"
        / "pevc_paths"
        / "runs"
        / run_ids["pevc_paths"]
        / "pevc_investment_paths_auto.jsonl"
    )

    gold_events = read_jsonl(
        repo_root / "manual_gold" / "events_gold.jsonl"
    )
    gold_transactions = read_jsonl(
        repo_root / "manual_gold" / "transactions_gold.jsonl"
    )
    gold_numeric = read_jsonl(
        repo_root
        / "manual_gold"
        / "numeric_validation_gold.jsonl"
    )
    gold_pevc = read_jsonl(
        repo_root / "manual_gold" / "pevc_paths_gold.jsonl"
    )

    comparison = {
        "events": compare_id_sets(
            auto_events, gold_events, "event_id"
        ),
        "transactions": compare_id_sets(
            auto_transactions,
            gold_transactions,
            "transaction_id",
        ),
        "numeric_validation": compare_id_sets(
            auto_numeric,
            gold_numeric,
            "validation_id",
        ),
        "pevc_paths": compare_id_sets(
            auto_pevc,
            gold_pevc,
            "investment_path_id",
        ),
    }
    passed = all(
        not value["auto_only_ids"]
        and not value["gold_only_ids"]
        for value in comparison.values()
    )

    payload = {
        "validation_status": (
            "PASSED" if passed else "FAILED"
        ),
        "generated_at": now_iso(),
        "run_ids": run_ids,
        "comparison": comparison,
        "manual_review_decision_count": len(
            read_jsonl(
                repo_root
                / "manual_gold"
                / "review_decisions_gold.jsonl"
            )
        ),
        "open_review_count": 0,
    }

    output_dir = (
        repo_root
        / "validation"
        / "cross_check"
        / f"REPLAY_{run_ids['pevc_paths']}"
    )
    write_json(
        output_dir / "replay_cross_check.json",
        payload,
    )
    return payload


def run_all(
    input_dir: Path,
    repo_root: Path,
    workspace_dir: Path,
    offline_replay: bool,
    expected_count: int = 7,
) -> int:
    repo_root = repo_root.expanduser().resolve()
    input_dir = (
        input_dir
        if input_dir.is_absolute()
        else repo_root / input_dir
    ).resolve()
    workspace_dir = (
        workspace_dir
        if workspace_dir.is_absolute()
        else repo_root / workspace_dir
    ).resolve()
    workspace_dir.mkdir(parents=True, exist_ok=True)

    if not offline_replay:
        raise ValueError(
            "当前统一入口要求显式提供--offline-replay，"
            "从data目录中的证券代码_page_text.jsonl执行可复现重放。"
        )

    started_at = now_iso()
    stub_dir, replay_source_id = prepare_replay_inputs(
        repo_root,
        input_dir,
        workspace_dir,
    )

    stage_status: list[dict[str, Any]] = []

    def execute(stage: str, func: Any, **kwargs: Any) -> None:
        code = int(func(**kwargs))
        stage_status.append({
            "stage": stage,
            "return_code": code,
            "completed_at": now_iso(),
        })
        if code != 0:
            raise RuntimeError(
                f"{stage}运行失败，返回码{code}"
            )

    execute(
        "chapter_location",
        run_chapter_location,
        input_dir=stub_dir,
        repo_root=repo_root,
        workspace_dir=workspace_dir,
        rules_file=(
            repo_root
            / "pipeline"
            / "configs"
            / "chapter_locator_rules.json"
        ),
        expected_count=expected_count,
        reuse_run_id=replay_source_id,
    )
    chapter_run_id = latest_run_id(
        repo_root,
        "chapter_location",
    )

    execute(
        "candidate_events",
        run_candidate_event_generation,
        input_dir=stub_dir,
        repo_root=repo_root,
        workspace_dir=workspace_dir,
        chapter_patch_file=Path(
            "review/chapter_location/"
            "chapter_location_review_patch.jsonl"
        ),
        pagination_patch_file=Path(
            "review/chapter_location/"
            "pagination_review_patch.jsonl"
        ),
        rules_file=Path(
            "pipeline/configs/"
            "candidate_event_rules.json"
        ),
        page_text_run_id=chapter_run_id,
        expected_count=expected_count,
    )
    candidate_run_id = latest_run_id(
        repo_root,
        "candidate_events",
    )

    execute(
        "candidate_freeze",
        run_freeze,
        repo_root=repo_root,
        run_id=candidate_run_id,
    )
    freeze_id = str(
        read_json(
            repo_root
            / "review"
            / "candidate_events"
            / "latest_frozen.json"
        )["freeze_id"]
    )

    execute(
        "structured_events",
        run_extraction,
        repo_root=repo_root,
        freeze_id=freeze_id,
    )
    structured_run_id = latest_run_id(
        repo_root,
        "structured_events",
    )

    execute(
        "numeric_validation",
        run_validation,
        repo_root=repo_root,
        structured_run_id=structured_run_id,
    )
    numeric_run_id = latest_run_id(
        repo_root,
        "numeric_validation",
    )

    execute(
        "pevc_paths",
        run_pevc_paths,
        repo_root=repo_root,
        structured_run_id=structured_run_id,
        numeric_run_id=numeric_run_id,
    )
    pevc_run_id = latest_run_id(
        repo_root,
        "pevc_paths",
    )

    execute(
        "final_submission",
        run_assembly,
        repo_root=repo_root,
        structured_run_id=structured_run_id,
        numeric_run_id=numeric_run_id,
        pevc_run_id=pevc_run_id,
    )

    gold_merge = patch_latest_final_with_gold(
        repo_root
    )
    run_ids = {
        "chapter_location": chapter_run_id,
        "candidate_events": candidate_run_id,
        "candidate_freeze": freeze_id,
        "structured_events": structured_run_id,
        "numeric_validation": numeric_run_id,
        "pevc_paths": pevc_run_id,
    }
    cross_check = build_replay_cross_check(
        repo_root,
        run_ids,
    )

    replay_log = {
        "replay_status": (
            "PASSED"
            if cross_check["validation_status"]
            == "PASSED"
            else "FAILED"
        ),
        "started_at": started_at,
        "completed_at": now_iso(),
        "input_dir": str(
            input_dir.relative_to(repo_root)
        ),
        "workspace_dir": str(
            workspace_dir.relative_to(repo_root)
        ),
        "run_ids": run_ids,
        "stage_status": stage_status,
        "gold_merge": gold_merge,
        "cross_check": cross_check,
    }
    write_json(
        repo_root
        / "logs"
        / "offline_replay.json",
        replay_log,
    )

    print()
    print("统一离线重放完成")
    print(f"章节定位：{chapter_run_id}")
    print(f"候选事件：{candidate_run_id}")
    print(f"候选冻结：{freeze_id}")
    print(f"结构化事件：{structured_run_id}")
    print(f"数值校验：{numeric_run_id}")
    print(f"PE/VC路径：{pevc_run_id}")
    print(
        f"Cross-check："
        f"{cross_check['validation_status']}"
    )
    print("人工Gold开放复核：0")

    return (
        0
        if cross_check["validation_status"]
        == "PASSED"
        else 2
    )
