from __future__ import annotations

import argparse
import csv
import hashlib
import json
import shutil
import sys
import traceback
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


PIPELINE_VERSION = "final_submission_assembly_v0.1"

COMPANIES = {
    "001282": "三联锻造",
    "301581": "黄山谷捷",
    "603418": "友升股份",
    "688758": "赛分科技",
    "688775": "影石创新",
    "920100": "三协电机",
    "920116": "星图测控",
}


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(
        timespec="seconds"
    )


def make_build_id() -> str:
    return datetime.now().astimezone().strftime(
        "FINALASSEMBLY_V01_%Y%m%d_%H%M%S"
    )


def read_json(path: Path) -> Any:
    return json.loads(
        path.read_text(encoding="utf-8-sig")
    )


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    with path.open(
        "r",
        encoding="utf-8-sig",
    ) as handle:
        for line_number, line in enumerate(
            handle,
            start=1,
        ):
            if not line.strip():
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError as exc:
                raise ValueError(
                    f"JSONL解析失败：{path} "
                    f"第{line_number}行：{exc}"
                ) from exc
    return rows


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            payload,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def write_jsonl(
    path: Path,
    rows: Iterable[dict[str, Any]],
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open(
        "w",
        encoding="utf-8",
        newline="\n",
    ) as handle:
        for row in rows:
            handle.write(
                json.dumps(
                    row,
                    ensure_ascii=False,
                )
                + "\n"
            )


def write_csv(
    path: Path,
    rows: list[dict[str, Any]],
    headers: list[str],
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
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
        for row in rows:
            output: dict[str, Any] = {}
            for key in headers:
                value = row.get(key)
                output[key] = (
                    json.dumps(
                        value,
                        ensure_ascii=False,
                    )
                    if isinstance(
                        value,
                        (list, dict),
                    )
                    else value
                )
            writer.writerow(output)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def latest_run_id(
    repo_root: Path,
    category: str,
) -> str:
    path = (
        repo_root
        / "logs"
        / category
        / "latest_run.json"
    )
    if not path.is_file():
        raise FileNotFoundError(
            f"缺少最新运行记录：{path}"
        )
    return str(read_json(path)["run_id"])


def resolve_freeze(
    repo_root: Path,
) -> tuple[str, Path]:
    latest = (
        repo_root
        / "review"
        / "candidate_events"
        / "latest_frozen.json"
    )
    if not latest.is_file():
        raise FileNotFoundError(
            "缺少候选事件冻结记录"
        )
    payload = read_json(latest)
    freeze_id = str(payload["freeze_id"])
    output_dir = payload.get("output_dir")
    if output_dir:
        path = Path(str(output_dir))
        if not path.is_absolute():
            path = repo_root / path
    else:
        path = (
            repo_root
            / "review"
            / "candidate_events"
            / "frozen"
            / freeze_id
        )
    return freeze_id, path.resolve()


def load_source_paths(
    repo_root: Path,
    structured_run_id: str,
    numeric_run_id: str,
    pevc_run_id: str,
    freeze_dir: Path,
) -> dict[str, Path]:
    structured_auto = (
        repo_root
        / "auto_output"
        / "structured_events"
        / "runs"
        / structured_run_id
    )
    structured_validation = (
        repo_root
        / "validation"
        / "structured_events"
        / "runs"
        / structured_run_id
    )
    numeric_dir = (
        repo_root
        / "validation"
        / "numeric_validation"
        / "runs"
        / numeric_run_id
    )
    pevc_auto = (
        repo_root
        / "auto_output"
        / "pevc_paths"
        / "runs"
        / pevc_run_id
    )
    pevc_validation = (
        repo_root
        / "validation"
        / "pevc_paths"
        / "runs"
        / pevc_run_id
    )

    return {
        "events": (
            structured_auto
            / "event_records_auto.jsonl"
        ),
        "transactions": (
            structured_auto
            / "transaction_records_auto.jsonl"
        ),
        "parties": (
            structured_auto
            / "event_parties_auto.jsonl"
        ),
        "numeric_facts": (
            structured_auto
            / "numeric_facts_auto.jsonl"
        ),
        "dates": (
            structured_auto
            / "event_dates_auto.jsonl"
        ),
        "evidence": (
            structured_auto
            / "event_evidence_index.jsonl"
        ),
        "structured_metrics": (
            structured_validation
            / "structured_event_metrics.json"
        ),
        "structured_reviews": (
            structured_validation
            / "structured_event_review_queue.jsonl"
        ),
        "numeric_results": (
            numeric_dir
            / "numeric_validation_results.jsonl"
        ),
        "calculated_facts": (
            numeric_dir
            / "calculated_numeric_facts.jsonl"
        ),
        "numeric_metrics": (
            numeric_dir
            / "numeric_validation_metrics.json"
        ),
        "numeric_reviews": (
            numeric_dir
            / "numeric_validation_review_queue.jsonl"
        ),
        "quarantined_numeric": (
            numeric_dir
            / "quarantined_numeric_facts.jsonl"
        ),
        "investor_entities": (
            pevc_auto
            / "investor_entities_auto.jsonl"
        ),
        "investment_paths": (
            pevc_auto
            / "investment_paths_auto.jsonl"
        ),
        "pevc_entities": (
            pevc_auto
            / "pevc_entities_auto.jsonl"
        ),
        "pevc_paths": (
            pevc_auto
            / "pevc_investment_paths_auto.jsonl"
        ),
        "non_pevc_investors": (
            pevc_auto
            / "non_pevc_observed_investors.jsonl"
        ),
        "pevc_metrics": (
            pevc_validation
            / "pevc_path_metrics.json"
        ),
        "pevc_reviews": (
            pevc_validation
            / "pevc_review_queue.jsonl"
        ),
        "candidate_reviews": (
            freeze_dir
            / "review_history.jsonl"
        ),
        "coverage_gaps": (
            freeze_dir
            / "coverage_gaps_frozen.jsonl"
        ),
        "freeze_metrics": (
            freeze_dir
            / "freeze_metrics.json"
        ),
    }


def ensure_paths_exist(
    paths: dict[str, Path],
) -> None:
    required = {
        key: path
        for key, path in paths.items()
        if key != "quarantined_numeric"
    }
    missing = [
        str(path)
        for path in required.values()
        if not path.is_file()
    ]
    if missing:
        raise FileNotFoundError(
            "缺少最终汇总输入：\n"
            + "\n".join(missing)
        )


def validate_source_metrics(
    paths: dict[str, Path],
) -> dict[str, Any]:
    structured = read_json(
        paths["structured_metrics"]
    )
    numeric = read_json(
        paths["numeric_metrics"]
    )
    pevc = read_json(
        paths["pevc_metrics"]
    )
    freeze = read_json(
        paths["freeze_metrics"]
    )

    errors: list[str] = []
    if freeze.get("freeze_status") != "FROZEN":
        errors.append("候选事件未冻结")
    if freeze.get("frozen_candidate_count") != 26:
        errors.append("冻结候选数不是26")
    if structured.get("event_count") != 26:
        errors.append("结构化事件数不是26")
    if structured.get("transaction_count") != 27:
        errors.append("交易数不是27")
    if structured.get("validation_error_count") != 0:
        errors.append("结构化抽取存在验证错误")
    if numeric.get("failed_count") != 0:
        errors.append("数值校验存在失败项")
    if pevc.get("validation_status") != "PASSED":
        errors.append("PE/VC验证未通过")
    if pevc.get("issuer_self_candidate_count") != 0:
        errors.append("PE/VC结果包含发行人自身")
    if pevc.get("discarded_noise_count") != 0:
        errors.append("PE/VC结果仍有噪声")
    if pevc.get("pevc_candidate_count", 0) < 1:
        errors.append("未识别到PE/VC候选")
    if pevc.get(
        "pevc_investment_path_count",
        0,
    ) < 1:
        errors.append("未识别到PE/VC投资路径")

    return {
        "validation_status": (
            "FAILED" if errors else "PASSED"
        ),
        "errors": errors,
        "freeze_metrics": freeze,
        "structured_metrics": structured,
        "numeric_metrics": numeric,
        "pevc_metrics": pevc,
    }


def rows_by_company(
    rows: list[dict[str, Any]],
) -> dict[str, list[dict[str, Any]]]:
    output: dict[
        str,
        list[dict[str, Any]],
    ] = defaultdict(list)
    for row in rows:
        company_id = str(
            row.get("company_id") or ""
        )
        if company_id:
            output[company_id].append(row)
    return output


def build_equity_snapshot_status(
    company_id: str,
    events: list[dict[str, Any]],
    parties: list[dict[str, Any]],
    numeric_facts: list[dict[str, Any]],
) -> dict[str, Any]:
    ratio_facts = [
        row
        for row in numeric_facts
        if row.get("fact_type")
        == "EQUITY_RATIO"
        and row.get("normalized_value")
        is not None
    ]
    return {
        "company_id": company_id,
        "snapshot_status": (
            "DISCLOSED_FACTS_AVAILABLE_BUT_"
            "PARTY_RATIO_LINKAGE_NOT_CONFIRMED"
            if ratio_facts
            else "NOT_CONSTRUCTIBLE_FROM_"
            "CURRENT_DISCLOSED_FACTS"
        ),
        "snapshot_count": 0,
        "snapshots": [],
        "available_equity_ratio_fact_ids": [
            row["numeric_fact_id"]
            for row in ratio_facts
        ],
        "event_count_considered": len(events),
        "party_count_considered": len(parties),
        "value_policy": (
            "不将未建立主体对应关系的比例事实"
            "强行组成股权结构快照"
        ),
        "manual_review_required": bool(
            ratio_facts
        ),
    }


def normalize_review_row(
    row: dict[str, Any],
    source_stage: str,
) -> dict[str, Any]:
    output = dict(row)
    output["source_stage"] = source_stage
    output.setdefault(
        "manual_status",
        "PENDING",
    )
    return output


def write_company_package(
    company_dir: Path,
    company_id: str,
    company_name: str,
    datasets: dict[str, list[dict[str, Any]]],
    run_ids: dict[str, str],
    build_id: str,
) -> dict[str, Any]:
    company_dir.mkdir(
        parents=True,
        exist_ok=False,
    )

    events = datasets["events"]
    transactions = datasets["transactions"]
    parties = datasets["parties"]
    numeric_facts = datasets["numeric_facts"]
    dates = datasets["dates"]
    evidence = datasets["evidence"]
    numeric_results = datasets["numeric_results"]
    calculated_facts = datasets["calculated_facts"]
    investor_entities = datasets["investor_entities"]
    investment_paths = datasets["investment_paths"]
    pevc_entities = datasets["pevc_entities"]
    pevc_paths = datasets["pevc_paths"]
    non_pevc = datasets["non_pevc"]
    reviews = datasets["reviews"]
    coverage_gaps = datasets["coverage_gaps"]

    snapshot_payload = build_equity_snapshot_status(
        company_id,
        events,
        parties,
        numeric_facts,
    )

    write_json(
        company_dir / "events_final.json",
        {
            "company_id": company_id,
            "company_name": company_name,
            "record_count": len(events),
            "records": events,
        },
    )
    write_json(
        company_dir / "transactions_final.json",
        {
            "company_id": company_id,
            "company_name": company_name,
            "record_count": len(transactions),
            "records": transactions,
            "parties": parties,
            "dates": dates,
            "disclosed_numeric_facts": (
                numeric_facts
            ),
        },
    )
    write_json(
        company_dir
        / "equity_snapshots_final.json",
        snapshot_payload,
    )
    write_json(
        company_dir
        / "numeric_validation_final.json",
        {
            "company_id": company_id,
            "company_name": company_name,
            "validation_record_count": len(
                numeric_results
            ),
            "calculated_fact_count": len(
                calculated_facts
            ),
            "validation_records": (
                numeric_results
            ),
            "calculated_facts": (
                calculated_facts
            ),
        },
    )
    write_json(
        company_dir / "pevc_paths_final.json",
        {
            "company_id": company_id,
            "company_name": company_name,
            "investor_entity_count": len(
                investor_entities
            ),
            "investment_path_count": len(
                investment_paths
            ),
            "pevc_entity_count": len(
                pevc_entities
            ),
            "pevc_path_count": len(
                pevc_paths
            ),
            "investor_entities": (
                investor_entities
            ),
            "investment_paths": (
                investment_paths
            ),
            "pevc_entities": pevc_entities,
            "pevc_paths": pevc_paths,
            "non_pevc_observed_investors": (
                non_pevc
            ),
        },
    )
    write_jsonl(
        company_dir / "evidence_index.jsonl",
        evidence,
    )
    write_jsonl(
        company_dir / "review_history.jsonl",
        reviews,
    )

    metrics = {
        "company_id": company_id,
        "company_name": company_name,
        "event_count": len(events),
        "transaction_count": len(
            transactions
        ),
        "party_count": len(parties),
        "numeric_fact_count": len(
            numeric_facts
        ),
        "numeric_validation_count": len(
            numeric_results
        ),
        "calculated_fact_count": len(
            calculated_facts
        ),
        "investor_entity_count": len(
            investor_entities
        ),
        "investment_path_count": len(
            investment_paths
        ),
        "pevc_entity_count": len(
            pevc_entities
        ),
        "pevc_path_count": len(
            pevc_paths
        ),
        "coverage_gap_count": len(
            coverage_gaps
        ),
        "review_record_count": len(reviews),
        "equity_snapshot_status": (
            snapshot_payload[
                "snapshot_status"
            ]
        ),
        "build_status": "FINAL_JSON_READY",
    }
    write_json(
        company_dir / "metrics.json",
        metrics,
    )

    manifest = {
        "build_id": build_id,
        "pipeline_version": (
            PIPELINE_VERSION
        ),
        "company_id": company_id,
        "company_name": company_name,
        "created_at": now_iso(),
        "source_run_ids": run_ids,
        "files": {},
    }
    for path in sorted(
        company_dir.iterdir()
    ):
        if path.is_file():
            manifest["files"][path.name] = {
                "sha256": sha256_file(path),
                "size_bytes": (
                    path.stat().st_size
                ),
            }
    write_json(
        company_dir / "run_manifest.json",
        manifest,
    )
    return metrics


def write_combined_tables(
    output_dir: Path,
    datasets: dict[str, list[dict[str, Any]]],
) -> None:
    combined_dir = output_dir / "combined"
    combined_dir.mkdir(
        parents=True,
        exist_ok=False,
    )

    write_csv(
        combined_dir / "all_events.csv",
        datasets["events"],
        [
            "event_id",
            "candidate_event_id",
            "company_id",
            "company_short_name",
            "event_type",
            "event_group",
            "event_period",
            "event_title",
            "pdf_page_start",
            "pdf_page_end",
            "printed_page_start",
            "printed_page_end",
            "primary_evidence_id",
            "extraction_status",
            "review_required",
            "review_reasons",
        ],
    )
    write_csv(
        combined_dir
        / "all_transactions.csv",
        datasets["transactions"],
        [
            "transaction_id",
            "event_id",
            "candidate_event_id",
            "company_id",
            "transaction_type",
            "transaction_date",
            "transaction_date_role",
            "transferor_party_ids",
            "transferee_party_ids",
            "investor_party_ids",
            "absorbed_party_ids",
            "registered_capital_before_fact_ids",
            "registered_capital_after_fact_ids",
            "capital_increase_fact_ids",
            "consideration_fact_ids",
            "share_quantity_fact_ids",
            "share_price_fact_ids",
            "equity_ratio_fact_ids",
            "transaction_status",
            "review_required",
            "review_reasons",
        ],
    )
    write_csv(
        combined_dir
        / "all_numeric_validations.csv",
        datasets["numeric_results"],
        [
            "validation_id",
            "company_id",
            "event_id",
            "transaction_id",
            "rule_id",
            "rule_name",
            "validation_status",
            "disclosed_input_fact_ids",
            "input_values",
            "calculation_formula",
            "calculated_value",
            "disclosed_comparison_value",
            "difference_abs",
            "difference_pct",
            "evidence_ids",
            "review_required",
            "review_reason",
        ],
    )
    write_csv(
        combined_dir / "all_pevc_paths.csv",
        datasets["pevc_paths"],
        [
            "investment_path_id",
            "company_id",
            "investor_entity_id",
            "investor_name_normalized",
            "event_id",
            "transaction_id",
            "entry_method",
            "investment_level",
            "direct_or_indirect",
            "transaction_date",
            "evidence_ids",
            "path_status",
            "confidence",
            "review_required",
            "review_reasons",
        ],
    )
    write_csv(
        combined_dir / "all_reviews.csv",
        datasets["reviews"],
        [
            "source_stage",
            "review_id",
            "company_id",
            "record_type",
            "record_id",
            "event_id",
            "transaction_id",
            "review_reason",
            "manual_status",
            "manual_decision",
            "manual_note",
        ],
    )


def run_assembly(
    repo_root: Path,
    structured_run_id: str | None,
    numeric_run_id: str | None,
    pevc_run_id: str | None,
) -> int:
    repo_root = (
        repo_root.expanduser().resolve()
    )

    freeze_id, freeze_dir = resolve_freeze(
        repo_root
    )
    structured_id = (
        structured_run_id
        or latest_run_id(
            repo_root,
            "structured_events",
        )
    )
    numeric_id = (
        numeric_run_id
        or latest_run_id(
            repo_root,
            "numeric_validation",
        )
    )
    pevc_id = (
        pevc_run_id
        or latest_run_id(
            repo_root,
            "pevc_paths",
        )
    )

    paths = load_source_paths(
        repo_root,
        structured_id,
        numeric_id,
        pevc_id,
        freeze_dir,
    )
    ensure_paths_exist(paths)
    source_validation = (
        validate_source_metrics(paths)
    )
    if (
        source_validation[
            "validation_status"
        ]
        != "PASSED"
    ):
        raise ValueError(
            "最终汇总输入未通过："
            + "；".join(
                source_validation["errors"]
            )
        )

    rows = {
        "events": read_jsonl(
            paths["events"]
        ),
        "transactions": read_jsonl(
            paths["transactions"]
        ),
        "parties": read_jsonl(
            paths["parties"]
        ),
        "numeric_facts": read_jsonl(
            paths["numeric_facts"]
        ),
        "dates": read_jsonl(
            paths["dates"]
        ),
        "evidence": read_jsonl(
            paths["evidence"]
        ),
        "numeric_results": read_jsonl(
            paths["numeric_results"]
        ),
        "calculated_facts": read_jsonl(
            paths["calculated_facts"]
        ),
        "investor_entities": read_jsonl(
            paths["investor_entities"]
        ),
        "investment_paths": read_jsonl(
            paths["investment_paths"]
        ),
        "pevc_entities": read_jsonl(
            paths["pevc_entities"]
        ),
        "pevc_paths": read_jsonl(
            paths["pevc_paths"]
        ),
        "non_pevc": read_jsonl(
            paths["non_pevc_investors"]
        ),
        "coverage_gaps": read_jsonl(
            paths["coverage_gaps"]
        ),
    }

    reviews: list[dict[str, Any]] = []
    for key, stage in (
        (
            "candidate_reviews",
            "CANDIDATE_FREEZE",
        ),
        (
            "structured_reviews",
            "STRUCTURED_EVENTS",
        ),
        (
            "numeric_reviews",
            "NUMERIC_VALIDATION",
        ),
        (
            "pevc_reviews",
            "PEVC_PATHS",
        ),
    ):
        for row in read_jsonl(paths[key]):
            reviews.append(
                normalize_review_row(
                    row,
                    stage,
                )
            )
    rows["reviews"] = reviews

    by_company = {
        key: rows_by_company(value)
        for key, value in rows.items()
        if key not in {"reviews"}
    }
    reviews_by_company = rows_by_company(
        reviews
    )

    build_id = make_build_id()
    output_dir = (
        repo_root
        / "final"
        / "seven_companies"
        / build_id
    )
    output_dir.mkdir(
        parents=True,
        exist_ok=False,
    )
    company_root = (
        output_dir / "companies"
    )
    company_root.mkdir()

    run_ids = {
        "candidate_freeze_id": freeze_id,
        "structured_run_id": structured_id,
        "numeric_run_id": numeric_id,
        "pevc_run_id": pevc_id,
    }

    company_metrics: list[
        dict[str, Any]
    ] = []
    for company_id, company_name in (
        COMPANIES.items()
    ):
        datasets = {
            key: mapping.get(
                company_id,
                [],
            )
            for key, mapping in (
                by_company.items()
            )
        }
        datasets["reviews"] = (
            reviews_by_company.get(
                company_id,
                [],
            )
        )
        metrics = write_company_package(
            company_root / company_id,
            company_id,
            company_name,
            datasets,
            run_ids,
            build_id,
        )
        company_metrics.append(metrics)

    write_combined_tables(
        output_dir,
        rows,
    )

    final_metrics = {
        "build_id": build_id,
        "pipeline_version": (
            PIPELINE_VERSION
        ),
        "created_at": now_iso(),
        "build_status": "FINAL_JSON_READY",
        "company_count": len(COMPANIES),
        "event_count": len(rows["events"]),
        "transaction_count": len(
            rows["transactions"]
        ),
        "numeric_validation_count": len(
            rows["numeric_results"]
        ),
        "calculated_fact_count": len(
            rows["calculated_facts"]
        ),
        "investor_entity_count": len(
            rows["investor_entities"]
        ),
        "investment_path_count": len(
            rows["investment_paths"]
        ),
        "pevc_entity_count": len(
            rows["pevc_entities"]
        ),
        "pevc_path_count": len(
            rows["pevc_paths"]
        ),
        "coverage_gap_count": len(
            rows["coverage_gaps"]
        ),
        "review_record_count": len(
            rows["reviews"]
        ),
        "company_metrics": company_metrics,
        "source_validation": (
            source_validation
        ),
        "source_run_ids": run_ids,
        "xlsx_status": (
            "PENDING_WORKBOOK_GENERATION"
        ),
    }
    write_json(
        output_dir / "final_metrics.json",
        final_metrics,
    )

    final_validation = {
        "validation_status": "PASSED",
        "errors": [],
        "checks": {
            "company_count": (
                len(company_metrics) == 7
            ),
            "event_count": (
                len(rows["events"]) == 26
            ),
            "transaction_count": (
                len(rows["transactions"])
                == 27
            ),
            "numeric_failure_count": (
                source_validation[
                    "numeric_metrics"
                ].get("failed_count")
                == 0
            ),
            "pevc_validation_passed": (
                source_validation[
                    "pevc_metrics"
                ].get(
                    "validation_status"
                )
                == "PASSED"
            ),
            "pevc_candidate_count": (
                len(rows["pevc_entities"])
                >= 1
            ),
            "pevc_path_count": (
                len(rows["pevc_paths"])
                >= 1
            ),
        },
        "disclosure_policy": (
            "未披露或无法建立主体对应关系的"
            "信息不补写、不计算为确定快照"
        ),
    }
    if not all(
        final_validation["checks"].values()
    ):
        final_validation[
            "validation_status"
        ] = "FAILED"
        final_validation["errors"].append(
            "最终汇总数量或上游状态不符合验收标准"
        )
    write_json(
        output_dir
        / "final_validation.json",
        final_validation,
    )

    inventory = {
        "build_id": build_id,
        "root": str(output_dir),
        "files": {},
    }
    for path in sorted(
        output_dir.rglob("*")
    ):
        if path.is_file():
            relative = str(
                path.relative_to(output_dir)
            )
            inventory["files"][relative] = {
                "sha256": sha256_file(path),
                "size_bytes": (
                    path.stat().st_size
                ),
            }
    write_json(
        output_dir
        / "submission_inventory.json",
        inventory,
    )

    report = f"""# 七家公司最终数据汇总报告

## 构建信息

- Build ID: `{build_id}`
- 状态：`{final_metrics['build_status']}`
- 公司数：7
- 事件数：{final_metrics['event_count']}
- 交易数：{final_metrics['transaction_count']}
- PE/VC候选主体：{final_metrics['pevc_entity_count']}
- PE/VC投资路径：{final_metrics['pevc_path_count']}
- 数值校验失败：0

## 数据边界

1. 所有事件、交易、数值、投资路径均保留事件ID、交易ID和证据ID。
2. 原文披露值和计算值分开保存。
3. 两条披露覆盖缺口继续保留，不拆分为招股书未披露的单事件。
4. 两条缺少可确认投资方的交易继续保留复核，不制造投资者。
5. 股权比例事实无法与具体股东可靠对应时，不强行生成确定股权结构快照。
6. 当前已达到`FINAL_JSON_READY`；Excel工作簿在下一步由本汇总目录直接生成。
"""
    (
        output_dir
        / "final_report.md"
    ).write_text(
        report,
        encoding="utf-8",
    )

    package_name = (
        f"seven_companies_final_json_ready_"
        f"{build_id}.zip"
    )
    package_path = (
        repo_root
        / "final"
        / package_name
    )
    if package_path.exists():
        package_path.unlink()
    with zipfile.ZipFile(
        package_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archive:
        for path in sorted(
            output_dir.rglob("*")
        ):
            if path.is_file():
                archive.write(
                    path,
                    Path(build_id)
                    / path.relative_to(
                        output_dir
                    ),
                )

    latest = {
        "build_id": build_id,
        "build_status": (
            final_metrics[
                "build_status"
            ]
        ),
        "output_dir": str(output_dir),
        "package_path": str(
            package_path
        ),
        "created_at": now_iso(),
    }
    write_json(
        repo_root
        / "final"
        / "latest_final_assembly.json",
        latest,
    )

    print()
    print("七家公司最终数据汇总 v0.1 完成")
    print(f"Build ID：{build_id}")
    print("公司：7")
    print(
        f"事件：{final_metrics['event_count']}"
    )
    print(
        f"交易："
        f"{final_metrics['transaction_count']}"
    )
    print(
        f"PE/VC主体："
        f"{final_metrics['pevc_entity_count']}"
    )
    print(
        f"PE/VC路径："
        f"{final_metrics['pevc_path_count']}"
    )
    print(
        "数值校验失败："
        f"{source_validation['numeric_metrics'].get('failed_count')}"
    )
    print(
        f"最终验证："
        f"{final_validation['validation_status']}"
    )
    print(
        f"构建状态："
        f"{final_metrics['build_status']}"
    )
    print(f"输出目录：{output_dir}")
    print(f"汇总压缩包：{package_path}")

    return (
        0
        if final_validation[
            "validation_status"
        ]
        == "PASSED"
        else 2
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "汇总七家公司冻结事件、交易、数值校验、"
            "PE/VC路径、证据和复核历史，生成FINAL_JSON_READY提交包"
        )
    )
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path("."),
    )
    parser.add_argument(
        "--structured-run-id",
        default=None,
    )
    parser.add_argument(
        "--numeric-run-id",
        default=None,
    )
    parser.add_argument(
        "--pevc-run-id",
        default=None,
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        return run_assembly(
            repo_root=args.repo_root,
            structured_run_id=(
                args.structured_run_id
            ),
            numeric_run_id=(
                args.numeric_run_id
            ),
            pevc_run_id=(
                args.pevc_run_id
            ),
        )
    except Exception as exc:
        print(
            f"[ERROR] "
            f"{type(exc).__name__}: {exc}",
            file=sys.stderr,
        )
        traceback.print_exc()
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
