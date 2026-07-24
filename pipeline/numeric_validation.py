from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import sys
import traceback
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


PIPELINE_VERSION = "transaction_numeric_validation_v0.1.2"


@dataclass
class ValidationResult:
    validation_id: str
    company_id: str
    event_id: str | None
    transaction_id: str | None
    rule_id: str
    rule_name: str
    validation_status: str
    disclosed_input_fact_ids: list[str]
    input_values: dict[str, Any]
    calculation_formula: str | None
    calculated_value: float | None
    disclosed_comparison_value: float | None
    difference_abs: float | None
    difference_pct: float | None
    tolerance_abs: float | None
    tolerance_pct: float | None
    evidence_ids: list[str]
    review_required: bool
    review_reason: str | None


@dataclass
class CalculatedFact:
    calculated_fact_id: str
    company_id: str
    event_id: str
    transaction_id: str
    fact_type: str
    formula: str
    input_fact_ids: list[str]
    input_values: dict[str, float]
    calculated_value: float
    unit: str | None
    currency: str | None
    value_type: str
    comparison_fact_id: str | None
    comparison_disclosed_value: float | None
    difference_abs: float | None
    difference_pct: float | None
    validation_status: str


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def make_run_id() -> str:
    return datetime.now().astimezone().strftime(
        "NUMVALID_V01_%Y%m%d_%H%M%S"
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


def write_csv(
    path: Path,
    rows: list[dict[str, Any]],
    headers: list[str],
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=headers,
            extrasaction="ignore",
        )
        writer.writeheader()
        for row in rows:
            output = {}
            for key in headers:
                value = row.get(key)
                if isinstance(value, (list, dict)):
                    output[key] = json.dumps(
                        value,
                        ensure_ascii=False,
                    )
                else:
                    output[key] = value
            writer.writerow(output)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def resolve_structured_run(
    repo_root: Path,
    structured_run_id: str | None,
) -> tuple[str, Path, Path]:
    if structured_run_id:
        auto_dir = (
            repo_root
            / "auto_output"
            / "structured_events"
            / "runs"
            / structured_run_id
        )
        validation_dir = (
            repo_root
            / "validation"
            / "structured_events"
            / "runs"
            / structured_run_id
        )
        return structured_run_id, auto_dir, validation_dir

    latest_path = (
        repo_root
        / "logs"
        / "structured_events"
        / "latest_run.json"
    )
    if not latest_path.is_file():
        raise FileNotFoundError(
            "未找到logs/structured_events/latest_run.json"
        )
    latest = read_json(latest_path)
    run_id = str(latest["run_id"])
    auto_dir = (
        repo_root
        / "auto_output"
        / "structured_events"
        / "runs"
        / run_id
    )
    validation_dir = (
        repo_root
        / "validation"
        / "structured_events"
        / "runs"
        / run_id
    )
    return run_id, auto_dir, validation_dir


def safe_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        result = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(result):
        return None
    return result


def relative_difference(
    calculated: float,
    disclosed: float,
) -> float | None:
    if disclosed == 0:
        return None
    return abs(calculated - disclosed) / abs(disclosed)


def within_tolerance(
    calculated: float,
    disclosed: float,
    tolerance_abs: float,
    tolerance_pct: float,
) -> tuple[bool, float, float | None]:
    difference_abs = abs(calculated - disclosed)
    difference_pct = relative_difference(
        calculated,
        disclosed,
    )
    passed = (
        difference_abs <= tolerance_abs
        or (
            difference_pct is not None
            and difference_pct <= tolerance_pct
        )
    )
    return passed, difference_abs, difference_pct


def evidence_ids_for_facts(
    facts: list[dict[str, Any]],
) -> list[str]:
    output: list[str] = []
    for fact in facts:
        value = str(fact.get("source_evidence_id") or "")
        if value and value not in output:
            output.append(value)
    return output


def facts_by_id(
    facts: list[dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    return {
        str(item["numeric_fact_id"]): item
        for item in facts
    }


def selected_facts(
    transaction: dict[str, Any],
    field_name: str,
    numeric_by_id: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for fact_id in transaction.get(field_name) or []:
        fact = numeric_by_id.get(str(fact_id))
        if fact is not None:
            output.append(fact)
    return output


def normalized_values(
    facts: list[dict[str, Any]],
) -> list[tuple[str, float]]:
    output: list[tuple[str, float]] = []
    for fact in facts:
        value = safe_float(fact.get("normalized_value"))
        if value is not None:
            output.append((
                str(fact["numeric_fact_id"]),
                value,
            ))
    return output


def best_single_value(
    facts: list[dict[str, Any]],
) -> tuple[str, float] | None:
    values = normalized_values(facts)
    if not values:
        return None
    return values[0]


def make_result(
    *,
    sequence: int,
    company_id: str,
    event_id: str | None,
    transaction_id: str | None,
    rule_id: str,
    rule_name: str,
    status: str,
    input_facts: list[dict[str, Any]],
    input_values: dict[str, Any],
    formula: str | None = None,
    calculated_value: float | None = None,
    disclosed_value: float | None = None,
    difference_abs: float | None = None,
    difference_pct: float | None = None,
    tolerance_abs: float | None = None,
    tolerance_pct: float | None = None,
    review_reason: str | None = None,
) -> ValidationResult:
    return ValidationResult(
        validation_id=f"VAL-{company_id}-{sequence:05d}",
        company_id=company_id,
        event_id=event_id,
        transaction_id=transaction_id,
        rule_id=rule_id,
        rule_name=rule_name,
        validation_status=status,
        disclosed_input_fact_ids=[
            str(item["numeric_fact_id"])
            for item in input_facts
        ],
        input_values=input_values,
        calculation_formula=formula,
        calculated_value=calculated_value,
        disclosed_comparison_value=disclosed_value,
        difference_abs=difference_abs,
        difference_pct=difference_pct,
        tolerance_abs=tolerance_abs,
        tolerance_pct=tolerance_pct,
        evidence_ids=evidence_ids_for_facts(
            input_facts
        ),
        review_required=status in {
            "FAILED",
            "REVIEW_REQUIRED",
        },
        review_reason=review_reason,
    )


def validate_capital_increase(
    transaction: dict[str, Any],
    numeric_by_id: dict[str, dict[str, Any]],
    sequence_start: int,
    tolerance_abs: float,
    tolerance_pct: float,
) -> tuple[list[ValidationResult], list[CalculatedFact]]:
    results: list[ValidationResult] = []
    calculated_facts: list[CalculatedFact] = []
    company_id = str(transaction["company_id"])
    event_id = str(transaction["event_id"])
    transaction_id = str(transaction["transaction_id"])

    before_facts = selected_facts(
        transaction,
        "registered_capital_before_fact_ids",
        numeric_by_id,
    )
    increase_facts = selected_facts(
        transaction,
        "capital_increase_fact_ids",
        numeric_by_id,
    )
    after_facts = selected_facts(
        transaction,
        "registered_capital_after_fact_ids",
        numeric_by_id,
    )
    before = best_single_value(before_facts)
    increase = best_single_value(increase_facts)
    after = best_single_value(after_facts)
    all_facts = (
        before_facts + increase_facts + after_facts
    )

    if not (before and increase and after):
        results.append(make_result(
            sequence=sequence_start,
            company_id=company_id,
            event_id=event_id,
            transaction_id=transaction_id,
            rule_id="NV-CAP-001",
            rule_name="增资前注册资本+增资额=增资后注册资本",
            status="NOT_APPLICABLE",
            input_facts=all_facts,
            input_values={
                "registered_capital_before": (
                    before[1] if before else None
                ),
                "capital_increase_amount": (
                    increase[1] if increase else None
                ),
                "registered_capital_after": (
                    after[1] if after else None
                ),
            },
            review_reason=(
                "原文未同时披露增资前注册资本、增资额和增资后注册资本"
            ),
        ))
        return results, calculated_facts

    calculated = before[1] + increase[1]
    passed, difference_abs, difference_pct = within_tolerance(
        calculated,
        after[1],
        tolerance_abs,
        tolerance_pct,
    )
    status = "PASSED" if passed else "FAILED"
    review_reason = (
        None
        if passed
        else "增资前注册资本加增资额与披露的增资后注册资本不一致"
    )
    results.append(make_result(
        sequence=sequence_start,
        company_id=company_id,
        event_id=event_id,
        transaction_id=transaction_id,
        rule_id="NV-CAP-001",
        rule_name="增资前注册资本+增资额=增资后注册资本",
        status=status,
        input_facts=all_facts,
        input_values={
            "registered_capital_before": before[1],
            "capital_increase_amount": increase[1],
            "registered_capital_after": after[1],
        },
        formula="registered_capital_before + capital_increase_amount",
        calculated_value=calculated,
        disclosed_value=after[1],
        difference_abs=difference_abs,
        difference_pct=difference_pct,
        tolerance_abs=tolerance_abs,
        tolerance_pct=tolerance_pct,
        review_reason=review_reason,
    ))
    calculated_facts.append(CalculatedFact(
        calculated_fact_id=f"CAL-{company_id}-{sequence_start:05d}",
        company_id=company_id,
        event_id=event_id,
        transaction_id=transaction_id,
        fact_type="CALCULATED_REGISTERED_CAPITAL_AFTER",
        formula="registered_capital_before + capital_increase_amount",
        input_fact_ids=[before[0], increase[0]],
        input_values={
            "registered_capital_before": before[1],
            "capital_increase_amount": increase[1],
        },
        calculated_value=calculated,
        unit="元",
        currency="CNY",
        value_type="CALCULATED",
        comparison_fact_id=after[0],
        comparison_disclosed_value=after[1],
        difference_abs=difference_abs,
        difference_pct=difference_pct,
        validation_status=status,
    ))
    return results, calculated_facts


def validate_price_quantity_consideration(
    transaction: dict[str, Any],
    numeric_by_id: dict[str, dict[str, Any]],
    sequence_start: int,
    tolerance_abs: float,
    tolerance_pct: float,
) -> tuple[list[ValidationResult], list[CalculatedFact]]:
    company_id = str(transaction["company_id"])
    event_id = str(transaction["event_id"])
    transaction_id = str(transaction["transaction_id"])
    quantity_facts = selected_facts(
        transaction,
        "share_quantity_fact_ids",
        numeric_by_id,
    )
    price_facts = selected_facts(
        transaction,
        "share_price_fact_ids",
        numeric_by_id,
    )
    consideration_facts = selected_facts(
        transaction,
        "consideration_fact_ids",
        numeric_by_id,
    )
    quantity = best_single_value(quantity_facts)
    price = best_single_value(price_facts)
    consideration = best_single_value(
        consideration_facts
    )
    all_facts = (
        quantity_facts + price_facts + consideration_facts
    )

    if not (quantity and price and consideration):
        return [make_result(
            sequence=sequence_start,
            company_id=company_id,
            event_id=event_id,
            transaction_id=transaction_id,
            rule_id="NV-PRICE-001",
            rule_name="股份数量×每股价格=交易对价",
            status="NOT_APPLICABLE",
            input_facts=all_facts,
            input_values={
                "share_quantity": (
                    quantity[1] if quantity else None
                ),
                "share_price": (
                    price[1] if price else None
                ),
                "consideration": (
                    consideration[1]
                    if consideration
                    else None
                ),
            },
            review_reason=(
                "原文未同时披露股份数量、每股价格和交易对价"
            ),
        )], []

    calculated = quantity[1] * price[1]
    passed, difference_abs, difference_pct = within_tolerance(
        calculated,
        consideration[1],
        tolerance_abs,
        tolerance_pct,
    )
    status = "PASSED" if passed else "FAILED"
    result = make_result(
        sequence=sequence_start,
        company_id=company_id,
        event_id=event_id,
        transaction_id=transaction_id,
        rule_id="NV-PRICE-001",
        rule_name="股份数量×每股价格=交易对价",
        status=status,
        input_facts=all_facts,
        input_values={
            "share_quantity": quantity[1],
            "share_price": price[1],
            "consideration": consideration[1],
        },
        formula="share_quantity * share_price",
        calculated_value=calculated,
        disclosed_value=consideration[1],
        difference_abs=difference_abs,
        difference_pct=difference_pct,
        tolerance_abs=tolerance_abs,
        tolerance_pct=tolerance_pct,
        review_reason=(
            None
            if passed
            else "股份数量乘每股价格与披露对价不一致"
        ),
    )
    calculated_fact = CalculatedFact(
        calculated_fact_id=f"CAL-{company_id}-{sequence_start:05d}",
        company_id=company_id,
        event_id=event_id,
        transaction_id=transaction_id,
        fact_type="CALCULATED_CONSIDERATION",
        formula="share_quantity * share_price",
        input_fact_ids=[quantity[0], price[0]],
        input_values={
            "share_quantity": quantity[1],
            "share_price": price[1],
        },
        calculated_value=calculated,
        unit="元",
        currency="CNY",
        value_type="CALCULATED",
        comparison_fact_id=consideration[0],
        comparison_disclosed_value=consideration[1],
        difference_abs=difference_abs,
        difference_pct=difference_pct,
        validation_status=status,
    )
    return [result], [calculated_fact]


def validate_equity_ratios(
    numeric_facts: list[dict[str, Any]],
    sequence_start: int,
) -> list[ValidationResult]:
    results: list[ValidationResult] = []
    sequence = sequence_start
    for fact in numeric_facts:
        if fact.get("fact_type") != "EQUITY_RATIO":
            continue
        company_id = str(fact["company_id"])
        value = safe_float(fact.get("normalized_value"))
        if value is None:
            status = "NOT_APPLICABLE"
            reason = (
                "无效的自动持股比例事实已隔离；"
                "不作为披露数值参与校验"
            )
        elif 0 <= value <= 1:
            status = "PASSED"
            reason = None
        else:
            status = "FAILED"
            reason = "持股比例超出0%至100%范围"
        results.append(make_result(
            sequence=sequence,
            company_id=company_id,
            event_id=str(fact["event_id"]),
            transaction_id=None,
            rule_id="NV-RATIO-001",
            rule_name="持股比例范围校验",
            status=status,
            input_facts=[fact],
            input_values={
                "equity_ratio_normalized": value,
            },
            review_reason=reason,
        ))
        sequence += 1
    return results


def validate_duplicate_numeric_facts(
    numeric_facts: list[dict[str, Any]],
    sequence_start: int,
) -> list[ValidationResult]:
    groups: dict[
        tuple[str, str, float | None, str],
        list[dict[str, Any]],
    ] = defaultdict(list)
    for fact in numeric_facts:
        normalized_value = safe_float(
            fact.get("normalized_value")
        )
        if normalized_value is None:
            continue
        key = (
            str(fact["event_id"]),
            str(fact["fact_type"]),
            normalized_value,
            str(fact.get("source_evidence_id") or ""),
        )
        groups[key].append(fact)

    results: list[ValidationResult] = []
    sequence = sequence_start
    for key, facts in groups.items():
        if len(facts) <= 1:
            continue
        company_id = str(facts[0]["company_id"])
        results.append(make_result(
            sequence=sequence,
            company_id=company_id,
            event_id=str(facts[0]["event_id"]),
            transaction_id=None,
            rule_id="NV-DUP-001",
            rule_name="同证据同类型同数值重复事实校验",
            status="REVIEW_REQUIRED",
            input_facts=facts,
            input_values={
                "duplicate_count": len(facts),
                "fact_type": key[1],
                "normalized_value": key[2],
            },
            review_reason="同一证据中生成重复数值事实",
        ))
        sequence += 1
    return results


def parse_sortable_date(value: str | None) -> tuple[int, int, int] | None:
    if not value:
        return None
    parts = str(value).split("-")
    try:
        year = int(parts[0])
        month = int(parts[1]) if len(parts) > 1 else 0
        day = int(parts[2]) if len(parts) > 2 else 0
    except (ValueError, IndexError):
        return None
    return year, month, day


def validate_company_chronology(
    events: list[dict[str, Any]],
    sequence_start: int,
) -> list[ValidationResult]:
    by_company: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for event in events:
        by_company[str(event["company_id"])].append(event)

    results: list[ValidationResult] = []
    sequence = sequence_start
    for company_id, company_events in sorted(by_company.items()):
        dated = [
            (
                parse_sortable_date(
                    event.get("event_period")
                ),
                event,
            )
            for event in company_events
        ]
        dated = [
            item for item in dated
            if item[0] is not None
        ]
        page_order = sorted(
            dated,
            key=lambda item: (
                int(item[1].get("pdf_page_start") or 0),
                str(item[1]["event_id"]),
            ),
        )
        dates_in_page_order = [
            item[0] for item in page_order
        ]
        monotonic = all(
            dates_in_page_order[index]
            <= dates_in_page_order[index + 1]
            for index in range(
                len(dates_in_page_order) - 1
            )
        )
        results.append(make_result(
            sequence=sequence,
            company_id=company_id,
            event_id=None,
            transaction_id=None,
            rule_id="NV-DATE-001",
            rule_name="公司事件时间顺序校验",
            status=(
                "PASSED"
                if monotonic
                else "INFORMATIONAL"
            ),
            input_facts=[],
            input_values={
                "event_order": [
                    {
                        "event_id": item[1]["event_id"],
                        "event_period": item[1].get(
                            "event_period"
                        ),
                        "pdf_page_start": item[1].get(
                            "pdf_page_start"
                        ),
                    }
                    for item in page_order
                ]
            },
            review_reason=(
                None
                if monotonic
                else (
                    "PDF页面顺序与事件日期顺序不一致；"
                    "该结果仅记录时间线与详细正文并存，不进入人工复核"
                )
            ),
        ))
        sequence += 1
    return results


def validate_transaction_linkage(
    transactions: list[dict[str, Any]],
    parties: list[dict[str, Any]],
    sequence_start: int,
) -> list[ValidationResult]:
    party_by_id = {
        str(item["party_id"]): item
        for item in parties
    }
    results: list[ValidationResult] = []
    sequence = sequence_start

    for transaction in transactions:
        company_id = str(transaction["company_id"])
        transaction_type = str(
            transaction["transaction_type"]
        )
        required_fields: list[str] = []
        if transaction_type == "EQUITY_TRANSFER":
            required_fields = [
                "transferor_party_ids",
                "transferee_party_ids",
            ]
        elif transaction_type == "ABSORPTION_MERGER":
            required_fields = ["absorbed_party_ids"]

        if not required_fields:
            continue

        missing = [
            field
            for field in required_fields
            if not transaction.get(field)
        ]
        referenced_ids = [
            str(party_id)
            for field in required_fields
            for party_id in (
                transaction.get(field) or []
            )
        ]
        nonexistent = [
            party_id
            for party_id in referenced_ids
            if party_id not in party_by_id
        ]
        if missing or nonexistent:
            status = "REVIEW_REQUIRED"
            reason = (
                f"缺失参与方字段：{missing}；"
                f"不存在的参与方ID：{nonexistent}"
            )
        else:
            status = "PASSED"
            reason = None

        results.append(make_result(
            sequence=sequence,
            company_id=company_id,
            event_id=str(transaction["event_id"]),
            transaction_id=str(
                transaction["transaction_id"]
            ),
            rule_id="NV-LINK-001",
            rule_name="交易参与方完整性校验",
            status=status,
            input_facts=[],
            input_values={
                "transaction_type": transaction_type,
                "required_fields": required_fields,
                "referenced_party_ids": referenced_ids,
            },
            review_reason=reason,
        ))
        sequence += 1

    return results


def invalid_numeric_facts(
    numeric_facts: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for fact in numeric_facts:
        if (
            fact.get("fact_type") == "EQUITY_RATIO"
            and safe_float(
                fact.get("normalized_value")
            ) is None
        ):
            output.append({
                "numeric_fact_id": fact["numeric_fact_id"],
                "company_id": fact["company_id"],
                "event_id": fact["event_id"],
                "fact_type": fact["fact_type"],
                "raw_value_text": fact.get(
                    "raw_value_text"
                ),
                "source_evidence_id": fact.get(
                    "source_evidence_id"
                ),
                "quarantine_reason": (
                    "持股比例自动抽取结果无法转换为数值；"
                    "不参与校验，不进入人工复核"
                ),
                "record_status": (
                    "QUARANTINED_INVALID_AUTO_FACT"
                ),
            })
    return output


def validate_source(
    metrics: dict[str, Any],
    events: list[dict[str, Any]],
    transactions: list[dict[str, Any]],
) -> list[str]:
    errors: list[str] = []
    if metrics.get("event_count") != 26:
        errors.append("源结构化事件数不是26")
    if metrics.get("transaction_count") != 27:
        errors.append("源交易记录数不是27")
    if metrics.get("validation_error_count") != 0:
        errors.append("源结构化抽取存在验证错误")
    if len(events) != 26:
        errors.append(
            f"事件文件记录数不是26：{len(events)}"
        )
    if len(transactions) != 27:
        errors.append(
            f"交易文件记录数不是27：{len(transactions)}"
        )
    return errors


def run_validation(
    repo_root: Path,
    structured_run_id: str | None,
) -> int:
    repo_root = repo_root.expanduser().resolve()
    (
        source_run_id,
        auto_dir,
        source_validation_dir,
    ) = resolve_structured_run(
        repo_root,
        structured_run_id,
    )

    paths = {
        "events": auto_dir / "event_records_auto.jsonl",
        "transactions": auto_dir / "transaction_records_auto.jsonl",
        "parties": auto_dir / "event_parties_auto.jsonl",
        "numerics": auto_dir / "numeric_facts_auto.jsonl",
        "dates": auto_dir / "event_dates_auto.jsonl",
        "evidence": auto_dir / "event_evidence_index.jsonl",
        "metrics": (
            source_validation_dir
            / "structured_event_metrics.json"
        ),
    }
    missing = [
        str(path)
        for path in paths.values()
        if not path.is_file()
    ]
    if missing:
        raise FileNotFoundError(
            "缺少结构化抽取输入：\n" + "\n".join(missing)
        )

    events = read_jsonl(paths["events"])
    transactions = read_jsonl(paths["transactions"])
    parties = read_jsonl(paths["parties"])
    numerics = read_jsonl(paths["numerics"])
    dates = read_jsonl(paths["dates"])
    evidence = read_jsonl(paths["evidence"])
    source_metrics = read_json(paths["metrics"])

    source_errors = validate_source(
        source_metrics,
        events,
        transactions,
    )
    if source_errors:
        raise ValueError("；".join(source_errors))

    quarantined_facts = invalid_numeric_facts(
        numerics
    )

    config_path = (
        repo_root
        / "pipeline"
        / "configs"
        / "numeric_validation_rules.json"
    )
    config = (
        read_json(config_path)
        if config_path.is_file()
        else {
            "capital_tolerance_abs": 1.0,
            "capital_tolerance_pct": 0.0001,
            "consideration_tolerance_abs": 1.0,
            "consideration_tolerance_pct": 0.005,
        }
    )
    capital_tolerance_abs = float(
        config["capital_tolerance_abs"]
    )
    capital_tolerance_pct = float(
        config["capital_tolerance_pct"]
    )
    consideration_tolerance_abs = float(
        config["consideration_tolerance_abs"]
    )
    consideration_tolerance_pct = float(
        config["consideration_tolerance_pct"]
    )

    run_id = make_run_id()
    output_dir = (
        repo_root
        / "validation"
        / "numeric_validation"
        / "runs"
        / run_id
    )
    logs_dir = (
        repo_root
        / "logs"
        / "numeric_validation"
        / "runs"
        / run_id
    )
    output_dir.mkdir(parents=True, exist_ok=False)
    logs_dir.mkdir(parents=True, exist_ok=False)

    numeric_by_id = facts_by_id(numerics)
    results: list[ValidationResult] = []
    calculated_facts: list[CalculatedFact] = []
    sequence = 1

    for transaction in transactions:
        transaction_type = str(
            transaction["transaction_type"]
        )
        if transaction_type in {
            "CAPITAL_INCREASE",
            "DIRECTIONAL_FINANCING",
        }:
            new_results, new_calculated = (
                validate_capital_increase(
                    transaction,
                    numeric_by_id,
                    sequence,
                    capital_tolerance_abs,
                    capital_tolerance_pct,
                )
            )
            results.extend(new_results)
            calculated_facts.extend(new_calculated)
            sequence += len(new_results)

        if transaction_type in {
            "EQUITY_TRANSFER",
            "CAPITAL_INCREASE",
            "DIRECTIONAL_FINANCING",
        }:
            new_results, new_calculated = (
                validate_price_quantity_consideration(
                    transaction,
                    numeric_by_id,
                    sequence,
                    consideration_tolerance_abs,
                    consideration_tolerance_pct,
                )
            )
            results.extend(new_results)
            calculated_facts.extend(new_calculated)
            sequence += len(new_results)

    ratio_results = validate_equity_ratios(
        numerics,
        sequence,
    )
    results.extend(ratio_results)
    sequence += len(ratio_results)

    duplicate_results = validate_duplicate_numeric_facts(
        numerics,
        sequence,
    )
    results.extend(duplicate_results)
    sequence += len(duplicate_results)

    chronology_results = validate_company_chronology(
        events,
        sequence,
    )
    results.extend(chronology_results)
    sequence += len(chronology_results)

    linkage_results = validate_transaction_linkage(
        transactions,
        parties,
        sequence,
    )
    results.extend(linkage_results)

    review_items = [
        asdict(item)
        for item in results
        if item.review_required
    ]
    failure_count = sum(
        item.validation_status == "FAILED"
        for item in results
    )
    review_count = sum(
        item.validation_status == "REVIEW_REQUIRED"
        for item in results
    )
    pass_count = sum(
        item.validation_status == "PASSED"
        for item in results
    )
    not_applicable_count = sum(
        item.validation_status == "NOT_APPLICABLE"
        for item in results
    )
    informational_count = sum(
        item.validation_status == "INFORMATIONAL"
        for item in results
    )

    batch_status = (
        "FAILED"
        if failure_count
        else (
            "READY_WITH_REVIEW"
            if review_items
            else "READY"
        )
    )

    metrics = {
        "metrics_version": "0.1.2",
        "run_id": run_id,
        "pipeline_version": PIPELINE_VERSION,
        "source_structured_run_id": source_run_id,
        "batch_status": batch_status,
        "validation_result_count": len(results),
        "passed_count": pass_count,
        "failed_count": failure_count,
        "review_required_count": review_count,
        "not_applicable_count": not_applicable_count,
        "informational_count": informational_count,
        "calculated_fact_count": len(calculated_facts),
        "review_queue_count": len(review_items),
        "quarantined_invalid_fact_count": len(
            quarantined_facts
        ),
        "rule_status_counts": {
            rule_id: dict(
                Counter(
                    item.validation_status
                    for item in results
                    if item.rule_id == rule_id
                )
            )
            for rule_id in sorted({
                item.rule_id
                for item in results
            })
        },
        "source_event_count": len(events),
        "source_transaction_count": len(transactions),
        "source_numeric_fact_count": len(numerics),
        "source_party_count": len(parties),
        "llm_called": False,
        "pevc_classification_performed": False,
        "final_generated": False,
        "note": (
            "v0.1.2将无法数值化的比例事实隔离，"
            "不再把空比例和其重复项送入人工复核；"
            "PDF页面顺序校验改为信息提示。"
        ),
    }

    write_jsonl(
        output_dir / "numeric_validation_results.jsonl",
        [asdict(item) for item in results],
    )
    write_jsonl(
        output_dir / "calculated_numeric_facts.jsonl",
        [asdict(item) for item in calculated_facts],
    )
    write_jsonl(
        output_dir / "numeric_validation_review_queue.jsonl",
        review_items,
    )
    write_jsonl(
        output_dir / "quarantined_numeric_facts.jsonl",
        quarantined_facts,
    )
    write_csv(
        output_dir / "numeric_validation_review_queue.csv",
        review_items,
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
            "review_reason",
        ],
    )
    write_csv(
        output_dir / "numeric_validation_results.csv",
        [asdict(item) for item in results],
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
            "tolerance_abs",
            "tolerance_pct",
            "evidence_ids",
            "review_required",
            "review_reason",
        ],
    )
    write_json(
        output_dir / "numeric_validation_metrics.json",
        metrics,
    )
    write_json(
        output_dir / "numeric_validation_summary.json",
        {
            "run_id": run_id,
            "source_structured_run_id": source_run_id,
            "batch_status": batch_status,
            "rule_summary": {
                rule_id: dict(
                    Counter(
                        item.validation_status
                        for item in results
                        if item.rule_id == rule_id
                    )
                )
                for rule_id in sorted({
                    item.rule_id
                    for item in results
                })
            },
        },
    )
    write_json(
        logs_dir / "run_manifest.json",
        {
            "run_id": run_id,
            "pipeline_version": PIPELINE_VERSION,
            "source_structured_run_id": source_run_id,
            "completed_at": now_iso(),
            "batch_status": batch_status,
            "source_files": {
                key: {
                    "path": str(path),
                    "sha256": sha256_file(path),
                }
                for key, path in paths.items()
            },
            "llm_called": False,
        },
    )
    write_json(
        (
            repo_root
            / "logs"
            / "numeric_validation"
            / "latest_run.json"
        ),
        {
            "run_id": run_id,
            "source_structured_run_id": source_run_id,
            "batch_status": batch_status,
            "completed_at": now_iso(),
        },
    )

    print()
    print("交易与数值校验 v0.1.2 完成")
    print(f"运行ID：{run_id}")
    print(f"源结构化运行：{source_run_id}")
    print(f"校验记录：{len(results)}")
    print(f"通过：{pass_count}")
    print(f"失败：{failure_count}")
    print(f"待复核：{review_count}")
    print(f"不适用/信息不足：{not_applicable_count}")
    print(f"信息提示：{informational_count}")
    print(
        f"隔离无效数值事实：{len(quarantined_facts)}"
    )
    print(f"计算值：{len(calculated_facts)}")
    print(f"复核队列：{len(review_items)}")
    print(f"批次状态：{batch_status}")

    return 2 if failure_count else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "对26条事件和27条交易执行注册资本、"
            "价格对价、持股比例、时间顺序和参与方链接校验。"
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
        help="可选；不提供时读取最新结构化运行",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        return run_validation(
            repo_root=args.repo_root,
            structured_run_id=args.structured_run_id,
        )
    except Exception as exc:
        print(
            f"[ERROR] {type(exc).__name__}: {exc}",
            file=sys.stderr,
        )
        traceback.print_exc()
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
