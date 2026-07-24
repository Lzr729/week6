from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import traceback
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


PIPELINE_VERSION = "pevc_path_identification_v0.3.2"

ELIGIBLE_TRANSACTION_TYPES = {
    "CAPITAL_INCREASE",
    "DIRECTIONAL_FINANCING",
    "EQUITY_TRANSFER",
}

STRONG_PEVC_TERMS = (
    "股权投资基金",
    "创业投资基金",
    "私募投资基金",
    "私募基金",
    "产业投资基金",
    "并购基金",
    "投资基金",
    "创业投资",
)

MEDIUM_PEVC_TERMS = (
    "股权投资",
    "创业投资",
    "创投",
    "投资管理",
    "资产管理",
    "资本管理",
    "投资合伙",
)

LEGAL_ENTITY_SUFFIXES = (
    "合伙企业（有限合伙）",
    "合伙企业(有限合伙)",
    "有限合伙企业",
    "股份有限公司",
    "有限责任公司",
    "有限公司",
    "有限合伙",
    "投资基金",
    "基金",
)

NOISE_TERMS = (
    "序号",
    "发行对象名称",
    "认购对象名称",
    "股东名称",
    "投资者名称",
    "数量（股",
    "数量(股",
    "金额（元",
    "金额(元",
    "认购数量",
    "认购金额",
    "新增股本",
    "新增股本的",
    "包括优先",
    "每股价格",
    "50元的价格",
    "合计",
    "比例",
    "持股数量",
    "名称",
)

NON_INVESTOR_TERMS = (
    "会计师事务所",
    "律师事务所",
    "资产评估事务所",
    "资产评估有限公司",
    "市场监督管理局",
    "工商行政管理局",
    "证券登记结算",
    "银行",
)


@dataclass
class InvestorEntity:
    investor_entity_id: str
    company_id: str
    investor_name_raw: str
    investor_name_normalized: str
    investor_type_candidate: str
    pevc_candidate_status: str
    classification_basis: list[str]
    source_party_ids: list[str]
    source_event_ids: list[str]
    source_transaction_ids: list[str]
    evidence_ids: list[str]
    extraction_source: str
    confidence: float
    review_required: bool
    review_reasons: list[str]


@dataclass
class InvestmentPath:
    investment_path_id: str
    company_id: str
    investor_entity_id: str
    investor_name_normalized: str
    event_id: str
    transaction_id: str
    entry_method: str
    investment_level: str
    direct_or_indirect: str
    transaction_date: str | None
    transferor_party_ids: list[str]
    evidence_ids: list[str]
    path_status: str
    confidence: float
    review_required: bool
    review_reasons: list[str]


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def make_run_id() -> str:
    return datetime.now().astimezone().strftime(
        "PEVCPATH_V032_%Y%m%d_%H%M%S"
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
                output[key] = (
                    json.dumps(value, ensure_ascii=False)
                    if isinstance(value, (list, dict))
                    else value
                )
            writer.writerow(output)


def normalize_name(value: str) -> str:
    text = re.sub(r"\s+", "", str(value or "").strip())
    # 保留“合伙企业（有限合伙）”等名称中的法定括号。
    text = text.strip("，。；、:：[]【】")
    for prefix in (
        "发行对象名称",
        "认购对象名称",
        "股东名称",
        "投资者名称",
        "发行对象",
        "认购对象",
        "受让方",
        "投资者",
        "股东",
    ):
        if text.startswith(prefix):
            text = text[len(prefix):]
    return text.strip("，。；、:：[]【】")


def identity_key(value: str) -> str:
    text = normalize_name(value)
    return re.sub(
        r"(?:股份有限公司|有限责任公司|有限公司|"
        r"合伙企业（有限合伙）|合伙企业\(有限合伙\)|"
        r"有限合伙企业|有限合伙|投资基金|基金)$",
        "",
        text,
    )


def is_noise_name(value: str) -> bool:
    text = normalize_name(value)
    if len(text) < 2 or len(text) > 60:
        return True
    if text in NOISE_TERMS:
        return True
    if any(term in text for term in NOISE_TERMS):
        return True
    if any(term in text for term in NON_INVESTOR_TERMS):
        return True
    if re.fullmatch(r"[\d.,，.%％元股万元亿元/（）()]+", text):
        return True
    if not re.search(r"[\u4e00-\u9fffA-Za-z]", text):
        return True
    return False


def is_legal_entity_name(value: str) -> bool:
    text = normalize_name(value)
    return any(text.endswith(suffix) for suffix in LEGAL_ENTITY_SUFFIXES)


COMPANY_CORE_SUFFIXES = (
    "股份有限公司",
    "有限责任公司",
    "有限公司",
    "股份公司",
    "集团公司",
    "公司",
    "股份",
)


def company_name_core(value: str) -> str:
    text = normalize_name(value)
    text = re.sub(
        r"^(?:安徽|江苏|浙江|上海|北京|深圳|广州|"
        r"芜湖|黄山|昆山|苏州|无锡|常州|杭州|南京|"
        r"共青城)",
        "",
        text,
    )
    for suffix in COMPANY_CORE_SUFFIXES:
        if text.endswith(suffix):
            text = text[:-len(suffix)]
            break
    return text


def issuer_alias_map(
    events: list[dict[str, Any]],
) -> dict[str, set[str]]:
    aliases: dict[str, set[str]] = defaultdict(set)
    for event in events:
        company_id = str(event["company_id"])
        short_name = normalize_name(
            str(event.get("company_short_name") or "")
        )
        if short_name:
            aliases[company_id].add(short_name)
            core = company_name_core(short_name)
            if core:
                aliases[company_id].add(core)
    return aliases


def is_issuer_name(
    value: str,
    company_aliases: set[str],
) -> bool:
    text = normalize_name(value)
    core = company_name_core(text)
    if not text or not core:
        return False
    for alias in company_aliases:
        alias_text = normalize_name(alias)
        alias_core = company_name_core(alias_text)
        if (
            text == alias_text
            or core == alias_core
            or (
                len(alias_core) >= 3
                and alias_core in core
                and len(core) - len(alias_core) <= 2
            )
        ):
            return True
    return False


def split_compound_names(value: str) -> list[str]:
    text = normalize_name(value)
    list_prefix_pattern = (
        r"^(?:前公司股东|原公司股东|公司股东|"
        r"原股东|现有股东|股东包括|股东为|"
        r"投资者包括|投资者为|发行对象包括|"
        r"认购对象包括)"
    )
    had_list_prefix = bool(
        re.search(list_prefix_pattern, text)
    )
    text = re.sub(
        list_prefix_pattern,
        "",
        text,
    )

    has_separator = bool(
        re.search(r"[、，,；;]|(?:以及)|(?:及)|(?:和)", text)
    )
    if not has_separator and not had_list_prefix:
        pieces = [text]
    else:
        pieces = re.split(
            r"[、，,；;]|(?:以及)|(?:及)|(?:和)",
            text,
        )

    output: list[str] = []
    terminal_pattern = re.compile(
        r"(?P<name>"
        r"[\u4e00-\u9fffA-Za-z0-9·（）()]{2,60}?"
        r"(?:合伙企业（有限合伙）|合伙企业\(有限合伙\)|"
        r"有限合伙企业|股份有限公司|有限责任公司|"
        r"有限公司|有限合伙|股权投资基金|创业投资基金|"
        r"私募投资基金|私募基金|产业投资基金|"
        r"并购基金|投资基金|基金|创业投资|"
        r"股权投资|投资管理|资产管理|资本管理)"
        r")"
    )

    for piece in pieces:
        name = normalize_name(piece)
        name = re.sub(
            r"^(?:前公司股东|原公司股东|公司股东|"
            r"原股东|股东|投资者|发行对象|认购对象)",
            "",
            name,
        )
        name = re.sub(
            r"(?:参与本次增资|参与增资|参与认购|"
            r"认购本次发行|认购|增资入股|受让股权|"
            r"受让股份|参与本次交易|参与交易|"
            r"持有公司股权|持有公司股份).*$",
            "",
            name,
        )
        match = terminal_pattern.search(name)
        if match:
            name = match.group("name")
        name = normalize_name(name)
        if name and name not in output:
            output.append(name)
    return output


def high_quality_investor_fragment(value: str) -> bool:
    text = normalize_name(value)
    if is_noise_name(text):
        return False
    if is_legal_entity_name(text):
        return True
    if is_natural_person_name(text):
        return True
    if any(
        term in text
        for term in (
            *STRONG_PEVC_TERMS,
            *MEDIUM_PEVC_TERMS,
        )
    ):
        return True
    return False


def is_natural_person_name(value: str) -> bool:
    text = normalize_name(value)
    return (
        bool(re.fullmatch(r"[\u4e00-\u9fff·]{2,4}", text))
        and not is_noise_name(text)
    )


def has_strong_pevc_name_evidence(
    value: str,
) -> bool:
    text = normalize_name(value)
    return (
        text.endswith("基金")
        or any(
            term in text
            for term in STRONG_PEVC_TERMS
        )
    )


def classify_investor_name(
    value: str,
) -> tuple[str, str, list[str], float]:
    text = normalize_name(value)

    if is_noise_name(text):
        return (
            "INVALID_TEXT_FRAGMENT",
            "EXCLUDED_NOISE",
            ["命中表头、数值或叙述片段排除规则"],
            0.99,
        )

    if has_strong_pevc_name_evidence(text):
        return (
            "PRIVATE_EQUITY_OR_VENTURE_ENTITY",
            "PEVC_CANDIDATE",
            [
                "名称命中强PE/VC关键词",
                (
                    "名称可能为招股书使用的基金简称"
                    if not is_legal_entity_name(text)
                    else "名称具有基金或投资机构法律形态"
                ),
            ],
            0.93,
        )

    if any(term in text for term in MEDIUM_PEVC_TERMS):
        return (
            "INVESTMENT_ENTITY_UNRESOLVED",
            "POSSIBLE_PEVC",
            ["名称命中投资机构关键词，但不足以确认基金属性"],
            0.75,
        )

    if is_natural_person_name(text):
        return (
            "NATURAL_PERSON",
            "NOT_PEVC",
            ["名称形态为自然人"],
            0.90,
        )

    if is_legal_entity_name(text):
        return (
            "LEGAL_ENTITY_OR_STRATEGIC_INVESTOR",
            "NOT_PEVC_OR_UNRESOLVED_STRATEGIC",
            ["名称为法人或合伙企业，但无PE/VC关键词"],
            0.78,
        )

    return (
        "UNRESOLVED_NAME",
        "EXCLUDED_LOW_QUALITY",
        ["名称不满足自然人或法律实体的高精度格式"],
        0.35,
    )


def resolve_latest_structured_run(
    repo_root: Path,
    structured_run_id: str | None,
) -> tuple[str, Path, Path]:
    if structured_run_id:
        return (
            structured_run_id,
            repo_root / "auto_output" / "structured_events" / "runs" / structured_run_id,
            repo_root / "validation" / "structured_events" / "runs" / structured_run_id,
        )

    latest_path = (
        repo_root / "logs" / "structured_events" / "latest_run.json"
    )
    if not latest_path.is_file():
        raise FileNotFoundError("未找到最新结构化运行记录")
    run_id = str(read_json(latest_path)["run_id"])
    return (
        run_id,
        repo_root / "auto_output" / "structured_events" / "runs" / run_id,
        repo_root / "validation" / "structured_events" / "runs" / run_id,
    )


def resolve_latest_numeric_run(
    repo_root: Path,
    numeric_run_id: str | None,
) -> tuple[str, Path]:
    if numeric_run_id:
        return (
            numeric_run_id,
            repo_root / "validation" / "numeric_validation" / "runs" / numeric_run_id,
        )

    latest_path = (
        repo_root / "logs" / "numeric_validation" / "latest_run.json"
    )
    if not latest_path.is_file():
        raise FileNotFoundError("未找到最新数值校验运行记录")
    run_id = str(read_json(latest_path)["run_id"])
    return (
        run_id,
        repo_root / "validation" / "numeric_validation" / "runs" / run_id,
    )


def entry_method(transaction_type: str) -> str:
    return {
        "CAPITAL_INCREASE": "CAPITAL_INCREASE_ENTRY",
        "DIRECTIONAL_FINANCING": "CAPITAL_INCREASE_ENTRY",
        "EQUITY_TRANSFER": "TRANSFER_ENTRY",
    }[transaction_type]


def evidence_text_by_event(
    events: list[dict[str, Any]],
    evidence: list[dict[str, Any]],
) -> tuple[
    dict[str, list[dict[str, Any]]],
    dict[str, str],
]:
    candidate_to_event = {
        str(item["candidate_event_id"]): str(item["event_id"])
        for item in events
    }
    by_event: dict[str, list[dict[str, Any]]] = defaultdict(list)
    text_by_event: dict[str, str] = defaultdict(str)
    for row in evidence:
        event_id = candidate_to_event.get(
            str(row["candidate_event_id"])
        )
        if not event_id:
            continue
        by_event[event_id].append(row)
        text_by_event[event_id] += (
            "\n" + str(row.get("evidence_text") or "")
        )
    return by_event, text_by_event


LEGAL_NAME_PATTERN = re.compile(
    r"(?P<name>"
    r"[\u4e00-\u9fffA-Za-z0-9·（）()]{2,70}?"
    r"(?:"
    r"合伙企业（有限合伙）"
    r"|合伙企业\(有限合伙\)"
    r"|有限合伙企业"
    r"|股份有限公司"
    r"|有限责任公司"
    r"|有限公司"
    r"|有限合伙"
    r"|投资基金"
    r"|基金"
    r")"
    r")"
)



def clean_legal_candidate(value: str) -> str:
    text = normalize_name(value)

    # 优先从明确投资角色标记之后截取，防止把整句前文吞入名称。
    role_markers = (
        "发行对象名称",
        "认购对象名称",
        "投资者名称",
        "股东名称",
        "受让方为",
        "转让给",
        "转让予",
        "发行对象",
        "认购对象",
        "新增股东",
        "投资者",
        "受让方",
    )
    best_position = -1
    best_marker = ""
    for marker in role_markers:
        position = text.rfind(marker)
        if position > best_position:
            best_position = position
            best_marker = marker
    if best_position >= 0:
        text = text[
            best_position + len(best_marker):
        ]

    text = re.sub(
        r"^(?:序号|名称|对象|为|由)+",
        "",
        text,
    )

    # 复杂合伙企业后缀必须作为一个整体保留。
    complex_match = re.search(
        r"(?P<name>"
        r"[\u4e00-\u9fffA-Za-z0-9·]{2,60}"
        r"(?:合伙企业（有限合伙）|合伙企业\(有限合伙\)|"
        r"有限合伙企业))",
        text,
    )
    if complex_match:
        return normalize_name(
            complex_match.group("name")
        )

    # 普通法人/基金名称取到第一个完整法律后缀为止。
    ordinary_match = re.search(
        r"(?P<name>"
        r"[\u4e00-\u9fffA-Za-z0-9·]{2,60}?"
        r"(?:股份有限公司|有限责任公司|有限公司|"
        r"有限合伙|投资基金|基金))",
        text,
    )
    if ordinary_match:
        return normalize_name(
            ordinary_match.group("name")
        )

    return normalize_name(text)


def evidence_entity_candidates(
    text: str,
    transaction_type: str,
    company_aliases: set[str] | None = None,
) -> list[tuple[str, str]]:
    company_aliases = (
        company_aliases or set()
    )
    role_terms = (
        (
            "认购",
            "增资",
            "发行对象",
            "认购对象",
            "新增股东",
            "投资者",
            "股东",
        )
        if transaction_type
        in {
            "CAPITAL_INCREASE",
            "DIRECTIONAL_FINANCING",
        }
        else (
            "转让给",
            "受让",
            "受让方",
            "股权转让",
            "股东",
        )
    )

    raw_segments = re.split(
        r"[\n\r。；;]",
        str(text or ""),
    )
    output: list[tuple[str, str]] = []

    for raw_segment in raw_segments:
        segment = normalize_name(raw_segment)
        if not segment:
            continue
        if not (
            any(term in segment for term in role_terms)
            or any(
                term in segment
                for term in (
                    *STRONG_PEVC_TERMS,
                    *MEDIUM_PEVC_TERMS,
                )
            )
        ):
            continue

        # 只在明确名单或存在分隔符时拆分并列主体。
        is_list_segment = bool(
            re.search(
                r"[、，,]|^(?:前公司股东|原公司股东|"
                r"公司股东|现有股东|"
                r"股东包括|股东为|投资者包括|"
                r"发行对象包括|认购对象包括)",
                segment,
            )
        )
        if is_list_segment:
            for name in split_compound_names(
                segment
            ):
                if (
                    high_quality_investor_fragment(
                        name
                    )
                    and not is_issuer_name(
                        name,
                        company_aliases,
                    )
                ):
                    pair = (
                        name,
                        "EVIDENCE_COMPOUND_LIST",
                    )
                    if pair not in output:
                        output.append(pair)

        # 先从完整角色语句中提取一个最长法律实体名称。
        direct_candidate = clean_legal_candidate(
            segment
        )
        if (
            direct_candidate
            and high_quality_investor_fragment(
                direct_candidate
            )
            and not is_issuer_name(
                direct_candidate,
                company_aliases,
            )
        ):
            pair = (
                direct_candidate,
                "EVIDENCE_ROLE_PHRASE",
            )
            if pair not in output:
                output.append(pair)

        # 再处理完整法律实体名称。
        for match in LEGAL_NAME_PATTERN.finditer(segment):
            candidate = clean_legal_candidate(
                match.group("name")
            )
            if (
                candidate
                and high_quality_investor_fragment(
                    candidate
                )
                and not is_issuer_name(
                    candidate,
                    company_aliases,
                )
            ):
                pair = (
                    candidate,
                    "EVIDENCE_LEGAL_ENTITY",
                )
                if pair not in output:
                    output.append(pair)

        # 单独识别基金或投资机构简称。
        keyword_pattern = re.compile(
            r"(?P<name>"
            r"[\u4e00-\u9fffA-Za-z0-9·]{2,30}?"
            r"(?:股权投资基金|创业投资基金|"
            r"私募投资基金|私募基金|产业投资基金|"
            r"并购基金|投资基金|基金|创业投资|"
            r"股权投资|投资管理|资产管理|资本管理)"
            r")"
        )
        for match in keyword_pattern.finditer(segment):
            candidate = normalize_name(
                match.group("name")
            )
            candidate = re.sub(
                r"^(?:前公司股东|原公司股东|公司股东|"
                r"原股东|股东|投资者|发行对象|认购对象)",
                "",
                candidate,
            )
            candidate = normalize_name(candidate)
            if (
                candidate
                and high_quality_investor_fragment(
                    candidate
                )
                and not is_issuer_name(
                    candidate,
                    company_aliases,
                )
            ):
                pair = (
                    candidate,
                    "EVIDENCE_INVESTMENT_KEYWORD",
                )
                if pair not in output:
                    output.append(pair)

    collapsed: list[tuple[str, str]] = []
    for name, source in sorted(
        output,
        key=lambda item: len(item[0]),
        reverse=True,
    ):
        if any(
            name in existing_name
            for existing_name, _ in collapsed
        ):
            continue
        collapsed.append((name, source))

    return list(reversed(collapsed))


def linked_party_candidates(
    transaction: dict[str, Any],
    party_by_id: dict[str, dict[str, Any]],
    company_aliases: set[str],
) -> list[tuple[str, str, str]]:
    fields = (
        (
            "investor_party_ids",
            "LINKED_INVESTOR_PARTY",
        ),
        (
            "transferee_party_ids",
            "LINKED_TRANSFEREE_PARTY",
        ),
    )
    output: list[tuple[str, str, str]] = []
    for field, source in fields:
        for party_id in transaction.get(field) or []:
            party = party_by_id.get(str(party_id))
            if not party:
                continue
            raw_name = normalize_name(
                str(
                    party.get("party_name_normalized")
                    or party.get("party_name_raw")
                    or ""
                )
            )
            for name in split_compound_names(raw_name):
                if (
                    high_quality_investor_fragment(
                        name
                    )
                    and not is_issuer_name(
                        name,
                        company_aliases,
                    )
                ):
                    item = (
                        name,
                        source,
                        str(party_id),
                    )
                    if item not in output:
                        output.append(item)
    return output


def build_entities_and_paths(
    events: list[dict[str, Any]],
    transactions: list[dict[str, Any]],
    parties: list[dict[str, Any]],
    evidence: list[dict[str, Any]],
) -> tuple[
    list[InvestorEntity],
    list[InvestmentPath],
    list[dict[str, Any]],
    list[dict[str, Any]],
]:
    party_by_id = {
        str(item["party_id"]): item
        for item in parties
    }
    aliases_by_company = issuer_alias_map(
        events
    )
    evidence_by_event, text_by_event = evidence_text_by_event(
        events,
        evidence,
    )

    entity_accumulator: dict[
        tuple[str, str], dict[str, Any]
    ] = {}
    path_inputs: list[dict[str, Any]] = []
    missing_reviews: list[dict[str, Any]] = []
    discarded: list[dict[str, Any]] = []

    for transaction in transactions:
        transaction_type = str(transaction["transaction_type"])
        if transaction_type not in ELIGIBLE_TRANSACTION_TYPES:
            continue

        company_id = str(transaction["company_id"])
        event_id = str(transaction["event_id"])
        transaction_id = str(transaction["transaction_id"])
        company_aliases = aliases_by_company.get(
            company_id,
            set(),
        )
        linked_candidates = linked_party_candidates(
            transaction,
            party_by_id,
            company_aliases,
        )
        candidates: list[
            tuple[str, str, str | None]
        ] = list(linked_candidates)

        for name, source in evidence_entity_candidates(
            text_by_event.get(event_id, ""),
            transaction_type,
            company_aliases,
        ):
            item = (name, source, None)
            if item not in candidates:
                candidates.append(item)

        accepted_candidates: list[
            tuple[str, str, str | None]
        ] = []
        for name, source, source_party_id in candidates:
            if is_issuer_name(
                name,
                company_aliases,
            ):
                discarded.append({
                    "company_id": company_id,
                    "event_id": event_id,
                    "transaction_id": transaction_id,
                    "name": name,
                    "source": source,
                    "reason": [
                        "主体与发行人名称或核心名称一致"
                    ],
                })
                continue

            (
                entity_type,
                status,
                basis,
                confidence,
            ) = classify_investor_name(name)
            if status in {
                "EXCLUDED_NOISE",
                "EXCLUDED_LOW_QUALITY",
            }:
                discarded.append({
                    "company_id": company_id,
                    "event_id": event_id,
                    "transaction_id": transaction_id,
                    "name": name,
                    "source": source,
                    "reason": basis,
                })
                continue
            accepted_candidates.append((name, source))

            key = (company_id, identity_key(name))
            acc = entity_accumulator.setdefault(key, {
                "company_id": company_id,
                "name": name,
                "type": entity_type,
                "status": status,
                "basis": list(basis),
                "party_ids": [],
                "event_ids": [],
                "transaction_ids": [],
                "evidence_ids": [],
                "sources": [],
                "confidence": confidence,
            })
            if (
                source_party_id
                and source_party_id
                not in acc["party_ids"]
            ):
                acc["party_ids"].append(
                    source_party_id
                )
            if event_id not in acc["event_ids"]:
                acc["event_ids"].append(event_id)
            if transaction_id not in acc["transaction_ids"]:
                acc["transaction_ids"].append(transaction_id)
            for row in evidence_by_event.get(event_id, []):
                evidence_id = str(row["evidence_id"])
                if evidence_id not in acc["evidence_ids"]:
                    acc["evidence_ids"].append(evidence_id)
            if source not in acc["sources"]:
                acc["sources"].append(source)

            path_inputs.append({
                "key": key,
                "name": name,
                "event_id": event_id,
                "transaction_id": transaction_id,
                "transaction_type": transaction_type,
                "transaction_date": transaction.get(
                    "transaction_date"
                ),
                "transferor_party_ids": list(
                    transaction.get("transferor_party_ids") or []
                ),
                "evidence_ids": [
                    str(row["evidence_id"])
                    for row in evidence_by_event.get(event_id, [])
                ],
                "source": source,
            })

        if not accepted_candidates:
            missing_reviews.append({
                "review_id": f"REV-MISSING-{transaction_id}",
                "company_id": company_id,
                "record_type": "MISSING_INVESTOR_FOR_TRANSACTION",
                "record_id": transaction_id,
                "event_id": event_id,
                "transaction_id": transaction_id,
                "review_reason": (
                    "增资、定向发行或股权转让事件未抽取到"
                    "可确认的投资方；不得用表头或数值片段代替主体"
                ),
                "manual_status": "PENDING",
            })

    entities: list[InvestorEntity] = []
    entity_id_by_key: dict[tuple[str, str], str] = {}
    for index, (key, acc) in enumerate(
        sorted(entity_accumulator.items()),
        start=1,
    ):
        entity_id = f"INV-{acc['company_id']}-{index:04d}"
        entity_id_by_key[key] = entity_id
        review_reasons: list[str] = []
        if acc["status"] == "POSSIBLE_PEVC":
            review_reasons.append(
                "名称显示投资机构属性，但不足以确认PE/VC"
            )
        entities.append(InvestorEntity(
            investor_entity_id=entity_id,
            company_id=acc["company_id"],
            investor_name_raw=acc["name"],
            investor_name_normalized=acc["name"],
            investor_type_candidate=acc["type"],
            pevc_candidate_status=acc["status"],
            classification_basis=acc["basis"],
            source_party_ids=acc["party_ids"],
            source_event_ids=acc["event_ids"],
            source_transaction_ids=acc["transaction_ids"],
            evidence_ids=acc["evidence_ids"],
            extraction_source=";".join(acc["sources"]),
            confidence=acc["confidence"],
            review_required=bool(review_reasons),
            review_reasons=review_reasons,
        ))

    entity_by_id: dict[str, InvestorEntity] = {
        item.investor_entity_id: item
        for item in entities
    }
    paths: list[InvestmentPath] = []
    seen_paths: set[tuple[str, str]] = set()
    for item in path_inputs:
        entity_id = entity_id_by_key[item["key"]]
        path_key = (entity_id, item["transaction_id"])
        if path_key in seen_paths:
            continue
        seen_paths.add(path_key)
        entity = entity_by_id[entity_id]
        path_review_reasons: list[str] = []
        if entity.pevc_candidate_status == "POSSIBLE_PEVC":
            path_review_reasons.append(
                "投资主体PE/VC属性待确认"
            )
        paths.append(InvestmentPath(
            investment_path_id=(
                f"PATH-{item['key'][0]}-{len(paths)+1:05d}"
            ),
            company_id=item["key"][0],
            investor_entity_id=entity_id,
            investor_name_normalized=item["name"],
            event_id=item["event_id"],
            transaction_id=item["transaction_id"],
            entry_method=entry_method(item["transaction_type"]),
            investment_level="ISSUER_LEVEL",
            direct_or_indirect="DIRECT",
            transaction_date=item["transaction_date"],
            transferor_party_ids=item["transferor_party_ids"],
            evidence_ids=item["evidence_ids"],
            path_status=(
                "AUTO_IDENTIFIED"
                if not path_review_reasons
                else "AUTO_IDENTIFIED_WITH_REVIEW"
            ),
            confidence=min(
                entity.confidence,
                0.92 if item["source"] == "EVIDENCE_LEGAL_ENTITY" else 0.88,
            ),
            review_required=bool(path_review_reasons),
            review_reasons=path_review_reasons,
        ))

    return entities, paths, missing_reviews, discarded


def issuer_self_entities_for_validation(
    entities: list[InvestorEntity],
    events: list[dict[str, Any]],
) -> list[InvestorEntity]:
    aliases_by_company = issuer_alias_map(events)
    return [
        entity
        for entity in entities
        if is_issuer_name(
            entity.investor_name_normalized,
            aliases_by_company.get(
                entity.company_id,
                set(),
            ),
        )
    ]


def validate_outputs(
    entities: list[InvestorEntity],
    paths: list[InvestmentPath],
) -> dict[str, Any]:
    errors: list[str] = []
    entity_ids = {
        item.investor_entity_id for item in entities
    }
    if len(entity_ids) != len(entities):
        errors.append("投资主体ID重复")

    company_aliases_from_entities: dict[
        str, set[str]
    ] = defaultdict(set)

    for entity in entities:
        if is_noise_name(entity.investor_name_normalized):
            errors.append(
                f"{entity.investor_entity_id}仍为噪声主体"
            )
        if entity.pevc_candidate_status == "PEVC_CANDIDATE":
            if not has_strong_pevc_name_evidence(
                entity.investor_name_normalized
            ):
                errors.append(
                    f"{entity.investor_entity_id}缺少强PE/VC名称证据"
                )

    for path in paths:
        if path.investor_entity_id not in entity_ids:
            errors.append(
                f"{path.investment_path_id}引用不存在主体"
            )
        if path.entry_method not in {
            "CAPITAL_INCREASE_ENTRY",
            "TRANSFER_ENTRY",
        }:
            errors.append(
                f"{path.investment_path_id}包含非投资进入交易"
            )
        if not path.evidence_ids:
            errors.append(
                f"{path.investment_path_id}缺少证据ID"
            )

    return {
        "validation_status": (
            "FAILED" if errors else "PASSED"
        ),
        "errors": errors,
        "investor_entity_count": len(entities),
        "investment_path_count": len(paths),
        "pevc_candidate_count": sum(
            item.pevc_candidate_status == "PEVC_CANDIDATE"
            for item in entities
        ),
        "possible_pevc_count": sum(
            item.pevc_candidate_status == "POSSIBLE_PEVC"
            for item in entities
        ),
        "non_pevc_or_strategic_count": sum(
            item.pevc_candidate_status
            in {
                "NOT_PEVC",
                "NOT_PEVC_OR_UNRESOLVED_STRATEGIC",
            }
            for item in entities
        ),
    }


def run(
    repo_root: Path,
    structured_run_id: str | None,
    numeric_run_id: str | None,
) -> int:
    repo_root = repo_root.expanduser().resolve()
    structured_id, auto_dir, structured_validation_dir = (
        resolve_latest_structured_run(
            repo_root,
            structured_run_id,
        )
    )
    numeric_id, numeric_dir = resolve_latest_numeric_run(
        repo_root,
        numeric_run_id,
    )

    paths = {
        "events": auto_dir / "event_records_auto.jsonl",
        "transactions": auto_dir / "transaction_records_auto.jsonl",
        "parties": auto_dir / "event_parties_auto.jsonl",
        "evidence": auto_dir / "event_evidence_index.jsonl",
        "structured_metrics": (
            structured_validation_dir
            / "structured_event_metrics.json"
        ),
        "numeric_metrics": (
            numeric_dir / "numeric_validation_metrics.json"
        ),
    }
    missing = [
        str(path)
        for path in paths.values()
        if not path.is_file()
    ]
    if missing:
        raise FileNotFoundError(
            "缺少输入文件：\n" + "\n".join(missing)
        )

    structured_metrics = read_json(
        paths["structured_metrics"]
    )
    numeric_metrics = read_json(paths["numeric_metrics"])
    if structured_metrics.get("event_count") != 26:
        raise ValueError("源事件数不是26")
    if structured_metrics.get("transaction_count") != 27:
        raise ValueError("源交易数不是27")
    if structured_metrics.get("validation_error_count") != 0:
        raise ValueError("源结构化抽取存在验证错误")
    if numeric_metrics.get("failed_count") != 0:
        raise ValueError("源数值校验存在失败项")

    events = read_jsonl(paths["events"])
    transactions = read_jsonl(paths["transactions"])
    parties = read_jsonl(paths["parties"])
    evidence = read_jsonl(paths["evidence"])

    (
        entities,
        investment_paths,
        missing_reviews,
        discarded,
    ) = build_entities_and_paths(
        events,
        transactions,
        parties,
        evidence,
    )
    validation = validate_outputs(
        entities,
        investment_paths,
    )
    issuer_self_entities = (
        issuer_self_entities_for_validation(
            entities,
            events,
        )
    )
    for entity in issuer_self_entities:
        validation["errors"].append(
            f"{entity.investor_entity_id}为发行人自身"
        )
    validation["validation_status"] = (
        "FAILED"
        if validation["errors"]
        else "PASSED"
    )

    reviews = [
        {
            "review_id": f"REV-{item.investor_entity_id}",
            "company_id": item.company_id,
            "record_type": "INVESTOR_ENTITY",
            "record_id": item.investor_entity_id,
            "event_id": None,
            "transaction_id": None,
            "review_reason": "；".join(item.review_reasons),
            "manual_status": "PENDING",
        }
        for item in entities
        if item.review_required
    ] + [
        {
            "review_id": f"REV-{item.investment_path_id}",
            "company_id": item.company_id,
            "record_type": "INVESTMENT_PATH",
            "record_id": item.investment_path_id,
            "event_id": item.event_id,
            "transaction_id": item.transaction_id,
            "review_reason": "；".join(item.review_reasons),
            "manual_status": "PENDING",
        }
        for item in investment_paths
        if item.review_required
    ] + missing_reviews

    run_id = make_run_id()
    auto_output_dir = (
        repo_root
        / "auto_output"
        / "pevc_paths"
        / "runs"
        / run_id
    )
    validation_output_dir = (
        repo_root
        / "validation"
        / "pevc_paths"
        / "runs"
        / run_id
    )
    logs_dir = (
        repo_root
        / "logs"
        / "pevc_paths"
        / "runs"
        / run_id
    )
    for directory in (
        auto_output_dir,
        validation_output_dir,
        logs_dir,
    ):
        directory.mkdir(parents=True, exist_ok=False)

    write_jsonl(
        auto_output_dir / "investor_entities_auto.jsonl",
        [asdict(item) for item in entities],
    )
    write_jsonl(
        auto_output_dir / "investment_paths_auto.jsonl",
        [asdict(item) for item in investment_paths],
    )
    write_jsonl(
        auto_output_dir / "pevc_entities_auto.jsonl",
        [
            asdict(item)
            for item in entities
            if item.pevc_candidate_status
            in {
                "PEVC_CANDIDATE",
                "POSSIBLE_PEVC",
            }
        ],
    )
    write_jsonl(
        auto_output_dir / "pevc_investment_paths_auto.jsonl",
        [
            asdict(path)
            for path in investment_paths
            if next(
                entity
                for entity in entities
                if entity.investor_entity_id
                == path.investor_entity_id
            ).pevc_candidate_status
            in {
                "PEVC_CANDIDATE",
                "POSSIBLE_PEVC",
            }
        ],
    )
    write_jsonl(
        auto_output_dir / "non_pevc_observed_investors.jsonl",
        [
            asdict(item)
            for item in entities
            if item.pevc_candidate_status
            not in {
                "PEVC_CANDIDATE",
                "POSSIBLE_PEVC",
            }
        ],
    )
    write_jsonl(
        validation_output_dir / "discarded_investor_noise.jsonl",
        discarded,
    )
    write_jsonl(
        validation_output_dir / "pevc_review_queue.jsonl",
        reviews,
    )
    write_csv(
        validation_output_dir / "pevc_review_queue.csv",
        reviews,
        [
            "review_id",
            "company_id",
            "record_type",
            "record_id",
            "event_id",
            "transaction_id",
            "review_reason",
            "manual_status",
        ],
    )
    write_csv(
        validation_output_dir / "investor_entities_preview.csv",
        [asdict(item) for item in entities],
        [
            "investor_entity_id",
            "company_id",
            "investor_name_normalized",
            "investor_type_candidate",
            "pevc_candidate_status",
            "classification_basis",
            "source_event_ids",
            "source_transaction_ids",
            "evidence_ids",
            "extraction_source",
            "confidence",
            "review_required",
            "review_reasons",
        ],
    )
    write_csv(
        validation_output_dir / "investment_paths_preview.csv",
        [asdict(item) for item in investment_paths],
        [
            "investment_path_id",
            "company_id",
            "investor_entity_id",
            "investor_name_normalized",
            "event_id",
            "transaction_id",
            "entry_method",
            "direct_or_indirect",
            "transaction_date",
            "evidence_ids",
            "path_status",
            "confidence",
            "review_required",
            "review_reasons",
        ],
    )

    metrics = {
        "metrics_version": "0.3.2",
        "run_id": run_id,
        "pipeline_version": PIPELINE_VERSION,
        "source_structured_run_id": structured_id,
        "source_numeric_run_id": numeric_id,
        "batch_status": (
            "FAILED"
            if validation["validation_status"] == "FAILED"
            else (
                "READY_WITH_REVIEW"
                if reviews
                else "READY"
            )
        ),
        **validation,
        "review_queue_count": len(reviews),
        "missing_investor_transaction_count": sum(
            item["record_type"]
            == "MISSING_INVESTOR_FOR_TRANSACTION"
            for item in reviews
        ),
        "discarded_noise_count": len(discarded),
        "pevc_investment_path_count": sum(
            next(
                entity
                for entity in entities
                if entity.investor_entity_id
                == path.investor_entity_id
            ).pevc_candidate_status
            in {
                "PEVC_CANDIDATE",
                "POSSIBLE_PEVC",
            }
            for path in investment_paths
        ),
        "issuer_self_candidate_count": len(
            issuer_self_entities
        ),
        "entry_method_counts": dict(
            Counter(
                item.entry_method
                for item in investment_paths
            )
        ),
        "llm_called": False,
        "final_generated": False,
        "note": (
            "v0.3.2统一PE/VC分类与验证标准；"
            "以“基金”结尾的招股书基金简称可作为强名称证据。"
        ),
    }
    write_json(
        validation_output_dir / "pevc_path_metrics.json",
        metrics,
    )
    write_json(
        validation_output_dir / "pevc_path_validation.json",
        validation,
    )
    write_json(
        logs_dir / "run_manifest.json",
        {
            "run_id": run_id,
            "pipeline_version": PIPELINE_VERSION,
            "source_structured_run_id": structured_id,
            "source_numeric_run_id": numeric_id,
            "completed_at": now_iso(),
            "batch_status": metrics["batch_status"],
            "llm_called": False,
        },
    )
    write_json(
        repo_root / "logs" / "pevc_paths" / "latest_run.json",
        {
            "run_id": run_id,
            "batch_status": metrics["batch_status"],
            "completed_at": now_iso(),
        },
    )

    print()
    print("PE/VC主体及投资路径识别 v0.3.2 完成")
    print(f"运行ID：{run_id}")
    print(f"投资主体：{len(entities)}")
    print(f"投资路径：{len(investment_paths)}")
    print(
        f"PE/VC候选："
        f"{validation['pevc_candidate_count']}"
    )
    print(
        f"可能PE/VC："
        f"{validation['possible_pevc_count']}"
    )
    print(
        f"非PE/VC或战略投资者："
        f"{validation['non_pevc_or_strategic_count']}"
    )
    print(
        f"PE/VC投资路径："
        f"{metrics['pevc_investment_path_count']}"
    )
    print(
        f"发行人自身误识别："
        f"{metrics['issuer_self_candidate_count']}"
    )
    print(
        f"缺少可确认投资方的交易："
        f"{metrics['missing_investor_transaction_count']}"
    )
    print(f"隔离噪声：{len(discarded)}")
    print(f"人工复核项：{len(reviews)}")
    print(f"验证错误：{len(validation['errors'])}")
    print(f"批次状态：{metrics['batch_status']}")
    return 0 if not validation["errors"] else 2


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "高精度识别增资、定向发行及股权受让中的"
            "PE/VC候选主体和直接投资路径"
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
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        return run(
            repo_root=args.repo_root,
            structured_run_id=args.structured_run_id,
            numeric_run_id=args.numeric_run_id,
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
