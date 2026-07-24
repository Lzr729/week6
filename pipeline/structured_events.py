from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
import traceback
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


PIPELINE_VERSION = "structured_event_extraction_v0.2"


@dataclass
class EventRecord:
    event_id: str
    candidate_event_id: str
    company_id: str
    company_short_name: str
    event_type: str
    event_type_candidates: list[str]
    event_group: str
    event_period: str | None
    event_date_text: str | None
    event_date_primary_role: str | None
    event_dates: dict[str, list[str]]
    event_title: str
    disclosure_scope: str
    ordinal_labels: list[str]
    pdf_page_start: int
    pdf_page_end: int
    printed_page_start: str | None
    printed_page_end: str | None
    printed_page_value_type: str
    primary_evidence_id: str
    supporting_evidence_ids: list[str]
    entity_scope: str
    candidate_confidence: float
    extraction_status: str
    review_required: bool
    review_reasons: list[str]


@dataclass
class TransactionRecord:
    transaction_id: str
    event_id: str
    candidate_event_id: str
    company_id: str
    transaction_type: str
    sequence_in_event: int
    transaction_date: str | None
    transaction_date_role: str | None
    transferor_party_ids: list[str]
    transferee_party_ids: list[str]
    investor_party_ids: list[str]
    absorbed_party_ids: list[str]
    registered_capital_before_fact_ids: list[str]
    registered_capital_after_fact_ids: list[str]
    capital_increase_fact_ids: list[str]
    consideration_fact_ids: list[str]
    share_quantity_fact_ids: list[str]
    share_price_fact_ids: list[str]
    equity_ratio_fact_ids: list[str]
    transaction_status: str
    review_required: bool
    review_reasons: list[str]


@dataclass
class PartyRecord:
    party_id: str
    event_id: str
    candidate_event_id: str
    company_id: str
    party_name_raw: str
    party_name_normalized: str
    party_role: str
    party_type_candidate: str
    source_evidence_id: str
    pdf_page_start: int
    pdf_page_end: int
    evidence_text: str
    confidence: float
    review_required: bool
    review_reason: str | None


@dataclass
class NumericFact:
    numeric_fact_id: str
    event_id: str
    candidate_event_id: str
    company_id: str
    fact_type: str
    raw_value_text: str
    numeric_value: float | None
    unit: str | None
    currency: str | None
    scale_multiplier: float
    normalized_value: float | None
    value_type: str
    source_evidence_id: str
    pdf_page_start: int
    pdf_page_end: int
    context_text: str
    confidence: float
    review_required: bool
    review_reason: str | None


@dataclass
class DateFact:
    date_fact_id: str
    event_id: str
    candidate_event_id: str
    company_id: str
    date_role: str
    date_text: str
    date_normalized: str
    source: str
    source_evidence_id: str | None
    pdf_page_start: int
    pdf_page_end: int
    confidence: float


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def make_run_id() -> str:
    return datetime.now().astimezone().strftime(
        "STRUCTEVENT_V01_%Y%m%d_%H%M%S"
    )


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "").strip())


def normalize_party_name(name: str) -> str:
    value = re.sub(r"\s+", "", name.strip())
    value = re.sub(
        r"^(?:由|与|和|及|其中|股东|公司|本公司|发行人)",
        "",
        value,
    )
    value = re.sub(
        r"(?:分别|共同|合计|持有|认购|受让|转让)$",
        "",
        value,
    )
    return value.strip("，。；、:：()（）")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


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


def resolve_frozen_dir(
    repo_root: Path,
    freeze_id: str | None,
) -> tuple[str, Path]:
    if freeze_id:
        path = (
            repo_root
            / "review"
            / "candidate_events"
            / "frozen"
            / freeze_id
        )
        return freeze_id, path

    latest_path = (
        repo_root
        / "review"
        / "candidate_events"
        / "latest_frozen.json"
    )
    if not latest_path.is_file():
        raise FileNotFoundError(
            "未找到review/candidate_events/latest_frozen.json；"
            "请先运行候选事件冻结Patch。"
        )
    latest = read_json(latest_path)
    latest_freeze_id = str(latest["freeze_id"])
    output_dir = latest.get("output_dir")
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
            / latest_freeze_id
        )
    return latest_freeze_id, path.resolve()


EVENT_GROUPS = {
    "LIMITED_COMPANY_ESTABLISHMENT": "ESTABLISHMENT",
    "JOINT_STOCK_COMPANY_ESTABLISHMENT": "ESTABLISHMENT",
    "OVERALL_CHANGE": "REORGANIZATION",
    "CAPITAL_INCREASE": "CAPITAL_CHANGE",
    "CAPITAL_REDUCTION": "CAPITAL_CHANGE",
    "SHARE_CAPITAL_CHANGE": "CAPITAL_CHANGE",
    "EQUITY_TRANSFER": "EQUITY_TRANSFER",
    "ABSORPTION_MERGER": "REORGANIZATION",
    "EQUITY_HOLDING_PROXY": "EQUITY_HOLDING_PROXY",
    "EQUITY_HOLDING_PROXY_RELEASE": "EQUITY_HOLDING_PROXY",
    "CONTROL_CHANGE": "CONTROL_CHANGE",
    "MAJOR_ASSET_RESTRUCTURING": "REORGANIZATION",
    "LISTING_OR_DIRECTIONAL_FINANCING": "FINANCING",
}


TRANSACTION_TYPES = {
    "LIMITED_COMPANY_ESTABLISHMENT": "ESTABLISHMENT",
    "JOINT_STOCK_COMPANY_ESTABLISHMENT": "JOINT_STOCK_ESTABLISHMENT",
    "OVERALL_CHANGE": "OVERALL_CHANGE",
    "CAPITAL_INCREASE": "CAPITAL_INCREASE",
    "CAPITAL_REDUCTION": "CAPITAL_REDUCTION",
    "SHARE_CAPITAL_CHANGE": "SHARE_CAPITAL_CHANGE",
    "EQUITY_TRANSFER": "EQUITY_TRANSFER",
    "ABSORPTION_MERGER": "ABSORPTION_MERGER",
    "EQUITY_HOLDING_PROXY": "EQUITY_HOLDING_PROXY",
    "EQUITY_HOLDING_PROXY_RELEASE": "EQUITY_HOLDING_PROXY_RELEASE",
    "CONTROL_CHANGE": "CONTROL_CHANGE",
    "MAJOR_ASSET_RESTRUCTURING": "MAJOR_ASSET_RESTRUCTURING",
    "LISTING_OR_DIRECTIONAL_FINANCING": "DIRECTIONAL_FINANCING",
}


DATE_ROLE_PATTERNS = {
    "resolution_date": [
        r"(?P<date>(?:19|20)\d{2}年\d{1,2}月\d{1,2}日)[^。；]{0,25}?(?:股东会|董事会|股东大会)[^。；]{0,15}?决议",
        r"(?P<date>(?:19|20)\d{2}年\d{1,2}月\d{1,2}日)[^。；]{0,20}?审议通过",
    ],
    "agreement_date": [
        r"(?P<date>(?:19|20)\d{2}年\d{1,2}月\d{1,2}日)[^。；]{0,25}?(?:签署|签订)[^。；]{0,20}?(?:协议|合同)",
    ],
    "payment_date": [
        r"(?P<date>(?:19|20)\d{2}年\d{1,2}月\d{1,2}日)[^。；]{0,25}?(?:支付|缴纳|缴付|出资到位)",
    ],
    "verification_date": [
        r"(?P<date>(?:19|20)\d{2}年\d{1,2}月\d{1,2}日)[^。；]{0,25}?(?:验资|审验|验资报告)",
    ],
    "registration_date": [
        r"(?P<date>(?:19|20)\d{2}年\d{1,2}月\d{1,2}日)[^。；]{0,30}?(?:工商变更|工商登记|营业执照|登记手续)",
    ],
    "effective_date": [
        r"(?P<date>(?:19|20)\d{2}年\d{1,2}月\d{1,2}日)[^。；]{0,20}?(?:完成|生效)",
    ],
}


NUMERIC_PATTERNS = [
    (
        "REGISTERED_CAPITAL_BEFORE",
        re.compile(
            r"注册资本(?:由|从)(?P<value>[\d,，.]+)"
            r"(?P<unit>亿元|万元|元)"
        ),
        0.94,
    ),
    (
        "REGISTERED_CAPITAL_AFTER",
        re.compile(
            r"注册资本(?:由|从)[\d,，.]+(?:亿元|万元|元)"
            r"(?:增加至|增至|增加为|变更为|变更至|变更|减少至|减至|减少为|至)"
            r"(?P<value>[\d,，.]+)(?P<unit>亿元|万元|元)"
        ),
        0.96,
    ),
    (
        "REGISTERED_CAPITAL",
        re.compile(
            r"注册资本(?:为|：|:)?(?P<value>[\d,，.]+)"
            r"(?P<unit>亿元|万元|元)"
        ),
        0.88,
    ),
    (
        "CAPITAL_INCREASE_AMOUNT",
        re.compile(
            r"(?:新增注册资本|增加注册资本|增资额(?:为)?|"
            r"本次增资(?:金额)?(?:为)?)"
            r"(?P<value>[\d,，.]+)(?P<unit>亿元|万元|元)"
        ),
        0.92,
    ),
    (
        "CONSIDERATION",
        re.compile(
            r"(?:转让价款|转让价格|交易对价|支付对价|"
            r"受让价款|认购价款|出资额(?:为)?)"
            r"(?P<value>[\d,，.]+)(?P<unit>亿元|万元|元)"
        ),
        0.88,
    ),
    (
        "SHARE_QUANTITY",
        re.compile(
            r"(?P<value>[\d,，.]+)(?P<unit>亿股|万股|股)"
        ),
        0.82,
    ),
    (
        "SHARE_PRICE",
        re.compile(
            r"(?:每股|发行价格(?:为)?|转让单价(?:为)?)"
            r"(?P<value>[\d,，.]+)(?P<unit>元/股|元)"
        ),
        0.91,
    ),
    (
        "EQUITY_RATIO",
        re.compile(
            r"(?P<value>[\d.]+)(?P<unit>%)"
        ),
        0.76,
    ),
    (
        "NET_ASSET_VALUE",
        re.compile(
            r"(?:净资产|经审计净资产)(?:为|：|:)?"
            r"(?P<value>[\d,，.]+)(?P<unit>亿元|万元|元)"
        ),
        0.86,
    ),
    (
        "APPRAISED_VALUE",
        re.compile(
            r"(?:评估值|评估净资产)(?:为|：|:)?"
            r"(?P<value>[\d,，.]+)(?P<unit>亿元|万元|元)"
        ),
        0.84,
    ),
]


def parse_number(value: str) -> float | None:
    cleaned = value.replace(",", "").replace("，", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def unit_metadata(unit: str | None) -> tuple[str | None, str | None, float]:
    if unit is None:
        return None, None, 1.0
    if unit == "亿元":
        return "CNY", "元", 100_000_000.0
    if unit == "万元":
        return "CNY", "元", 10_000.0
    if unit == "元":
        return "CNY", "元", 1.0
    if unit == "亿股":
        return None, "股", 100_000_000.0
    if unit == "万股":
        return None, "股", 10_000.0
    if unit == "股":
        return None, "股", 1.0
    if unit == "元/股":
        return "CNY", "元/股", 1.0
    if unit == "%":
        return None, "%", 0.01
    return None, unit, 1.0


def context_window(text: str, start: int, end: int, size: int = 55) -> str:
    return text[max(0, start - size): min(len(text), end + size)]


def extract_numeric_facts(
    event_id: str,
    candidate: dict[str, Any],
    evidence_rows: list[dict[str, Any]],
) -> list[NumericFact]:
    facts: list[NumericFact] = []
    seen: set[tuple[str, str, str]] = set()
    for evidence in evidence_rows:
        text = str(evidence.get("evidence_text") or "")
        normalized = re.sub(r"\s+", "", text)
        for fact_type, pattern, confidence in NUMERIC_PATTERNS:
            for match in pattern.finditer(normalized):
                raw_value = match.group("value")
                unit_raw = match.groupdict().get("unit")
                key = (
                    fact_type,
                    raw_value,
                    str(unit_raw),
                )
                if key in seen:
                    continue
                seen.add(key)
                value = parse_number(raw_value)
                currency, unit, multiplier = unit_metadata(
                    unit_raw
                )
                normalized_value = (
                    value * multiplier
                    if value is not None
                    else None
                )
                review_reason = None
                review_required = False
                if fact_type == "EQUITY_RATIO" and value is not None and value > 100:
                    review_required = True
                    review_reason = "持股比例大于100%，疑似文本顺序或单位错误"
                facts.append(NumericFact(
                    numeric_fact_id=(
                        f"NUM-{candidate['company_id']}-"
                        f"{len(facts)+1:05d}"
                    ),
                    event_id=event_id,
                    candidate_event_id=str(
                        candidate["candidate_event_id"]
                    ),
                    company_id=str(candidate["company_id"]),
                    fact_type=fact_type,
                    raw_value_text=match.group(0),
                    numeric_value=value,
                    unit=unit,
                    currency=currency,
                    scale_multiplier=multiplier,
                    normalized_value=normalized_value,
                    value_type="DISCLOSED",
                    source_evidence_id=str(
                        evidence["evidence_id"]
                    ),
                    pdf_page_start=int(
                        evidence["pdf_page_start"]
                    ),
                    pdf_page_end=int(
                        evidence["pdf_page_end"]
                    ),
                    context_text=context_window(
                        normalized,
                        match.start(),
                        match.end(),
                    ),
                    confidence=confidence,
                    review_required=review_required,
                    review_reason=review_reason,
                ))
    return facts


def party_type_candidate(name: str) -> str:
    normalized = normalize_party_name(name)
    if re.fullmatch(r"[\u4e00-\u9fff]{2,4}", normalized):
        return "NATURAL_PERSON"
    if any(
        suffix in normalized
        for suffix in (
            "有限公司",
            "股份有限公司",
            "有限合伙",
            "合伙企业",
            "投资基金",
            "基金",
            "集团",
            "公司",
        )
    ):
        return "LEGAL_ENTITY_OR_FUND"
    return "UNRESOLVED"


def plausible_party_name(name: str) -> bool:
    normalized = normalize_party_name(name)
    if len(normalized) < 2 or len(normalized) > 45:
        return False
    if any(
        term in normalized
        for term in (
            "股权",
            "股份",
            "注册资本",
            "本次",
            "协议",
            "工商",
            "董事会",
            "股东会",
            "万元",
            "万股",
            "持股",
            "营业执照",
            "市场监督管理局",
            "工商登记",
            "基本情况",
            "法定程序",
            "财务状况",
            "资产负债",
            "对发行人",
            "吸收合并前",
            "吸收合并后",
            "换发",
        )
    ):
        return False
    return bool(
        re.search(r"[\u4e00-\u9fffA-Za-z]", normalized)
    )


PARTY_PATTERNS = [
    (
        "TRANSFEROR",
        re.compile(
            r"(?P<name>[\u4e00-\u9fffA-Za-z0-9（）()·]{2,35}?)"
            r"(?:将其持有的|将所持|转让其持有的)"
        ),
        0.86,
    ),
    (
        "TRANSFEREE",
        re.compile(
            r"(?:转让给|受让方为|转让予)"
            r"(?P<name>[\u4e00-\u9fffA-Za-z0-9（）()·]{2,35}?)"
            r"(?=，|。|；|持有|支付|受让|$)"
        ),
        0.88,
    ),
    (
        "TRANSFEREE",
        re.compile(
            r"(?P<name>[\u4e00-\u9fffA-Za-z0-9（）()·]{2,35}?)"
            r"(?:受让|取得)(?:了)?[^。；]{0,10}?"
            r"(?:公司)?(?:股权|股份)"
        ),
        0.82,
    ),
    (
        "INVESTOR",
        re.compile(
            r"(?P<name>[\u4e00-\u9fffA-Za-z0-9（）()·]{2,35}?)"
            r"(?:认购|缴纳出资|增资入股)"
        ),
        0.84,
    ),
    (
        "ABSORBED_ENTITY",
        re.compile(
            r"(?:"
            r"吸收合并(?!前|后|的|情况|程序|影响)"
            r"|被吸收合并方(?:为)?"
            r"|吸收合并对象为"
            r")"
            r"(?P<name>[\u4e00-\u9fffA-Za-z0-9（）()·]{2,30}?)"
            r"(?=，|。|；|注册资本|注销|被注销|后|$)"
        ),
        0.92,
    ),
    (
        "FOUNDER_OR_CONTRIBUTOR",
        re.compile(
            r"由(?P<name>[\u4e00-\u9fffA-Za-z0-9（）()·、和及与]{2,80}?)"
            r"(?:共同)?(?:出资|投资)设立"
        ),
        0.90,
    ),
]


def party_identity_key(name: str) -> str:
    normalized = normalize_party_name(name)
    return re.sub(
        r"(?:股份有限公司|有限责任公司|有限公司|有限)$",
        "",
        normalized,
    )


def role_specific_party_name(
    role: str,
    name: str,
) -> str:
    normalized = normalize_party_name(name)
    if role == "ABSORBED_ENTITY":
        normalized = re.split(
            r"(?:注册资本|基本情况|注销|财务|资产|负债|"
            r"业务|影响|程序|股东|后)",
            normalized,
            maxsplit=1,
        )[0]
    return normalized.strip("，。；、:：()（）")


def party_match_excluded(
    role: str,
    matched_text: str,
) -> bool:
    normalized = normalize_text(matched_text)
    if role == "TRANSFEREE" and any(
        term in normalized
        for term in (
            "营业执照",
            "市场监督管理局",
            "换发",
            "工商登记",
        )
    ):
        return True
    if role == "ABSORBED_ENTITY" and any(
        term in normalized
        for term in (
            "吸收合并前",
            "吸收合并后",
            "吸收合并的影响",
            "吸收合并程序",
        )
    ):
        return True
    return False



def split_party_names(value: str) -> list[str]:
    names = re.split(r"[、，,；;和及与]", value)
    return [
        normalize_party_name(name)
        for name in names
        if plausible_party_name(name)
    ]


def extract_parties(
    event_id: str,
    candidate: dict[str, Any],
    evidence_rows: list[dict[str, Any]],
) -> list[PartyRecord]:
    parties: list[PartyRecord] = []
    seen: set[tuple[str, str]] = set()

    for evidence in evidence_rows:
        text = re.sub(
            r"\s+",
            "",
            str(evidence.get("evidence_text") or ""),
        )
        for role, pattern, confidence in PARTY_PATTERNS:
            for match in pattern.finditer(text):
                matched_text = match.group(0)
                if party_match_excluded(
                    role,
                    matched_text,
                ):
                    continue
                raw = match.group("name")
                names = (
                    split_party_names(raw)
                    if role == "FOUNDER_OR_CONTRIBUTOR"
                    else [
                        role_specific_party_name(
                            role,
                            raw,
                        )
                    ]
                )
                for name in names:
                    if not plausible_party_name(name):
                        continue
                    key = (
                        role,
                        party_identity_key(name),
                    )
                    if key in seen:
                        continue
                    seen.add(key)
                    parties.append(PartyRecord(
                        party_id=(
                            f"PTY-{candidate['company_id']}-"
                            f"{len(parties)+1:05d}"
                        ),
                        event_id=event_id,
                        candidate_event_id=str(
                            candidate["candidate_event_id"]
                        ),
                        company_id=str(candidate["company_id"]),
                        party_name_raw=raw,
                        party_name_normalized=name,
                        party_role=role,
                        party_type_candidate=party_type_candidate(
                            name
                        ),
                        source_evidence_id=str(
                            evidence["evidence_id"]
                        ),
                        pdf_page_start=int(
                            evidence["pdf_page_start"]
                        ),
                        pdf_page_end=int(
                            evidence["pdf_page_end"]
                        ),
                        evidence_text=matched_text,
                        confidence=confidence,
                        review_required=confidence < 0.8,
                        review_reason=(
                            "参与方由弱规则抽取，需确认"
                            if confidence < 0.8
                            else None
                        ),
                    ))

    for service_name in candidate.get(
        "service_provider_mentions", []
    ) or []:
        name = normalize_party_name(str(service_name))
        if not plausible_party_name(name):
            continue
        key = ("SERVICE_PROVIDER", name)
        if key in seen:
            continue
        seen.add(key)
        primary = evidence_rows[0] if evidence_rows else {}
        parties.append(PartyRecord(
            party_id=(
                f"PTY-{candidate['company_id']}-"
                f"{len(parties)+1:05d}"
            ),
            event_id=event_id,
            candidate_event_id=str(
                candidate["candidate_event_id"]
            ),
            company_id=str(candidate["company_id"]),
            party_name_raw=str(service_name),
            party_name_normalized=name,
            party_role="SERVICE_PROVIDER",
            party_type_candidate="PROFESSIONAL_SERVICE_PROVIDER",
            source_evidence_id=str(
                primary.get("evidence_id") or ""
            ),
            pdf_page_start=int(
                primary.get("pdf_page_start")
                or candidate["pdf_page_start"]
            ),
            pdf_page_end=int(
                primary.get("pdf_page_end")
                or candidate["pdf_page_end"]
            ),
            evidence_text=str(service_name),
            confidence=0.95,
            review_required=False,
            review_reason=None,
        ))

    return parties


def normalize_date_text(value: str) -> str:
    normalized = normalize_text(value)
    match = re.match(
        r"(?P<year>(?:19|20)\d{2})年"
        r"(?:(?P<month>\d{1,2})月)?"
        r"(?:(?P<day>\d{1,2})日)?",
        normalized,
    )
    if match:
        parts = [match.group("year")]
        if match.group("month"):
            parts.append(match.group("month").zfill(2))
        if match.group("day"):
            parts.append(match.group("day").zfill(2))
        return "-".join(parts)
    return normalized


def extract_date_facts(
    event_id: str,
    candidate: dict[str, Any],
    evidence_rows: list[dict[str, Any]],
) -> list[DateFact]:
    output: list[DateFact] = []
    seen: set[tuple[str, str]] = set()

    for role, values in (
        candidate.get("event_dates") or {}
    ).items():
        for value in values or []:
            key = (str(role), str(value))
            if key in seen:
                continue
            seen.add(key)
            output.append(DateFact(
                date_fact_id=(
                    f"DTE-{candidate['company_id']}-"
                    f"{len(output)+1:05d}"
                ),
                event_id=event_id,
                candidate_event_id=str(
                    candidate["candidate_event_id"]
                ),
                company_id=str(candidate["company_id"]),
                date_role=str(role),
                date_text=str(value),
                date_normalized=normalize_date_text(
                    str(value)
                ),
                source="CANDIDATE_DATE_ROLE",
                source_evidence_id=None,
                pdf_page_start=int(candidate["pdf_page_start"]),
                pdf_page_end=int(candidate["pdf_page_end"]),
                confidence=0.96,
            ))

    if candidate.get("event_period"):
        key = (
            "event_period",
            str(candidate["event_period"]),
        )
        if key not in seen:
            seen.add(key)
            output.append(DateFact(
                date_fact_id=(
                    f"DTE-{candidate['company_id']}-"
                    f"{len(output)+1:05d}"
                ),
                event_id=event_id,
                candidate_event_id=str(
                    candidate["candidate_event_id"]
                ),
                company_id=str(candidate["company_id"]),
                date_role="event_period",
                date_text=str(candidate["event_period"]),
                date_normalized=str(
                    candidate["event_period"]
                ),
                source="CANDIDATE_PRIMARY_DATE",
                source_evidence_id=None,
                pdf_page_start=int(candidate["pdf_page_start"]),
                pdf_page_end=int(candidate["pdf_page_end"]),
                confidence=0.98,
            ))

    for evidence in evidence_rows:
        text = re.sub(
            r"\s+",
            "",
            str(evidence.get("evidence_text") or ""),
        )
        for role, patterns in DATE_ROLE_PATTERNS.items():
            for pattern in patterns:
                for match in re.finditer(pattern, text):
                    value = match.group("date")
                    key = (role, value)
                    if key in seen:
                        continue
                    seen.add(key)
                    output.append(DateFact(
                        date_fact_id=(
                            f"DTE-{candidate['company_id']}-"
                            f"{len(output)+1:05d}"
                        ),
                        event_id=event_id,
                        candidate_event_id=str(
                            candidate["candidate_event_id"]
                        ),
                        company_id=str(candidate["company_id"]),
                        date_role=role,
                        date_text=value,
                        date_normalized=normalize_date_text(
                            value
                        ),
                        source="EVIDENCE_REGEX",
                        source_evidence_id=str(
                            evidence["evidence_id"]
                        ),
                        pdf_page_start=int(
                            evidence["pdf_page_start"]
                        ),
                        pdf_page_end=int(
                            evidence["pdf_page_end"]
                        ),
                        confidence=0.88,
                    ))
    return output


def transaction_date(
    candidate: dict[str, Any],
    date_facts: list[DateFact],
) -> tuple[str | None, str | None]:
    preferred = [
        str(candidate.get("event_date_primary_role") or ""),
        "registration_date",
        "effective_date",
        "resolution_date",
        "agreement_date",
        "event_period",
    ]
    for role in preferred:
        if not role:
            continue
        values = [
            item.date_normalized
            for item in date_facts
            if item.date_role == role
        ]
        if values:
            return values[0], role
    return (
        str(candidate.get("event_period"))
        if candidate.get("event_period")
        else None,
        "event_period"
        if candidate.get("event_period")
        else None,
    )


def facts_by_type(
    facts: list[NumericFact],
) -> dict[str, list[str]]:
    result: dict[str, list[str]] = defaultdict(list)
    for fact in facts:
        result[fact.fact_type].append(
            fact.numeric_fact_id
        )
    return result


def parties_by_role(
    parties: list[PartyRecord],
) -> dict[str, list[str]]:
    result: dict[str, list[str]] = defaultdict(list)
    for party in parties:
        result[party.party_role].append(
            party.party_id
        )
    return result


def transaction_review_reasons(
    transaction_type: str,
    role_map: dict[str, list[str]],
    fact_map: dict[str, list[str]],
    candidate: dict[str, Any],
) -> list[str]:
    reasons: list[str] = []
    if transaction_type == "EQUITY_TRANSFER":
        if not role_map.get("TRANSFEROR"):
            reasons.append("股权转让未抽取到转让方")
        if not role_map.get("TRANSFEREE"):
            reasons.append("股权转让未抽取到受让方")
    elif transaction_type in {
        "CAPITAL_INCREASE",
        "DIRECTIONAL_FINANCING",
    }:
        if not (
            fact_map.get("CAPITAL_INCREASE_AMOUNT")
            or fact_map.get("REGISTERED_CAPITAL_AFTER")
            or fact_map.get("SHARE_QUANTITY")
        ):
            reasons.append("增资或发行未抽取到金额、注册资本或股份数量")
    elif transaction_type == "ABSORPTION_MERGER":
        if not role_map.get("ABSORBED_ENTITY"):
            reasons.append("吸收合并未抽取到被吸收方")
    elif transaction_type == "OVERALL_CHANGE":
        if not (
            fact_map.get("REGISTERED_CAPITAL")
            or fact_map.get("SHARE_QUANTITY")
            or fact_map.get("NET_ASSET_VALUE")
        ):
            reasons.append("整体变更未抽取到注册资本、折股数或净资产")
    elif transaction_type == "ESTABLISHMENT":
        if not (
            fact_map.get("REGISTERED_CAPITAL")
            or role_map.get("FOUNDER_OR_CONTRIBUTOR")
        ):
            reasons.append("设立事件未抽取到注册资本或出资人")

    if float(candidate.get("candidate_confidence") or 0) < 0.75:
        reasons.append("源候选置信度低于0.75")
    if (
        candidate.get("entity_scope_candidate")
        == "SUBSIDIARY_OR_OTHER_ENTITY_RISK"
    ):
        reasons.append("源候选存在其他主体风险")
    return reasons


def canonical_event_types(
    candidate: dict[str, Any],
) -> list[str]:
    values = list(
        candidate.get("event_type_candidates") or []
    )
    primary = str(
        candidate.get("event_type_candidate") or ""
    )
    if primary and primary not in values:
        values.insert(0, primary)

    # “股份公司设立”是整体变更的法律结果，不再生成第二条交易。
    if "OVERALL_CHANGE" in values:
        values = [
            value
            for value in values
            if value
            != "JOINT_STOCK_COMPANY_ESTABLISHMENT"
        ]

    # 解除代持已经包含代持关系的终止动作，避免同一候选重复交易。
    if "EQUITY_HOLDING_PROXY_RELEASE" in values:
        values = [
            value
            for value in values
            if value != "EQUITY_HOLDING_PROXY"
        ]

    output: list[str] = []
    for value in values:
        if value and value not in output:
            output.append(value)
    return output


def build_transactions(
    event_id: str,
    candidate: dict[str, Any],
    facts: list[NumericFact],
    parties: list[PartyRecord],
    dates: list[DateFact],
) -> list[TransactionRecord]:
    event_types = canonical_event_types(
        candidate
    )
    primary = str(
        candidate.get("event_type_candidate") or ""
    )

    transaction_types: list[str] = []
    for event_type in event_types:
        transaction_type = TRANSACTION_TYPES.get(event_type)
        if transaction_type and transaction_type not in transaction_types:
            transaction_types.append(transaction_type)

    if not transaction_types and primary:
        transaction_types.append(primary)

    date_value, date_role = transaction_date(
        candidate,
        dates,
    )
    fact_map = facts_by_type(facts)
    role_map = parties_by_role(parties)
    output: list[TransactionRecord] = []

    for sequence, transaction_type in enumerate(
        transaction_types,
        start=1,
    ):
        reasons = transaction_review_reasons(
            transaction_type,
            role_map,
            fact_map,
            candidate,
        )
        output.append(TransactionRecord(
            transaction_id=(
                f"TRX-{candidate['company_id']}-"
                f"{event_id.split('-')[-1]}-{sequence:02d}"
            ),
            event_id=event_id,
            candidate_event_id=str(
                candidate["candidate_event_id"]
            ),
            company_id=str(candidate["company_id"]),
            transaction_type=transaction_type,
            sequence_in_event=sequence,
            transaction_date=date_value,
            transaction_date_role=date_role,
            transferor_party_ids=role_map.get(
                "TRANSFEROR", []
            ),
            transferee_party_ids=role_map.get(
                "TRANSFEREE", []
            ),
            investor_party_ids=(
                role_map.get("INVESTOR", [])
                + role_map.get(
                    "FOUNDER_OR_CONTRIBUTOR", []
                )
            ),
            absorbed_party_ids=role_map.get(
                "ABSORBED_ENTITY", []
            ),
            registered_capital_before_fact_ids=(
                fact_map.get(
                    "REGISTERED_CAPITAL_BEFORE", []
                )
            ),
            registered_capital_after_fact_ids=(
                fact_map.get(
                    "REGISTERED_CAPITAL_AFTER", []
                )
            ),
            capital_increase_fact_ids=(
                fact_map.get(
                    "CAPITAL_INCREASE_AMOUNT", []
                )
            ),
            consideration_fact_ids=(
                fact_map.get("CONSIDERATION", [])
            ),
            share_quantity_fact_ids=(
                fact_map.get("SHARE_QUANTITY", [])
            ),
            share_price_fact_ids=(
                fact_map.get("SHARE_PRICE", [])
            ),
            equity_ratio_fact_ids=(
                fact_map.get("EQUITY_RATIO", [])
            ),
            transaction_status=(
                "AUTO_STRUCTURED"
                if not reasons
                else "AUTO_STRUCTURED_WITH_REVIEW"
            ),
            review_required=bool(reasons),
            review_reasons=reasons,
        ))
    return output


def validate_outputs(
    events: list[EventRecord],
    transactions: list[TransactionRecord],
    parties: list[PartyRecord],
    numerics: list[NumericFact],
    dates: list[DateFact],
    coverage_gaps: list[dict[str, Any]],
) -> dict[str, Any]:
    errors: list[str] = []
    warnings: list[str] = []

    event_ids = [item.event_id for item in events]
    if len(event_ids) != len(set(event_ids)):
        errors.append("事件ID重复")
    if len(events) != 26:
        errors.append(
            f"冻结事件数量应为26，实际为{len(events)}"
        )

    event_id_set = set(event_ids)
    transactions_by_event: dict[
        str, list[TransactionRecord]
    ] = defaultdict(list)
    for transaction in transactions:
        transactions_by_event[
            transaction.event_id
        ].append(transaction)
        if transaction.event_id not in event_id_set:
            errors.append(
                f"{transaction.transaction_id}引用不存在的事件"
            )

    for event_id, event_transactions in (
        transactions_by_event.items()
    ):
        transaction_types = {
            item.transaction_type
            for item in event_transactions
        }
        if (
            "OVERALL_CHANGE" in transaction_types
            and "JOINT_STOCK_ESTABLISHMENT"
            in transaction_types
        ):
            errors.append(
                f"{event_id}同时生成整体变更和股份公司设立重复交易"
            )

    for party in parties:
        event_transaction_types = {
            item.transaction_type
            for item in transactions_by_event.get(
                party.event_id,
                [],
            )
        }
        if (
            party.party_role == "TRANSFEREE"
            and "EQUITY_TRANSFER"
            not in event_transaction_types
        ):
            errors.append(
                f"{party.party_id}受让方角色不属于股权转让事件"
            )
        if (
            party.party_role == "ABSORBED_ENTITY"
            and "ABSORPTION_MERGER"
            not in event_transaction_types
        ):
            errors.append(
                f"{party.party_id}被吸收方角色不属于吸收合并事件"
            )
    for collection, label in [
        (parties, "参与方"),
        (numerics, "数值事实"),
        (dates, "日期事实"),
    ]:
        for item in collection:
            if item.event_id not in event_id_set:
                errors.append(
                    f"{label}记录引用不存在的事件：{item}"
                )

    coverage = sum(
        event.printed_page_start is not None
        and event.printed_page_end is not None
        for event in events
    ) / max(len(events), 1)
    if coverage != 1.0:
        errors.append(
            f"事件正文页码覆盖率不是100%：{coverage:.2%}"
        )

    if coverage_gaps:
        warnings.append(
            f"保留{len(coverage_gaps)}条招股书披露覆盖缺口"
        )

    return {
        "validation_status": (
            "FAILED" if errors else "PASSED"
        ),
        "errors": errors,
        "warnings": warnings,
        "event_count": len(events),
        "transaction_count": len(transactions),
        "party_count": len(parties),
        "numeric_fact_count": len(numerics),
        "date_fact_count": len(dates),
        "printed_page_coverage_rate": round(
            coverage,
            4,
        ),
    }


def review_rows(
    events: list[EventRecord],
    transactions: list[TransactionRecord],
    parties: list[PartyRecord],
    numerics: list[NumericFact],
    coverage_gaps: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []

    for transaction in transactions:
        if transaction.review_required:
            output.append({
                "review_id": (
                    f"REV-{transaction.transaction_id}"
                ),
                "company_id": transaction.company_id,
                "record_type": "TRANSACTION",
                "record_id": transaction.transaction_id,
                "event_id": transaction.event_id,
                "review_reason": "；".join(
                    transaction.review_reasons
                ),
                "auto_value": asdict(transaction),
                "manual_status": "PENDING",
                "manual_decision": None,
                "manual_note": None,
            })

    for party in parties:
        if party.review_required:
            output.append({
                "review_id": f"REV-{party.party_id}",
                "company_id": party.company_id,
                "record_type": "PARTY",
                "record_id": party.party_id,
                "event_id": party.event_id,
                "review_reason": party.review_reason,
                "auto_value": asdict(party),
                "manual_status": "PENDING",
                "manual_decision": None,
                "manual_note": None,
            })

    for fact in numerics:
        if fact.review_required:
            output.append({
                "review_id": (
                    f"REV-{fact.numeric_fact_id}"
                ),
                "company_id": fact.company_id,
                "record_type": "NUMERIC_FACT",
                "record_id": fact.numeric_fact_id,
                "event_id": fact.event_id,
                "review_reason": fact.review_reason,
                "auto_value": asdict(fact),
                "manual_status": "PENDING",
                "manual_decision": None,
                "manual_note": None,
            })

    for gap in coverage_gaps:
        output.append({
            "review_id": (
                f"REV-{gap['coverage_gap_id']}"
            ),
            "company_id": str(gap["company_id"]),
            "record_type": "DISCLOSURE_COVERAGE_GAP",
            "record_id": str(
                gap["coverage_gap_id"]
            ),
            "event_id": None,
            "review_reason": str(
                gap.get("gap_reason")
                or "招股书未逐项披露"
            ),
            "auto_value": gap,
            "manual_status": "PENDING",
            "manual_decision": (
                "ACCEPT_DISCLOSURE_LIMITATION"
            ),
            "manual_note": None,
        })

    output.sort(
        key=lambda item: (
            item["company_id"],
            item["record_type"],
            item["record_id"],
        )
    )
    return output


def run_extraction(
    repo_root: Path,
    freeze_id: str | None,
) -> int:
    repo_root = repo_root.expanduser().resolve()
    resolved_freeze_id, frozen_dir = resolve_frozen_dir(
        repo_root,
        freeze_id,
    )

    input_paths = {
        "events": (
            frozen_dir / "candidate_events_frozen.jsonl"
        ),
        "evidence": (
            frozen_dir / "candidate_evidence_frozen.jsonl"
        ),
        "coverage": (
            frozen_dir / "coverage_gaps_frozen.jsonl"
        ),
        "freeze_metrics": (
            frozen_dir / "freeze_metrics.json"
        ),
    }
    missing = [
        str(path)
        for path in input_paths.values()
        if not path.is_file()
    ]
    if missing:
        raise FileNotFoundError(
            "缺少冻结输入文件：\n" + "\n".join(missing)
        )

    freeze_metrics = read_json(
        input_paths["freeze_metrics"]
    )
    if freeze_metrics.get("freeze_status") != "FROZEN":
        raise ValueError("候选事件尚未成功冻结")

    candidates = read_jsonl(input_paths["events"])
    evidence = read_jsonl(input_paths["evidence"])
    coverage_gaps = read_jsonl(
        input_paths["coverage"]
    )
    evidence_by_candidate: dict[
        str, list[dict[str, Any]]
    ] = defaultdict(list)
    for row in evidence:
        evidence_by_candidate[
            str(row["candidate_event_id"])
        ].append(row)

    run_id = make_run_id()
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
    logs_dir = (
        repo_root
        / "logs"
        / "structured_events"
        / "runs"
        / run_id
    )
    for path in (auto_dir, validation_dir, logs_dir):
        path.mkdir(parents=True, exist_ok=False)

    company_sequences: Counter[str] = Counter()
    events: list[EventRecord] = []
    transactions: list[TransactionRecord] = []
    parties: list[PartyRecord] = []
    numerics: list[NumericFact] = []
    dates: list[DateFact] = []
    logs: list[dict[str, Any]] = []

    for candidate in sorted(
        candidates,
        key=lambda item: (
            str(item["company_id"]),
            int(item.get("pdf_page_start") or 0),
            str(item["candidate_event_id"]),
        ),
    ):
        company_id = str(candidate["company_id"])
        company_sequences[company_id] += 1
        event_id = (
            f"EVT-{company_id}-"
            f"{company_sequences[company_id]:04d}"
        )
        candidate_id = str(
            candidate["candidate_event_id"]
        )
        event_evidence = evidence_by_candidate.get(
            candidate_id,
            [],
        )
        extracted_parties = extract_parties(
            event_id,
            candidate,
            event_evidence,
        )
        extracted_numerics = extract_numeric_facts(
            event_id,
            candidate,
            event_evidence,
        )
        extracted_dates = extract_date_facts(
            event_id,
            candidate,
            event_evidence,
        )
        extracted_transactions = build_transactions(
            event_id,
            candidate,
            extracted_numerics,
            extracted_parties,
            extracted_dates,
        )

        event_reasons: list[str] = []
        if not event_evidence:
            event_reasons.append("冻结事件缺少证据记录")
        if not extracted_transactions:
            event_reasons.append("未生成交易记录")
        if (
            candidate.get("disclosure_scope", "")
            .startswith("PARTIAL")
        ):
            event_reasons.append("事件来源属于部分披露范围")

        primary_type = str(
            candidate.get("event_type_candidate") or ""
        )
        events.append(EventRecord(
            event_id=event_id,
            candidate_event_id=candidate_id,
            company_id=company_id,
            company_short_name=str(
                candidate.get("company_short_name") or ""
            ),
            event_type=primary_type,
            event_type_candidates=list(
                candidate.get(
                    "event_type_candidates"
                ) or [primary_type]
            ),
            event_group=EVENT_GROUPS.get(
                primary_type,
                "OTHER",
            ),
            event_period=(
                str(candidate["event_period"])
                if candidate.get("event_period")
                else None
            ),
            event_date_text=(
                str(candidate["event_date_text"])
                if candidate.get("event_date_text")
                else None
            ),
            event_date_primary_role=(
                str(candidate["event_date_primary_role"])
                if candidate.get(
                    "event_date_primary_role"
                )
                else None
            ),
            event_dates=dict(
                candidate.get("event_dates") or {}
            ),
            event_title=str(
                candidate.get("event_title") or ""
            ),
            disclosure_scope=str(
                candidate.get("disclosure_scope") or ""
            ),
            ordinal_labels=list(
                candidate.get("ordinal_labels") or []
            ),
            pdf_page_start=int(
                candidate["pdf_page_start"]
            ),
            pdf_page_end=int(
                candidate["pdf_page_end"]
            ),
            printed_page_start=(
                str(candidate["printed_page_start"])
                if candidate.get(
                    "printed_page_start"
                ) is not None
                else None
            ),
            printed_page_end=(
                str(candidate["printed_page_end"])
                if candidate.get(
                    "printed_page_end"
                ) is not None
                else None
            ),
            printed_page_value_type=str(
                candidate.get(
                    "printed_page_value_type"
                ) or ""
            ),
            primary_evidence_id=str(
                candidate["primary_evidence_id"]
            ),
            supporting_evidence_ids=list(
                candidate.get(
                    "supporting_evidence_ids"
                ) or []
            ),
            entity_scope=str(
                candidate.get(
                    "entity_scope_candidate"
                ) or ""
            ),
            candidate_confidence=float(
                candidate.get(
                    "candidate_confidence"
                ) or 0
            ),
            extraction_status=(
                "AUTO_STRUCTURED"
                if not event_reasons
                else "AUTO_STRUCTURED_WITH_REVIEW"
            ),
            review_required=bool(event_reasons),
            review_reasons=event_reasons,
        ))

        parties.extend(extracted_parties)
        numerics.extend(extracted_numerics)
        dates.extend(extracted_dates)
        transactions.extend(extracted_transactions)
        logs.append({
            "timestamp": now_iso(),
            "level": "INFO",
            "event": "event_structured",
            "run_id": run_id,
            "freeze_id": resolved_freeze_id,
            "company_id": company_id,
            "candidate_event_id": candidate_id,
            "event_id": event_id,
            "party_count": len(extracted_parties),
            "numeric_fact_count": len(
                extracted_numerics
            ),
            "date_fact_count": len(
                extracted_dates
            ),
            "transaction_count": len(
                extracted_transactions
            ),
        })

    validation = validate_outputs(
        events,
        transactions,
        parties,
        numerics,
        dates,
        coverage_gaps,
    )
    reviews = review_rows(
        events,
        transactions,
        parties,
        numerics,
        coverage_gaps,
    )

    metrics = {
        "metrics_version": "0.2",
        "run_id": run_id,
        "pipeline_version": PIPELINE_VERSION,
        "source_freeze_id": resolved_freeze_id,
        "batch_status": (
            "FAILED"
            if validation["validation_status"]
            == "FAILED"
            else (
                "READY_WITH_REVIEW"
                if reviews
                else "READY"
            )
        ),
        "event_count": len(events),
        "transaction_count": len(transactions),
        "party_count": len(parties),
        "numeric_fact_count": len(numerics),
        "date_fact_count": len(dates),
        "review_queue_count": len(reviews),
        "coverage_gap_count": len(coverage_gaps),
        "event_type_counts": dict(
            Counter(
                event.event_type
                for event in events
            )
        ),
        "transaction_type_counts": dict(
            Counter(
                transaction.transaction_type
                for transaction in transactions
            )
        ),
        "numeric_fact_type_counts": dict(
            Counter(
                fact.fact_type
                for fact in numerics
            )
        ),
        "party_role_counts": dict(
            Counter(
                party.party_role
                for party in parties
            )
        ),
        "printed_page_coverage_rate": (
            validation[
                "printed_page_coverage_rate"
            ]
        ),
        "validation_error_count": len(
            validation["errors"]
        ),
        "llm_called": False,
        "pevc_classification_performed": False,
        "numeric_validation_performed": False,
        "final_generated": False,
        "note": (
            "v0.2将整体变更与股份公司设立标准化为单一交易，"
            "收紧受让方和被吸收方抽取并提高明确出资人的置信度；"
            "未披露字段保持为空，不判断PE/VC，不生成Final。"
        ),
    }

    write_jsonl(
        auto_dir / "event_records_auto.jsonl",
        [asdict(item) for item in events],
    )
    write_jsonl(
        auto_dir / "transaction_records_auto.jsonl",
        [asdict(item) for item in transactions],
    )
    write_jsonl(
        auto_dir / "event_parties_auto.jsonl",
        [asdict(item) for item in parties],
    )
    write_jsonl(
        auto_dir / "numeric_facts_auto.jsonl",
        [asdict(item) for item in numerics],
    )
    write_jsonl(
        auto_dir / "event_dates_auto.jsonl",
        [asdict(item) for item in dates],
    )
    write_jsonl(
        auto_dir / "event_evidence_index.jsonl",
        evidence,
    )
    write_jsonl(
        validation_dir
        / "structured_event_review_queue.jsonl",
        reviews,
    )
    write_csv(
        validation_dir
        / "structured_event_review_queue.csv",
        reviews,
        [
            "review_id",
            "company_id",
            "record_type",
            "record_id",
            "event_id",
            "review_reason",
            "auto_value",
            "manual_status",
            "manual_decision",
            "manual_note",
        ],
    )
    write_csv(
        validation_dir
        / "event_records_preview.csv",
        [asdict(item) for item in events],
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
            "disclosure_scope",
            "extraction_status",
            "review_required",
            "review_reasons",
        ],
    )
    write_csv(
        validation_dir
        / "transaction_records_preview.csv",
        [asdict(item) for item in transactions],
        [
            "transaction_id",
            "event_id",
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
    write_json(
        validation_dir
        / "structured_event_metrics.json",
        metrics,
    )
    write_json(
        validation_dir
        / "structured_event_validation.json",
        validation,
    )
    write_jsonl(
        logs_dir / "structured_event_extraction.jsonl",
        logs,
    )
    write_json(
        logs_dir / "run_manifest.json",
        {
            "run_id": run_id,
            "pipeline_version": PIPELINE_VERSION,
            "source_freeze_id": resolved_freeze_id,
            "source_frozen_dir": str(frozen_dir),
            "started_and_completed_at": now_iso(),
            "batch_status": metrics["batch_status"],
            "source_events_sha256": sha256_file(
                input_paths["events"]
            ),
            "source_evidence_sha256": sha256_file(
                input_paths["evidence"]
            ),
            "llm_called": False,
        },
    )
    write_json(
        (
            repo_root
            / "logs"
            / "structured_events"
            / "latest_run.json"
        ),
        {
            "run_id": run_id,
            "batch_status": metrics["batch_status"],
            "source_freeze_id": resolved_freeze_id,
            "completed_at": now_iso(),
        },
    )

    print()
    print("冻结事件结构化字段抽取 v0.2 完成")
    print(f"运行ID：{run_id}")
    print(f"源Freeze ID：{resolved_freeze_id}")
    print(f"事件主表：{len(events)}")
    print(f"交易记录：{len(transactions)}")
    print(f"参与方：{len(parties)}")
    print(f"数值事实：{len(numerics)}")
    print(f"日期事实：{len(dates)}")
    print(f"覆盖缺口：{len(coverage_gaps)}")
    print(f"人工复核项：{len(reviews)}")
    print(
        "正文页码覆盖率："
        f"{validation['printed_page_coverage_rate']:.2%}"
    )
    print(
        f"验证错误：{len(validation['errors'])}"
    )
    print(f"批次状态：{metrics['batch_status']}")

    return (
        0
        if validation["validation_status"] == "PASSED"
        else 2
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "读取26条冻结候选，抽取事件主表、交易、参与方、"
            "日期及数值事实。"
        )
    )
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path("."),
    )
    parser.add_argument(
        "--freeze-id",
        default=None,
        help=(
            "可选；不提供时自动读取"
            "review/candidate_events/latest_frozen.json"
        ),
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        return run_extraction(
            repo_root=args.repo_root,
            freeze_id=args.freeze_id,
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
