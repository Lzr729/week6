from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from pipeline.structured_events import (
    build_transactions,
    canonical_event_types,
    extract_date_facts,
    extract_numeric_facts,
    extract_parties,
    normalize_date_text,
    validate_outputs,
)


def candidate(
    event_type: str,
    event_types: list[str] | None = None,
) -> dict:
    return {
        "candidate_event_id": "CE-123456-0001",
        "company_id": "123456",
        "company_short_name": "测试公司",
        "event_type_candidate": event_type,
        "event_type_candidates": (
            event_types or [event_type]
        ),
        "event_period": "2021-04",
        "event_date_text": "2021年4月",
        "event_date_primary_role": "event_period",
        "event_dates": {
            "event_period": ["2021年4月"]
        },
        "event_title": "测试事件",
        "disclosure_scope": "FULL",
        "ordinal_labels": [],
        "pdf_page_start": 10,
        "pdf_page_end": 11,
        "printed_page_start": "9",
        "printed_page_end": "10",
        "printed_page_value_type": "DISCLOSED",
        "primary_evidence_id": "E1",
        "supporting_evidence_ids": [],
        "entity_scope_candidate": "ISSUER_OR_UNRESOLVED",
        "candidate_confidence": 0.92,
        "service_provider_mentions": [],
    }


def evidence(text: str) -> list[dict]:
    return [{
        "evidence_id": "E1",
        "candidate_event_id": "CE-123456-0001",
        "company_id": "123456",
        "pdf_page_start": 10,
        "pdf_page_end": 11,
        "evidence_text": text,
    }]


class StructuredEventExtractionV02Tests(unittest.TestCase):
    def test_numeric_capital_before_after(self) -> None:
        facts = extract_numeric_facts(
            "EVT-123456-0001",
            candidate("CAPITAL_INCREASE"),
            evidence(
                "注册资本由1,000万元增加至1,500万元，"
                "新增注册资本500万元。"
            ),
        )
        by_type = {}
        for fact in facts:
            by_type.setdefault(
                fact.fact_type, []
            ).append(fact)
        self.assertEqual(
            by_type["REGISTERED_CAPITAL_BEFORE"][0]
            .normalized_value,
            10_000_000,
        )
        self.assertEqual(
            by_type["REGISTERED_CAPITAL_AFTER"][0]
            .normalized_value,
            15_000_000,
        )
        self.assertEqual(
            by_type["CAPITAL_INCREASE_AMOUNT"][0]
            .normalized_value,
            5_000_000,
        )

    def test_equity_transfer_parties(self) -> None:
        parties = extract_parties(
            "EVT-123456-0001",
            candidate("EQUITY_TRANSFER"),
            evidence(
                "张三将其持有的公司10%股权转让给测试投资有限公司。"
            ),
        )
        roles = {
            party.party_role: party.party_name_normalized
            for party in parties
        }
        self.assertEqual(roles["TRANSFEROR"], "张三")
        self.assertEqual(
            roles["TRANSFEREE"],
            "测试投资有限公司",
        )

    def test_absorption_merger_party(self) -> None:
        parties = extract_parties(
            "EVT-123456-0001",
            candidate("ABSORPTION_MERGER"),
            evidence(
                "公司吸收合并昆山测试有限公司，"
                "昆山测试依法注销。"
            ),
        )
        absorbed = [
            item for item in parties
            if item.party_role == "ABSORBED_ENTITY"
        ]
        self.assertEqual(len(absorbed), 1)

    def test_combined_event_creates_two_transactions(self) -> None:
        cand = candidate(
            "EQUITY_TRANSFER",
            ["EQUITY_TRANSFER", "CAPITAL_INCREASE"],
        )
        facts = extract_numeric_facts(
            "EVT-123456-0001",
            cand,
            evidence(
                "张三将股权转让给李四，"
                "新增注册资本100万元。"
            ),
        )
        parties = extract_parties(
            "EVT-123456-0001",
            cand,
            evidence(
                "张三将其持有的股权转让给李四。"
            ),
        )
        dates = extract_date_facts(
            "EVT-123456-0001",
            cand,
            evidence("2021年4月完成工商变更。"),
        )
        transactions = build_transactions(
            "EVT-123456-0001",
            cand,
            facts,
            parties,
            dates,
        )
        self.assertEqual(
            {
                item.transaction_type
                for item in transactions
            },
            {"EQUITY_TRANSFER", "CAPITAL_INCREASE"},
        )

    def test_date_normalization(self) -> None:
        self.assertEqual(
            normalize_date_text("2021年4月3日"),
            "2021-04-03",
        )

    def test_validation_requires_26_events(self) -> None:
        result = validate_outputs(
            [],
            [],
            [],
            [],
            [],
            [],
        )
        self.assertEqual(
            result["validation_status"],
            "FAILED",
        )
        self.assertTrue(result["errors"])

    def test_overall_change_suppresses_joint_stock_transaction(self) -> None:
        cand = candidate(
            "OVERALL_CHANGE",
            [
                "OVERALL_CHANGE",
                "JOINT_STOCK_COMPANY_ESTABLISHMENT",
            ],
        )
        self.assertEqual(
            canonical_event_types(cand),
            ["OVERALL_CHANGE"],
        )
        transactions = build_transactions(
            "EVT-123456-0001",
            cand,
            [],
            [],
            [],
        )
        self.assertEqual(len(transactions), 1)
        self.assertEqual(
            transactions[0].transaction_type,
            "OVERALL_CHANGE",
        )

    def test_license_context_not_transferee(self) -> None:
        parties = extract_parties(
            "EVT-123456-0001",
            candidate("OVERALL_CHANGE"),
            evidence(
                "赛分科技取得了江苏省市场监督管理局换发的营业执照。"
            ),
        )
        self.assertFalse(
            any(
                item.party_role == "TRANSFEREE"
                for item in parties
            )
        )

    def test_absorbed_entity_deduplicated(self) -> None:
        parties = extract_parties(
            "EVT-123456-0001",
            candidate("ABSORPTION_MERGER"),
            evidence(
                "公司吸收合并昆山测试有限公司，"
                "被吸收合并方为昆山测试有限公司；"
                "吸收合并前昆山测试基本情况如下。"
            ),
        )
        absorbed = [
            item
            for item in parties
            if item.party_role == "ABSORBED_ENTITY"
        ]
        self.assertEqual(len(absorbed), 1)

    def test_explicit_founder_not_reviewed(self) -> None:
        parties = extract_parties(
            "EVT-123456-0001",
            candidate("LIMITED_COMPANY_ESTABLISHMENT"),
            evidence(
                "由甲公司、乙公司共同出资设立。"
            ),
        )
        founders = [
            item for item in parties
            if item.party_role
            == "FOUNDER_OR_CONTRIBUTOR"
        ]
        self.assertEqual(len(founders), 2)
        self.assertTrue(
            all(
                not item.review_required
                for item in founders
            )
        )


if __name__ == "__main__":
    unittest.main()
