from __future__ import annotations

import unittest

from pipeline.numeric_validation import (
    validate_capital_increase,
    validate_company_chronology,
    validate_duplicate_numeric_facts,
    validate_equity_ratios,
    invalid_numeric_facts,
    validate_price_quantity_consideration,
)


def numeric_fact(
    fact_id: str,
    fact_type: str,
    value: float,
    event_id: str = "EVT-1",
) -> dict:
    return {
        "numeric_fact_id": fact_id,
        "event_id": event_id,
        "candidate_event_id": "CE-1",
        "company_id": "123456",
        "fact_type": fact_type,
        "normalized_value": value,
        "source_evidence_id": "E1",
    }


def transaction() -> dict:
    return {
        "transaction_id": "TRX-1",
        "event_id": "EVT-1",
        "company_id": "123456",
        "transaction_type": "CAPITAL_INCREASE",
        "registered_capital_before_fact_ids": ["B"],
        "registered_capital_after_fact_ids": ["A"],
        "capital_increase_fact_ids": ["I"],
        "share_quantity_fact_ids": [],
        "share_price_fact_ids": [],
        "consideration_fact_ids": [],
    }


class NumericValidationV012Tests(unittest.TestCase):
    def test_capital_increase_passes(self) -> None:
        facts = {
            "B": numeric_fact(
                "B", "REGISTERED_CAPITAL_BEFORE", 10_000_000
            ),
            "I": numeric_fact(
                "I", "CAPITAL_INCREASE_AMOUNT", 5_000_000
            ),
            "A": numeric_fact(
                "A", "REGISTERED_CAPITAL_AFTER", 15_000_000
            ),
        }
        results, calculated = validate_capital_increase(
            transaction(),
            facts,
            1,
            1.0,
            0.0001,
        )
        self.assertEqual(
            results[0].validation_status,
            "PASSED",
        )
        self.assertEqual(
            calculated[0].calculated_value,
            15_000_000,
        )

    def test_capital_increase_mismatch_fails(self) -> None:
        facts = {
            "B": numeric_fact(
                "B", "REGISTERED_CAPITAL_BEFORE", 10_000_000
            ),
            "I": numeric_fact(
                "I", "CAPITAL_INCREASE_AMOUNT", 5_000_000
            ),
            "A": numeric_fact(
                "A", "REGISTERED_CAPITAL_AFTER", 14_000_000
            ),
        }
        results, _ = validate_capital_increase(
            transaction(),
            facts,
            1,
            1.0,
            0.0001,
        )
        self.assertEqual(
            results[0].validation_status,
            "FAILED",
        )

    def test_price_quantity_consideration_passes(self) -> None:
        tx = transaction()
        tx["share_quantity_fact_ids"] = ["Q"]
        tx["share_price_fact_ids"] = ["P"]
        tx["consideration_fact_ids"] = ["C"]
        facts = {
            "Q": numeric_fact("Q", "SHARE_QUANTITY", 1_000_000),
            "P": numeric_fact("P", "SHARE_PRICE", 2.5),
            "C": numeric_fact("C", "CONSIDERATION", 2_500_000),
        }
        results, calculated = (
            validate_price_quantity_consideration(
                tx,
                facts,
                1,
                1.0,
                0.005,
            )
        )
        self.assertEqual(
            results[0].validation_status,
            "PASSED",
        )
        self.assertEqual(
            calculated[0].calculated_value,
            2_500_000,
        )

    def test_ratio_out_of_range_fails(self) -> None:
        facts = [
            numeric_fact(
                "R",
                "EQUITY_RATIO",
                1.2,
            )
        ]
        results = validate_equity_ratios(
            facts,
            1,
        )
        self.assertEqual(
            results[0].validation_status,
            "FAILED",
        )

    def test_duplicate_fact_requires_review(self) -> None:
        facts = [
            numeric_fact("X1", "EQUITY_RATIO", 0.5),
            numeric_fact("X2", "EQUITY_RATIO", 0.5),
        ]
        results = validate_duplicate_numeric_facts(
            facts,
            1,
        )
        self.assertEqual(len(results), 1)
        self.assertEqual(
            results[0].validation_status,
            "REVIEW_REQUIRED",
        )

    def test_chronology_disorder_is_informational_legacy(self) -> None:
        events = [
            {
                "event_id": "E1",
                "company_id": "123456",
                "event_period": "2021-01",
                "pdf_page_start": 20,
            },
            {
                "event_id": "E2",
                "company_id": "123456",
                "event_period": "2020-01",
                "pdf_page_start": 21,
            },
        ]
        results = validate_company_chronology(
            events,
            1,
        )
        self.assertEqual(
            results[0].validation_status,
            "INFORMATIONAL",
        )

    def test_rule_status_counts_are_json_serializable(self) -> None:
        from collections import Counter
        import json

        statuses = [
            ("NV-CAP-001", "PASSED"),
            ("NV-CAP-001", "NOT_APPLICABLE"),
            ("NV-RATIO-001", "PASSED"),
        ]
        payload = {
            rule_id: dict(
                Counter(
                    status
                    for current_rule, status in statuses
                    if current_rule == rule_id
                )
            )
            for rule_id in sorted({
                rule_id for rule_id, _ in statuses
            })
        }
        encoded = json.dumps(payload, ensure_ascii=False)
        self.assertIn("NV-CAP-001", encoded)
        self.assertNotIn("(", encoded)

    def test_null_ratio_is_not_applicable(self) -> None:
        fact = numeric_fact(
            "R-NULL",
            "EQUITY_RATIO",
            0.5,
        )
        fact["normalized_value"] = None
        results = validate_equity_ratios(
            [fact],
            1,
        )
        self.assertEqual(
            results[0].validation_status,
            "NOT_APPLICABLE",
        )
        self.assertFalse(
            results[0].review_required
        )

    def test_null_ratio_is_quarantined(self) -> None:
        fact = numeric_fact(
            "R-NULL",
            "EQUITY_RATIO",
            0.5,
        )
        fact["normalized_value"] = None
        quarantined = invalid_numeric_facts(
            [fact]
        )
        self.assertEqual(len(quarantined), 1)
        self.assertEqual(
            quarantined[0]["record_status"],
            "QUARANTINED_INVALID_AUTO_FACT",
        )

    def test_duplicate_null_facts_are_ignored(self) -> None:
        facts = [
            numeric_fact(
                "R1", "EQUITY_RATIO", 0.5
            ),
            numeric_fact(
                "R2", "EQUITY_RATIO", 0.5
            ),
        ]
        for fact in facts:
            fact["normalized_value"] = None
        results = validate_duplicate_numeric_facts(
            facts,
            1,
        )
        self.assertEqual(results, [])

    def test_chronology_disorder_is_informational(self) -> None:
        events = [
            {
                "event_id": "E1",
                "company_id": "123456",
                "event_period": "2021-01",
                "pdf_page_start": 20,
            },
            {
                "event_id": "E2",
                "company_id": "123456",
                "event_period": "2020-01",
                "pdf_page_start": 21,
            },
        ]
        results = validate_company_chronology(
            events,
            1,
        )
        self.assertEqual(
            results[0].validation_status,
            "INFORMATIONAL",
        )
        self.assertFalse(
            results[0].review_required
        )


if __name__ == "__main__":
    unittest.main()
