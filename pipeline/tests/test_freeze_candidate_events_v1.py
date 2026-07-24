from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from pipeline.freeze_candidate_events import (
    EXPECTED_COUNTS,
    apply_review_patch,
    validate_frozen,
)


def candidate(
    company_id: str,
    candidate_id: str,
    event_type: str,
    period: str,
    title: str,
    recovery: bool = False,
) -> dict:
    signals = (
        ["page_timeline_recovery"]
        if recovery
        else ["heading_section"]
    )
    return {
        "candidate_event_id": candidate_id,
        "company_id": company_id,
        "event_type_candidate": event_type,
        "event_type_candidates": [event_type],
        "event_period": period,
        "event_title": title,
        "pdf_page_start": 10,
        "pdf_page_end": 10,
        "printed_page_start": "9",
        "printed_page_end": "9",
        "printed_page_value_type": "DISCLOSED",
        "primary_evidence_id": f"E-{candidate_id}",
        "supporting_evidence_ids": [],
        "matched_signals": signals,
    }


class FreezeCandidateEventsV1Tests(unittest.TestCase):
    def build_candidates(self) -> list[dict]:
        rows: list[dict] = []

        rows.extend([
            candidate(
                "001282", "A1",
                "LIMITED_COMPANY_ESTABLISHMENT",
                "2004-06-18",
                "有限公司设立情况",
            ),
            candidate(
                "001282", "A2",
                "LIMITED_COMPANY_ESTABLISHMENT",
                "1999",
                "1999，三联有限成立",
                True,
            ),
            candidate(
                "001282", "A3",
                "OVERALL_CHANGE",
                "2018-11",
                "整体变更",
            ),
            candidate(
                "001282", "A4",
                "CAPITAL_INCREASE",
                "2019-01",
                "增资",
            ),
        ])

        rows.extend([
            candidate(
                "301581", "B1",
                "LIMITED_COMPANY_ESTABLISHMENT",
                "2012-06",
                "2012-06，发行人前身设立",
                True,
            ),
            candidate(
                "301581", "B2",
                "LIMITED_COMPANY_ESTABLISHMENT",
                "2009-07-06",
                "昆山谷捷有限公司成立",
            ),
        ])
        for index, event_type in enumerate([
            "ABSORPTION_MERGER",
            "EQUITY_TRANSFER",
            "CAPITAL_INCREASE",
            "CAPITAL_INCREASE",
            "OVERALL_CHANGE",
        ], start=3):
            rows.append(candidate(
                "301581",
                f"B{index}",
                event_type,
                f"2021-{index:02d}",
                event_type,
            ))

        rows.extend([
            candidate(
                "603418", "C1",
                "LIMITED_COMPANY_ESTABLISHMENT",
                "1992-12-04",
                "有限责任公司的设立情况",
            ),
            candidate(
                "603418", "C2",
                "LIMITED_COMPANY_ESTABLISHMENT",
                "1992-10-22",
                "1992-10-22，友升有限成立",
                True,
            ),
        ])
        for index, event_type in enumerate([
            "OVERALL_CHANGE",
            "CAPITAL_INCREASE",
            "CAPITAL_INCREASE",
            "EQUITY_TRANSFER",
        ], start=3):
            rows.append(candidate(
                "603418",
                f"C{index}",
                event_type,
                f"2020-{index:02d}",
                event_type,
            ))

        for company_id, count in {
            "688758": 7,
            "688775": 3,
            "920100": 1,
            "920116": 1,
        }.items():
            for index in range(1, count + 1):
                rows.append(candidate(
                    company_id,
                    f"{company_id}-{index}",
                    "CAPITAL_INCREASE",
                    f"2020-{index:02d}",
                    "事件",
                ))

        return rows

    def test_review_patch_freezes_expected_counts(self) -> None:
        candidates = self.build_candidates()
        frozen, patches = apply_review_patch(
            candidates,
            "SOURCE",
            "FREEZE",
        )
        counts = {}
        for company_id in EXPECTED_COUNTS:
            counts[company_id] = sum(
                item["company_id"] == company_id
                for item in frozen
            )
        self.assertEqual(counts, EXPECTED_COUNTS)
        self.assertEqual(len(frozen), 26)
        self.assertEqual(len(candidates) - len(frozen), 3)

    def test_selected_establishment_events(self) -> None:
        frozen, _ = apply_review_patch(
            self.build_candidates(),
            "SOURCE",
            "FREEZE",
        )
        periods = {
            item["company_id"]: item["event_period"]
            for item in frozen
            if item["event_type_candidate"]
            == "LIMITED_COMPANY_ESTABLISHMENT"
        }
        self.assertEqual(periods["001282"], "2004-06-18")
        self.assertEqual(periods["301581"], "2012-06")
        self.assertEqual(periods["603418"], "1992-12-04")

    def test_validation_passes_with_complete_evidence(self) -> None:
        frozen, _ = apply_review_patch(
            self.build_candidates(),
            "SOURCE",
            "FREEZE",
        )
        evidence = [
            {
                "evidence_id": item["primary_evidence_id"],
                "candidate_event_id": item[
                    "candidate_event_id"
                ],
            }
            for item in frozen
        ]
        result = validate_frozen(frozen, evidence)
        self.assertEqual(
            result["validation_status"],
            "PASSED",
        )
        self.assertEqual(
            result["printed_page_coverage_rate"],
            1.0,
        )


if __name__ == "__main__":
    unittest.main()
