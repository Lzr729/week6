import tempfile
import unittest
from pathlib import Path

from pipeline.final_submission import (
    build_equity_snapshot_status,
    rows_by_company,
)


class FinalSubmissionAssemblyV01Tests(
    unittest.TestCase
):
    def test_rows_grouped_by_company(self) -> None:
        grouped = rows_by_company([
            {
                "company_id": "001282",
                "event_id": "E1",
            },
            {
                "company_id": "001282",
                "event_id": "E2",
            },
            {
                "company_id": "603418",
                "event_id": "E3",
            },
        ])
        self.assertEqual(
            len(grouped["001282"]),
            2,
        )
        self.assertEqual(
            len(grouped["603418"]),
            1,
        )

    def test_snapshot_not_invented_without_linkage(
        self,
    ) -> None:
        payload = build_equity_snapshot_status(
            "001282",
            [{"event_id": "E1"}],
            [{"party_id": "P1"}],
            [{
                "numeric_fact_id": "N1",
                "fact_type": "EQUITY_RATIO",
                "normalized_value": 0.25,
            }],
        )
        self.assertEqual(
            payload["snapshot_count"],
            0,
        )
        self.assertIn(
            "PARTY_RATIO_LINKAGE_NOT_CONFIRMED",
            payload["snapshot_status"],
        )

    def test_snapshot_absent_without_ratio(
        self,
    ) -> None:
        payload = build_equity_snapshot_status(
            "001282",
            [],
            [],
            [],
        )
        self.assertEqual(
            payload["snapshot_count"],
            0,
        )
        self.assertIn(
            "NOT_CONSTRUCTIBLE",
            payload["snapshot_status"],
        )


if __name__ == "__main__":
    unittest.main()
