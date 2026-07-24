from __future__ import annotations

import unittest
from pathlib import Path

from pipeline.offline_replay import (
    compare_id_sets,
    load_gold_decisions,
)


class OfflineReplayTests(unittest.TestCase):
    def test_compare_id_sets(self) -> None:
        result = compare_id_sets(
            [{"id": "A"}, {"id": "B"}],
            [{"id": "A"}, {"id": "B"}],
            "id",
        )
        self.assertEqual(
            result["common_count"],
            2,
        )
        self.assertEqual(
            result["auto_only_ids"],
            [],
        )
        self.assertEqual(
            result["gold_only_ids"],
            [],
        )

    def test_gold_decisions_have_no_open_items(
        self,
    ) -> None:
        repo_root = Path(__file__).resolve().parents[2]
        decisions = load_gold_decisions(
            repo_root
        )
        self.assertEqual(len(decisions), 19)
        self.assertTrue(
            all(
                row["human_review_completed"]
                for row in decisions.values()
            )
        )


if __name__ == "__main__":
    unittest.main()
