from __future__ import annotations

import unittest

from pipeline.chapter_location import (
    PaginationRecord,
    assign_hierarchy,
    build_all_candidates,
    build_heading_records,
    build_review_items,
    enrich_from_section_context,
    infer_pagination,
    load_json,
    locate_issuer_master_section,
    pagination_candidates_for_page,
)


class ChapterLocationV05Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.rules = load_json(
            __import__("pathlib").Path(
                "pipeline/configs/chapter_locator_rules.json"
            )
        )

    def test_prospectus_header_number_with_following_text(self) -> None:
        candidates = pagination_candidates_for_page(
            [
                "芜湖三联锻造股份有限公司 "
                "招股说明书 29 第四节 发行人基本情况"
            ],
            self.rules,
            476,
        )
        self.assertEqual(candidates[0]["number"], 29)
        self.assertEqual(
            candidates[0]["pattern"],
            "PROSPECTUS_HEADER_BARE",
        )

    def test_prefixed_header_number_with_following_text(self) -> None:
        candidates = pagination_candidates_for_page(
            [
                "黄山谷捷股份有限公司 "
                "招股说明书 1-1-37 第四节"
            ],
            self.rules,
            343,
        )
        self.assertEqual(candidates[0]["prefix"], "1-1")
        self.assertEqual(candidates[0]["number"], 37)

    def test_date_and_table_range_are_not_page_numbers(self) -> None:
        self.assertEqual(
            pagination_candidates_for_page(
                ["基金备案日期2016-12-30"],
                self.rules,
                398,
            ),
            [],
        )
        self.assertEqual(
            pagination_candidates_for_page(
                ["房屋及建筑物5-20 15-40 5-20"],
                self.rules,
                398,
            ),
            [],
        )

    def test_dominant_offset_inference(self) -> None:
        records = [
            PaginationRecord(
                "TEST", 10, "1", None, 1, "HEADER",
                "DIRECT", "CONFIRMED", 0.97, [], "", None
            ),
            PaginationRecord(
                "TEST", 11, None, None, None, None,
                "NONE", "MISSING", 0.0, [], None, None
            ),
            PaginationRecord(
                "TEST", 12, "3", None, 3, "HEADER",
                "DIRECT", "CONFIRMED", 0.97, [], "", None
            ),
            PaginationRecord(
                "TEST", 13, "4", None, 4, "HEADER",
                "DIRECT", "CONFIRMED", 0.97, [], "", None
            ),
        ]
        infer_pagination(records, 3, 0.60)
        self.assertEqual(records[1].mapping_status, "INFERRED")
        self.assertEqual(records[1].printed_page_number, 2)

    def test_master_section_restricts_required_candidates(self) -> None:
        page_lines = {
            3: [
                "目录",
                "第四节 发行人基本情况........30",
                "第五节 业务与技术........80",
            ],
            30: [
                "测试公司 招股说明书 29 "
                "第四节 发行人基本情况",
                "一、发行人基本信息",
                "注册资本1000万元",
                "法定代表人甲",
            ],
            31: [
                "二、发行人设立及股本变化情况",
                "2008年5月公司设立",
            ],
            32: [
                "2009年6月公司增资",
                "2010年8月股权转让",
            ],
            40: [
                "三、发行人的股权结构",
                "截至本招股说明书签署日",
            ],
            45: ["四、发行人控股子公司情况"],
            50: [
                "五、发行人股本情况",
                "前十名股东持股比例",
            ],
            80: ["第五节 业务与技术", "一、主营业务"],
            90: ["2021年公司增资", "2022年股权转让"],
        }
        master = locate_issuer_master_section(
            "TEST",
            page_lines,
            100,
            self.rules,
        )
        self.assertEqual(master.start_pdf_page, 30)
        self.assertEqual(master.end_pdf_page, 79)

        headings, _ = build_heading_records(
            "TEST",
            page_lines,
            self.rules,
        )
        assign_hierarchy(headings, 100)
        enrich_from_section_context(
            headings,
            page_lines,
            self.rules,
        )

        pagination = []
        for page in range(1, 101):
            pagination.append(
                PaginationRecord(
                    "TEST",
                    page,
                    str(page - 1) if page > 1 else None,
                    None,
                    page - 1 if page > 1 else None,
                    "HEADER",
                    "DIRECT" if page > 1 else "NONE",
                    "CONFIRMED" if page > 1 else "MISSING",
                    0.97 if page > 1 else 0.0,
                    [],
                    None,
                    None,
                )
            )

        candidates = build_all_candidates(
            "TEST",
            headings,
            page_lines,
            pagination,
            100,
            self.rules,
            master,
        )
        required = [
            item
            for item in candidates
            if (
                item.is_primary
                and item.chapter_type
                in {"equity_history", "shareholders"}
            )
        ]
        self.assertEqual(len(required), 2)
        self.assertTrue(
            all(
                30 <= item.start_pdf_page <= 79
                and item.end_pdf_page_candidate <= 79
                for item in required
            )
        )

        reviews = build_review_items(
            "TEST",
            pagination,
            candidates,
            self.rules,
            master,
        )
        required_reviews = [
            item
            for item in reviews
            if item["review_type"]
            == "REQUIRED_CHAPTER_PRIMARY_REVIEW"
        ]
        self.assertEqual(len(required_reviews), 2)


if __name__ == "__main__":
    unittest.main()
