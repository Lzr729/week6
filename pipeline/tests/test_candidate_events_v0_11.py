from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from pipeline.candidate_events import (
    EventSection,
    FrozenRange,
    PrintedPageResolver,
    build_candidate_records,
    build_fallback_sections,
    build_heading_sections,
    build_shareholder_evidence,
    build_coverage_gaps,
    disclosed_event_counts,
    build_text_units,
    collect_summary_disclosures,
    event_types_semantically_equivalent,
    other_entity_summary_disclosures_from_sections,
    recover_missing_issuer_establishment,
    is_service_provider_aggregate_history,
    issuer_establishment_mentions,
    longest_event_type_hits,
    normalize_text,
    predecessor_subject_score,
    section_contains_exact_issuer_establishment,
    subject_has_exact_issuer_core,
    absorption_counterparty_names,
    generic_timeline_establishment_units,
    load_json,
    run_candidate_event_generation,
)


class CandidateEventGenerationV11Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.rules = load_json(
            Path("pipeline/configs/candidate_event_rules.json")
        )

    def frozen(
        self,
        chapter_type: str = "equity_history",
        scope: str = "FULL_RANGE_WITHIN_PROSPECTUS",
    ) -> FrozenRange:
        return FrozenRange(
            "PATCH",
            "123456",
            "测试公司",
            chapter_type,
            10,
            20,
            "9",
            "19",
            "DISCLOSED",
            scope,
            "FROZEN_FOR_EVENT_EXTRACTION",
            "ACCEPTED",
            "SOURCE",
        )

    def sections(
        self,
        pages: dict[int, str],
        scope: str = "FULL_RANGE_WITHIN_PROSPECTUS",
    ):
        frozen = self.frozen(scope=scope)
        units = build_text_units(
            frozen.company_id,
            pages,
            10,
            20,
        )
        heading, negatives = build_heading_sections(
            frozen,
            units,
            self.rules,
        )
        fallback, diagnostics = build_fallback_sections(
            frozen,
            units,
            heading,
            self.rules,
        )
        return (
            frozen,
            units,
            heading + fallback,
            negatives,
            diagnostics,
        )

    def candidates(
        self,
        frozen: FrozenRange,
        sections,
    ):
        resolver = PrintedPageResolver(frozen, None)
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        return build_candidate_records(
            frozen,
            sections,
            resolver,
            Path(temp.name),
            self.rules,
        )

    def test_negative_heading_and_body_filtered(self) -> None:
        frozen, units, sections, negatives, _ = self.sections({
            10: (
                "（九）报告期内重大资产重组情况\n"
                "报告期内，公司未发生重大资产重组。"
            )
        })
        self.assertEqual(sections, [])
        self.assertEqual(len(negatives), 1)

    def test_transfer_system_is_not_equity_transfer(self) -> None:
        frozen, units, sections, negatives, _ = self.sections({
            10: (
                "2023年5月，公司新增股份在全国中小企业"
                "股份转让系统挂牌并公开转让。"
            )
        })
        self.assertTrue(
            all(
                section.event_type_candidate
                != "EQUITY_TRANSFER"
                for section in sections
            )
        )

    def test_longest_match_is_exclusive(self) -> None:
        frozen, units, sections, _, _ = self.sections({
            10: (
                "2、股份公司设立情况\n"
                "2020年2月，公司完成股份有限公司设立。"
            )
        })
        self.assertEqual(
            sections[0].event_type_candidate,
            "JOINT_STOCK_COMPANY_ESTABLISHMENT",
        )
        self.assertNotIn(
            "LIMITED_COMPANY_ESTABLISHMENT",
            sections[0].event_type_candidates,
        )

    def test_procedural_dates_stay_inside_parent_event(self) -> None:
        pages = {
            10: (
                "2、2020年9月，整体变更为股份有限公司\n"
                "2020年8月25日，会计师事务所对本次整体变更"
                "进行审验。\n"
                "2020年9月16日，公司办理完毕工商变更登记。"
            )
        }
        frozen, units, sections, _, diagnostics = self.sections(
            pages
        )
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        self.assertEqual(len(candidates), 1)
        self.assertEqual(
            candidates[0].event_period,
            "2020-09",
        )
        self.assertIn(
            "registration_date",
            candidates[0].event_dates,
        )

    def test_establishment_uses_registration_date(self) -> None:
        pages = {
            10: (
                "（一）有限公司设立情况\n"
                "2004年6月1日，股东签署发起协议。\n"
                "2004年6月10日，会计师事务所出具验资报告。\n"
                "2004年6月18日，公司取得营业执照。\n"
                "2006年9月28日，公司召开股东会。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        self.assertEqual(
            candidates[0].event_period,
            "2004-06-18",
        )
        self.assertEqual(
            candidates[0].event_date_primary_role,
            "registration_date",
        )

    def test_different_ordinals_do_not_merge(self) -> None:
        pages = {
            10: (
                "1、2021年11月，第一次增资\n"
                "2021年11月22日，公司签署增资协议。\n"
                "2、2022年5月，第二次增资\n"
                "2022年5月24日，公司签署增资协议。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        self.assertEqual(len(candidates), 2)
        self.assertEqual(
            {tuple(item.ordinal_labels) for item in candidates},
            {("第一次",), ("第二次",)},
        )

    def test_combined_event_absorbs_procedural_children(self) -> None:
        pages = {
            10: (
                "1、2021年4月，第一次股权转让及第一次增资\n"
                "2021年1月，股东签署股权转让协议。\n"
                "2021年3月，受让方支付股权转让款。\n"
                "（2）报告期内第一次增资\n"
                "2021年2月25日，公司召开股东会审议增资。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        self.assertEqual(len(candidates), 1)
        self.assertIn(
            "EQUITY_TRANSFER",
            candidates[0].event_type_candidates,
        )
        self.assertIn(
            "CAPITAL_INCREASE",
            candidates[0].event_type_candidates,
        )

    def test_timeline_row_is_summary_when_detail_exists(self) -> None:
        pages = {
            10: (
                "1、有限公司设立情况\n"
                "2015年7月9日，公司取得营业执照。\n"
                "1 2020年1月7日 有限公司设立及其他事项"
            )
        }
        frozen, units, sections, _, diagnostics = self.sections(
            pages
        )
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        summaries = collect_summary_disclosures(
            frozen,
            units,
            self.rules,
        )
        self.assertEqual(len(candidates), 1)
        self.assertTrue(
            any(
                item.summary_kind
                == "TIMELINE_SUMMARY_ROW"
                for item in summaries
            )
        )
        self.assertEqual(len(candidates), 1)

    def test_aggregate_summary_is_not_candidate(self) -> None:
        pages = {
            10: (
                "三联有限自设立至报告期期初进行了四次股权转让"
                "和四次增资，具体情况参见相关文件。"
            )
        }
        frozen, units, sections, _, diagnostics = self.sections(
            pages
        )
        summaries = collect_summary_disclosures(
            frozen,
            units,
            self.rules,
        )
        self.assertEqual(sections, [])
        self.assertTrue(
            any(
                item.summary_kind
                == "AGGREGATE_EVENT_SUMMARY"
                for item in summaries
            )
        )

    def test_metadata_date_field_is_not_candidate(self) -> None:
        pages = {
            10: (
                "整体变更设立日期 2020年2月26日\n"
                "2、股份公司设立情况\n"
                "2020年2月26日，公司取得营业执照。"
            )
        }
        frozen, units, sections, _, diagnostics = self.sections(
            pages
        )
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        summaries = collect_summary_disclosures(
            frozen,
            units,
            self.rules,
        )
        self.assertEqual(len(candidates), 1)
        self.assertTrue(
            any(
                item.summary_kind
                == "METADATA_DATE_FIELD"
                for item in summaries
            )
        )

    def test_cross_reference_only_is_not_candidate(self) -> None:
        pages = {
            10: (
                "之“二、发行人挂牌期间的基本情况”之"
                "“（八）报告期内发行融资情况”之"
                "“2、2023年股票发行”"
            )
        }
        frozen, units, sections, _, diagnostics = self.sections(
            pages,
            scope="REFERENCED_NOT_FULLY_DISCLOSED",
        )
        summaries = collect_summary_disclosures(
            frozen,
            units,
            self.rules,
        )
        self.assertEqual(sections, [])
        self.assertTrue(
            any(
                item.summary_kind
                == "CROSS_REFERENCE_ONLY"
                for item in summaries
            )
        )

    def test_classification_wrapper_not_separate_candidate(self) -> None:
        pages = {
            10: (
                "（一）发行人报告期内的重大资产重组情况\n"
                "2021年4月，吸收合并\n"
                "2021年4月7日，公司召开股东会通过吸收合并。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        summaries = collect_summary_disclosures(
            frozen,
            units,
            self.rules,
        )
        self.assertEqual(len(candidates), 1)
        self.assertEqual(
            candidates[0].event_type_candidate,
            "ABSORPTION_MERGER",
        )
        self.assertTrue(
            any(
                item.summary_kind
                == "CLASSIFICATION_WRAPPER"
                for item in summaries
            )
        )

    def test_proxy_disclosure_may_be_undated_without_review(self) -> None:
        pages = {
            10: (
                "（五）发行人历史沿革中存在股份代持的情况\n"
                "相关代持关系已经规范，不存在纠纷。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        self.assertEqual(len(candidates), 1)
        self.assertFalse(candidates[0].review_required)

    def test_joint_stock_and_overall_change_merge(self) -> None:
        pages = {
            10: (
                "（二）股份公司设立情况\n"
                "2018年10月11日，公司召开股东会审议整体变更。\n"
                "2018年11月26日，公司办理完毕整体变更为股份有限公司的工商登记。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(frozen, sections)
        self.assertEqual(len(candidates), 1)
        self.assertEqual(candidates[0].event_type_candidate, "OVERALL_CHANGE")
        self.assertIn(
            "JOINT_STOCK_COMPANY_ESTABLISHMENT",
            candidates[0].event_type_candidates,
        )
        self.assertEqual(candidates[0].event_period, "2018-11-26")


    def test_service_provider_overall_change_is_not_separate_candidate(self) -> None:
        pages = {
            10: (
                "2021年8月25日，容诚会计师事务所对本次整体变更进行审验。\n"
                "2021年9月16日，公司整体变更为股份有限公司并办理完毕工商登记。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(frozen, sections)
        summaries = collect_summary_disclosures(frozen, units, self.rules)
        self.assertEqual(len(candidates), 1)
        self.assertTrue(any(
            item.summary_kind == "SERVICE_PROVIDER_PROCEDURE"
            for item in summaries
        ))


    def test_timeline_overall_change_suppressed_by_joint_stock_detail(self) -> None:
        pages = {
            10: (
                "2、股份公司设立情况\n"
                "2020年2月26日，公司取得股份有限公司营业执照。\n"
                "2 2020年1月18日 整体变更 容诚会计师事务所"
            )
        }
        frozen, units, sections, _, diagnostics = self.sections(pages)
        candidates, evidence, reviews = self.candidates(frozen, sections)
        self.assertEqual(len(candidates), 1)
        self.assertTrue(event_types_semantically_equivalent(
            "JOINT_STOCK_COMPANY_ESTABLISHMENT",
            "OVERALL_CHANGE",
            self.rules,
        ))


    def test_other_company_capital_change_is_suppressed(self) -> None:
        pages = {
            10: (
                "2024年5月8日，中科星图2023年年度股东大会审议通过资本公积转增股本方案。\n"
                "中科星图总股本变更为54332万股，2024年7月4日实施完成。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        other_sections = [
            section for section in sections
            if section.entity_scope_candidate == "SUBSIDIARY_OR_OTHER_ENTITY_RISK"
        ]
        business_sections = [
            section for section in sections
            if section.entity_scope_candidate != "SUBSIDIARY_OR_OTHER_ENTITY_RISK"
        ]
        candidates, evidence, reviews = self.candidates(frozen, business_sections)
        summaries = other_entity_summary_disclosures_from_sections(
            frozen, other_sections, self.rules, 1
        )
        self.assertEqual(candidates, [])
        self.assertEqual(len(summaries), 1)
        self.assertEqual(summaries[0].summary_kind, "OTHER_ENTITY_EVENT")


    def test_partial_scope_alone_does_not_require_review(self) -> None:
        pages = {
            10: (
                "2023年3月，公司进行第一次股票定向发行。\n"
                "2023年5月19日，新增股份完成登记。"
            )
        }
        frozen, units, sections, _, _ = self.sections(
            pages,
            scope="PARTIAL_DISCLOSURE_LISTING_AND_REPORTING_PERIOD",
        )
        candidates, evidence, reviews = self.candidates(frozen, sections)
        self.assertEqual(len(candidates), 1)
        self.assertFalse(candidates[0].review_required)


    def test_reference_scope_alone_does_not_require_review(self) -> None:
        pages = {
            10: "2021年3月，上述股权代持已解除，涉及各方不存在纠纷。"
        }
        frozen, units, sections, _, _ = self.sections(
            pages,
            scope="REFERENCED_NOT_FULLY_DISCLOSED",
        )
        candidates, evidence, reviews = self.candidates(frozen, sections)
        self.assertEqual(len(candidates), 1)
        self.assertFalse(candidates[0].review_required)

    def test_shareholder_evidence_deduplicates(self) -> None:
        frozen = self.frozen(chapter_type="shareholders")
        pages = {
            10: (
                "四、发行人股东及实际控制人情况\n"
                "控股股东为甲。实际控制人为乙。\n"
                "前十名股东持股数量及持股比例如下。"
            ),
            11: (
                "控股股东甲持股60%，"
                "实际控制人乙持股20%。"
            ),
        }
        units = build_text_units(
            frozen.company_id,
            pages,
            10,
            20,
        )
        records = build_shareholder_evidence(
            frozen,
            units,
            PrintedPageResolver(frozen, None),
            self.rules,
        )
        self.assertLessEqual(len(records), 6)

    def test_full_runner_cache_only(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            repo = root / "repo"
            workspace = root / "runs"
            pdf_dir = root / "pdf"
            repo.mkdir()
            workspace.mkdir()
            pdf_dir.mkdir()
            (
                pdf_dir
                / "123456_测试公司_IPO招股说明书.pdf"
            ).write_bytes(b"not-opened")

            review = repo / "review/chapter_location"
            review.mkdir(parents=True)
            config = repo / "pipeline/configs"
            config.mkdir(parents=True)
            (
                config / "candidate_event_rules.json"
            ).write_text(
                json.dumps(
                    self.rules,
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            patches = [
                {
                    "patch_id": "EQ",
                    "company_id": "123456",
                    "company_short_name": "测试公司",
                    "chapter_type": "equity_history",
                    "final_start_pdf_page": 2,
                    "final_end_pdf_page": 3,
                    "final_start_printed_page_raw": "1",
                    "final_end_printed_page_raw": "2",
                    "printed_page_value_type": "DISCLOSED",
                    "disclosure_scope": "FULL_RANGE_WITHIN_PROSPECTUS",
                    "final_status": "FROZEN_FOR_EVENT_EXTRACTION",
                    "decision": "ACCEPTED",
                    "source_run_id": "SOURCE",
                },
                {
                    "patch_id": "SH",
                    "company_id": "123456",
                    "company_short_name": "测试公司",
                    "chapter_type": "shareholders",
                    "final_start_pdf_page": 4,
                    "final_end_pdf_page": 4,
                    "final_start_printed_page_raw": "3",
                    "final_end_printed_page_raw": "3",
                    "printed_page_value_type": "DISCLOSED",
                    "disclosure_scope": "MAJOR_SHAREHOLDERS",
                    "final_status": "FROZEN_FOR_EVENT_EXTRACTION",
                    "decision": "ACCEPTED",
                    "source_run_id": "SOURCE",
                },
            ]
            chapter_patch = (
                review
                / "chapter_location_review_patch.jsonl"
            )
            with chapter_patch.open(
                "w",
                encoding="utf-8",
            ) as handle:
                for patch in patches:
                    handle.write(
                        json.dumps(
                            patch,
                            ensure_ascii=False,
                        )
                        + "\n"
                    )
            (
                review / "pagination_review_patch.jsonl"
            ).write_text("", encoding="utf-8")

            page_dir = workspace / "SOURCE/123456"
            page_dir.mkdir(parents=True)
            rows = [
                {"pdf_page": 1, "text": "封面"},
                {
                    "pdf_page": 2,
                    "text": (
                        "1、2019年1月，有限公司设立\n"
                        "2019年1月，公司取得营业执照。"
                    ),
                },
                {
                    "pdf_page": 3,
                    "text": (
                        "公司成立以来进行了两次增资，"
                        "具体情况如下。"
                    ),
                },
                {
                    "pdf_page": 4,
                    "text": (
                        "前十名股东 持股数量 持股比例"
                    ),
                },
            ]
            with (
                page_dir / "page_text.jsonl"
            ).open("w", encoding="utf-8") as handle:
                for row in rows:
                    handle.write(
                        json.dumps(
                            row,
                            ensure_ascii=False,
                        )
                        + "\n"
                    )

            result = run_candidate_event_generation(
                pdf_dir,
                repo,
                workspace,
                Path(
                    "review/chapter_location/"
                    "chapter_location_review_patch.jsonl"
                ),
                Path(
                    "review/chapter_location/"
                    "pagination_review_patch.jsonl"
                ),
                Path(
                    "pipeline/configs/"
                    "candidate_event_rules.json"
                ),
                "SOURCE",
                1,
            )
            self.assertEqual(result, 0)
            latest = json.loads(
                (
                    repo
                    / "logs/candidate_events/latest_run.json"
                ).read_text(encoding="utf-8")
            )
            auto_dir = (
                repo
                / "auto_output/candidate_events/runs"
                / latest["run_id"]
            )
            self.assertTrue(
                (
                    auto_dir
                    / "summary_disclosures_auto.jsonl"
                ).is_file()
            )


    def test_same_type_overall_change_same_month_merges_across_pages(self) -> None:
        pages = {
            10: (
                "2021年9月16日，公司就整体变更为股份有限公司事项"
                "办理完毕工商登记。"
            ),
            16: (
                "3、2021年9月，整体变更为股份公司\n"
                "本次整体变更的具体情况详见前述。"
            ),
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen, sections
        )
        self.assertEqual(len(candidates), 1)
        self.assertEqual(
            candidates[0].event_type_candidate,
            "OVERALL_CHANGE",
        )
        self.assertEqual(candidates[0].event_period, "2021-09-16")

    def test_predecessor_limited_name_is_issuer_event(self) -> None:
        pages = {
            10: (
                "1、2021年4月，第一次股权转让及第一次增资\n"
                "2021年2月22日，测试有限召开股东会并作出决议，"
                "同意股东转让股权。\n"
                "2021年2月25日，公司召开股东会，新增注册资本。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        self.assertTrue(sections)
        self.assertTrue(all(
            item.entity_scope_candidate == "ISSUER_OR_UNRESOLVED"
            for item in sections
        ))
        candidates, evidence, reviews = self.candidates(
            frozen, sections
        )
        self.assertEqual(len(candidates), 1)
        self.assertIn(
            "EQUITY_TRANSFER",
            candidates[0].event_type_candidates,
        )
        self.assertIn(
            "CAPITAL_INCREASE",
            candidates[0].event_type_candidates,
        )

    def test_location_prefix_predecessor_is_not_other_entity(self) -> None:
        pages = {
            10: (
                "2021年4月，吸收合并\n"
                "2021年4月7日，谷捷有限召开股东会，"
                "吸收合并后谷捷有限的注册资本变更为1200万元。"
            )
        }
        frozen = self.frozen()
        frozen.company_short_name = "黄山谷捷"
        units = build_text_units(
            frozen.company_id, pages, 10, 20
        )
        heading, negatives = build_heading_sections(
            frozen, units, self.rules
        )
        fallback, diagnostics = build_fallback_sections(
            frozen, units, heading, self.rules
        )
        sections = heading + fallback
        self.assertTrue(sections)
        self.assertTrue(all(
            item.entity_scope_candidate == "ISSUER_OR_UNRESOLVED"
            for item in sections
        ))

    def test_parent_company_without_limited_suffix_remains_other_entity(self) -> None:
        pages = {
            10: (
                "2024年5月8日，中科星图2023年年度股东大会"
                "审议通过资本公积转增股本方案。\n"
                "中科星图总股本变更为54332万股。"
            )
        }
        frozen = self.frozen()
        frozen.company_short_name = "星图测控"
        units = build_text_units(
            frozen.company_id, pages, 10, 20
        )
        heading, negatives = build_heading_sections(
            frozen, units, self.rules
        )
        fallback, diagnostics = build_fallback_sections(
            frozen, units, heading, self.rules
        )
        sections = heading + fallback
        self.assertTrue(sections)
        self.assertTrue(all(
            item.entity_scope_candidate
            == "SUBSIDIARY_OR_OTHER_ENTITY_RISK"
            for item in sections
        ))

    def test_joint_stock_conversion_canonicalizes_to_overall_change(self) -> None:
        pages = {
            10: (
                "2、股份公司设立情况\n"
                "全体股东同意将有限公司变更为股份有限公司，"
                "以经审计净资产折合股本。\n"
                "2020年2月26日，公司取得股份有限公司营业执照。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen, sections
        )
        self.assertEqual(len(candidates), 1)
        self.assertEqual(
            candidates[0].event_type_candidate,
            "OVERALL_CHANGE",
        )
        self.assertIn(
            "OVERALL_CHANGE",
            candidates[0].event_type_candidates,
        )
    def test_absorption_merger_capital_change_consequence_attaches(self) -> None:
        pages = {
            10: (
                "2021年4月，吸收合并\n"
                "2021年4月7日，谷捷有限召开股东会，"
                "吸收合并后谷捷有限的注册资本变更为1200万元。"
            )
        }
        frozen = self.frozen()
        frozen.company_short_name = "黄山谷捷"
        units = build_text_units(
            frozen.company_id, pages, 10, 20
        )
        heading, negatives = build_heading_sections(
            frozen, units, self.rules
        )
        fallback, diagnostics = build_fallback_sections(
            frozen, units, heading, self.rules
        )
        candidates, evidence, reviews = self.candidates(
            frozen, heading + fallback
        )
        self.assertEqual(len(candidates), 1)
        self.assertEqual(
            candidates[0].event_type_candidate,
            "ABSORPTION_MERGER",
        )
        self.assertIn(
            "SHARE_CAPITAL_CHANGE",
            candidates[0].event_type_candidates,
        )

    def test_named_predecessor_establishment_heading_detected(self) -> None:
        pages = {
            10: (
                "（一）公司前身赛分有限的设立情况\n"
                "2009年3月16日，赛分有限取得营业执照。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        self.assertEqual(len(candidates), 1)
        self.assertEqual(
            candidates[0].event_type_candidate,
            "LIMITED_COMPANY_ESTABLISHMENT",
        )
        self.assertEqual(candidates[0].event_period, "2009-03-16")

    def test_timeline_predecessor_establishment_not_hidden_by_later_event(self) -> None:
        pages = {
            10: (
                "2012年6月，谷捷有限成立\n"
                "2022年9月，谷捷有限整体变更为股份公司"
            )
        }
        frozen, units, sections, _, diagnostics = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        self.assertEqual(
            {
                item.event_type_candidate
                for item in candidates
            },
            {
                "LIMITED_COMPANY_ESTABLISHMENT",
                "OVERALL_CHANGE",
            },
        )

    def test_absorbed_company_establishment_is_other_entity(self) -> None:
        pages = {
            10: (
                "2012年6月，昆山谷捷有限公司成立。\n"
                "2021年4月，发行人吸收合并昆山谷捷。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        establishment = [
            item for item in sections
            if item.event_type_candidate
            == "LIMITED_COMPANY_ESTABLISHMENT"
        ]
        self.assertEqual(len(establishment), 1)
        self.assertEqual(
            establishment[0].entity_scope_candidate,
            "SUBSIDIARY_OR_OTHER_ENTITY_RISK",
        )

    def test_aggregate_event_counts_are_parsed(self) -> None:
        counts = disclosed_event_counts(
            "公司进行了四次股权转让和四次增资，另有一次资本公积转增股本。"
        )
        self.assertEqual(counts["EQUITY_TRANSFER"], 4)
        self.assertEqual(counts["CAPITAL_INCREASE"], 4)
        self.assertEqual(counts["SHARE_CAPITAL_CHANGE"], 1)

    def test_coverage_gap_detects_unexpanded_history(self) -> None:
        pages = {
            10: (
                "公司自设立至报告期期初进行了四次股权转让和四次增资，"
                "具体情况参见申报文件。"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        summaries = collect_summary_disclosures(
            frozen,
            units,
            self.rules,
        )
        gaps = build_coverage_gaps(
            frozen,
            summaries,
            [],
        )
        self.assertEqual(len(gaps), 1)
        self.assertEqual(
            gaps[0].gap_type,
            "REFERENCE_LIMITED_AGGREGATE_HISTORY",
        )
        self.assertEqual(
            gaps[0].missing_event_counts["EQUITY_TRANSFER"],
            4,
        )
        self.assertEqual(
            gaps[0].missing_event_counts["CAPITAL_INCREASE"],
            4,
        )

    def test_coverage_gap_not_created_when_counts_are_represented(self) -> None:
        pages = {
            10: (
                "报告期内，公司进行了2次增资。\n"
                "1、2021年1月，第一次增资\n"
                "2、2021年6月，第二次增资"
            )
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        summaries = collect_summary_disclosures(
            frozen,
            units,
            self.rules,
        )
        gaps = build_coverage_gaps(
            frozen,
            summaries,
            candidates,
        )
        self.assertEqual(gaps, [])

    def test_generic_limited_company_heading_is_not_other_entity(self) -> None:
        frozen = FrozenRange(
            "PATCH",
            "001282",
            "三联锻造",
            "equity_history",
            10,
            20,
            "9",
            "19",
            "DISCLOSED",
            "FULL_RANGE_WITHIN_PROSPECTUS",
            "FROZEN_FOR_EVENT_EXTRACTION",
            "ACCEPTED",
            "SOURCE",
        )
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "（一）有限公司设立情况\n"
                    "2004年6月1日，股东签署协议设立三联有限。\n"
                    "2004年6月18日，三联有限取得营业执照。"
                )
            },
            10,
            20,
        )
        sections, negatives = build_heading_sections(
            frozen,
            units,
            self.rules,
        )
        self.assertEqual(len(sections), 1)
        self.assertEqual(
            sections[0].entity_scope_candidate,
            "ISSUER_OR_UNRESOLVED",
        )

    def test_generic_limited_liability_heading_is_root_event(self) -> None:
        frozen, units, sections, _, _ = self.sections({
            10: (
                "（一）有限责任公司的设立情况\n"
                "1992年12月4日，友升有限取得营业执照。\n"
                "注册资本400万美元。"
            )
        })
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        self.assertEqual(len(candidates), 1)
        self.assertGreaterEqual(
            candidates[0].candidate_confidence,
            0.7,
        )
        self.assertEqual(reviews, [])

    def test_page_timeline_recovery_for_scrambled_columns(self) -> None:
        frozen = FrozenRange(
            "PATCH",
            "301581",
            "黄山谷捷",
            "equity_history",
            10,
            20,
            "9",
            "19",
            "DISCLOSED",
            "REPORTING_PERIOD_EQUITY_AND_SHAREHOLDER_CHANGES",
            "FROZEN_FOR_EVENT_EXTRACTION",
            "ACCEPTED",
            "SOURCE",
        )
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "2012年6月\n"
                    "2021年4月\n"
                    "2022年9月\n"
                    "谷捷有限成立\n"
                    "第一次股权转让\n"
                    "整体变更为股份公司"
                )
            },
            10,
            20,
        )
        recovered = recover_missing_issuer_establishment(
            frozen,
            units,
            [],
            self.rules,
        )
        self.assertEqual(len(recovered), 1)
        self.assertEqual(
            recovered[0].event_type_candidate,
            "LIMITED_COMPANY_ESTABLISHMENT",
        )
        self.assertEqual(recovered[0].event_period, "2012-06")

    def test_service_provider_history_row_is_not_candidate(self) -> None:
        pages = {
            10: (
                "1 2020年1月7日 有限公司设立及六次增资扩股\n"
                "容诚会计师事务所（特殊普通合伙）\n"
                "容诚验字[2020]518Z0005号\n"
                "根据容诚会计师出具的验资报告进行审验。"
            )
        }
        frozen, units, sections, _, diagnostics = self.sections(
            pages
        )
        self.assertEqual(sections, [])
        self.assertTrue(
            is_service_provider_aggregate_history(
                units[0],
                units,
                self.rules,
            )
        )

    def test_duplicate_service_provider_coverage_gap_is_merged(self) -> None:
        pages = {
            10: (
                "1 2020年1月7日 有限公司设立及六次增资扩股\n"
                "容诚会计师事务所 容诚验字[2020]1号\n"
                "2 2020年1月18日 整体变更\n"
                "根据验资报告，有限公司设立及六次增资扩股、整体变更。"
            )
        }
        frozen = self.frozen()
        units = build_text_units(
            frozen.company_id,
            pages,
            10,
            20,
        )
        summaries = collect_summary_disclosures(
            frozen,
            units,
            self.rules,
        )
        gaps = build_coverage_gaps(
            frozen,
            summaries,
            [],
        )
        self.assertEqual(len(gaps), 1)
        self.assertEqual(
            gaps[0].missing_event_counts,
            {"CAPITAL_INCREASE": 6},
        )

    def test_joint_stock_heading_does_not_match_limited_establishment(self) -> None:
        types, terms, combined = longest_event_type_hits(
            normalize_text("（二）股份有限公司的设立情况"),
            self.rules,
        )
        self.assertIn(
            "JOINT_STOCK_COMPANY_ESTABLISHMENT",
            types,
        )
        self.assertNotIn(
            "LIMITED_COMPANY_ESTABLISHMENT",
            types,
        )

    def test_joint_stock_heading_canonicalizes_to_overall_change(self) -> None:
        pages = {
            10: (
                "（二）股份有限公司的设立情况\n"
                "友升有限以经审计的净资产折合股本，"
                "整体变更为股份有限公司。\n"
                "2020年9月9日完成工商登记。"
            ),
            11: (
                "2、2020年9月，整体变更为股份有限公司"
            ),
        }
        frozen, units, sections, _, _ = self.sections(pages)
        candidates, evidence, reviews = self.candidates(
            frozen,
            sections,
        )
        overall = [
            item
            for item in candidates
            if item.event_type_candidate == "OVERALL_CHANGE"
        ]
        limited = [
            item
            for item in candidates
            if item.event_type_candidate
            == "LIMITED_COMPANY_ESTABLISHMENT"
        ]
        self.assertEqual(len(overall), 1)
        self.assertEqual(limited, [])

    def test_issuer_core_establishment_mentions_scrambled_timeline(self) -> None:
        names = issuer_establishment_mentions(
            (
                "2012年6月2021年4月2022年9月"
                "谷捷有限成立第一次股权转让整体变更"
            ),
            "黄山谷捷",
            self.rules,
        )
        self.assertIn("谷捷有限", names)

    def test_recovery_not_blocked_by_unrelated_establishment_section(self) -> None:
        frozen = FrozenRange(
            "PATCH",
            "301581",
            "黄山谷捷",
            "equity_history",
            10,
            20,
            "9",
            "19",
            "DISCLOSED",
            "REPORTING_PERIOD_EQUITY_AND_SHAREHOLDER_CHANGES",
            "FROZEN_FOR_EVENT_EXTRACTION",
            "ACCEPTED",
            "SOURCE",
        )
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "2012年6月\n"
                    "2021年4月\n"
                    "2022年9月\n"
                    "谷捷有限成立\n"
                    "第一次股权转让\n"
                    "整体变更"
                )
            },
            10,
            20,
        )
        unrelated = EventSection(
            section_id="OTHER",
            company_id=frozen.company_id,
            source_patch_id=frozen.patch_id,
            source_kind="HEADING_SECTION",
            title="昆山谷捷有限公司成立",
            title_unit_index=0,
            heading_level=3,
            event_type_candidate="LIMITED_COMPANY_ESTABLISHMENT",
            event_type_candidates=["LIMITED_COMPANY_ESTABLISHMENT"],
            explicit_combined_event=False,
            ordinal_labels=[],
            timeline_summary=False,
            units=[units[0]],
            date_roles={"event_period": ["2018-01"]},
            event_period="2018-01",
            event_date_primary_role="event_period",
            event_date_selection_basis="TEST",
            negative_disclosure=False,
            negative_reason=None,
            entity_scope_candidate="SUBSIDIARY_OR_OTHER_ENTITY_RISK",
            service_provider_mentions=[],
            signals=[],
            confidence=0.8,
            review_reasons=[],
        )
        recovered = recover_missing_issuer_establishment(
            frozen,
            units,
            [unrelated],
            self.rules,
        )
        self.assertEqual(len(recovered), 1)
        self.assertEqual(recovered[0].event_period, "2012-06")

    def test_recovery_skips_already_represented_period(self) -> None:
        frozen = self.frozen()
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "2012年6月\n"
                    "测试有限成立"
                )
            },
            10,
            20,
        )
        existing = EventSection(
            section_id="EXISTING",
            company_id=frozen.company_id,
            source_patch_id=frozen.patch_id,
            source_kind="HEADING_SECTION",
            title="有限公司设立",
            title_unit_index=0,
            heading_level=3,
            event_type_candidate="LIMITED_COMPANY_ESTABLISHMENT",
            event_type_candidates=["LIMITED_COMPANY_ESTABLISHMENT"],
            explicit_combined_event=False,
            ordinal_labels=[],
            timeline_summary=False,
            units=[units[0]],
            date_roles={"event_period": ["2012-06"]},
            event_period="2012-06",
            event_date_primary_role="event_period",
            event_date_selection_basis="TEST",
            negative_disclosure=False,
            negative_reason=None,
            entity_scope_candidate="ISSUER_OR_UNRESOLVED",
            service_provider_mentions=[],
            signals=[],
            confidence=0.9,
            review_reasons=[],
        )
        recovered = recover_missing_issuer_establishment(
            frozen,
            units,
            [existing],
            self.rules,
        )
        self.assertEqual(recovered, [])

    def test_establishment_date_metadata_not_a_mention(self) -> None:
        names = issuer_establishment_mentions(
            "公司名称昆山谷捷金属制品有限公司成立日期2009年7月6日",
            "黄山谷捷",
            self.rules,
        )
        self.assertEqual(names, set())

    def test_detailed_establishment_blocks_timeline_recovery(self) -> None:
        frozen = self.frozen()
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "（一）有限公司设立情况\n"
                    "2004年6月18日，测试有限取得营业执照。\n"
                    "1999年，测试有限其他背景信息。"
                )
            },
            10,
            20,
        )
        existing = EventSection(
            section_id="DETAIL",
            company_id=frozen.company_id,
            source_patch_id=frozen.patch_id,
            source_kind="HEADING_SECTION",
            title="（一）有限公司设立情况",
            title_unit_index=0,
            heading_level=3,
            event_type_candidate="LIMITED_COMPANY_ESTABLISHMENT",
            event_type_candidates=["LIMITED_COMPANY_ESTABLISHMENT"],
            explicit_combined_event=False,
            ordinal_labels=[],
            timeline_summary=False,
            units=units,
            date_roles={"registration_date": ["2004-06-18"]},
            event_period="2004-06-18",
            event_date_primary_role="registration_date",
            event_date_selection_basis="TEST",
            negative_disclosure=False,
            negative_reason=None,
            entity_scope_candidate="ISSUER_OR_UNRESOLVED",
            service_provider_mentions=[],
            signals=[],
            confidence=0.9,
            review_reasons=[],
        )
        self.assertEqual(
            recover_missing_issuer_establishment(
                frozen,
                units,
                [existing],
                self.rules,
            ),
            [],
        )

    def test_exact_predecessor_name_outranks_longer_name(self) -> None:
        short_score = predecessor_subject_score(
            "谷捷有限",
            "黄山谷捷",
            self.rules,
        )
        long_score = predecessor_subject_score(
            "谷捷金属制品有限公司",
            "黄山谷捷",
            self.rules,
        )
        self.assertLess(short_score, long_score)

    def test_huangshan_timeline_prefers_gujie_limited(self) -> None:
        frozen = FrozenRange(
            "PATCH",
            "301581",
            "黄山谷捷",
            "equity_history",
            10,
            20,
            "9",
            "19",
            "DISCLOSED",
            "REPORTING_PERIOD_EQUITY_AND_SHAREHOLDER_CHANGES",
            "FROZEN_FOR_EVENT_EXTRACTION",
            "ACCEPTED",
            "SOURCE",
        )
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "2012年6月\n"
                    "2021年4月\n"
                    "2022年9月\n"
                    "谷捷有限成立\n"
                    "第一次股权转让\n"
                    "整体变更"
                ),
                15: (
                    "公司名称 昆山谷捷金属制品有限公司\n"
                    "成立日期 2009年7月6日\n"
                    "发行人吸收合并昆山谷捷"
                ),
            },
            10,
            20,
        )
        recovered = recover_missing_issuer_establishment(
            frozen,
            units,
            [],
            self.rules,
        )
        self.assertEqual(len(recovered), 1)
        self.assertEqual(
            recovered[0].title,
            "2012-06，谷捷有限成立",
        )
        self.assertEqual(
            recovered[0].event_period,
            "2012-06",
        )

    def test_absorber_named_company_detects_counterparty(self) -> None:
        names = absorption_counterparty_names(
            "谷捷有限吸收合并昆山谷捷，注册资本变为1,200万元。",
            self.rules,
        )
        self.assertIn("昆山谷捷", names)

    def test_exact_issuer_core_subject(self) -> None:
        self.assertTrue(
            subject_has_exact_issuer_core(
                "谷捷有限",
                "黄山谷捷",
                self.rules,
            )
        )
        self.assertFalse(
            subject_has_exact_issuer_core(
                "昆山谷捷金属制品有限公司",
                "黄山谷捷",
                self.rules,
            )
        )

    def test_absorbed_company_section_does_not_block_recovery(self) -> None:
        frozen = FrozenRange(
            "PATCH",
            "301581",
            "黄山谷捷",
            "equity_history",
            10,
            20,
            "9",
            "19",
            "DISCLOSED",
            "REPORTING_PERIOD_EQUITY_AND_SHAREHOLDER_CHANGES",
            "FROZEN_FOR_EVENT_EXTRACTION",
            "ACCEPTED",
            "SOURCE",
        )
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "2012年6月，谷捷有限成立\n"
                    "2021年4月，第一次股权转让\n"
                    "2022年9月，整体变更"
                ),
                15: (
                    "公司名称 昆山谷捷金属制品有限公司\n"
                    "成立日期 2009年7月6日\n"
                    "2009年6月30日，昆山谷捷召开股东会设立公司。\n"
                    "谷捷有限吸收合并昆山谷捷，注册资本变为1,200万元。"
                ),
            },
            10,
            20,
        )
        absorbed = EventSection(
            section_id="ABSORBED",
            company_id=frozen.company_id,
            source_patch_id=frozen.patch_id,
            source_kind="HEADING_SECTION",
            title="昆山谷捷有限公司设立情况",
            title_unit_index=3,
            heading_level=3,
            event_type_candidate="LIMITED_COMPANY_ESTABLISHMENT",
            event_type_candidates=["LIMITED_COMPANY_ESTABLISHMENT"],
            explicit_combined_event=False,
            ordinal_labels=[],
            timeline_summary=False,
            units=units[3:],
            date_roles={"other_date": ["2009-07-06"]},
            event_period="2009-07-06",
            event_date_primary_role="other_date",
            event_date_selection_basis="TEST",
            negative_disclosure=False,
            negative_reason=None,
            entity_scope_candidate="ISSUER_OR_UNRESOLVED",
            service_provider_mentions=[],
            signals=[],
            confidence=0.8,
            review_reasons=[],
        )
        self.assertFalse(
            section_contains_exact_issuer_establishment(
                absorbed,
                frozen.company_short_name,
                self.rules,
            )
        )
        recovered = recover_missing_issuer_establishment(
            frozen,
            units,
            [absorbed],
            self.rules,
        )
        self.assertEqual(len(recovered), 1)
        self.assertEqual(
            recovered[0].title,
            "2012-06，谷捷有限成立",
        )

    def test_exact_detailed_section_still_blocks_recovery(self) -> None:
        frozen = self.frozen()
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "（一）有限公司设立情况\n"
                    "2004年6月18日，测试有限取得营业执照。"
                )
            },
            10,
            20,
        )
        detailed = EventSection(
            section_id="DETAIL",
            company_id=frozen.company_id,
            source_patch_id=frozen.patch_id,
            source_kind="HEADING_SECTION",
            title="（一）有限公司设立情况",
            title_unit_index=0,
            heading_level=3,
            event_type_candidate="LIMITED_COMPANY_ESTABLISHMENT",
            event_type_candidates=["LIMITED_COMPANY_ESTABLISHMENT"],
            explicit_combined_event=False,
            ordinal_labels=[],
            timeline_summary=False,
            units=units,
            date_roles={"registration_date": ["2004-06-18"]},
            event_period="2004-06-18",
            event_date_primary_role="registration_date",
            event_date_selection_basis="TEST",
            negative_disclosure=False,
            negative_reason=None,
            entity_scope_candidate="ISSUER_OR_UNRESOLVED",
            service_provider_mentions=[],
            signals=[],
            confidence=0.9,
            review_reasons=[],
        )
        self.assertTrue(
            section_contains_exact_issuer_establishment(
                detailed,
                frozen.company_short_name,
                self.rules,
            )
        )
        self.assertEqual(
            recover_missing_issuer_establishment(
                frozen,
                units,
                [detailed],
                self.rules,
            ),
            [],
        )

    def test_generic_capital_setup_row_detected(self) -> None:
        frozen = FrozenRange(
            "PATCH",
            "301581",
            "黄山谷捷",
            "equity_history",
            10,
            20,
            "9",
            "19",
            "DISCLOSED",
            "REPORTING_PERIOD_EQUITY_AND_SHAREHOLDER_CHANGES",
            "FROZEN_FOR_EVENT_EXTRACTION",
            "ACCEPTED",
            "SOURCE",
        )
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "2012年6月\n"
                    "2021年4月\n"
                    "2022年9月\n"
                    "注册资本1,000万元，由昆山谷捷出资设立\n"
                    "第一次股权转让\n"
                    "整体变更为股份公司"
                )
            },
            10,
            20,
        )
        generic = generic_timeline_establishment_units(
            units,
            self.rules,
        )
        self.assertEqual(len(generic), 1)

    def test_generic_capital_setup_row_recovers_establishment(self) -> None:
        frozen = FrozenRange(
            "PATCH",
            "301581",
            "黄山谷捷",
            "equity_history",
            10,
            20,
            "9",
            "19",
            "DISCLOSED",
            "REPORTING_PERIOD_EQUITY_AND_SHAREHOLDER_CHANGES",
            "FROZEN_FOR_EVENT_EXTRACTION",
            "ACCEPTED",
            "SOURCE",
        )
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "2012年6月\n"
                    "2021年4月\n"
                    "2022年9月\n"
                    "注册资本1,000万元，由昆山谷捷出资设立\n"
                    "昆山谷捷将其所持谷捷有限股权转让\n"
                    "谷捷有限吸收合并昆山谷捷\n"
                    "整体变更为股份公司"
                )
            },
            10,
            20,
        )
        recovered = recover_missing_issuer_establishment(
            frozen,
            units,
            [],
            self.rules,
        )
        self.assertEqual(len(recovered), 1)
        self.assertEqual(
            recovered[0].event_period,
            "2012-06",
        )
        self.assertEqual(
            recovered[0].title,
            "2012-06，发行人前身设立",
        )
        self.assertIn(
            "generic_capital_setup_row",
            recovered[0].signals,
        )

    def test_absorbed_long_section_opening_does_not_block(self) -> None:
        frozen = FrozenRange(
            "PATCH",
            "301581",
            "黄山谷捷",
            "equity_history",
            10,
            20,
            "9",
            "19",
            "DISCLOSED",
            "REPORTING_PERIOD_EQUITY_AND_SHAREHOLDER_CHANGES",
            "FROZEN_FOR_EVENT_EXTRACTION",
            "ACCEPTED",
            "SOURCE",
        )
        units = build_text_units(
            frozen.company_id,
            {
                10: (
                    "吸收合并前昆山谷捷基本情况\n"
                    "公司名称 昆山谷捷金属制品有限公司\n"
                    "成立日期 2009年7月6日\n"
                    "注册资本200万元\n"
                    "经营范围金属制品生产\n"
                    "股东构成黄山供销集团持股78%\n"
                    "谷捷有限吸收合并昆山谷捷"
                ),
                15: (
                    "2012年6月\n"
                    "2021年4月\n"
                    "注册资本1,000万元，由昆山谷捷出资设立\n"
                    "第一次股权转让"
                ),
            },
            10,
            20,
        )
        absorbed = EventSection(
            section_id="ABSORBED",
            company_id=frozen.company_id,
            source_patch_id=frozen.patch_id,
            source_kind="HEADING_SECTION",
            title="吸收合并前昆山谷捷基本情况",
            title_unit_index=0,
            heading_level=3,
            event_type_candidate="LIMITED_COMPANY_ESTABLISHMENT",
            event_type_candidates=["LIMITED_COMPANY_ESTABLISHMENT"],
            explicit_combined_event=False,
            ordinal_labels=[],
            timeline_summary=False,
            units=units[:7],
            date_roles={"other_date": ["2009-07-06"]},
            event_period="2009-07-06",
            event_date_primary_role="other_date",
            event_date_selection_basis="TEST",
            negative_disclosure=False,
            negative_reason=None,
            entity_scope_candidate="ISSUER_OR_UNRESOLVED",
            service_provider_mentions=[],
            signals=[],
            confidence=0.8,
            review_reasons=[],
        )
        self.assertFalse(
            section_contains_exact_issuer_establishment(
                absorbed,
                frozen.company_short_name,
                self.rules,
            )
        )
        recovered = recover_missing_issuer_establishment(
            frozen,
            units,
            [absorbed],
            self.rules,
        )
        self.assertEqual(len(recovered), 1)
        self.assertEqual(
            recovered[0].event_period,
            "2012-06",
        )


if __name__ == "__main__":
    unittest.main()
