import unittest

from pipeline.pevc_paths import (
    build_entities_and_paths,
    classify_investor_name,
    evidence_entity_candidates,
    is_noise_name,
    issuer_alias_map,
    issuer_self_entities_for_validation,
    has_strong_pevc_name_evidence,
    validate_outputs,
    is_issuer_name,
    split_compound_names,
)


class PevcPathIdentificationV032Tests(unittest.TestCase):
    def test_table_headers_are_noise(self) -> None:
        for value in (
            "序号发行对象名称",
            "数量（股",
            "金额（元",
            "新增股本的",
            "50元的价格",
        ):
            self.assertTrue(is_noise_name(value))

    def test_strong_fund_is_pevc(self) -> None:
        _, status, _, confidence = classify_investor_name(
            "上海测试股权投资基金合伙企业（有限合伙）"
        )
        self.assertEqual(status, "PEVC_CANDIDATE")
        self.assertGreaterEqual(confidence, 0.9)

    def test_evidence_extracts_legal_investor(self) -> None:
        values = evidence_entity_candidates(
            (
                "发行对象名称上海测试股权投资基金"
                "合伙企业（有限合伙）认购100万股"
            ),
            "DIRECTIONAL_FINANCING",
        )
        self.assertIn(
            "上海测试股权投资基金合伙企业（有限合伙）",
            [name for name, _ in values],
        )

    def test_reorganization_is_not_investment_path(self) -> None:
        events = [{
            "event_id": "EVT-1",
            "candidate_event_id": "CE-1",
            "company_id": "123456",
        }]
        transactions = [{
            "transaction_id": "TRX-1",
            "event_id": "EVT-1",
            "company_id": "123456",
            "transaction_type": "OVERALL_CHANGE",
            "investor_party_ids": ["PTY-1"],
            "transferee_party_ids": [],
            "transferor_party_ids": [],
            "transaction_date": "2021-01",
        }]
        parties = [{
            "party_id": "PTY-1",
            "party_name_normalized": "测试投资有限公司",
            "party_name_raw": "测试投资有限公司",
        }]
        evidence = [{
            "candidate_event_id": "CE-1",
            "evidence_id": "E1",
            "evidence_text": "整体变更",
        }]
        entities, paths, reviews, discarded = (
            build_entities_and_paths(
                events,
                transactions,
                parties,
                evidence,
            )
        )
        self.assertEqual(entities, [])
        self.assertEqual(paths, [])

    def test_missing_investor_creates_review_not_fake_entity(self) -> None:
        events = [{
            "event_id": "EVT-1",
            "candidate_event_id": "CE-1",
            "company_id": "123456",
        }]
        transactions = [{
            "transaction_id": "TRX-1",
            "event_id": "EVT-1",
            "company_id": "123456",
            "transaction_type": "CAPITAL_INCREASE",
            "investor_party_ids": ["PTY-1"],
            "transferee_party_ids": [],
            "transferor_party_ids": [],
            "transaction_date": "2021-01",
        }]
        parties = [{
            "party_id": "PTY-1",
            "party_name_normalized": "序号发行对象名称",
            "party_name_raw": "序号发行对象名称",
        }]
        evidence = [{
            "candidate_event_id": "CE-1",
            "evidence_id": "E1",
            "evidence_text": "本次增资情况如下",
        }]
        entities, paths, reviews, discarded = (
            build_entities_and_paths(
                events,
                transactions,
                parties,
                evidence,
            )
        )
        self.assertEqual(entities, [])
        self.assertEqual(paths, [])
        self.assertEqual(len(reviews), 1)
        self.assertEqual(
            reviews[0]["record_type"],
            "MISSING_INVESTOR_FOR_TRANSACTION",
        )

    def test_valid_transfer_builds_direct_path(self) -> None:
        events = [{
            "event_id": "EVT-1",
            "candidate_event_id": "CE-1",
            "company_id": "123456",
        }]
        transactions = [{
            "transaction_id": "TRX-1",
            "event_id": "EVT-1",
            "company_id": "123456",
            "transaction_type": "EQUITY_TRANSFER",
            "investor_party_ids": [],
            "transferee_party_ids": ["PTY-1"],
            "transferor_party_ids": ["PTY-2"],
            "transaction_date": "2021-01",
        }]
        parties = [{
            "party_id": "PTY-1",
            "party_name_normalized": "测试投资有限公司",
            "party_name_raw": "测试投资有限公司",
        }]
        evidence = [{
            "candidate_event_id": "CE-1",
            "evidence_id": "E1",
            "evidence_text": (
                "原股东将股权转让给测试投资有限公司"
            ),
        }]
        entities, paths, reviews, discarded = (
            build_entities_and_paths(
                events,
                transactions,
                parties,
                evidence,
            )
        )
        self.assertEqual(len(entities), 1)
        self.assertEqual(len(paths), 1)
        self.assertEqual(
            paths[0].entry_method,
            "TRANSFER_ENTRY",
        )
        self.assertEqual(
            paths[0].direct_or_indirect,
            "DIRECT",
        )

    def test_issuer_full_name_is_excluded(self) -> None:
        events = [{
            "event_id": "EVT-1",
            "candidate_event_id": "CE-1",
            "company_id": "001282",
            "company_short_name": "三联锻造",
        }]
        aliases = issuer_alias_map(events)
        self.assertTrue(
            is_issuer_name(
                "芜湖三联锻造股份有限公司",
                aliases["001282"],
            )
        )

    def test_compound_shareholders_are_split(self) -> None:
        values = split_compound_names(
            "前公司股东泽升贸易、罗世兵、共青城泽升、达晨创联基金"
        )
        self.assertEqual(
            values,
            [
                "泽升贸易",
                "罗世兵",
                "共青城泽升",
                "达晨创联基金",
            ],
        )

    def test_fund_short_name_is_pevc(self) -> None:
        _, status, _, confidence = classify_investor_name(
            "达晨创联基金"
        )
        self.assertEqual(
            status,
            "PEVC_CANDIDATE",
        )
        self.assertGreaterEqual(confidence, 0.9)

    def test_compound_evidence_creates_fund_entity_only(self) -> None:
        events = [{
            "event_id": "EVT-1",
            "candidate_event_id": "CE-1",
            "company_id": "603418",
            "company_short_name": "友升股份",
        }]
        transactions = [{
            "transaction_id": "TRX-1",
            "event_id": "EVT-1",
            "company_id": "603418",
            "transaction_type": "CAPITAL_INCREASE",
            "investor_party_ids": [],
            "transferee_party_ids": [],
            "transferor_party_ids": [],
            "transaction_date": "2020-09",
        }]
        parties = []
        evidence = [{
            "candidate_event_id": "CE-1",
            "evidence_id": "E1",
            "evidence_text": (
                "前公司股东泽升贸易、罗世兵、"
                "共青城泽升、达晨创联基金参与本次增资"
            ),
        }]
        entities, paths, reviews, discarded = (
            build_entities_and_paths(
                events,
                transactions,
                parties,
                evidence,
            )
        )
        pevc_names = [
            item.investor_name_normalized
            for item in entities
            if item.pevc_candidate_status
            == "PEVC_CANDIDATE"
        ]
        self.assertIn(
            "达晨创联基金",
            pevc_names,
        )
        self.assertFalse(
            any(
                "、" in item.investor_name_normalized
                for item in entities
            )
        )

    def test_issuer_evidence_does_not_create_path(self) -> None:
        events = [{
            "event_id": "EVT-1",
            "candidate_event_id": "CE-1",
            "company_id": "001282",
            "company_short_name": "三联锻造",
        }]
        transactions = [{
            "transaction_id": "TRX-1",
            "event_id": "EVT-1",
            "company_id": "001282",
            "transaction_type": "CAPITAL_INCREASE",
            "investor_party_ids": [],
            "transferee_party_ids": [],
            "transferor_party_ids": [],
            "transaction_date": "2019-12",
        }]
        evidence = [{
            "candidate_event_id": "CE-1",
            "evidence_id": "E1",
            "evidence_text": (
                "芜湖三联锻造股份有限公司完成本次增资"
            ),
        }]
        entities, paths, reviews, discarded = (
            build_entities_and_paths(
                events,
                transactions,
                [],
                evidence,
            )
        )
        self.assertEqual(entities, [])
        self.assertEqual(paths, [])

    def test_scope_safe_issuer_self_validation(self) -> None:
        from pipeline.pevc_paths import InvestorEntity

        events = [{
            "event_id": "EVT-1",
            "candidate_event_id": "CE-1",
            "company_id": "001282",
            "company_short_name": "三联锻造",
        }]
        entities = [InvestorEntity(
            investor_entity_id="INV-001282-0001",
            company_id="001282",
            investor_name_raw="芜湖三联锻造股份有限公司",
            investor_name_normalized="芜湖三联锻造股份有限公司",
            investor_type_candidate="LEGAL_ENTITY_OR_STRATEGIC_INVESTOR",
            pevc_candidate_status="NOT_PEVC_OR_UNRESOLVED_STRATEGIC",
            classification_basis=[],
            source_party_ids=[],
            source_event_ids=["EVT-1"],
            source_transaction_ids=["TRX-1"],
            evidence_ids=["E1"],
            extraction_source="TEST",
            confidence=0.9,
            review_required=False,
            review_reasons=[],
        )]
        result = issuer_self_entities_for_validation(
            entities,
            events,
        )
        self.assertEqual(len(result), 1)
        self.assertEqual(
            result[0].investor_entity_id,
            "INV-001282-0001",
        )

    def test_fund_short_name_is_strong_validation_evidence(self) -> None:
        self.assertTrue(
            has_strong_pevc_name_evidence(
                "达晨创联基金"
            )
        )

    def test_fund_short_name_candidate_passes_validation(self) -> None:
        from pipeline.pevc_paths import (
            InvestorEntity,
            InvestmentPath,
        )

        entity = InvestorEntity(
            investor_entity_id="INV-603418-0001",
            company_id="603418",
            investor_name_raw="达晨创联基金",
            investor_name_normalized="达晨创联基金",
            investor_type_candidate=(
                "PRIVATE_EQUITY_OR_VENTURE_ENTITY"
            ),
            pevc_candidate_status="PEVC_CANDIDATE",
            classification_basis=[
                "名称命中强PE/VC关键词",
                "名称可能为招股书使用的基金简称",
            ],
            source_party_ids=[],
            source_event_ids=["EVT-603418-0001"],
            source_transaction_ids=["TRX-603418-0001"],
            evidence_ids=["E1"],
            extraction_source="EVIDENCE_COMPOUND_LIST",
            confidence=0.93,
            review_required=False,
            review_reasons=[],
        )
        path = InvestmentPath(
            investment_path_id="PATH-603418-00001",
            company_id="603418",
            investor_entity_id="INV-603418-0001",
            investor_name_normalized="达晨创联基金",
            event_id="EVT-603418-0001",
            transaction_id="TRX-603418-0001",
            entry_method="CAPITAL_INCREASE_ENTRY",
            investment_level="ISSUER_LEVEL",
            direct_or_indirect="DIRECT",
            transaction_date="2020-09",
            transferor_party_ids=[],
            evidence_ids=["E1"],
            path_status="AUTO_IDENTIFIED",
            confidence=0.92,
            review_required=False,
            review_reasons=[],
        )
        result = validate_outputs(
            [entity],
            [path],
        )
        self.assertEqual(
            result["validation_status"],
            "PASSED",
        )
        self.assertEqual(result["errors"], [])


if __name__ == "__main__":
    unittest.main()
