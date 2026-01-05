"""Tests for DateInferenceEngine."""
import pytest
from datetime import datetime

from backend.app.date_inference import (
    DateInferenceEngine,
    DateInferenceContext,
    DateInferenceResult,
    InferenceMethod,
    DateInterpretation,
)


class TestDateInferenceEngine:
    """DateInferenceEngine のテストスイート"""

    @pytest.fixture
    def engine(self):
        return DateInferenceEngine()

    # ============================================================
    # Step 1: 確実な判定（32以上は西暦2桁）
    # ============================================================
    def test_definite_seireki_32(self, engine):
        """32 → 2032年（確実に西暦2桁）"""
        result = engine.infer_date(32, 1, 1)
        assert result.year == 2032
        assert result.inference_method == InferenceMethod.DEFINITE_SEIREKI
        assert result.confidence == 1.0
        assert not result.is_ambiguous

    def test_definite_seireki_99(self, engine):
        """99 → 2099年（確実に西暦2桁）"""
        result = engine.infer_date(99, 12, 31)
        assert result.year == 2099
        assert result.inference_method == InferenceMethod.DEFINITE_SEIREKI

    # ============================================================
    # Step 2: 高確率判定（8-31は平成の可能性が高い）
    # ============================================================
    def test_high_prob_heisei_17(self, engine):
        """17 → H17 = 2005年（高確率で平成）"""
        result = engine.infer_date(17, 11, 24)
        assert result.year == 2005
        assert result.inference_method == InferenceMethod.HIGH_PROB_HEISEI
        assert result.confidence >= 0.8

    def test_high_prob_heisei_31(self, engine):
        """31 → H31 = 2019年（高確率で平成）"""
        result = engine.infer_date(31, 4, 30)
        assert result.year == 2019
        assert result.inference_method == InferenceMethod.HIGH_PROB_HEISEI

    def test_high_prob_heisei_8(self, engine):
        """8 → H8 = 1996年（高確率で平成）"""
        result = engine.infer_date(8, 6, 15)
        assert result.year == 1996
        assert result.inference_method == InferenceMethod.HIGH_PROB_HEISEI

    # ============================================================
    # Step 3: 曖昧ゾーン（1-7）
    # ============================================================
    def test_ambiguous_zone_1_defaults_to_reiwa(self, engine):
        """1 → R1 = 2019年（デフォルトは令和）"""
        result = engine.infer_date(1, 12, 6)
        assert result.year == 2019
        assert result.is_ambiguous or result.confidence < 0.9
        # alternatives should contain heisei option
        if result.alternatives:
            heisei_alt = next(
                (a for a in result.alternatives if a.interpretation == DateInterpretation.HEISEI),
                None
            )
            if heisei_alt:
                assert heisei_alt.year == 1989  # H1

    def test_ambiguous_zone_2(self, engine):
        """2 → R2 = 2020年（デフォルトは令和）"""
        result = engine.infer_date(2, 12, 15)
        assert result.year == 2020

    def test_ambiguous_zone_7(self, engine):
        """7 → R7 = 2025年（現在の令和年以下なら令和）"""
        current_reiwa = datetime.now().year - 2018
        result = engine.infer_date(7, 1, 1)
        if 7 <= current_reiwa:
            assert result.year == 2025
        else:
            # If current year is before R7, might fall back to Heisei
            assert result.year in [1995, 2025]

    # ============================================================
    # 金融機関コードによる判定
    # ============================================================
    def test_known_wareki_bank_ufj(self, engine):
        """三菱UFJ銀行（0005）は和暦"""
        context = DateInferenceContext(bank_code="0005")
        result = engine.infer_date(1, 12, 6, context)
        # UFJは和暦なので1は令和1年または平成1年
        assert result.year in [2019, 1989]
        assert result.inference_method == InferenceMethod.BANK_LOOKUP

    def test_known_seireki_bank_mizuho(self, engine):
        """みずほ銀行（0001）は西暦"""
        context = DateInferenceContext(bank_code="0001")
        result = engine.infer_date(17, 11, 24, context)
        # みずほは西暦なので17は2017年
        assert result.year == 2017
        assert result.inference_method == InferenceMethod.BANK_LOOKUP

    def test_known_seireki_bank_rokin(self, engine):
        """中央労働金庫（2952）は西暦"""
        context = DateInferenceContext(bank_code="2952")
        result = engine.infer_date(5, 3, 15, context)
        # 労金は西暦なので5は2005年
        assert result.year == 2005
        assert result.inference_method == InferenceMethod.BANK_LOOKUP

    # ============================================================
    # バッチ処理とコンテキスト推論
    # ============================================================
    def test_batch_inference_with_context(self, engine):
        """バッチ処理でコンテキストを共有"""
        dates = [
            (17, 11, 24),  # H17 = 2005
            (18, 3, 15),   # H18 = 2006
            (1, 12, 6),    # 曖昧だがコンテキストから平成と推論？
        ]
        results = engine.infer_dates_batch(dates)
        assert len(results) == 3
        assert results[0].year == 2005
        assert results[1].year == 2006
        # 3番目は前後のコンテキストから判断される

    # ============================================================
    # エッジケース
    # ============================================================
    def test_february_29_leap_year(self, engine):
        """2月29日の閏年チェック"""
        # 2020年は閏年（R2）
        result = engine.infer_date(2, 2, 29)
        assert result.year == 2020  # R2は閏年

    def test_invalid_date_handling(self, engine):
        """無効な日付でも推論自体は行う"""
        # 2月30日は存在しない
        result = engine.infer_date(5, 2, 30)
        # 年の推論は行われる
        assert result.year is not None

    # ============================================================
    # 学習機能
    # ============================================================
    def test_learn_bank_format(self, engine):
        """金融機関フォーマットの学習"""
        engine.learn_bank_format("9999", DateInterpretation.SEIREKI)
        context = DateInferenceContext(bank_code="9999")
        result = engine.infer_date(17, 11, 24, context)
        # 学習済みなので西暦として解釈
        assert result.year == 2017
        assert result.inference_method == InferenceMethod.USER_CONFIRMED

    def test_get_learned_formats(self, engine):
        """学習済みフォーマットの取得"""
        engine.learn_bank_format("1234", DateInterpretation.HEISEI)
        formats = engine.get_learned_formats()
        assert "1234" in formats
        assert formats["1234"] == DateInterpretation.HEISEI


class TestRealWorldScenarios:
    """実際の使用シナリオに基づくテスト"""

    @pytest.fixture
    def engine(self):
        return DateInferenceEngine()

    def test_scenario_01_12_06_as_reiwa1(self, engine):
        """01-12-06 → R1/12/06 = 2019-12-06（相続税案件の典型例）"""
        result = engine.infer_date(1, 12, 6)
        # デフォルトでは令和を優先
        assert result.year == 2019, f"Expected 2019 (R1), got {result.year}"

    def test_scenario_17_11_24_as_heisei17(self, engine):
        """17-11-24 → H17/11/24 = 2005-11-24（平成17年）"""
        result = engine.infer_date(17, 11, 24)
        assert result.year == 2005, f"Expected 2005 (H17), got {result.year}"

    def test_scenario_02_12_15_as_reiwa2(self, engine):
        """02-12-15 → R2/12/15 = 2020-12-15（令和2年）"""
        result = engine.infer_date(2, 12, 15)
        assert result.year == 2020, f"Expected 2020 (R2), got {result.year}"

    def test_scenario_26_01_29_as_heisei26(self, engine):
        """26-01-29 → H26/01/29 = 2014-01-29（平成26年）"""
        result = engine.infer_date(26, 1, 29)
        assert result.year == 2014, f"Expected 2014 (H26), got {result.year}"
