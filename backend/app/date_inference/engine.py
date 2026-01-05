"""Date inference engine for Japanese passbook 2-digit years."""
from __future__ import annotations

from datetime import datetime
from typing import List, Optional, Tuple

from .constants import KNOWN_SEIREKI_BANKS, KNOWN_WAREKI_BANKS, ERA_START_YEARS
from .types import (
    DateInferenceResult,
    DateInferenceContext,
    InferenceMethod,
    DateInterpretation,
    DateAlternative,
)


class DateInferenceEngine:
    """
    日本の金融機関通帳における2桁年号の推論エンジン。

    判定優先順位:
    1. 確実な判定: 32以上 → 西暦2桁 (2032年以降)
    2. 高確率判定: 8-31 → 平成の可能性が高い
    3. 曖昧ゾーン: 1-7 → 金融機関コードとコンテキストで判断
    """

    # 現在の令和年（動的に計算）
    CURRENT_YEAR = datetime.now().year
    CURRENT_REIWA_YEAR = CURRENT_YEAR - 2018  # 2024 → R6

    # 年号の境界値
    DEFINITE_SEIREKI_THRESHOLD = 32  # これ以上は確実に西暦2桁
    HIGH_PROB_HEISEI_MIN = 8         # 平成8年(1996)以上は高確率で平成
    AMBIGUOUS_MAX = 7                # 1-7は曖昧ゾーン

    def __init__(self, learned_formats: Optional[dict] = None):
        """
        Args:
            learned_formats: 学習済みの金融機関フォーマット {bank_code: DateInterpretation}
        """
        self._learned_formats = learned_formats or {}

    def infer_date(
        self,
        two_digit_year: int,
        month: int,
        day: int,
        context: Optional[DateInferenceContext] = None,
    ) -> DateInferenceResult:
        """
        2桁年号から西暦4桁年を推論する。

        Args:
            two_digit_year: 2桁の年（0-99）
            month: 月（1-12）
            day: 日（1-31）
            context: 推論コンテキスト（金融機関コード、前後の日付など）

        Returns:
            DateInferenceResult: 推論結果
        """
        context = context or DateInferenceContext()

        # Step 1: 確実な判定（32以上は西暦2桁）
        if two_digit_year >= self.DEFINITE_SEIREKI_THRESHOLD:
            return self._create_definite_seireki_result(two_digit_year, month, day)

        # Step 2: 金融機関コードによる判定
        if context.bank_code:
            bank_result = self._infer_from_bank_code(
                two_digit_year, month, day, context.bank_code
            )
            if bank_result:
                return bank_result

        # Step 3: 高確率判定（8-31は平成の可能性が高い）
        if two_digit_year >= self.HIGH_PROB_HEISEI_MIN:
            return self._create_high_prob_heisei_result(two_digit_year, month, day)

        # Step 4: 曖昧ゾーン（1-7）- コンテキスト解析
        if context.surrounding_dates:
            context_result = self._infer_from_context(
                two_digit_year, month, day, context
            )
            if context_result:
                return context_result

        # Step 5: デフォルト - 現在の令和年以下なら令和、超えたら平成
        return self._create_default_result(two_digit_year, month, day)

    def infer_dates_batch(
        self,
        dates: List[Tuple[int, int, int]],
        context: Optional[DateInferenceContext] = None,
    ) -> List[DateInferenceResult]:
        """
        複数の日付をバッチ処理し、コンテキストを共有して推論精度を向上させる。

        Args:
            dates: [(year, month, day), ...] のリスト
            context: 共通コンテキスト

        Returns:
            推論結果のリスト
        """
        context = context or DateInferenceContext()

        # 全日付を文字列化してコンテキストに追加
        date_strings = [f"{y:02d}-{m:02d}-{d:02d}" for y, m, d in dates]
        context.surrounding_dates = date_strings

        results = []
        for i, (year, month, day) in enumerate(dates):
            context.current_index = i
            result = self.infer_date(year, month, day, context)
            results.append(result)

        return results

    def _create_definite_seireki_result(
        self, two_digit_year: int, month: int, day: int
    ) -> DateInferenceResult:
        """32以上の確実な西暦2桁判定"""
        full_year = 2000 + two_digit_year
        return DateInferenceResult(
            year=full_year,
            month=month,
            day=day,
            confidence=1.0,
            inference_method=InferenceMethod.DEFINITE_SEIREKI,
            is_ambiguous=False,
            original_year_digits=2,
        )

    def _create_high_prob_heisei_result(
        self, two_digit_year: int, month: int, day: int
    ) -> DateInferenceResult:
        """8-31の高確率平成判定"""
        heisei_year = 1988 + two_digit_year  # H1 = 1989

        # 代替候補も生成
        alternatives = []

        # 西暦2桁の可能性
        seireki_year = 2000 + two_digit_year
        if self._is_valid_date(seireki_year, month, day):
            alternatives.append(DateAlternative(
                year=seireki_year,
                interpretation=DateInterpretation.SEIREKI,
            ))

        return DateInferenceResult(
            year=heisei_year,
            month=month,
            day=day,
            confidence=0.85,
            inference_method=InferenceMethod.HIGH_PROB_HEISEI,
            is_ambiguous=len(alternatives) > 0,
            original_year_digits=2,
            alternatives=alternatives,
        )

    def _infer_from_bank_code(
        self, two_digit_year: int, month: int, day: int, bank_code: str
    ) -> Optional[DateInferenceResult]:
        """金融機関コードから日付形式を判定"""

        # 学習済みフォーマットを優先
        if bank_code in self._learned_formats:
            interpretation = self._learned_formats[bank_code]
            return self._create_result_from_interpretation(
                two_digit_year, month, day, interpretation,
                InferenceMethod.USER_CONFIRMED, confidence=0.95
            )

        # 既知の西暦銀行
        if bank_code in KNOWN_SEIREKI_BANKS:
            seireki_year = 2000 + two_digit_year
            if self._is_valid_date(seireki_year, month, day):
                return DateInferenceResult(
                    year=seireki_year,
                    month=month,
                    day=day,
                    confidence=0.9,
                    inference_method=InferenceMethod.BANK_LOOKUP,
                    is_ambiguous=False,
                    original_year_digits=2,
                )

        # 既知の和暦銀行
        if bank_code in KNOWN_WAREKI_BANKS:
            # 令和か平成かを判定
            if two_digit_year <= self.CURRENT_REIWA_YEAR:
                # 令和として解釈可能
                reiwa_year = 2018 + two_digit_year
                heisei_year = 1988 + two_digit_year

                # 両方有効な日付なら曖昧
                reiwa_valid = self._is_valid_date(reiwa_year, month, day)
                heisei_valid = self._is_valid_date(heisei_year, month, day)

                if reiwa_valid and heisei_valid:
                    # コンテキストがない場合は令和を優先（最近の取引が多いため）
                    alternatives = [DateAlternative(
                        year=heisei_year,
                        interpretation=DateInterpretation.HEISEI,
                    )]
                    return DateInferenceResult(
                        year=reiwa_year,
                        month=month,
                        day=day,
                        confidence=0.7,
                        inference_method=InferenceMethod.BANK_LOOKUP,
                        is_ambiguous=True,
                        original_year_digits=2,
                        alternatives=alternatives,
                    )
                elif reiwa_valid:
                    return DateInferenceResult(
                        year=reiwa_year,
                        month=month,
                        day=day,
                        confidence=0.85,
                        inference_method=InferenceMethod.BANK_LOOKUP,
                        is_ambiguous=False,
                        original_year_digits=2,
                    )

            # 令和範囲外なら平成
            heisei_year = 1988 + two_digit_year
            if self._is_valid_date(heisei_year, month, day):
                return DateInferenceResult(
                    year=heisei_year,
                    month=month,
                    day=day,
                    confidence=0.85,
                    inference_method=InferenceMethod.BANK_LOOKUP,
                    is_ambiguous=False,
                    original_year_digits=2,
                )

        return None

    def _infer_from_context(
        self, two_digit_year: int, month: int, day: int,
        context: DateInferenceContext
    ) -> Optional[DateInferenceResult]:
        """前後の日付コンテキストから推論"""

        if not context.surrounding_dates or context.current_index is None:
            return None

        # 前後の日付から年号の傾向を分析
        surrounding_years = self._extract_years_from_context(context)
        if not surrounding_years:
            return None

        # 連続性チェック：前後の日付との整合性を確認
        interpretation = self._determine_interpretation_from_context(
            two_digit_year, surrounding_years, context.current_index
        )

        if interpretation:
            return self._create_result_from_interpretation(
                two_digit_year, month, day, interpretation,
                InferenceMethod.CONTEXT_BASED, confidence=0.8
            )

        return None

    def _extract_years_from_context(
        self, context: DateInferenceContext
    ) -> List[Tuple[int, Optional[int]]]:
        """
        コンテキストから年情報を抽出。

        Returns:
            [(2桁年, 推論済み4桁年 or None), ...]
        """
        results = []
        for date_str in context.surrounding_dates:
            parts = date_str.replace("/", "-").split("-")
            if len(parts) >= 1 and parts[0].isdigit():
                two_digit = int(parts[0])
                # 32以上は確実に西暦
                if two_digit >= self.DEFINITE_SEIREKI_THRESHOLD:
                    results.append((two_digit, 2000 + two_digit))
                # 8-31は高確率で平成
                elif two_digit >= self.HIGH_PROB_HEISEI_MIN:
                    results.append((two_digit, 1988 + two_digit))
                else:
                    results.append((two_digit, None))
        return results

    def _determine_interpretation_from_context(
        self, two_digit_year: int,
        surrounding_years: List[Tuple[int, Optional[int]]],
        current_index: int
    ) -> Optional[DateInterpretation]:
        """コンテキストから最適な解釈を決定"""

        # 確定済みの年から傾向を判断
        confirmed_years = [y for _, y in surrounding_years if y is not None]
        if not confirmed_years:
            return None

        # 年の範囲を確認
        min_year = min(confirmed_years)
        max_year = max(confirmed_years)

        # 令和候補
        reiwa_year = 2018 + two_digit_year
        # 平成候補
        heisei_year = 1988 + two_digit_year

        # 範囲内に収まるかチェック
        # 通帳は通常時系列順なので、前後の日付との連続性を重視
        reiwa_fits = min_year <= reiwa_year <= max_year + 5  # 少し余裕を持たせる
        heisei_fits = min_year - 5 <= heisei_year <= max_year

        if reiwa_fits and not heisei_fits:
            return DateInterpretation.REIWA
        elif heisei_fits and not reiwa_fits:
            return DateInterpretation.HEISEI

        # 両方フィットする場合は、より連続性の高い方を選択
        if current_index > 0:
            prev_year = surrounding_years[current_index - 1][1]
            if prev_year:
                # 前の日付との差が小さい方を選択
                reiwa_diff = abs(reiwa_year - prev_year)
                heisei_diff = abs(heisei_year - prev_year)
                if reiwa_diff < heisei_diff:
                    return DateInterpretation.REIWA
                elif heisei_diff < reiwa_diff:
                    return DateInterpretation.HEISEI

        return None

    def _create_default_result(
        self, two_digit_year: int, month: int, day: int
    ) -> DateInferenceResult:
        """デフォルトの推論結果を生成"""

        # 現在の令和年以下なら令和を優先
        if two_digit_year <= self.CURRENT_REIWA_YEAR:
            reiwa_year = 2018 + two_digit_year
            heisei_year = 1988 + two_digit_year

            alternatives = []
            if self._is_valid_date(heisei_year, month, day):
                alternatives.append(DateAlternative(
                    year=heisei_year,
                    interpretation=DateInterpretation.HEISEI,
                ))

            return DateInferenceResult(
                year=reiwa_year,
                month=month,
                day=day,
                confidence=0.6,
                inference_method=InferenceMethod.DEFAULT_REIWA,
                is_ambiguous=True,
                original_year_digits=2,
                alternatives=alternatives,
            )

        # 令和範囲外なら平成
        heisei_year = 1988 + two_digit_year
        return DateInferenceResult(
            year=heisei_year,
            month=month,
            day=day,
            confidence=0.7,
            inference_method=InferenceMethod.HIGH_PROB_HEISEI,
            is_ambiguous=False,
            original_year_digits=2,
        )

    def _create_result_from_interpretation(
        self, two_digit_year: int, month: int, day: int,
        interpretation: DateInterpretation,
        method: InferenceMethod,
        confidence: float,
    ) -> DateInferenceResult:
        """解釈に基づいて結果を生成"""

        if interpretation == DateInterpretation.SEIREKI:
            full_year = 2000 + two_digit_year
        elif interpretation == DateInterpretation.REIWA:
            full_year = 2018 + two_digit_year
        elif interpretation == DateInterpretation.HEISEI:
            full_year = 1988 + two_digit_year
        elif interpretation == DateInterpretation.SHOWA:
            full_year = 1925 + two_digit_year
        else:
            full_year = 2000 + two_digit_year

        return DateInferenceResult(
            year=full_year,
            month=month,
            day=day,
            confidence=confidence,
            inference_method=method,
            is_ambiguous=False,
            original_year_digits=2,
        )

    def _is_valid_date(self, year: int, month: int, day: int) -> bool:
        """日付が有効かどうかをチェック"""
        try:
            datetime(year, month, day)
            return True
        except ValueError:
            return False

    def learn_bank_format(
        self, bank_code: str, interpretation: DateInterpretation
    ) -> None:
        """金融機関の日付形式を学習"""
        self._learned_formats[bank_code] = interpretation

    def get_learned_formats(self) -> dict:
        """学習済みフォーマットを取得"""
        return self._learned_formats.copy()
