"""Type definitions for date inference."""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class InferenceMethod(str, Enum):
    """推論方法"""
    DEFINITE_SEIREKI = "definite_seireki"      # 確実に西暦（32以上）
    DEFINITE_WAREKI = "definite_wareki"        # 確実に和暦
    HIGH_PROB_HEISEI = "high_prob_heisei"      # 高確率で平成（8-31）
    CONTEXT_BASED = "context_based"            # コンテキストから判断
    BANK_LOOKUP = "bank_lookup"                # 金融機関コードから判断
    USER_CONFIRMED = "user_confirmed"          # ユーザー確認済み
    DEFAULT_REIWA = "default_reiwa"            # デフォルト令和推論


class DateInterpretation(str, Enum):
    """日付の解釈"""
    SEIREKI = "seireki"   # 西暦
    HEISEI = "heisei"     # 平成
    REIWA = "reiwa"       # 令和
    SHOWA = "showa"       # 昭和


@dataclass
class DateAlternative:
    """曖昧な場合の代替候補"""
    year: int
    interpretation: DateInterpretation


@dataclass
class DateInferenceResult:
    """日付推論の結果"""
    year: int                                    # 推論された西暦年（4桁）
    month: int                                   # 月（1-12）
    day: int                                     # 日（1-31）
    confidence: float                            # 推論の確信度（0.0-1.0）
    inference_method: InferenceMethod            # 推論方法
    is_ambiguous: bool                           # 曖昧性があるか
    original_year_digits: int                    # 元の年の桁数（2 or 4）
    alternatives: List[DateAlternative] = field(default_factory=list)  # 曖昧な場合の候補

    def to_iso_date(self) -> str:
        """YYYY-MM-DD形式の文字列を返す"""
        return f"{self.year:04d}-{self.month:02d}-{self.day:02d}"

    def to_wareki_display(self) -> str:
        """和暦表示を返す（例: R6/12/25, H30/05/01）"""
        if self.year >= 2019:
            era_year = self.year - 2018
            return f"R{era_year}/{self.month:02d}/{self.day:02d}"
        elif self.year >= 1989:
            era_year = self.year - 1988
            return f"H{era_year}/{self.month:02d}/{self.day:02d}"
        elif self.year >= 1926:
            era_year = self.year - 1925
            return f"S{era_year}/{self.month:02d}/{self.day:02d}"
        else:
            return f"{self.year}/{self.month:02d}/{self.day:02d}"


@dataclass
class DateInferenceContext:
    """日付推論のコンテキスト情報"""
    bank_code: Optional[str] = None                # 金融機関コード（4桁）
    bank_name: Optional[str] = None                # 金融機関名
    surrounding_dates: List[str] = field(default_factory=list)  # 前後の日付リスト
    current_index: Optional[int] = None            # 現在の日付のインデックス
    user_confirmed_format: Optional[DateInterpretation] = None  # ユーザー確認済みの形式


@dataclass
class LearnedBankFormat:
    """学習済み金融機関フォーマット"""
    bank_code: str
    bank_name: Optional[str]
    format: DateInterpretation
    confirmed_at: str  # ISO format datetime
    sample_date: str   # 確認に使用した日付
