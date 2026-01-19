"""Shared helpers for normalizing bank transaction descriptions."""
from __future__ import annotations

import re
import unicodedata
from typing import Any

DESCRIPTION_CLEANUPS = (":selected:",)
# Pattern to remove branch info prefix (supports both half-width and full-width)
CARD_PREFIX_PATTERN = re.compile(
    r"^(?:取扱店|取扱店番号|店番|店舗番号|取扱局)[\s:：]*[0-9０-９]+[\s　]*",
    re.IGNORECASE
)
# ゆうちょ銀行の取扱店番号パターン（先頭の4-5桁の数字）
# 例: "03050 1,000,000通帳" → "1,000,000通帳"
YUUCHO_BRANCH_PREFIX_PATTERN = re.compile(
    r"^[0-9０-９]{4,5}[\s　]+(?=[0-9０-９,，]+)"
)
NUMERIC_BRACKETS_PATTERN = re.compile(r"[（(]\s*[0-9０-９]+\s*[）)]")
PAYPAY_PATTERN = re.compile(r"RT.*ペイペイ")
CARD_EXACT_PATTERN = re.compile(r"^カード(?:\s.*)?$")
PAYMENT_KEYWORDS = ("払込", "払込み", "払込金", "払込料")

REPLACEMENTS = (
    # 銀行・金融機関
    ("ﾌﾘｺﾐ", "振込"),
    ("フリコミ", "振込"),
    ("ﾌﾘｶｴ", "振替"),
    ("フリカエ", "振替"),
    ("ﾐﾂｲｽﾐﾄﾓ", "三井住友"),
    ("ミツイスミトモ", "三井住友"),
    ("ＳＭＢＣ", "三井住友"),
    ("SMBC", "三井住友"),
    ("ﾐｽﾞﾎ", "みずほ"),
    ("ミズホ", "みずほ"),
    ("ﾅﾝﾄ", "南都"),
    ("ナント", "南都"),
    ("ﾕｳﾁｮ", "ゆうちょ"),
    ("ユウチョ", "ゆうちょ"),
    ("ｿﾞｷﾝ", "送金"),
    ("ソウキン", "送金"),
    ("ﾋｷｵﾄｼ", "引落"),
    ("ヒキオトシ", "引落"),
    ("ﾃｽｳﾘｮｳ", "手数料"),
    ("テスウリョウ", "手数料"),
    # 証券・保険
    ("ｼｮｳｹﾝ", "証券"),
    ("ショウケン", "証券"),
    ("ﾎｹﾝ", "保険"),
    ("ホケン", "保険"),
    ("ｾｲﾒｲﾎｹﾝ", "生命保険"),
    ("セイメイホケン", "生命保険"),
    # 税金・行政
    ("ｼｴﾝ", "支援"),
    ("シエン", "支援"),
    ("ｶﾝﾌﾟ", "還付"),
    ("カンプ", "還付"),
    ("ﾘｿｸ", "利息"),
    ("リソク", "利息"),
    ("ｷﾝﾘ", "金利"),
    ("キンリ", "金利"),
    ("ﾈﾝｷﾝ", "年金"),
    ("ネンキン", "年金"),
    ("ｺﾞｾﾝﾀｸ", "税"),
    ("ゼイキン", "税金"),
    # 公共料金
    ("ﾃﾞﾝｷ", "電気"),
    ("デンキ", "電気"),
    ("ｶﾞｽ", "ガス"),
    ("ｽｲﾄﾞｳ", "水道"),
    ("スイドウ", "水道"),
    # 給与・報酬
    ("ｷｭｳﾖ", "給与"),
    ("キュウヨ", "給与"),
    ("ﾎｳｼｭｳ", "報酬"),
    ("ホウシュウ", "報酬"),
    ("ﾎﾞｰﾅｽ", "賞与"),
    ("ボーナス", "賞与"),
)


def normalize_description(value: Any) -> str:
    """Normalize raw bank description strings for CSV / JSON outputs."""
    if value is None:
        return ""
    text = str(value)
    if not text.strip():
        return ""

    normalized = unicodedata.normalize("NFKC", text)
    normalized = normalized.replace("\u3000", " ")
    for token in DESCRIPTION_CLEANUPS:
        normalized = normalized.replace(token, "")
    normalized = re.sub(r"\s+", " ", normalized).strip()
    normalized = CARD_PREFIX_PATTERN.sub("", normalized).strip()
    # ゆうちょ銀行の取扱店番号を除去
    normalized = YUUCHO_BRANCH_PREFIX_PATTERN.sub("", normalized).strip()

    for before, after in REPLACEMENTS:
        normalized = normalized.replace(before, after)

    normalized = NUMERIC_BRACKETS_PATTERN.sub("", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()

    if PAYPAY_PATTERN.search(normalized) and "ペイペイ" in normalized:
        return "RT (ペイペイ)"

    for keyword in PAYMENT_KEYWORDS:
        if keyword in normalized:
            return "払込み"

    if normalized == "カード" or CARD_EXACT_PATTERN.fullmatch(normalized):
        return "カード"

    return normalized

__all__ = ["normalize_description"]
