from pathlib import Path
import sys

import pytest

sys.path.append(str(Path(__file__).resolve().parents[1]))

from description_utils import normalize_description  # noqa: E402


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("RT 普通預金 ペイペイ", "RT (ペイペイ)"),
        ("取扱店: 51317 カード", "カード"),
        ("取扱店: 51317 払込み", "払込み"),
        (None, ""),
    ],
)
def test_normalize_description(raw, expected):
    assert normalize_description(raw) == expected
