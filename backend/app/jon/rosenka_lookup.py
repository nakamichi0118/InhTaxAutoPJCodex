"""
国税庁路線価図URLルックアップ
GCSに保存されたスクレイピング結果から路線価図URLを検索
"""
from __future__ import annotations

import logging
import re
import unicodedata
from typing import Dict, List, Optional

import httpx

logger = logging.getLogger("uvicorn.error")

# GCSの公開URL
ROSENKA_DATA_URL = "https://storage.googleapis.com/souzoku-browser/rosenka/rosenka-flat.json"

# キャッシュ（サーバー起動中は保持）
_rosenka_cache: Optional[List[Dict]] = None
_rosenka_index: Optional[Dict[str, List[str]]] = None


def normalize_text(text: str) -> str:
    """テキストを正規化（全角→半角、トリム）"""
    text = unicodedata.normalize("NFKC", text)
    return text.strip()


def extract_district_base(district: str) -> str:
    """丁目部分を除いた町名を抽出（例: 梅田3丁目 → 梅田）"""
    # 丁目、番地などを除去
    match = re.match(r'^(.+?)[\d０-９一二三四五六七八九十]+', district)
    if match:
        return match.group(1)
    return district


async def load_rosenka_data() -> List[Dict]:
    """GCSから路線価データを読み込み（キャッシュあり）"""
    global _rosenka_cache

    if _rosenka_cache is not None:
        return _rosenka_cache

    try:
        logger.info("路線価データをGCSから読み込み中...")
        async with httpx.AsyncClient(timeout=60.0) as client:
            response = await client.get(ROSENKA_DATA_URL)
            response.raise_for_status()
            _rosenka_cache = response.json()
            logger.info(f"路線価データ読み込み完了: {len(_rosenka_cache)}件")
            return _rosenka_cache
    except Exception as e:
        logger.error(f"路線価データ読み込みエラー: {e}")
        return []


async def build_rosenka_index() -> Dict[str, List[str]]:
    """検索用インデックスを構築"""
    global _rosenka_index

    if _rosenka_index is not None:
        return _rosenka_index

    data = await load_rosenka_data()
    if not data:
        return {}

    # インデックス構築: "都道府県/市区町村/町名" → [URL, ...]
    index: Dict[str, List[str]] = {}

    for item in data:
        pref = normalize_text(item.get("prefecture", ""))
        city = normalize_text(item.get("city", ""))
        district = normalize_text(item.get("district", ""))
        url = item.get("url", "")

        if not url:
            continue

        # フルキー（丁目付き）
        full_key = f"{pref}/{city}/{district}"
        if full_key not in index:
            index[full_key] = []
        if url not in index[full_key]:
            index[full_key].append(url)

        # 町名ベースキー（丁目なし）
        district_base = extract_district_base(district)
        if district_base != district:
            base_key = f"{pref}/{city}/{district_base}"
            if base_key not in index:
                index[base_key] = []
            if url not in index[base_key]:
                index[base_key].append(url)

    _rosenka_index = index
    logger.info(f"路線価インデックス構築完了: {len(index)}キー")
    return index


async def lookup_rosenka_urls(
    prefecture: str,
    city: str,
    district: str,
) -> List[str]:
    """
    住所から路線価図URLを検索

    Args:
        prefecture: 都道府県（例: 大阪府, 大阪）
        city: 市区町村（例: 大阪市北区）
        district: 町名（例: 梅田3丁目, 梅田）

    Returns:
        路線価図URLのリスト
    """
    index = await build_rosenka_index()
    if not index:
        return []

    # 正規化
    pref = normalize_text(prefecture)
    city = normalize_text(city)
    district = normalize_text(district)

    # 都道府県の「県」「府」「都」を除去してマッチング
    pref_variations = [pref]
    for suffix in ["県", "府", "都", "道"]:
        if pref.endswith(suffix):
            pref_variations.append(pref[:-1])
        else:
            pref_variations.append(pref + suffix)

    # 丁目の数字を抽出
    district_match = re.search(r'(\d+)', district)
    district_num = district_match.group(1) if district_match else None
    district_base = extract_district_base(district)

    # 検索キーの候補
    search_keys = []
    for p in pref_variations:
        # フルマッチ
        search_keys.append(f"{p}/{city}/{district}")
        # 丁目付きバリエーション
        if district_num:
            search_keys.append(f"{p}/{city}/{district_base}{district_num}")
        # 町名ベースマッチ
        search_keys.append(f"{p}/{city}/{district_base}")

    # 検索
    for key in search_keys:
        if key in index:
            return index[key]

    # 部分一致検索（町名のみ）
    for key, urls in index.items():
        parts = key.split("/")
        if len(parts) >= 3:
            idx_pref, idx_city, idx_district = parts[0], parts[1], parts[2]
            # 都道府県と市区町村が一致し、町名が含まれる
            if any(p in idx_pref or idx_pref in p for p in pref_variations):
                if city in idx_city or idx_city in city:
                    if district_base in idx_district or idx_district in district_base:
                        return urls

    return []


def clear_cache():
    """キャッシュをクリア（テスト用）"""
    global _rosenka_cache, _rosenka_index
    _rosenka_cache = None
    _rosenka_index = None
