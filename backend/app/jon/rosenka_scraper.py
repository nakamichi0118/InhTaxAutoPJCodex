"""
国税庁路線価図URLスクレイパー
PowerAutomateのRosenkaフローをPythonで実装
"""
from __future__ import annotations

import logging
import re
import unicodedata
from typing import Dict, List, Optional
from urllib.parse import urljoin

import httpx
from bs4 import BeautifulSoup

logger = logging.getLogger("uvicorn.error")

# 都道府県コード（国税庁サイト用）
PREF_CODES = {
    "北海道": "hokkaido",
    "青森": "aomori", "青森県": "aomori",
    "岩手": "iwate", "岩手県": "iwate",
    "宮城": "miyagi", "宮城県": "miyagi",
    "秋田": "akita", "秋田県": "akita",
    "山形": "yamagata", "山形県": "yamagata",
    "福島": "fukushima", "福島県": "fukushima",
    "茨城": "ibaraki", "茨城県": "ibaraki",
    "栃木": "tochigi", "栃木県": "tochigi",
    "群馬": "gunma", "群馬県": "gunma",
    "埼玉": "saitama", "埼玉県": "saitama",
    "千葉": "chiba", "千葉県": "chiba",
    "東京": "tokyo", "東京都": "tokyo",
    "神奈川": "kanagawa", "神奈川県": "kanagawa",
    "新潟": "niigata", "新潟県": "niigata",
    "富山": "toyama", "富山県": "toyama",
    "石川": "ishikawa", "石川県": "ishikawa",
    "福井": "fukui", "福井県": "fukui",
    "山梨": "yamanashi", "山梨県": "yamanashi",
    "長野": "nagano", "長野県": "nagano",
    "岐阜": "gifu", "岐阜県": "gifu",
    "静岡": "shizuoka", "静岡県": "shizuoka",
    "愛知": "aichi", "愛知県": "aichi",
    "三重": "mie", "三重県": "mie",
    "滋賀": "shiga", "滋賀県": "shiga",
    "京都": "kyoto", "京都府": "kyoto",
    "大阪": "osaka", "大阪府": "osaka",
    "兵庫": "hyogo", "兵庫県": "hyogo",
    "奈良": "nara", "奈良県": "nara",
    "和歌山": "wakayama", "和歌山県": "wakayama",
    "鳥取": "tottori", "鳥取県": "tottori",
    "島根": "shimane", "島根県": "shimane",
    "岡山": "okayama", "岡山県": "okayama",
    "広島": "hiroshima", "広島県": "hiroshima",
    "山口": "yamaguchi", "山口県": "yamaguchi",
    "徳島": "tokushima", "徳島県": "tokushima",
    "香川": "kagawa", "香川県": "kagawa",
    "愛媛": "ehime", "愛媛県": "ehime",
    "高知": "kochi", "高知県": "kochi",
    "福岡": "fukuoka", "福岡県": "fukuoka",
    "佐賀": "saga", "佐賀県": "saga",
    "長崎": "nagasaki", "長崎県": "nagasaki",
    "熊本": "kumamoto", "熊本県": "kumamoto",
    "大分": "oita", "大分県": "oita",
    "宮崎": "miyazaki", "宮崎県": "miyazaki",
    "鹿児島": "kagoshima", "鹿児島県": "kagoshima",
    "沖縄": "okinawa", "沖縄県": "okinawa",
}

# 国税庁路線価図の基本URL（年度により変わる）
NTA_BASE_URL = "https://www.rosenka.nta.go.jp"
NTA_YEAR_PATH = "main_r07"  # 令和7年版


def normalize_text(text: str) -> str:
    """テキストを正規化（全角→半角、トリム）"""
    # 全角数字→半角数字、全角英字→半角英字
    text = unicodedata.normalize("NFKC", text)
    return text.strip()


def extract_choumei(address: str) -> str:
    """住所から町名（丁目より前）を抽出"""
    # 「丁目」で分割して前半を取得
    if "丁目" in address:
        return address.split("丁目")[0]
    return address


class RosenkaScraper:
    """国税庁路線価図URLスクレイパー"""

    def __init__(self, timeout: float = 30.0):
        self.timeout = timeout
        self._client: Optional[httpx.AsyncClient] = None

    async def _get_client(self) -> httpx.AsyncClient:
        if self._client is None:
            self._client = httpx.AsyncClient(
                timeout=self.timeout,
                follow_redirects=True,
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
                    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
                    "Referer": "https://www.rosenka.nta.go.jp/",
                    "Connection": "keep-alive",
                }
            )
        return self._client

    async def close(self):
        if self._client:
            await self._client.aclose()
            self._client = None

    async def get_rosenka_urls(
        self,
        prefecture: str,
        city: str,
        town: str,
    ) -> List[str]:
        """
        住所から路線価図URLを取得

        Args:
            prefecture: 都道府県（例: 大阪府, 大阪）
            city: 市区町村（例: 大阪市北区）
            town: 町名（例: 梅田, 梅田3丁目）

        Returns:
            路線価図URLのリスト
        """
        try:
            # 1. 都道府県コードを取得
            pref_code = self._get_pref_code(prefecture)
            if not pref_code:
                logger.warning(f"都道府県コードが見つかりません: {prefecture}")
                return []

            # 2. 都道府県ページから路線価図ページURLを取得
            city_list_url = await self._get_city_list_url(pref_code)
            if not city_list_url:
                logger.warning(f"路線価図ページが見つかりません: {prefecture}")
                return []

            # 3. 市区町村ページURLを取得
            city_url = await self._get_city_url(city_list_url, city)
            if not city_url:
                logger.warning(f"市区町村が見つかりません: {city}")
                return []

            # 4. 町名から路線価図URLを取得
            choumei = extract_choumei(town)
            urls = await self._get_town_urls(city_url, choumei)

            return urls

        except Exception as e:
            logger.exception(f"路線価図URL取得エラー: {e}")
            return []

    def _get_pref_code(self, prefecture: str) -> Optional[str]:
        """都道府県名からコードを取得"""
        # 完全一致
        if prefecture in PREF_CODES:
            return PREF_CODES[prefecture]

        # 部分一致
        for name, code in PREF_CODES.items():
            if name in prefecture or prefecture in name:
                return code

        return None

    async def _get_city_list_url(self, pref_code: str) -> Optional[str]:
        """都道府県の路線価図市区町村一覧ページURLを取得"""
        # 路線価図ページURL形式: https://www.rosenka.nta.go.jp/main_r07/{pref_code}/{pref_code}/
        pref_url = f"{NTA_BASE_URL}/{NTA_YEAR_PATH}/{pref_code}/"

        client = await self._get_client()
        try:
            response = await client.get(pref_url)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, "html.parser")

            # 「路線価図」リンクを探す
            for link in soup.find_all("a"):
                text = link.get_text(strip=True)
                if "路線価図" in text:
                    href = link.get("href")
                    if href:
                        return urljoin(pref_url, href)

            # 直接市区町村一覧ページを試す
            city_list_url = f"{NTA_BASE_URL}/{NTA_YEAR_PATH}/{pref_code}/{pref_code}/"
            response = await client.get(city_list_url)
            if response.status_code == 200:
                return city_list_url

        except Exception as e:
            logger.warning(f"都道府県ページ取得エラー: {e}")

        return None

    async def _get_city_url(self, city_list_url: str, city: str) -> Optional[str]:
        """市区町村一覧から該当市区町村のURLを取得"""
        client = await self._get_client()
        try:
            response = await client.get(city_list_url)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, "html.parser")

            # 市区町村リンクを探す
            city_links: Dict[str, str] = {}

            for link in soup.find_all("a"):
                href = link.get("href")
                if not href:
                    continue

                # title属性またはテキストから市区町村名を取得
                link_text = link.get("title") or link.get_text(strip=True)
                if link_text:
                    link_text = normalize_text(link_text)
                    city_links[link_text] = urljoin(city_list_url, href)

            # 完全一致
            normalized_city = normalize_text(city)
            if normalized_city in city_links:
                return city_links[normalized_city]

            # 部分一致（大阪市北区 → 北区）
            # 政令指定都市の区を探す
            for name, url in city_links.items():
                if name in normalized_city or normalized_city in name:
                    return url

            # 区単位で探す（大阪市北区 → 北区）
            if "区" in city:
                ku = city.split("市")[-1] if "市" in city else city
                for name, url in city_links.items():
                    if name == ku or ku in name:
                        return url

        except Exception as e:
            logger.warning(f"市区町村ページ取得エラー: {e}")

        return None

    async def _get_town_urls(self, city_url: str, choumei: str) -> List[str]:
        """市区町村ページから町名の路線価図URLを取得"""
        client = await self._get_client()
        try:
            response = await client.get(city_url)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, "html.parser")

            # テーブルから町名とURLを抽出
            town_urls: Dict[str, List[str]] = {}

            for table in soup.find_all("table"):
                rows = table.find_all("tr")
                for row in rows:
                    cells = row.find_all(["td", "th"])
                    if len(cells) >= 2:
                        # 1列目が町名、2列目以降がリンク
                        town_name_cell = cells[1] if len(cells) > 1 else cells[0]
                        town_name = normalize_text(town_name_cell.get_text(strip=True))

                        # リンクを収集
                        urls = []
                        for cell in cells:
                            for link in cell.find_all("a"):
                                href = link.get("href")
                                if href and "prices" in href:
                                    urls.append(urljoin(city_url, href))

                        if town_name and urls:
                            town_urls[town_name] = urls

            # 町名で検索
            normalized_choumei = normalize_text(choumei)

            # 完全一致
            if normalized_choumei in town_urls:
                return town_urls[normalized_choumei]

            # 部分一致
            for name, urls in town_urls.items():
                if normalized_choumei in name or name in normalized_choumei:
                    return urls

        except Exception as e:
            logger.warning(f"町名ページ取得エラー: {e}")

        return []


async def get_rosenka_url(
    prefecture: str,
    city: str,
    town: str,
) -> List[str]:
    """
    住所から国税庁路線価図URLを取得（ユーティリティ関数）

    Args:
        prefecture: 都道府県
        city: 市区町村
        town: 町名

    Returns:
        路線価図URLのリスト
    """
    scraper = RosenkaScraper()
    try:
        return await scraper.get_rosenka_urls(prefecture, city, town)
    finally:
        await scraper.close()


# テスト用
if __name__ == "__main__":
    import asyncio

    async def test():
        urls = await get_rosenka_url("大阪府", "大阪市北区", "梅田3丁目")
        print(f"梅田3丁目: {urls}")

        urls = await get_rosenka_url("大阪", "大阪市旭区", "高殿2丁目")
        print(f"高殿2丁目: {urls}")

    asyncio.run(test())
