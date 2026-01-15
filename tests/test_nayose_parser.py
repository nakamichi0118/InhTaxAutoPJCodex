"""Tests for nayose (property tax roll) parser."""
import pytest
from backend.app.parsers.nayose_parser import (
    parse_nayose_response,
    _normalize_land_type,
    _normalize_building_type,
    _parse_number,
)


class TestParseNayoseResponse:
    """Test cases for parse_nayose_response function."""

    def test_parse_land_and_building(self):
        """Test parsing a response with both land and building."""
        response = '''
        {
          "municipality": "大阪市都島区",
          "properties": [
            {
              "property_type": "land",
              "location": "都島区中野町1丁目",
              "lot_number": "123-4",
              "land_category": "宅地",
              "area": 150.25,
              "valuation_amount": 25000000,
              "notes": null
            },
            {
              "property_type": "building",
              "location": "都島区中野町1丁目",
              "lot_number": "123",
              "structure": "木造2階建居宅",
              "area": 98.5,
              "valuation_amount": 8000000,
              "built_year": "平成10年",
              "floors": "2階建",
              "notes": null
            }
          ]
        }
        '''
        assets = parse_nayose_response(response, "test.pdf")

        assert len(assets) == 2

        # Land asset
        land = assets[0]
        assert land.category == "land"
        assert land.type == "residential"
        assert land.identifier_primary == "123-4"
        assert land.valuation_amount == 25000000.0
        assert land.location_municipality == "大阪市都島区"
        assert "宅地" in land.notes or "課税地目" in land.notes

        # Building asset
        building = assets[1]
        assert building.category == "building"
        assert building.type == "residence"
        assert building.identifier_primary == "123"
        assert building.valuation_amount == 8000000.0
        assert "木造" in building.notes or "構造" in building.notes

    def test_parse_empty_properties(self):
        """Test handling of empty properties array."""
        response = '{"municipality": "test", "properties": []}'
        assets = parse_nayose_response(response, "test.pdf")
        assert len(assets) == 0

    def test_parse_no_json(self):
        """Test handling of response without JSON."""
        response = "This is not JSON"
        assets = parse_nayose_response(response, "test.pdf")
        assert len(assets) == 0

    def test_parse_invalid_json(self):
        """Test handling of invalid JSON."""
        response = '{"municipality": "test", invalid}'
        assets = parse_nayose_response(response, "test.pdf")
        assert len(assets) == 0

    def test_multiple_lands(self):
        """Test parsing multiple land records."""
        response = '''
        {
          "municipality": "熱海市",
          "properties": [
            {"property_type": "land", "lot_number": "1-1", "land_category": "宅地", "area": 100, "valuation_amount": 10000000},
            {"property_type": "land", "lot_number": "1-2", "land_category": "田", "area": 500, "valuation_amount": 5000000},
            {"property_type": "land", "lot_number": "1-3", "land_category": "山林", "area": 2000, "valuation_amount": 2000000}
          ]
        }
        '''
        assets = parse_nayose_response(response, "test.pdf")

        assert len(assets) == 3
        assert assets[0].type == "residential"
        assert assets[1].type == "agricultural"
        assert assets[2].type == "forest"


class TestNormalizeLandType:
    """Test cases for land type normalization."""

    def test_residential(self):
        assert _normalize_land_type("宅地") == "residential"

    def test_agricultural(self):
        assert _normalize_land_type("田") == "agricultural"
        assert _normalize_land_type("畑") == "agricultural"

    def test_forest(self):
        assert _normalize_land_type("山林") == "forest"

    def test_miscellaneous(self):
        assert _normalize_land_type("雑種地") == "miscellaneous"

    def test_unknown(self):
        assert _normalize_land_type("特殊用地") == "other"
        assert _normalize_land_type(None) == "other"


class TestNormalizeBuildingType:
    """Test cases for building type normalization."""

    def test_residence(self):
        assert _normalize_building_type("居宅") == "residence"
        assert _normalize_building_type("木造2階建居宅") == "residence"

    def test_store(self):
        assert _normalize_building_type("店舗") == "store"

    def test_warehouse(self):
        assert _normalize_building_type("倉庫") == "warehouse"

    def test_office(self):
        assert _normalize_building_type("事務所") == "office"

    def test_unknown(self):
        assert _normalize_building_type("その他") == "other"
        assert _normalize_building_type(None) == "other"


class TestParseNumber:
    """Test cases for number parsing."""

    def test_integer(self):
        assert _parse_number(100) == 100.0

    def test_float(self):
        assert _parse_number(150.25) == 150.25

    def test_string_with_commas(self):
        assert _parse_number("25,000,000") == 25000000.0

    def test_string_with_unit(self):
        assert _parse_number("100.5㎡") == 100.5

    def test_string_with_yen(self):
        assert _parse_number("1000円") == 1000.0

    def test_none(self):
        assert _parse_number(None) is None

    def test_invalid_string(self):
        assert _parse_number("abc") is None
