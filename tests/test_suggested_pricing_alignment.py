from __future__ import annotations

import sys

from adapters.common import MASTER_CODE_DIR, load_fresh_module
from defaults.suggested_rules import rule_counts


def test_built_in_rule_counts_match_former_excel_defaults():
    assert rule_counts() == {
        "removed_parties": 1562,
        "change_format": 142,
        "tolerance": 1,
        "no_need": 260,
    }


def test_pricing_engine_uses_new_buckets_without_breaking_regular_profile_buckets(tmp_path):
    regular = load_fresh_module(MASTER_CODE_DIR / "Suggested_Process.py", "regular_alignment")
    pricing = load_fresh_module(MASTER_CODE_DIR / "Pricing_Suggested_Process.py", "pricing_alignment")
    try:
        # Existing permanent Suggested profiles still use the broad historic buckets.
        assert regular.get_carat_range_string(1.05, country="USA") == "1.00-1.19"
        assert regular.get_carat_range_string(0.55, country="USA") == "0.50-0.69"

        # Pricing Last-5-BID profiles are built with the narrower buckets from the
        # attached single-pipeline source, so they must use the matching engine below.
        assert pricing.ENABLE_FLO_MAPPING is True
        assert pricing.get_carat_range_string(0.55, country="USA") == "0.50-0.59"
        assert pricing.get_carat_range_string(0.65, country="USA") == "0.60-0.69"
        assert pricing.get_carat_range_string(1.05, country="USA") == "1.00-1.09"
        assert pricing.get_carat_range_string(1.15, country="USA") == "1.10-1.19"
        assert pricing.get_carat_range_string(1.22, country="USA") == "1.20-1.24"
        assert pricing.get_carat_range_string(1.30, country="USA") == "1.25-1.49"

        removed, change_format, tolerance = pricing.load_filter_data(tmp_path / "missing_filters.xlsx")
        no_need = pricing.load_exclusion_data(tmp_path / "missing_no_need.xlsx")
        assert len(removed) == 1562
        assert len(change_format) == 142
        assert len(tolerance) == 1
        assert len(no_need) == 260
    finally:
        sys.modules.pop(regular.__name__, None)
        sys.modules.pop(pricing.__name__, None)
