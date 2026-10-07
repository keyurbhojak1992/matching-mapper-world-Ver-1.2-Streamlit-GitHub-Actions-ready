from __future__ import annotations

import io
from datetime import datetime

import pandas as pd

from adapters.pricing_last5_bid_adapter import (
    DEFAULT_LAST_N_BIDS,
    DEFAULT_OUTPUT_MODE,
    PRICING_LAST5_DEFAULTS,
    build_confirm_last_n_bid_from_summary,
    build_master_db_from_confirm,
    build_party_profile_temp,
)


def _bid_summary_bytes() -> bytes:
    rows = []
    # Intentionally unsorted to prove selection is based on numeric Bid value.
    for bid_no, carat in [(276, 1.05), (270, 0.35), (275, 1.22), (272, 2.60), (274, 8.50), (271, 0.55), (273, 15.2)]:
        rows.append({
            "Bid": f" Bid-{bid_no} ",
            "Bid no": bid_no,
            "Serial No": 100000 + bid_no,
            "Shape": " RBC ",
            "Carat": carat,
            "Color": " G ",
            "Clarity": " VS1 ",
            "Fluo": " NON ",
            "Lab": " GIA ",
        })
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, index=False, sheet_name="Sheet1")
    return buffer.getvalue()


def test_requested_pricing_defaults_are_preserved():
    assert DEFAULT_LAST_N_BIDS == 5
    assert DEFAULT_OUTPUT_MODE == "DIRECT_PARTY_ZIP"
    assert PRICING_LAST5_DEFAULTS["TOLERANCE_LEVEL"] == 1
    assert PRICING_LAST5_DEFAULTS["ENABLE_FLO_MAPPING"] is False
    assert PRICING_LAST5_DEFAULTS["ENABLE_COLOR_TOLERANCE"] is True
    assert PRICING_LAST5_DEFAULTS["ENABLE_POLISH_TOLERANCE"] is False
    assert PRICING_LAST5_DEFAULTS["ENABLE_SYMMETRY_TOLERANCE"] is False
    assert PRICING_LAST5_DEFAULTS["SHOW_REF_STONE_ID"] is True
    assert PRICING_LAST5_DEFAULTS["SHOW_BASIC_SUPERTAG"] is False
    assert PRICING_LAST5_DEFAULTS["TEST_MODE"] is False
    assert PRICING_LAST5_DEFAULTS["TEST_SALESPEOPLE"] == [
        "Mr. Vivek Chalodiya", "Mr. Hrithik Hundia", "Ms. Sonali Sonavane"
    ]
    assert PRICING_LAST5_DEFAULTS["MATCH_STATUS_FILTER_MODE"] == "ALL"
    assert PRICING_LAST5_DEFAULTS["MATCH_STATUS_ALLOWED_TYPES"] == ["Exact", "Near by"]


def test_latest_five_bids_are_selected_numerically_and_cleaned():
    confirm, summary = build_confirm_last_n_bid_from_summary(
        _bid_summary_bytes(),
        last_n=5,
        run_datetime=datetime(2026, 10, 7, 11, 47, 0),
    )
    assert summary["selected_bids"] == ["Bid-272", "Bid-273", "Bid-274", "Bid-275", "Bid-276"]
    assert set(confirm["Party_type"].dropna()) == {"Bid-272", "Bid-273", "Bid-274", "Bid-275", "Bid-276"}
    assert len(confirm) == 5
    assert set(confirm["Shp"]) == {"RBC"}
    assert set(confirm["Sold Party"]) == {"Pricing"}
    assert set(confirm["Country"]) == {"INDIA"}
    assert set(confirm["City"]) == {"MUMBAI"}
    assert set(confirm["SALE_PERSON"]) == {"Pricing"}
    assert set(confirm["Type"]) == {"BID"}
    assert set(confirm["Business Type"]) == {"Pricing"}
    assert set(confirm["MTD"]) == {"Oct--26"}
    assert set(confirm["Date"].dt.strftime("%Y-%m-%d")) == {"2026-10-07"}

    carat_groups = dict(zip(confirm["Carat"], confirm["Carat_grp"]))
    assert carat_groups[1.05] == "1.00-1.09"
    assert carat_groups[1.22] == "1.20-1.24"
    assert carat_groups[2.60] == "2.50-2.99"
    assert carat_groups[8.50] == "8.00-14.99"
    assert carat_groups[15.2] == "15+"


def test_generated_temp_profile_has_expected_schema_and_any_advanced_fields():
    confirm, _ = build_confirm_last_n_bid_from_summary(
        _bid_summary_bytes(),
        last_n=5,
        run_datetime=datetime(2026, 10, 7, 11, 47, 0),
    )
    master = build_master_db_from_confirm(confirm)
    profile = build_party_profile_temp(master)

    assert master["Super Tag"].str.contains("ANY", regex=False).all()
    assert set(profile.columns) == {
        "Salesperson", "Sold Party Original", "Super Tag", "Count", "BidSaleCount",
        "Types_Found", "Stone_Id", "Match_Logic_Type", "SoldParty", "Country",
        "City", "MTD_Label",
    }
    assert set(profile["Salesperson"]) == {"Pricing"}
    assert set(profile["SoldParty"]) == {"Pricing"}
    assert set(profile["Types_Found"]) == {"BID"}
    assert set(profile["Match_Logic_Type"]) == {"STANDARD"}


def test_app_contains_pricing_last5_workspace_and_saved_bid_default():
    from pathlib import Path
    root = Path(__file__).resolve().parents[1]
    text = (root / "app.py").read_text(encoding="utf-8")
    assert '"Pricing Last 5 BID Suggested"' in text
    assert '_repository_default_exists("bid_summary")' in text
    assert (root / "data" / "Bid_Summary.xlsx").is_file()
