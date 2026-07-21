from __future__ import annotations

import io

import pandas as pd

from ui.config_controls import (
    combine_selected_and_custom,
    default_parties,
    default_salespeople,
    discover_suggested_names,
    discover_suggested_names_from_streams,
    parse_delimited_values,
    parse_type_weights,
)


def test_default_scope_lists_include_requested_names():
    assert "Mr. Bhavik Hadiya" in default_salespeople()
    assert "Ms. Zeal Harsora" in default_salespeople()
    assert "NS LUXURY" in default_parties()
    assert "AL BUTA JEWELLERY LLC" in default_parties()


def test_custom_scope_and_weight_parsing():
    assert parse_delimited_values("A, B; C\nD") == ["A", "B", "C", "D"]
    assert combine_selected_and_custom(["A"], "a, B") == ["A", "B"]
    assert parse_type_weights("SALE=20, BID=15\nVIEW=4") == {"SALE": 20, "BID": 15, "VIEW": 4}


def test_name_discovery_from_profiles_and_filters():
    profile = pd.DataFrame(
        [{"Salesperson": "Mr. New User", "SoldParty": "NEW PARTY"}]
    )
    profile_buffer = io.BytesIO()
    profile.to_parquet(profile_buffer, index=False)

    filter_buffer = io.BytesIO()
    with pd.ExcelWriter(filter_buffer, engine="openpyxl") as writer:
        pd.DataFrame(
            [{"SALES_PERSON": "Ms. Filter User", "SOLD PARTY": "FILTER PARTY"}]
        ).to_excel(writer, index=False, sheet_name="Change Format")

    salespeople, parties = discover_suggested_names(
        profile_buffer.getvalue(), None, filter_buffer.getvalue()
    )
    assert "Mr. New User" in salespeople
    assert "Ms. Filter User" in salespeople
    assert "NEW PARTY" in parties
    assert "FILTER PARTY" in parties


def test_stream_name_discovery_rewinds_uploads():
    profile = pd.DataFrame(
        [{"Salesperson": "Mr. Stream User", "SoldParty": "STREAM PARTY"}]
    )
    profile_buffer = io.BytesIO()
    profile.to_parquet(profile_buffer, index=False)
    profile_buffer.seek(7)

    filter_buffer = io.BytesIO()
    with pd.ExcelWriter(filter_buffer, engine="openpyxl") as writer:
        pd.DataFrame(
            [{"SALES_PERSON": "Ms. Stream Filter", "SOLD PARTY": "STREAM FILTER PARTY"}]
        ).to_excel(writer, index=False, sheet_name="Change Format")
    filter_buffer.seek(11)

    salespeople, parties = discover_suggested_names_from_streams(
        profile_buffer, None, filter_buffer
    )
    assert "Mr. Stream User" in salespeople
    assert "Ms. Stream Filter" in salespeople
    assert "STREAM PARTY" in parties
    assert "STREAM FILTER PARTY" in parties
    assert profile_buffer.tell() == 0
    assert filter_buffer.tell() == 0
