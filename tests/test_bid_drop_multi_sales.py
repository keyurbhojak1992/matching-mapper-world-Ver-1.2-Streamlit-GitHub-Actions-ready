import io
from pathlib import Path

import pandas as pd

from adapters.bid_drop_adapter import _materialize_sales_workbook


def _excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buffer.getvalue()


def test_bid_drop_sales_multiple_workbooks_are_combined(tmp_path: Path):
    first = pd.DataFrame([
        {"Sold Party": "PARTY A", "Col": "D", "Qua": "VS1", "Sales_Person": "USER A"}
    ])
    second = pd.DataFrame([
        {"Sold Party": "PARTY B", "Col": "E", "Qua": "VVS2", "Sales_Person": "USER B"}
    ])

    output = _materialize_sales_workbook(
        work=tmp_path,
        sales_files=[
            ("sales_a.xlsx", _excel_bytes(first)),
            ("sales_b.xlsx", _excel_bytes(second)),
        ],
    )

    combined = pd.read_excel(output)
    assert combined["Sold Party"].tolist() == ["PARTY A", "PARTY B"]
    assert combined["Sales_Person"].tolist() == ["USER A", "USER B"]


def test_bid_drop_single_sales_workbook_keeps_original_bytes(tmp_path: Path):
    original = _excel_bytes(pd.DataFrame([{"Sold Party": "PARTY A"}]))
    output = _materialize_sales_workbook(work=tmp_path, sales_bytes=original)
    assert output.read_bytes() == original
