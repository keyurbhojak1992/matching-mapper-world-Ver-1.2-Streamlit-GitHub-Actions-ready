from __future__ import annotations

import io
import zipfile

import pandas as pd
from openpyxl import load_workbook

from adapters.usa_suggested_stock_mapper_adapter import run_usa_suggested_stock_mapper


def _xlsx_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()


def test_usa_suggested_mapper_exact_and_nearby_output():
    stock = pd.DataFrame(
        [
            {
                "Serial No": "S-EXACT",
                "Shp": "RBC",
                "Carat_grp": "1.00-1.49",
                "Color": "D",
                "Qua": "VVS1",
                "Flo": "NON",
                "Lab": "GIA",
                "Carat": 1.10,
                "Reserve%": -20.0,
            },
            {
                "Serial No": "S-NEAR",
                "Shp": "RBC",
                "Carat_grp": "1.00-1.49",
                "Color": "E",
                "Qua": "VVS1",
                "Flo": "NON",
                "Lab": "GIA",
                "Carat": 1.05,
                "Reserve%": -18.0,
            },
            {
                "Serial No": "S-NO-MATCH",
                "Shp": "PEAR",
                "Carat_grp": "2.00-2.49",
                "Color": "H",
                "Qua": "SI1",
                "Flo": "MED",
                "Lab": "IGI",
                "Carat": 2.10,
                "Reserve%": -25.0,
            },
        ]
    )
    sales = pd.DataFrame(
        [
            {
                "SHP": "RBC",
                "CARAT_GRP": "1.00-1.49",
                "COL": "D",
                "QUA": "VVS1",
                "FLO": "NON",
                "LAB": "GIA",
                "SOLD PARTY": "Acme Diamonds LLC",
                "SALE_PERSON": "USA SELLER",
            }
        ]
    )

    data, filename, logs = run_usa_suggested_stock_mapper(
        stock_bytes=_xlsx_bytes(stock, "USA Stock"),
        sales_history_files=[("usa_sales.xlsx", _xlsx_bytes(sales, "Sales"))],
        top_n_supertags=251,
    )

    assert filename.startswith("USA_Suggested_Stock_") and filename.endswith(".zip")
    assert "mapped to temporary 'Sheet1'" in logs
    assert "Loaded 3 stock items" in logs
    assert "Loaded 1 sales records" in logs

    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        names = zf.namelist()
        assert names == ["Usa Seller/Acme Diamonds LLC.xlsx"]
        report_bytes = zf.read(names[0])

    wb = load_workbook(io.BytesIO(report_bytes))
    ws = wb["Suggested_Stones"]
    headers = [cell.value for cell in ws[1]]
    assert "Is_Nearby" not in headers
    serial_col = headers.index("Serial No") + 1
    serials = [ws.cell(row=r, column=serial_col).value for r in range(2, ws.max_row + 1)]
    assert serials == ["S-EXACT", "S-NEAR"]
    # Exact match row is green in the supplied USA mapper.
    assert ws.cell(row=2, column=1).fill.fgColor.rgb in {"00CCFFCC", "FFCCFFCC"}


def test_usa_suggested_mapper_requires_sales_history():
    stock = pd.DataFrame(
        [{"Shp": "RBC", "Carat_grp": "1.00-1.49", "Color": "D", "Qua": "VVS1", "Flo": "NON", "Lab": "GIA"}]
    )
    try:
        run_usa_suggested_stock_mapper(
            stock_bytes=_xlsx_bytes(stock, "Sheet1"),
            sales_history_files=[],
        )
    except ValueError as exc:
        assert "sales history" in str(exc).lower()
    else:
        raise AssertionError("Expected a missing-sales-history ValueError")
