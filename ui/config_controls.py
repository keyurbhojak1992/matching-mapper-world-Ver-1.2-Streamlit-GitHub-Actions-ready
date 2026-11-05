from __future__ import annotations

import ast
import io
import re
from functools import lru_cache
from pathlib import Path
from typing import BinaryIO, Iterable

from openpyxl import load_workbook


ROOT_DIR = Path(__file__).resolve().parents[1]
SUGGESTED_MASTER_PATH = ROOT_DIR / "original_master_code" / "Suggested_Process.py"
PARTIES_FILE_PATH = ROOT_DIR / "data" / "parties.txt"

# Kept outside the master script so UI choices can be extended without changing
# the preserved matching engine.
FALLBACK_SALESPEOPLE = [
    "Mr. Bhavik Hadiya",
    "Mr. Dhaval Moradiya",
    "Mr. Parva Shah",
    "Mr. Hiren Rathod",
    "Mr. Jagdish Tarsariya",
    "Mr. Shubham Patil",
    "Mr. Aakash Parmar",
    "Mr. Aatish Sarodiya",
    "Mr. Amit Tarsaria",
    "Mr. Hrithik Hundia",
    "Mr. Kalash Sanghavi",
    "Mr. Mukul Shah",
    "Mr. Nikunj Shankar",
    "Mr. Shubham Tarsariya",
    "Mr. Sunil Bhatiya",
    "Ms. Zeal Harsora",
    "Mr. Ridham Lakhani",
    "Mr. Pranav Chalodiya",
]

FALLBACK_PARTIES = [
    "ABHAY GEMS",
    "NS LUXURY",
    "VIDHYA GEMS",
    "NOOR JEWELS",
    "AL BUTA JEWELLERY LLC",
]

DEFAULT_TYPE_SEQUENCE = [
    "Type_SALE",
    "Type_BID",
    "Type_SEARCH",
    "Type_WISHLIST",
    "Type_DETAILS",
    "Type_VIEW",
    "Type_CART",
    "Type_OFFER",
]

DEFAULT_PRIORITY_TYPES = ["PLATINUM", "GOLD", "DIAMOND", "PREMIUM", "VIP"]


def unique_clean(values: Iterable[object]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = str(value).strip() if value is not None else ""
        if not text or text.lower() == "nan":
            continue
        key = text.casefold()
        if key not in seen:
            seen.add(key)
            result.append(text)
    return result


def parse_delimited_values(text: str | None) -> list[str]:
    if not text:
        return []
    return unique_clean(re.split(r"[,;\n]+", text))


def parse_type_weights(text: str | None) -> dict[str, int]:
    result: dict[str, int] = {}
    for item in parse_delimited_values(text):
        if "=" not in item:
            continue
        name, raw_value = item.split("=", 1)
        name = name.strip().upper()
        if not name:
            continue
        try:
            result[name] = int(float(raw_value.strip()))
        except ValueError:
            continue
    return result


def format_type_weights(weights: dict[str, object]) -> str:
    return ", ".join(f"{key}={value}" for key, value in weights.items())


def _extract_commented_or_active_lists(variable_prefix: str) -> list[str]:
    """Read single-line list assignments from the preserved source, including comments."""
    if not SUGGESTED_MASTER_PATH.exists():
        return []
    values: list[str] = []
    for line in SUGGESTED_MASTER_PATH.read_text(encoding="utf-8").splitlines():
        candidate = line.strip()
        if candidate.startswith("#"):
            candidate = candidate[1:].strip()
        match = re.match(rf"({re.escape(variable_prefix)}\w*)\s*=\s*(\[.*\])\s*$", candidate)
        if not match:
            continue
        try:
            parsed = ast.literal_eval(match.group(2))
        except (SyntaxError, ValueError):
            continue
        if isinstance(parsed, list):
            values.extend(parsed)
    return unique_clean(values)


@lru_cache(maxsize=1)
def _load_parties_file() -> list[str]:
    """Load repository-managed party names without changing existing callers."""
    if not PARTIES_FILE_PATH.exists():
        return []
    try:
        return unique_clean(PARTIES_FILE_PATH.read_text(encoding="utf-8-sig").splitlines())
    except OSError:
        # Preserve the existing fallback behavior if the deployed file cannot be read.
        return []


def default_salespeople() -> list[str]:
    return unique_clean(FALLBACK_SALESPEOPLE + _extract_commented_or_active_lists("TEST_SALESPEOPLE"))


def default_parties() -> list[str]:
    return unique_clean(
        FALLBACK_PARTIES
        + _load_parties_file()
        + _extract_commented_or_active_lists("TEST_PARTIES")
    )



def _normalized_column(value: object) -> str:
    return "".join(character for character in str(value or "").casefold() if character.isalnum())


def _rewind(source: BinaryIO) -> None:
    try:
        source.seek(0)
    except (AttributeError, OSError):
        pass


def _parquet_values_stream(source: BinaryIO, candidates: set[str]) -> list[str]:
    """Read only matching Parquet columns from a seekable upload stream."""
    import pyarrow.parquet as pq
    _rewind(source)
    try:
        parquet = pq.ParquetFile(source)
        normalized_candidates = {_normalized_column(value) for value in candidates}
        matching = [
            name
            for name in parquet.schema_arrow.names
            if _normalized_column(name) in normalized_candidates
        ]
        if not matching:
            return []
        table = parquet.read(columns=matching)
        values: list[object] = []
        for name in matching:
            values.extend(table.column(name).to_pylist())
        return unique_clean(values)
    except Exception:
        return []
    finally:
        _rewind(source)


def _excel_values_stream(
    source: BinaryIO,
    candidates: set[str],
    *,
    title_case: bool = False,
) -> list[str]:
    """Read selected Excel columns in read-only mode without copying the upload bytes."""
    _rewind(source)
    try:
        workbook = load_workbook(source, read_only=True, data_only=True)
    except Exception:
        _rewind(source)
        return []
    try:
        normalized_candidates = {_normalized_column(value) for value in candidates}
        values: list[str] = []
        for worksheet in workbook.worksheets:
            header = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
            indexes = [
                index
                for index, value in enumerate(header)
                if _normalized_column(value) in normalized_candidates
            ]
            if not indexes:
                continue
            for row in worksheet.iter_rows(min_row=2, values_only=True):
                for index in indexes:
                    if index >= len(row) or row[index] in (None, ""):
                        continue
                    value = str(row[index]).strip()
                    values.append(value.title() if title_case else value)
        return unique_clean(values)
    finally:
        workbook.close()
        _rewind(source)


def discover_suggested_names_from_streams(
    profile_perm_file: BinaryIO | None,
    profile_temp_file: BinaryIO | None,
    filter_file: BinaryIO | None,
) -> tuple[list[str], list[str]]:
    """Discover current names from upload streams while reading only required columns."""
    salespeople: list[str] = []
    parties: list[str] = []

    for profile_file in (profile_perm_file, profile_temp_file):
        if profile_file is None:
            continue
        salespeople.extend(
            _parquet_values_stream(profile_file, {"Salesperson", "Sales Person", "Sales_Person"})
        )
        parties.extend(
            _parquet_values_stream(
                profile_file,
                {
                    "SoldParty",
                    "Sold Party",
                    "Sold Party Original",
                    "Normalized Party",
                    "Party Name",
                },
            )
        )

    if filter_file is not None:
        salespeople.extend(
            _excel_values_stream(
                filter_file,
                {"SALES_PERSON", "Sales Person", "Sales_Person"},
                title_case=True,
            )
        )
        parties.extend(
            _excel_values_stream(filter_file, {"SOLD PARTY", "Sold Party", "Party Name"})
        )
    else:
        from defaults.suggested_rules import DEFAULT_FILTER_RULES

        for sheet_name in ("Change Format", "Tolerance"):
            payload = DEFAULT_FILTER_RULES.get(sheet_name, {})
            columns = [str(value) for value in payload.get("columns", [])]
            normalized = [_normalized_column(value) for value in columns]
            sales_indexes = [i for i, value in enumerate(normalized) if value == "salesperson"]
            party_indexes = [i for i, value in enumerate(normalized) if value in {"soldparty", "partyname"}]
            for row in payload.get("rows", []):
                for index in sales_indexes:
                    if index < len(row) and row[index] not in (None, ""):
                        salespeople.append(str(row[index]).strip().title())
                for index in party_indexes:
                    if index < len(row) and row[index] not in (None, ""):
                        parties.append(str(row[index]).strip())

    return unique_clean(salespeople), unique_clean(parties)


def discover_suggested_names(
    profile_perm_bytes: bytes | None,
    profile_temp_bytes: bytes | None,
    filter_bytes: bytes | None,
) -> tuple[list[str], list[str]]:
    """Backwards-compatible bytes wrapper used by local tests and callers."""
    return discover_suggested_names_from_streams(
        io.BytesIO(profile_perm_bytes) if profile_perm_bytes else None,
        io.BytesIO(profile_temp_bytes) if profile_temp_bytes else None,
        io.BytesIO(filter_bytes) if filter_bytes else None,
    )


def combine_selected_and_custom(selected: Iterable[str], custom_text: str | None) -> list[str]:
    return unique_clean(list(selected) + parse_delimited_values(custom_text))
