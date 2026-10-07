from __future__ import annotations

import gc
import io
import re
import zipfile
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd

from .common import repository_data_path, safe_filename


PRICING_LAST5_DEFAULTS: dict[str, object] = {
    "TOLERANCE_LEVEL": 1,
    "ENABLE_FLO_MAPPING": True,
    "INCLUDE_COLOR": True,
    "INCLUDE_CLARITY": True,
    "INCLUDE_FLORO": True,
    "INCLUDE_CUT": False,
    "INCLUDE_POLISH": False,
    "INCLUDE_SYMMETRY": False,
    "INCLUDE_SHADE": False,
    "INCLUDE_MILKY": False,
    "INCLUDE_TABLE_BLACK": False,
    "INCLUDE_SIDE_BLACK": False,
    "INCLUDE_TABLE_OPEN": False,
    "INCLUDE_SIDE_OPEN": False,
    "INCLUDE_GIRDLE_OPEN": False,
    "INCLUDE_PAVILION_OPEN": False,
    "ENABLE_COLOR_TOLERANCE": True,
    "ENABLE_CLARITY_TOLERANCE": True,
    "ENABLE_FLORO_TOLERANCE": True,
    "ENABLE_CUT_TOLERANCE": False,
    "ENABLE_POLISH_TOLERANCE": False,
    "ENABLE_SYMMETRY_TOLERANCE": False,
    "ENABLE_SHADE_TOLERANCE": True,
    "ENABLE_MILKY_TOLERANCE": True,
    "ENABLE_TABLE_BLACK_TOLERANCE": True,
    "ENABLE_SIDE_BLACK_TOLERANCE": True,
    "ENABLE_TABLE_OPEN_TOLERANCE": True,
    "ENABLE_SIDE_OPEN_TOLERANCE": True,
    "ENABLE_GIRDLE_OPEN_TOLERANCE": True,
    "ENABLE_PAVILION_OPEN_TOLERANCE": True,
    "SHOW_REF_STONE_ID": True,
    "SHOW_BASIC_SUPERTAG": False,
    "TEST_MODE": False,
    "TEST_SALESPEOPLE": ["Mr. Vivek Chalodiya", "Mr. Hrithik Hundia", "Ms. Sonali Sonavane"],
    "TEST_MODE_PARTY": False,
    "TEST_PARTIES": ["SHWET GEMS LTD"],
    "MATCH_STATUS_FILTER_MODE": "ALL",
    "MATCH_STATUS_ALLOWED_TYPES": ["Exact", "Near by"],
    "MATCH_STATUS_SPECIFIC_PARTIES": ["NONE"],
}

DEFAULT_LAST_N_BIDS = 5
DEFAULT_OUTPUT_MODE = "DIRECT_PARTY_ZIP"
OUTPUT_MODES = ("REGULAR_NAV_ZIP", "DIRECT_PARTY_ZIP", "BOTH")

# These values reproduce the existing Confirm Stone Last 5 BID structure.
FIXED_SOLD_PARTY = "Pricing"
FIXED_COUNTRY = "INDIA"
FIXED_CITY = "MUMBAI"
FIXED_SALE_PERSON = "Pricing"
FIXED_TYPE = "BID"
FIXED_BUSINESS_TYPE = "Pricing"


def _clean_text_value(value):
    if pd.isna(value):
        return np.nan
    text = str(value).replace("\u00a0", " ").replace("\u200b", "")
    text = re.sub(r"\s+", " ", text).strip()
    return text if text else np.nan


def _clean_object_columns(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result.columns = [
        re.sub(r"\s+", " ", str(column).replace("\u00a0", " ")).strip()
        for column in result.columns
    ]
    for column in result.select_dtypes(include=["object", "string"]).columns:
        result[column] = result[column].map(_clean_text_value)
    return result


def pricing_carat_group(value) -> str:
    try:
        if pd.isna(value):
            return ""
        carat = float(value)
    except Exception:
        return ""

    if carat >= 15.00:
        return "15+"
    if 8.00 <= carat <= 14.99:
        return "8.00-14.99"
    if 5.00 <= carat <= 7.99:
        return "5.00-7.99"
    if 4.50 <= carat <= 4.99:
        return "4.50-4.99"
    if 4.00 <= carat <= 4.49:
        return "4.00-4.49"
    if 3.50 <= carat <= 3.99:
        return "3.50-3.99"
    if 3.00 <= carat <= 3.49:
        return "3.00-3.49"
    if 2.50 <= carat <= 2.99:
        return "2.50-2.99"
    if 2.00 <= carat <= 2.49:
        return "2.00-2.49"
    if 1.70 <= carat <= 1.99:
        return "1.70-1.99"
    if 1.50 <= carat <= 1.69:
        return "1.50-1.69"
    if 1.25 <= carat <= 1.49:
        return "1.25-1.49"
    if 1.20 <= carat <= 1.24:
        return "1.20-1.24"
    if 1.10 <= carat <= 1.19:
        return "1.10-1.19"
    if 1.00 <= carat <= 1.09:
        return "1.00-1.09"
    if 0.90 <= carat <= 0.99:
        return "0.90-0.99"
    if 0.80 <= carat <= 0.89:
        return "0.80-0.89"
    if 0.70 <= carat <= 0.79:
        return "0.70-0.79"
    if 0.60 <= carat <= 0.69:
        return "0.60-0.69"
    if 0.50 <= carat <= 0.59:
        return "0.50-0.59"
    if 0.40 <= carat <= 0.49:
        return "0.40-0.49"
    if 0.30 <= carat <= 0.39:
        return "0.30-0.39"
    return ""


def _pricing_mtd_label(run_datetime: datetime) -> str:
    month = run_datetime.strftime("%b")
    if run_datetime.month == 9:
        month = "Sept"
    return f"{month}--{run_datetime.strftime('%y')}"


def _current_pricing_datetime() -> datetime:
    return datetime.now(ZoneInfo("Asia/Kolkata")).replace(tzinfo=None, microsecond=0)


_BID_SUMMARY_COLUMNS = {
    "Bid", "Bid no", "Serial No", "Shape", "Carat", "Color", "Clarity", "Fluo", "Lab"
}


def _read_bid_summary(payload: bytes) -> pd.DataFrame:
    """Read only the columns used by the Last-BID profile builder.

    Bid Summary can contain many extra pricing/action columns. Loading those columns
    into pandas used a large amount of memory on Streamlit Cloud even though this
    workspace never reads them. The callable usecols keeps the selected-row logic
    identical while materially reducing peak memory.
    """
    kwargs = {"usecols": lambda column: str(column).strip() in _BID_SUMMARY_COLUMNS}
    try:
        return pd.read_excel(io.BytesIO(payload), sheet_name="Sheet1", **kwargs)
    except ValueError:
        # Some exported workbooks use a different worksheet name. Keep the file
        # usable by falling back to the first sheet while leaving other logic unchanged.
        return pd.read_excel(io.BytesIO(payload), sheet_name=0, **kwargs)


def build_confirm_last_n_bid_from_summary(
    bid_summary_bytes: bytes,
    *,
    last_n: int = DEFAULT_LAST_N_BIDS,
    run_datetime: datetime | None = None,
) -> tuple[pd.DataFrame, dict[str, object]]:
    if int(last_n) < 1:
        raise ValueError("Last N bids must be at least 1.")

    print("[1/4] Reading Bid Summary")
    # First load only the columns this feature uses. Clean text only after reducing
    # the workbook to the latest N bids so the same output is produced with much
    # lower temporary memory usage.
    dataframe = _read_bid_summary(bid_summary_bytes)
    dataframe.columns = [
        re.sub(r"\s+", " ", str(column).replace("\u00a0", " ")).strip()
        for column in dataframe.columns
    ]
    required = ["Bid", "Serial No", "Shape", "Carat", "Color", "Clarity", "Fluo", "Lab"]
    missing = [column for column in required if column not in dataframe.columns]
    if missing:
        raise ValueError(f"Bid_Summary is missing required columns: {', '.join(missing)}")

    if "Bid no" in dataframe.columns:
        bid_numbers = pd.to_numeric(dataframe["Bid no"], errors="coerce")
    else:
        bid_numbers = pd.Series(np.nan, index=dataframe.index, dtype="float64")

    parsed = pd.to_numeric(
        dataframe["Bid"].astype(str).str.extract(r"(\d+)", expand=False),
        errors="coerce",
    )
    dataframe["_BID_NO_NUM"] = bid_numbers.fillna(parsed)
    distinct_numbers = sorted(dataframe["_BID_NO_NUM"].dropna().unique().tolist())
    if not distinct_numbers:
        raise ValueError("No numeric Bid numbers were found in Bid_Summary.")

    selected_numbers = distinct_numbers[-int(last_n):]
    selected = dataframe[dataframe["_BID_NO_NUM"].isin(selected_numbers)].copy()
    selected.drop(columns=["_BID_NO_NUM"], inplace=True)
    selected.reset_index(drop=True, inplace=True)
    selected = _clean_object_columns(selected)

    run_dt = (run_datetime or _current_pricing_datetime()).replace(microsecond=0)
    output = pd.DataFrame(index=selected.index)
    output["Serial No"] = pd.to_numeric(selected["Serial No"], errors="coerce").astype("Int64")
    output["Shp"] = selected["Shape"].map(_clean_text_value).astype("string").str.upper()
    output["Carat"] = pd.to_numeric(selected["Carat"], errors="coerce")
    output["Carat_grp"] = output["Carat"].map(pricing_carat_group)
    output["Col"] = selected["Color"].map(_clean_text_value).astype("string").str.upper()
    output["Qua"] = selected["Clarity"].map(_clean_text_value).astype("string").str.upper()
    output["Flo"] = selected["Fluo"].map(_clean_text_value).astype("string").str.upper()
    output["Lab"] = selected["Lab"].map(_clean_text_value).astype("string").str.upper()

    for column in [
        "Cut", "Pol", "Sym", "Shade", "Milky", "TBINC_NAME", "TOIN_NAME",
        "SBINC_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "Ratio",
        "Disc%", "RAP", "RATE", "AMT",
    ]:
        output[column] = np.nan

    output["Sold Party"] = FIXED_SOLD_PARTY
    output["Country"] = FIXED_COUNTRY
    output["City"] = FIXED_CITY
    output["SALE_PERSON"] = FIXED_SALE_PERSON
    output["Type"] = FIXED_TYPE
    output["MTD"] = _pricing_mtd_label(run_dt)
    output["Date"] = pd.Timestamp(run_dt)
    output["Party_type"] = selected["Bid"].map(_clean_text_value)
    output["Business Type"] = FIXED_BUSINESS_TYPE

    output = output[[
        "Serial No", "Shp", "Carat", "Carat_grp", "Col", "Qua", "Flo", "Lab",
        "Cut", "Pol", "Sym", "Shade", "Milky", "TBINC_NAME", "TOIN_NAME",
        "SBINC_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "Ratio",
        "Disc%", "RAP", "RATE", "AMT", "Sold Party", "Country", "City",
        "SALE_PERSON", "Type", "MTD", "Date", "Party_type", "Business Type",
    ]]

    bid_rows: list[dict[str, object]] = []
    selected_names: list[str] = []
    for bid_number in selected_numbers:
        names = (
            dataframe.loc[dataframe["_BID_NO_NUM"] == bid_number, "Bid"]
            .dropna()
            .map(_clean_text_value)
        )
        clean_names = [value for value in names.tolist() if value]
        name = str(clean_names[0]) if clean_names else f"Bid-{int(bid_number)}"
        rows = int((dataframe["_BID_NO_NUM"] == bid_number).sum())
        selected_names.append(name)
        bid_rows.append({"bid": name, "rows": rows, "number": int(bid_number)})

    print("=" * 68)
    print(f"LATEST {len(selected_names)} BIDS SELECTED FOR THIS RUN")
    print("=" * 68)
    print(f"  Selected bids    : {', '.join(selected_names)}")
    print("-" * 68)
    for index, item in enumerate(bid_rows, start=1):
        print(f"  {index}. {item['bid']:<20}  Rows: {int(item['rows']):,}")
    print("-" * 68)
    print(f"  Total rows used : {len(output):,}")
    print(f"  Run date        : {run_dt.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 68)
    if len(selected_names) < int(last_n):
        print(
            f"WARNING: Requested latest {last_n} bids, but only "
            f"{len(selected_names)} distinct bid numbers were available."
        )

    summary = {
        "selected_bids": selected_names,
        "bid_rows": bid_rows,
        "total_rows": int(len(output)),
        "run_datetime": run_dt.isoformat(sep=" "),
        "requested_last_n": int(last_n),
    }
    return output, summary


def _normalize_party_name(name: object) -> str:
    if not isinstance(name, str):
        return ""
    text = re.sub(r"[^\w\s]", "", name).upper().strip()
    return re.sub(r"\s+", " ", text).strip()


def build_master_db_from_confirm(confirm_df: pd.DataFrame) -> pd.DataFrame:
    print("[2/4] Building MASTER_DB dataframe")
    master = confirm_df.rename(columns={
        "Serial No": "Stone_Id",
        "SALE_PERSON": "Salesperson",
        "Ratio": "RATIO",
        "Disc%": "DISC%",
        "Date": "DATE",
        "Party_type": "PARTY_TYPE",
        "Business Type": "BUSINESS TYPE",
    }).copy()

    tag_columns = [
        "Shp", "Carat_grp", "Col", "Qua", "Flo", "Lab", "Cut", "Pol", "Sym",
        "Shade", "Milky", "TBINC_NAME", "SBINC_NAME", "TOIN_NAME", "SOIN_NAME",
        "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME",
    ]
    for column in tag_columns:
        if column not in master.columns:
            master[column] = "ANY"
        values = master[column].map(_clean_text_value).astype("string").str.upper()
        master[column] = values.fillna("ANY").replace("", "ANY")

    for column in ["Sold Party", "Country", "City", "Salesperson", "MTD", "PARTY_TYPE", "BUSINESS TYPE"]:
        master[column] = master[column].map(_clean_text_value)
    master["Type"] = master["Type"].map(_clean_text_value).astype("string").str.upper()
    master["Sold Party Original"] = master["Sold Party"]
    master["Normalized Party"] = master["Sold Party Original"].apply(_normalize_party_name)
    master["Super Tag"] = master[tag_columns].astype(str).agg("|".join, axis=1)
    master["Source_File"] = "Bid_Summary.xlsx (latest bids)"

    columns = [
        "Stone_Id", "Shp", "Carat", "Carat_grp", "Col", "Qua", "Flo", "Lab",
        "Cut", "Pol", "Sym", "Shade", "Milky", "TBINC_NAME", "TOIN_NAME",
        "SBINC_NAME", "SOIN_NAME", "GIRDLE_OPEN_NAME", "PAV_OPEN_NAME", "RATIO",
        "DISC%", "RAP", "RATE", "AMT", "Sold Party", "Country", "City",
        "Salesperson", "Type", "MTD", "DATE", "PARTY_TYPE", "BUSINESS TYPE",
        "Sold Party Original", "Normalized Party", "Super Tag", "Source_File",
    ]
    master = master[columns]
    print(f"  MASTER_DB rows: {len(master):,}; unique Super Tags: {master['Super Tag'].nunique():,}")
    return master


def build_party_profile_temp(master_df: pd.DataFrame) -> pd.DataFrame:
    print("[3/4] Building PARTY_PROFILE_TEMP dataframe")

    def join_types(series: pd.Series) -> str:
        found: list[str] = []
        for value in series.dropna().astype(str):
            text = value.strip().upper()
            if text and text not in found:
                found.append(text)
        return ",".join(found)

    def first_nonblank(series: pd.Series, default=""):
        for value in series:
            if pd.notna(value) and str(value).strip() != "":
                return value
        return default

    working = master_df.copy()
    working["_BIDSALE"] = working["Type"].astype(str).str.upper().isin({"BID", "BID_SALE", "SALE"}).astype(int)
    profile = (
        working.groupby(["Salesperson", "Sold Party Original", "Super Tag"], sort=False, dropna=False)
        .agg(
            Count=("Super Tag", "size"),
            BidSaleCount=("_BIDSALE", "sum"),
            Types_Found=("Type", join_types),
            Stone_Id=("Stone_Id", lambda values: str(first_nonblank(values, ""))),
            SoldParty=("Sold Party", first_nonblank),
            Country=("Country", first_nonblank),
            City=("City", first_nonblank),
            MTD_Label=("MTD", first_nonblank),
        )
        .reset_index()
    )
    profile["Match_Logic_Type"] = "STANDARD"
    profile = profile[[
        "Salesperson", "Sold Party Original", "Super Tag", "Count", "BidSaleCount",
        "Types_Found", "Stone_Id", "Match_Logic_Type", "SoldParty", "Country", "City",
        "MTD_Label",
    ]]
    print(f"  PARTY_PROFILE_TEMP rows: {len(profile):,}")
    return profile


def _to_parquet_bytes(dataframe: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    dataframe.to_parquet(buffer, index=False)
    return buffer.getvalue()


def prepare_pricing_profile(
    bid_summary_bytes: bytes,
    *,
    last_n_bids: int = DEFAULT_LAST_N_BIDS,
    run_datetime: datetime | None = None,
    include_master_bytes: bool = True,
) -> tuple[bytes, bytes, dict[str, object]]:
    confirm, summary = build_confirm_last_n_bid_from_summary(
        bid_summary_bytes,
        last_n=last_n_bids,
        run_datetime=run_datetime,
    )
    master = build_master_db_from_confirm(confirm)
    profile = build_party_profile_temp(master)
    profile_bytes = _to_parquet_bytes(profile)
    master_bytes = _to_parquet_bytes(master) if include_master_bytes else b""
    return master_bytes, profile_bytes, summary


def _repository_bid_summary_bytes() -> bytes:
    path = repository_data_path("bid_summary")
    if not path.is_file():
        raise FileNotFoundError(
            "No Bid Summary upload was provided and the saved default is missing: data/Bid_Summary.xlsx"
        )
    return path.read_bytes()


def _repack_regular(raw_zip_bytes: bytes, *, date_label: str) -> tuple[bytes, str]:
    destination = io.BytesIO()
    count = 0
    with zipfile.ZipFile(io.BytesIO(raw_zip_bytes), "r") as source, zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():
            if name.lower().endswith(".xlsx"):
                target.writestr(Path(name).name, source.read(name))
                count += 1
    if count == 0:
        raise RuntimeError("No regular salesperson workbooks were generated.")
    return destination.getvalue(), f"Suggest_Stock_Last5BID_REGULAR_{date_label}.zip"


def _repack_direct(raw_zip_bytes: bytes, *, date_label: str, prefix: str = "") -> tuple[bytes, str, int]:
    destination = io.BytesIO()
    count = 0
    with zipfile.ZipFile(io.BytesIO(raw_zip_bytes), "r") as source, zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as target:
        for outer_name in source.namelist():
            if not outer_name.endswith("_Separated_Parties.zip"):
                continue
            salesperson = safe_filename(
                Path(outer_name).stem.replace("_Separated_Parties", "").strip(),
                fallback="Salesperson",
            )
            with zipfile.ZipFile(io.BytesIO(source.read(outer_name)), "r") as nested:
                for party_name in nested.namelist():
                    if not party_name.lower().endswith(".xlsx"):
                        continue
                    member = f"{prefix}{salesperson}/{Path(party_name).name}"
                    target.writestr(member, nested.read(party_name))
                    count += 1
    if count == 0:
        raise RuntimeError("No separated party XLSX files were generated; direct-party output would be empty.")
    return destination.getvalue(), f"Suggest_Stock_Last5BID_DIRECT_PARTIES_{date_label}.zip", count


def _repack_both(raw_zip_bytes: bytes, *, date_label: str) -> tuple[bytes, str, int, int]:
    destination = io.BytesIO()
    regular_count = 0
    direct_count = 0
    with zipfile.ZipFile(io.BytesIO(raw_zip_bytes), "r") as source, zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():
            if name.lower().endswith(".xlsx"):
                target.writestr(f"REGULAR_NAV/{Path(name).name}", source.read(name))
                regular_count += 1
            elif name.endswith("_Separated_Parties.zip"):
                salesperson = safe_filename(
                    Path(name).stem.replace("_Separated_Parties", "").strip(),
                    fallback="Salesperson",
                )
                with zipfile.ZipFile(io.BytesIO(source.read(name)), "r") as nested:
                    for party_name in nested.namelist():
                        if party_name.lower().endswith(".xlsx"):
                            target.writestr(
                                f"DIRECT_PARTIES/{salesperson}/{Path(party_name).name}",
                                nested.read(party_name),
                            )
                            direct_count += 1
    if regular_count == 0 and direct_count == 0:
        raise RuntimeError("No Suggested Stock output files were generated.")
    return (
        destination.getvalue(),
        f"Suggest_Stock_Last5BID_BOTH_{date_label}.zip",
        regular_count,
        direct_count,
    )


def run_pricing_last5_bid_suggested(
    *,
    stock_bytes: bytes,
    bid_summary_bytes: bytes | None = None,
    profile_perm_bytes: bytes | None = None,
    filter_bytes: bytes | None = None,
    clean_bytes: bytes | None = None,
    sales_history_files=(),
    runtime_overrides: dict[str, object] | None = None,
    last_n_bids: int = DEFAULT_LAST_N_BIDS,
    output_mode: str = DEFAULT_OUTPUT_MODE,
) -> tuple[bytes, str, str, dict[str, object]]:
    mode = str(output_mode).strip().upper()
    if mode not in OUTPUT_MODES:
        raise ValueError(f"Output mode must be one of {', '.join(OUTPUT_MODES)}.")

    bid_payload = bid_summary_bytes if bid_summary_bytes is not None else _repository_bid_summary_bytes()
    run_dt = _current_pricing_datetime()
    prep_log = io.StringIO()
    with redirect_stdout(prep_log), redirect_stderr(prep_log):
        _master_bytes, profile_temp_bytes, summary = prepare_pricing_profile(
            bid_payload,
            last_n_bids=int(last_n_bids),
            run_datetime=run_dt,
            include_master_bytes=False,
        )
        print("[4/4] Running existing Suggested Stock engine")

    # The Last-BID profile created above belongs to the synthetic Pricing salesperson.
    # Without this scope guard, the preserved engine also processes every salesperson
    # from PARTY_PROFILE_PERM, turning this pricing job into a full Suggested Stock run
    # and causing long runtimes / high memory usage. Keep the matching engine itself
    # untouched and limit only this new workspace to the profile it just created.
    runtime = dict(PRICING_LAST5_DEFAULTS)
    if runtime_overrides:
        runtime.update(runtime_overrides)

    runtime["TEST_MODE"] = True
    runtime["TEST_SALESPEOPLE"] = [FIXED_SALE_PERSON]
    runtime["TEST_MODE_PARTY"] = True
    runtime["TEST_PARTIES"] = [FIXED_SOLD_PARTY]
    summary["processing_scope"] = f"{FIXED_SALE_PERSON} / {FIXED_SOLD_PARTY}"

    # Release the original Bid Summary payload before the Excel-heavy preserved
    # engine starts. This matters on memory-constrained Streamlit workers.
    bid_payload = None
    _master_bytes = None
    gc.collect()

    # Output-mode controls are managed here so the preserved matching engine stays untouched.
    runtime["GENERATE_INDIVIDUAL_REPORTS"] = True
    runtime["GENERATE_MERGED_MASTER"] = False
    runtime["ENABLE_OUTPUT_FILE_SEPARATED"] = mode in {"DIRECT_PARTY_ZIP", "BOTH"}

    # Import lazily so Bid Summary preparation remains independently testable;
    # the deployed app already includes pyarrow for the preserved Suggested adapter.
    from .suggested_adapter import PRICING_MASTER_PATH, run_suggested_process

    raw_zip, _raw_name, engine_logs = run_suggested_process(
        stock_bytes=stock_bytes,
        profile_perm_bytes=profile_perm_bytes,
        profile_temp_bytes=profile_temp_bytes,
        filter_bytes=filter_bytes,
        clean_bytes=clean_bytes,
        sales_history_files=sales_history_files,
        runtime_overrides=runtime,
        master_path=PRICING_MASTER_PATH,
    )

    date_label = run_dt.strftime("%Y-%m-%d")
    if mode == "REGULAR_NAV_ZIP":
        data, filename = _repack_regular(raw_zip, date_label=date_label)
        summary["regular_workbooks"] = len(zipfile.ZipFile(io.BytesIO(data)).namelist())
    elif mode == "DIRECT_PARTY_ZIP":
        data, filename, direct_count = _repack_direct(raw_zip, date_label=date_label)
        summary["direct_party_files"] = direct_count
    else:
        data, filename, regular_count, direct_count = _repack_both(raw_zip, date_label=date_label)
        summary["regular_workbooks"] = regular_count
        summary["direct_party_files"] = direct_count

    summary["output_mode"] = mode
    logs = prep_log.getvalue() + "\n" + (engine_logs or "")
    return data, filename, logs, summary
