from __future__ import annotations

import io
import json
import re
import tempfile
import zipfile
from contextlib import ExitStack
from pathlib import Path

import streamlit as st

from adapters.auto_pricing_adapter import ACTION_NAMES as AUTO_PRICING_ACTION_NAMES, run_auto_pricing
from adapters.auto_pricing_pipeline_adapter import (
    default_remove_party_bytes,
    run_auto_pricing_full_pipeline_from_files,
)
from adapters.bid_drop_adapter import run_bid_drop_checking
from adapters.common import repository_data_path
from adapters.dealmaker_adapter import bundled_dealmaker_control_bytes, run_dealmaker
from adapters.filter_demand_adapter import run_filter_demand
from adapters.layout_stock_matcher_adapter import run_layout_stock_matcher
from adapters.separate_adapter import run_separate_excel_files
from adapters.sort_adapter import run_sorting_process
from adapters.suggested_list_splitter_adapter import run_suggested_list_splitter
from adapters.usa_suggested_stock_mapper_adapter import run_usa_suggested_stock_mapper
from adapters.suggested_adapter import run_suggested_process
from cloud_jobs.bundle import (
    FAILURE_ASSET_NAME,
    build_job_bundle,
    new_job_tag,
)

from cloud_jobs.github_backend import (
    GitHubCloudConfig,
    GitHubJobError,
    GitHubReleaseBackend,
)
from ui.theme import apply_theme, hero, page_footer, section_header, sidebar_brand, sidebar_footer
from ui.config_controls import (
    DEFAULT_PRIORITY_TYPES,
    DEFAULT_TYPE_SEQUENCE,
    combine_selected_and_custom,
    default_parties,
    default_salespeople,
    discover_suggested_names_from_streams,
    format_type_weights,
    parse_delimited_values,
    parse_type_weights,
    unique_clean,
)

# Keep the Streamlit UI compatible with both the original GitHub worker
# result name and the final-output worker. Defining these names here avoids
# a startup failure when a repository is updated only partially.
RESULT_ASSET_NAME = "suggested_reports_final_combined.zip"
LEGACY_RESULT_ASSET_NAME = "suggested_reports_all_batches.zip"
APP_VERSION = (Path(__file__).resolve().parent / "VERSION.txt").read_text(encoding="utf-8").strip()


st.set_page_config(
    page_title="Matching Mapper World",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded",
)
apply_theme()
sidebar_brand()

PAGES = [
    "Overview",
    "Suggested Stock Mapper",
    "USA Suggested Stock Mapper",
    "Auto Pricing",
    "Dealmaker",
    "Layout Stock Matcher",
    "Demand Filter Mapper",
    "Sort Suggested Output",
    "Suggested List Splitter",
    "BID Drop Checking Process",
    "Separate Excel Sheets",
]
page = st.sidebar.radio("WORKSPACE", PAGES, label_visibility="visible")
sidebar_footer()


def _uploaded_file_identity(uploaded_file) -> tuple[str, str, int] | None:
    if uploaded_file is None:
        return None
    return (
        str(getattr(uploaded_file, "file_id", "")),
        str(getattr(uploaded_file, "name", "")),
        int(getattr(uploaded_file, "size", 0) or 0),
    )


def _save_result(
    key: str,
    data: bytes,
    filename: str,
    logs: str,
    mime: str = "application/zip",
) -> None:
    st.session_state[key] = {
        "data": data, "filename": filename, "logs": logs, "mime": mime
    }


def _repository_default_exists(key: str) -> bool:
    return repository_data_path(key).is_file()


def _input_or_default_stream(stack: ExitStack, uploaded_file, key: str):
    if uploaded_file is not None:
        return uploaded_file
    path = repository_data_path(key)
    return stack.enter_context(path.open("rb")) if path.is_file() else None


def _input_identity(uploaded_file, key: str) -> tuple[str, str, int] | None:
    uploaded = _uploaded_file_identity(uploaded_file)
    if uploaded is not None:
        return uploaded
    path = repository_data_path(key)
    if not path.is_file():
        return None
    stat = path.stat()
    return ("repository-default", path.name, int(stat.st_size))


def _friendly_message(value: object, fallback: str = "The request could not be completed.") -> str:
    text = str(value or "").strip()
    if not text:
        return fallback
    replacements = (
        ("GitHub Actions", "the processing service"),
        ("GitHub", "the processing service"),
        ("Streamlit Community Cloud", "the online application"),
        ("Streamlit", "the application"),
        ("repository", "secure workspace"),
        ("Repository", "Secure workspace"),
        ("background worker", "processing service"),
        ("worker", "processor"),
        ("runner", "processor"),
        ("REST API", "service"),
        ("API", "service"),
        ("release asset", "result file"),
        ("release", "request"),
        ("secrets", "settings"),
        ("workflow", "process"),
        ("manifest", "summary"),
        ("batches", "groups"),
        ("batch", "group"),
    )
    for old, new in replacements:
        text = text.replace(old, new)
    text = re.sub(r"https?://\S+", "", text)
    text = re.sub(r"(?:[A-Za-z]:)?[/\\](?:[^\s:/\\]+[/\\])+[^\s:/\\]+", "an internal file", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or fallback


def _show_user_error(context: str, exc: object | None = None) -> None:
    detail = _friendly_message(exc, "") if exc is not None else ""
    st.error(f"{context}{': ' + detail if detail else '.'}")


def _show_result(key: str, label: str) -> None:
    result = st.session_state.get(key)
    if not result:
        return
    c1, c2, c3 = st.columns(3)
    c1.metric("Output status", "Ready")
    c2.metric("File size", f"{len(result['data']) / (1024 * 1024):,.2f} MB")
    if str(result["filename"]).lower().endswith(".zip"):
        try:
            with zipfile.ZipFile(io.BytesIO(result["data"])) as zf:
                count = len([n for n in zf.namelist() if not n.endswith("/")])
        except Exception:
            count = 1
    else:
        count = 1
    c3.metric("Files packaged", f"{count:,}")
    st.success("Processing completed. Your output is ready to download.")
    st.download_button(
        label,
        data=result["data"],
        file_name=result["filename"],
        mime=result.get("mime", "application/zip"),
        type="primary",
        use_container_width=True,
        key=f"download_{key}",
        on_click="ignore",
    )
    with st.expander("Processing details", expanded=False):
        st.write("Your files were processed successfully and the download package is ready.")


def _github_job_settings() -> GitHubCloudConfig | None:
    try:
        values = st.secrets.get("github", {})
        return GitHubCloudConfig.from_mapping(values)
    except Exception:
        return None


def _show_github_job_panel() -> None:
    job = st.session_state.get("github_suggested_job")
    if not job:
        return

    st.markdown("#### Submitted full processing request")
    settings = _github_job_settings()
    if settings is None:
        st.error("Full processing is not available right now. Please contact the administrator.")
        return

    backend = GitHubReleaseBackend(settings)
    snapshot = None
    try:
        snapshot = backend.job_snapshot(int(job["release_id"]))
    except Exception as exc:
        st.warning(f"Unable to refresh this request right now: {_friendly_message(exc)}")

    status_payload = snapshot.get("status", {}) if snapshot else {}
    state = str(status_payload.get("state") or job.get("status", "submitted"))
    columns = st.columns(4)
    columns[0].metric("Request status", state.replace("_", " ").title())
    columns[1].metric("Salespeople per group", str(job.get("batch_size", 3)))
    columns[2].metric("Salespeople", str(status_payload.get("total_salespeople", job.get("salespeople_count", "Pending"))))
    current_batch = status_payload.get("current_batch")
    total_batches = status_payload.get("total_batches")
    if isinstance(total_batches, int) and total_batches > 0:
        if state in {"completed", "partial_success"}:
            current_batch = total_batches
        completed_count = current_batch if isinstance(current_batch, int) else 0
        progress_label = f"{completed_count}/{total_batches}"
    else:
        progress_label = "Queued" if state in {"queued", "submitted", "starting"} else "Preparing"
    columns[3].metric("Group progress", progress_label)
    st.caption(f"Reference ID: {job['tag_name']}")
    message = str(status_payload.get("message", "Waiting for processing to start."))
    st.caption(_friendly_message(message, "Processing status is being updated."))
    if isinstance(current_batch, int) and isinstance(total_batches, int) and total_batches > 0:
        st.progress(min(max(current_batch / total_batches, 0.0), 1.0))

    refresh, prepare, cleanup = st.columns(3)
    if refresh.button("Refresh status", use_container_width=True, key="refresh_github_job"):
        st.rerun()

    assets_by_name = snapshot.get("assets_by_name", {}) if snapshot else {}
    result_asset_name = (
        RESULT_ASSET_NAME
        if RESULT_ASSET_NAME in assets_by_name
        else LEGACY_RESULT_ASSET_NAME
        if LEGACY_RESULT_ASSET_NAME in assets_by_name
        else None
    )
    result_ready = result_asset_name is not None
    failure_ready = FAILURE_ASSET_NAME in assets_by_name

    if prepare.button(
        "Create result download link",
        use_container_width=True,
        disabled=not result_ready,
        key="prepare_github_result",
    ):
        try:
            assert result_asset_name is not None
            asset = assets_by_name[result_asset_name]
            secure_url = backend.temporary_asset_download_url(asset)
            if secure_url:
                st.session_state["github_suggested_download_url"] = secure_url
                st.session_state.pop("github_suggested_download", None)
            else:
                target = Path(tempfile.gettempdir()) / f"{job['tag_name']}_{result_asset_name}"
                with st.spinner("Preparing your result for download..."):
                    backend.download_asset_to_path(asset, target)
                st.session_state["github_suggested_download"] = str(target)
                st.session_state.pop("github_suggested_download_url", None)
            st.rerun()
        except Exception as exc:
            _show_user_error("Unable to complete this action", exc)

    confirm_delete = cleanup.checkbox("Allow cleanup", key="confirm_github_cleanup")
    if cleanup.button(
        "Delete job files",
        use_container_width=True,
        disabled=not confirm_delete,
        key="delete_github_job",
        help="Deletes temporary processing files after you have downloaded the result.",
    ):
        try:
            backend.delete_job(int(job["release_id"]), tag_name=str(job.get("tag_name", "")))
            download_path = st.session_state.pop("github_suggested_download", None)
            st.session_state.pop("github_suggested_download_url", None)
            if download_path:
                Path(download_path).unlink(missing_ok=True)
            st.session_state.pop("github_suggested_job", None)
            st.success("Temporary processing files were deleted.")
            st.rerun()
        except Exception as exc:
            _show_user_error("Unable to complete this action", exc)

    download_url = st.session_state.get("github_suggested_download_url")
    if download_url:
        st.success("A short-lived private download link is ready. Open it promptly because it may expire.")
        st.link_button(
            "Download final combined result ZIP",
            download_url,
            type="primary",
            use_container_width=True,
        )

    download_path = st.session_state.get("github_suggested_download")
    if download_path and Path(download_path).exists():
        result_path = Path(download_path)
        st.success("The combined result is ready for browser download.")
        with result_path.open("rb") as handle:
            st.download_button(
                "Download final combined result ZIP",
                data=handle,
                file_name=RESULT_ASSET_NAME,
                mime="application/zip",
                type="primary",
                use_container_width=True,
                key="download_github_full_result",
                on_click="ignore",
            )

    if failure_ready or state == "failed":
        st.error("Processing could not be completed. Please retry or contact support with the reference ID shown above.")
    elif state == "partial_success":
        st.warning("The result is ready, but one or more groups could not be completed. You may download the available result or try again.")


def _show_github_recovery_panel(settings: GitHubCloudConfig | None) -> None:
    if settings is None:
        return
    with st.expander("Recover a previous full processing request", expanded=False):
        backend = GitHubReleaseBackend(settings)
        manual_tag = st.text_input(
            "Reference ID",
            placeholder="mmw-job-YYYYMMDD-HHMMSS-xxxxxxxx",
            key="recover_github_tag",
        )
        a, b = st.columns(2)
        if a.button("Load reference", use_container_width=True, disabled=not manual_tag.strip()):
            try:
                release = backend.find_job_by_tag(manual_tag.strip())
                st.session_state["github_suggested_job"] = {
                    "tag_name": release["tag_name"],
                    "release_id": release["id"],
                    "status": "submitted",
                    "batch_size": "?",
                    "salespeople_count": "Pending",
                }
                st.rerun()
            except Exception as exc:
                _show_user_error("Unable to complete this action", exc)

        if b.button("Find recent requests", use_container_width=True):
            try:
                releases = backend.list_recent_jobs(limit=20)
                st.session_state["recent_github_jobs"] = [
                    {
                        "tag_name": release["tag_name"],
                        "name": release.get("name") or release["tag_name"],
                        "release_id": release["id"],
                    }
                    for release in releases
                ]
            except Exception as exc:
                _show_user_error("Unable to complete this action", exc)

        recent = st.session_state.get("recent_github_jobs", [])
        if recent:
            labels = [f"{item['name']} | {item['tag_name']}" for item in recent]
            selected_label = st.selectbox("Recent requests", labels, key="recent_github_job_select")
            selected = recent[labels.index(selected_label)]
            if st.button("Load selected request", use_container_width=True):
                st.session_state["github_suggested_job"] = {
                    "tag_name": selected["tag_name"],
                    "release_id": selected["release_id"],
                    "status": "submitted",
                    "batch_size": "?",
                    "salespeople_count": "Pending",
                }
                st.rerun()


def overview_page() -> None:
    hero()
    metrics = st.columns(4)
    metrics[0].metric("Available services", "10")
    metrics[1].metric("Core operations", "10")
    metrics[2].metric("Processing", "Secure full processing")
    metrics[3].metric("Primary outputs", "XLSX / ZIP")

    st.markdown("### Choose a service")
    cards = [
        ("PROFILE MATCHING", "Suggested Stock Mapper", "Output: Suggested report ZIP"),
        ("USA SALES HISTORY", "USA Suggested Stock Mapper", "Output: Salesperson / party suggested ZIP"),
        ("PRICING INTELLIGENCE", "Auto Pricing", "Output: Pricing revision + cross-verification ZIP"),
        ("DEAL ALLOCATION", "Dealmaker", "Output: Production review ZIP"),
        ("LAYOUT MATCHING", "Layout Stock Matcher", "Output: Salesperson layout suggestions ZIP"),
        ("DEMAND RULES", "Demand Filter Mapper", "Output: Salesperson report ZIP"),
        ("PRIORITIZATION", "Sort Suggested Output", "Output: Sorted suggested ZIP"),
        ("LIST SPLITTING", "Suggested List Splitter", "Output: Under 10 / 10+ ZIPs"),
        ("BID DIAGNOSTICS", "BID Drop Checking Process", "Output: Deep gap analysis XLSX"),
        ("WORKBOOK DELIVERY", "Separate Excel Sheets", "Output: Nested separated ZIP"),
    ]
    for row_start in range(0, len(cards), 2):
        columns = st.columns(2)
        for column, (kicker, title, output) in zip(columns, cards[row_start:row_start + 2]):
            with column:
                st.markdown(
                    f"""
                    <article class="mm-card"><div class="mm-kicker">{kicker}</div><h3>{title}</h3><div class="mm-card-foot">{output}</div></article>
                    """,
                    unsafe_allow_html=True,
                )


def suggested_page() -> None:
    section_header(
        "Profile, demand & history intelligence",
        "Suggested Stock Mapper",
        "Run the complete stock recommendation process with all matching, report, scope, styling, and sorting options available in one place.",
        ["Upload inputs", "Configure matching", "Select scope", "Generate & download"],
    )

    tab_files, tab_matching, tab_display, tab_scope, tab_status, tab_sorting = st.tabs(
        [
            "Input files",
            "Matching dimensions",
            "Display & outputs",
            "Salesperson / party scope",
            "Status, styling & cleanup",
            "Advanced sorting",
        ]
    )

    with tab_files:
        left, right = st.columns(2)
        with left:
            stock = st.file_uploader("Stock Actual workbook *", type=["xlsx"], key="suggest_stock")
            profile_perm = st.file_uploader("Permanent party profile (optional override)", type=["parquet"], key="suggest_perm")
            profile_temp = st.file_uploader("Temporary party profile (optional override)", type=["parquet"], key="suggest_temp")
        with right:
            filters = st.file_uploader("Filters workbook (optional override)", type=["xlsx"], key="suggest_filters")
            no_need = st.file_uploader("No-need / exclusion workbook (optional override)", type=["xlsx"], key="suggest_noneed")
            sales_history = st.file_uploader(
                "Sales history workbook(s)", type=["xlsx"], accept_multiple_files=True, key="suggest_sales"
            )
        st.caption(
            "Stock Actual must be uploaded. Omitted profiles/rule files use: "
            "data/PARTY_PROFILE_PERM.parquet, data/PARTY_PROFILE_TEMP.parquet, "
            "data/filters - EBID.xlsx, and data/no_need.xlsx."
        )

        discovery_identity = (
            _input_identity(profile_perm, "profile_perm"),
            _input_identity(profile_temp, "profile_temp"),
            _input_identity(filters, "filters"),
        )
        scan_names = st.button(
            "Scan uploaded/default files for salesperson and party names (optional)",
            use_container_width=True,
            disabled=not any((
                profile_perm, profile_temp, filters,
                _repository_default_exists("profile_perm"),
                _repository_default_exists("profile_temp"),
                _repository_default_exists("filters"),
            )),
            help=("Use this only when you need names from the available files in the manual selection lists. Full processing reads the names automatically, so this scan is optional."),
        )
        if scan_names:
            with st.spinner("Reading only salesperson and party columns..."):
                with ExitStack() as stack:
                    found_salespeople, found_parties = discover_suggested_names_from_streams(
                        _input_or_default_stream(stack, profile_perm, "profile_perm"),
                        _input_or_default_stream(stack, profile_temp, "profile_temp"),
                        _input_or_default_stream(stack, filters, "filters"),
                    )
            st.session_state["suggested_name_discovery"] = {
                "identity": discovery_identity,
                "salespeople": found_salespeople,
                "parties": found_parties,
            }
            st.success(
                f"Found {len(found_salespeople)} salesperson name(s) and {len(found_parties)} party name(s)."
            )

    discovery = st.session_state.get("suggested_name_discovery", {})
    if discovery.get("identity") == discovery_identity:
        discovered_salespeople = list(discovery.get("salespeople") or [])
        discovered_parties = list(discovery.get("parties") or [])
    else:
        discovered_salespeople = []
        discovered_parties = []
    salesperson_options = unique_clean(default_salespeople() + discovered_salespeople)
    party_options = unique_clean(default_parties() + discovered_parties)

    include_fields = {
        "COLOR": True,
        "CLARITY": True,
        "FLORO": True,
        "CUT": False,
        "POLISH": False,
        "SYMMETRY": False,
        "SHADE": False,
        "MILKY": False,
        "TABLE_BLACK": False,
        "SIDE_BLACK": False,
        "TABLE_OPEN": False,
        "SIDE_OPEN": False,
        "GIRDLE_OPEN": False,
        "PAVILION_OPEN": False,
    }
    tolerance_defaults = {
        "COLOR": True,
        "CLARITY": True,
        "FLORO": True,
        "CUT": False,
        "POLISH": True,
        "SYMMETRY": True,
        "SHADE": True,
        "MILKY": True,
        "TABLE_BLACK": True,
        "SIDE_BLACK": True,
        "TABLE_OPEN": True,
        "SIDE_OPEN": True,
        "GIRDLE_OPEN": True,
        "PAVILION_OPEN": True,
    }
    runtime: dict[str, object] = {}

    with tab_matching:
        c1, c2, c3 = st.columns(3)
        runtime["TOLERANCE_LEVEL"] = int(
            c1.number_input("Default tolerance level", min_value=0, max_value=3, value=1, step=1)
        )
        runtime["ENABLE_FLO_MAPPING"] = c2.toggle("Group VSL / SL with FNT", value=True)
        use_top_n = c3.toggle("Limit top supertags", value=False)
        runtime["TOP_N_SUPERTAGS"] = (
            int(c3.number_input("Top N supertags", min_value=1, value=20, step=1, disabled=not use_top_n))
            if use_top_n
            else None
        )

        st.markdown("#### Matching dimensions and tolerance switches")
        h1, h2, h3 = st.columns([1.5, 1, 1])
        h1.markdown("**Dimension**")
        h2.markdown("**Include in matching**")
        h3.markdown("**Allow tolerance**")
        for label, default in include_fields.items():
            display_name = label.replace("_", " ").title()
            a, b, c = st.columns([1.5, 1, 1])
            a.write(display_name)
            runtime[f"INCLUDE_{label}"] = b.checkbox(
                "Include",
                value=default,
                key=f"suggest_inc_{label}",
                label_visibility="collapsed",
            )
            runtime[f"ENABLE_{label}_TOLERANCE"] = c.checkbox(
                "Tolerance",
                value=tolerance_defaults[label],
                key=f"suggest_tol_{label}",
                label_visibility="collapsed",
                disabled=not runtime[f"INCLUDE_{label}"],
            )

    with tab_display:
        st.markdown("#### Sales supertag and reference columns")
        d1, d2, d3, d4 = st.columns(4)
        runtime["SHOW_BASIC_SUPERTAG"] = d1.toggle("Show basic supertag", value=True)
        runtime["SHOW_ADVANCED_SUPERTAG"] = d2.toggle("Show advanced supertag", value=False)
        runtime["SHOW_REF_STONE_ID"] = d3.toggle("Show reference Stone ID", value=True)
        runtime["SHOW_CONCAT_STATUS"] = d4.toggle("Show concatenated status", value=False)

        st.markdown("#### Report formats")
        r1, r2, r3 = st.columns(3)
        runtime["GENERATE_MERGED_MASTER"] = r1.toggle("Generate merged master", value=True)
        runtime["GENERATE_INDIVIDUAL_REPORTS"] = r1.toggle("Generate salesperson workbooks", value=False)
        runtime["ENABLE_OUTPUT_FILE_SEPARATED"] = r2.toggle(
            "Create separated party ZIPs",
            value=False,
            disabled=not runtime["GENERATE_INDIVIDUAL_REPORTS"],
            help="Used only while individual salesperson reports are enabled.",
        )
        runtime["ENABLE_TYPE_COLUMNS"] = r2.toggle("Add Type_* columns", value=False)
        runtime["ENABLE_SALES_SUPERTAG_COLUMN"] = r3.toggle("Keep Matched_Sales_Tag column", value=False)
        runtime["MAX_ROWS_PER_SHEET"] = int(
            r3.number_input(
                "Intermediate batch-sheet row limit",
                min_value=1_000,
                max_value=1_048_576,
                value=1_000_000,
                step=10_000,
                help=("This limit is used only while large requests are processed in smaller groups. The final downloaded MASTER_CONSOLIDATED_REPORT.xlsx is assembled into one Master_Data worksheet when it fits Excel's row limit."),
            )
        )

        type_sequence_text = st.text_area(
            "Type column sequence",
            value=", ".join(DEFAULT_TYPE_SEQUENCE),
            disabled=not runtime["ENABLE_TYPE_COLUMNS"],
            help="Comma-separated order used by STATIC_TYPE_SEQUENCE.",
        )
        parsed_type_sequence = parse_delimited_values(type_sequence_text)
        runtime["STATIC_TYPE_SEQUENCE"] = parsed_type_sequence or list(DEFAULT_TYPE_SEQUENCE)

    with tab_scope:
        st.markdown("#### Salesperson output scope")
        s1, s2 = st.columns([1, 2])
        runtime["TEST_MODE"] = s1.toggle(
            "Enable salesperson selection",
            value=True,
            help="When off, every salesperson found in the profiles and filters is processed.",
        )
        selected_salespeople = s2.multiselect(
            "Select one or multiple salespeople",
            options=salesperson_options,
            default=default_salespeople()[:1],
            disabled=not runtime["TEST_MODE"],
            help="Options combine the available default lists with names found in uploaded files.",
        )
        custom_salespeople = st.text_area(
            "Additional salesperson names",
            placeholder="Mr. New Person, Ms. Another Person",
            disabled=not runtime["TEST_MODE"],
            help="Comma, semicolon or new-line separated. This makes future names usable without code changes.",
        )
        runtime["TEST_SALESPEOPLE"] = combine_selected_and_custom(selected_salespeople, custom_salespeople)

        st.markdown("#### Party output scope")
        p1, p2 = st.columns([1, 2])
        runtime["TEST_MODE_PARTY"] = p1.toggle(
            "Enable party selection",
            value=False,
            help="When enabled, output is restricted to the selected party names after normalization.",
        )
        selected_parties = p2.multiselect(
            "Select one or multiple parties",
            options=party_options,
            default=[],
            disabled=not runtime["TEST_MODE_PARTY"],
        )
        custom_parties = st.text_area(
            "Additional party names",
            placeholder="NS LUXURY, VIDHYA GEMS, NOOR JEWELS, AL BUTA JEWELLERY LLC",
            disabled=not runtime["TEST_MODE_PARTY"],
            help="Comma, semicolon or new-line separated.",
        )
        runtime["TEST_PARTIES"] = combine_selected_and_custom(selected_parties, custom_parties)

        st.caption(
            f"Available options: {len(salesperson_options)} salesperson name(s) and "
            f"{len(party_options)} party name(s), including names discovered from current uploads."
        )

    with tab_status:
        st.markdown("#### Match-status output filtering")
        f1, f2 = st.columns([1, 2])
        runtime["MATCH_STATUS_FILTER_MODE"] = f1.selectbox(
            "Filter mode",
            ["NONE", "ALL", "SPECIFIC"],
            index=0,
            help="NONE keeps all statuses. ALL filters every party. SPECIFIC applies only to named parties.",
        )
        runtime["MATCH_STATUS_ALLOWED_TYPES"] = f2.multiselect(
            "Allowed match statuses",
            ["Exact", "Near by", "May be", "Filter"],
            default=["Exact", "Near by", "Filter"],
            disabled=runtime["MATCH_STATUS_FILTER_MODE"] == "NONE",
        )
        selected_status_parties = st.multiselect(
            "Parties using status filtering",
            options=party_options,
            default=[],
            disabled=runtime["MATCH_STATUS_FILTER_MODE"] != "SPECIFIC",
        )
        custom_status_parties = st.text_area(
            "Additional status-filter party names",
            placeholder="Party One, Party Two",
            disabled=runtime["MATCH_STATUS_FILTER_MODE"] != "SPECIFIC",
        )
        runtime["MATCH_STATUS_SPECIFIC_PARTIES"] = (
            combine_selected_and_custom(selected_status_parties, custom_status_parties) or ["NONE"]
        )

        st.markdown("#### Workbook styling and cleanup")
        c1, c2, c3 = st.columns(3)
        runtime["ENABLE_STYLING"] = c1.toggle("Enable Excel styling", value=True)
        runtime["ENABLE_NEW_FILL_LOGIC"] = c1.toggle("Use new fill-color logic", value=True)
        runtime["ENABLE_CLEAN_FILE"] = c2.toggle("Apply no-need rules", value=True)
        runtime["ENABLE_PAIR_MODE"] = c2.toggle("Protect and synchronize pairs", value=True)
        runtime["ENABLE_TOP_CUSTOMERS_HIGHLIGHT"] = c3.toggle(
            "Highlight top customers",
            value=True,
            disabled=not runtime["GENERATE_INDIVIDUAL_REPORTS"],
        )
        runtime["TOP_CUSTOMERS_N"] = int(
            c3.number_input(
                "Top customers count",
                min_value=1,
                value=10,
                step=1,
                disabled=not runtime["ENABLE_TOP_CUSTOMERS_HIGHLIGHT"],
            )
        )

        runtime["ENABLE_SMALL_OUTPUT_DELETE"] = st.toggle(
            "Delete small party outputs",
            value=False,
            disabled=not runtime["GENERATE_INDIVIDUAL_REPORTS"],
        )
        runtime["SMALL_OUTPUT_THRESHOLD"] = int(
            st.number_input(
                "Small-output row threshold",
                min_value=0,
                value=5,
                step=1,
                disabled=not runtime["ENABLE_SMALL_OUTPUT_DELETE"],
            )
        )

    with tab_sorting:
        runtime["ENABLE_ADVANCED_SORTING"] = st.toggle(
            "Enable advanced sales-history sorting and limiting",
            value=False,
            help="This module is used for individual party worksheets when their row count reaches the configured threshold.",
        )
        a1, a2, a3 = st.columns(3)
        runtime["SORT_COUNT_THRESHOLD"] = int(
            a1.number_input(
                "Minimum rows before sorting",
                min_value=1,
                value=250,
                step=10,
                disabled=not runtime["ENABLE_ADVANCED_SORTING"],
            )
        )
        runtime["SORT_NORMAL_PARTY_LIMIT"] = int(
            a2.number_input(
                "Normal-party row limit",
                min_value=1,
                value=120,
                step=10,
                disabled=not runtime["ENABLE_ADVANCED_SORTING"],
            )
        )
        runtime["SORT_HIGH_PRIORITY_PARTY_LIMIT"] = int(
            a3.number_input(
                "High-priority row limit",
                min_value=1,
                value=150,
                step=10,
                disabled=not runtime["ENABLE_ADVANCED_SORTING"],
            )
        )

        b1, b2, b3 = st.columns(3)
        runtime["SORT_EXACT_CAP"] = int(
            b1.number_input("Exact family cap", min_value=0, value=3, step=1, disabled=not runtime["ENABLE_ADVANCED_SORTING"])
        )
        runtime["SORT_NEAR_CAP"] = int(
            b2.number_input("Nearby family cap", min_value=0, value=2, step=1, disabled=not runtime["ENABLE_ADVANCED_SORTING"])
        )
        runtime["SORT_MAYBE_CAP"] = int(
            b3.number_input("Maybe family cap", min_value=0, value=1, step=1, disabled=not runtime["ENABLE_ADVANCED_SORTING"])
        )

        priority_types_text = st.text_input(
            "High-priority party types",
            value=", ".join(DEFAULT_PRIORITY_TYPES),
            disabled=not runtime["ENABLE_ADVANCED_SORTING"],
        )
        runtime["SORT_HIGH_PRIORITY_PARTY_TYPES"] = set(
            value.upper() for value in (parse_delimited_values(priority_types_text) or DEFAULT_PRIORITY_TYPES)
        )
        runtime["SORT_HIGH_PRIORITY_TOTAL_HISTORY_THRESHOLD"] = int(
            st.number_input(
                "High-priority total-history threshold",
                min_value=0,
                value=200,
                step=10,
                disabled=not runtime["ENABLE_ADVANCED_SORTING"],
            )
        )

    missing: list[str] = []
    if stock is None:
        missing.append("Stock Actual")
    if filters is None and not _repository_default_exists("filters"):
        missing.append("Filters upload or data/filters - EBID.xlsx")
    if (
        profile_perm is None
        and profile_temp is None
        and not _repository_default_exists("profile_perm")
        and not _repository_default_exists("profile_temp")
    ):
        missing.append("party profile file")
    if runtime["ENABLE_CLEAN_FILE"] and no_need is None and not _repository_default_exists("no_need"):
        missing.append("No-need upload or data/no_need.xlsx")
    if not runtime["GENERATE_MERGED_MASTER"] and not runtime["GENERATE_INDIVIDUAL_REPORTS"]:
        missing.append("at least one report format")
    if runtime["TEST_MODE"] and not runtime["TEST_SALESPEOPLE"]:
        missing.append("one or more selected salespeople")
    if runtime["TEST_MODE_PARTY"] and not runtime["TEST_PARTIES"]:
        missing.append("one or more selected parties")
    if runtime["MATCH_STATUS_FILTER_MODE"] != "NONE" and not runtime["MATCH_STATUS_ALLOWED_TYPES"]:
        missing.append("one or more allowed match statuses")
    if runtime["MATCH_STATUS_FILTER_MODE"] == "SPECIFIC" and runtime["MATCH_STATUS_SPECIFIC_PARTIES"] == ["NONE"]:
        missing.append("one or more status-filter parties")
    if runtime["ENABLE_ADVANCED_SORTING"] and not sales_history:
        st.warning("Advanced sorting is enabled without sales-history files. History-based limits will be skipped when no sales history is available.")

    st.markdown("#### Processing option")
    processing_mode = st.radio(
        "Choose where the Suggested Stock Mapper runs",
        [
            "Full processing (recommended)",
            "Quick browser test (up to 3 salespeople)",
        ],
        horizontal=True,
        help="Full processing is recommended for large files and all-salespeople requests.",
    )

    if processing_mode.startswith("Full processing"):
        settings = _github_job_settings()
        c1, c2, c3 = st.columns(3)
        batch_size = int(c1.number_input("Salespeople per group", min_value=1, max_value=20, value=1, step=1))
        continue_after_failure = c2.toggle("Continue after an unsuccessful group", value=True)
        scope_mode = c3.selectbox(
            "People to include",
            ["All salespeople discovered from current inputs", "Only salespeople selected above"],
        )
        requested_by = st.text_input(
            "Requested by (optional)",
            placeholder="Name or work email for the request record",
        )

        cloud_scope = "all" if scope_mode.startswith("All") else "selected"
        cloud_salespeople = [] if cloud_scope == "all" else list(runtime.get("TEST_SALESPEOPLE") or [])

        cloud_missing = [
            item
            for item in missing
            if not (cloud_scope == "all" and item == "one or more selected salespeople")
        ]
        if cloud_scope == "selected" and not cloud_salespeople:
            cloud_missing.append("one or more selected salespeople")
        if settings is None:
            cloud_missing.append("full processing connection")
            st.warning("Full processing is not connected yet. Please contact the administrator to complete the application setup.")

        if cloud_scope == "all":
            st.caption(f"The application will read salesperson names from the available profile and Filters files, then process them in groups of {batch_size}. All matching, filtering, styling, and output choices above are applied to every group.")
        else:
            estimated_batches = (len(cloud_salespeople) + batch_size - 1) // batch_size if cloud_salespeople else 0
            st.caption(f"This request will process {len(cloud_salespeople)} selected salesperson(s) in approximately {estimated_batches} group(s). All other choices remain the same for every group.")
        st.info("Large requests are processed in smaller groups and combined into one final ZIP. When Merged Master is enabled, one global MASTER_CONSOLIDATED_REPORT.xlsx is included with any enabled salesperson workbooks and separated-party files.")

        if st.button(
            "Start full processing",
            type="primary",
            use_container_width=True,
            disabled=bool(cloud_missing),
            key="submit_github_full_run",
        ):
            try:
                assert settings is not None
                backend = GitHubReleaseBackend(settings)
                with st.spinner("Checking the secure processing connection..."):
                    backend.check_connection(require_private=True)
                    active_jobs = backend.list_active_jobs(limit=30)
                    if active_jobs:
                        active_tag = str(active_jobs[0].get("tag_name", "another job"))
                        raise GitHubJobError(
                            "A Matching Mapper request is already queued or running: "
                            f"{active_tag}. Please let it finish or remove the previous request before starting another."
                        )
                job_tag = new_job_tag()
                with tempfile.TemporaryDirectory(prefix="matching_mapper_submit_") as tmp:
                    bundle_path = Path(tmp) / "input_bundle.zip"
                    with st.spinner("Preparing files and starting full processing..."):
                        build_job_bundle(
                            output_path=bundle_path,
                            job_id=job_tag,
                            runtime_overrides=runtime,
                            scope=cloud_scope,
                            salespeople=cloud_salespeople,
                            batch_size=batch_size,
                            continue_after_failed_batch=continue_after_failure,
                            stock_file=(stock.name, stock),
                            filter_file=(filters.name, filters) if filters else None,
                            profile_perm_file=(profile_perm.name, profile_perm) if profile_perm else None,
                            profile_temp_file=(profile_temp.name, profile_temp) if profile_temp else None,
                            clean_file=(no_need.name, no_need) if no_need else None,
                            sales_history_files=[(file.name, file) for file in (sales_history or [])],
                            requested_by=requested_by,
                        )
                        release = backend.submit_job(
                            job_tag=job_tag,
                            bundle_path=bundle_path,
                            requested_by=requested_by,
                        )
                st.session_state["github_suggested_job"] = {
                    "tag_name": release["tag_name"],
                    "release_id": release["id"],
                    "status": "queued",
                    "batch_size": batch_size,
                    "salespeople_count": len(cloud_salespeople) if cloud_scope == "selected" else "Discovering",
                }
                st.success("Full processing started. It will continue even if this page is closed.")
                st.rerun()
            except GitHubJobError as exc:
                _show_user_error("Full processing could not be started", exc)
            except Exception as exc:
                _show_user_error("Unable to complete this action", exc)

        if cloud_missing:
            st.info("Complete the following before submission: " + "; ".join(dict.fromkeys(cloud_missing)))
        _show_github_recovery_panel(settings)
        _show_github_job_panel()
    else:
        local_missing = list(missing)
        local_salespeople = list(runtime.get("TEST_SALESPEOPLE") or [])
        if not runtime.get("TEST_MODE"):
            local_missing.append("enable salesperson selection for the local test")
        elif len(local_salespeople) > 3:
            local_missing.append("select no more than 3 salespeople for the local test")
        st.warning("The quick browser test is limited to at most 3 selected salespeople. Use full processing for all-salespeople or large requests.")
        if st.button(
            "Generate test reports",
            type="primary",
            use_container_width=True,
            disabled=bool(local_missing),
            key="run_suggested_locally",
        ):
            try:
                with st.spinner("Generating the selected reports. Large profile files may take several minutes..."):
                    data, filename, logs = run_suggested_process(
                        stock_bytes=stock.getvalue(),
                        profile_perm_bytes=profile_perm.getvalue() if profile_perm else None,
                        profile_temp_bytes=profile_temp.getvalue() if profile_temp else None,
                        filter_bytes=filters.getvalue() if filters else None,
                        clean_bytes=no_need.getvalue() if no_need else None,
                        sales_history_files=[(file.name, file.getvalue()) for file in (sales_history or [])],
                        runtime_overrides=runtime,
                    )
                    _save_result("suggested_result", data, filename, logs)
                st.rerun()
            except Exception as exc:
                _show_user_error("Unable to complete this action", exc)

        if local_missing:
            st.info(
                "Complete the following before processing: "
                + "; ".join(dict.fromkeys(local_missing))
            )
        _show_result("suggested_result", "Download suggested reports ZIP")


def auto_pricing_page() -> None:
    section_header(
        "Sales, actions & price intelligence",
        "Auto Pricing",
        "Upload the latest files and the available previous files. The app updates the data, removes excluded parties, and creates the final Auto Pricing output.",
        ["Add latest files", "Add previous files", "Review selected files", "Create & download"],
    )

    tab_full, tab_prepared = st.tabs([
        "Create Latest Output",
        "Use Ready Files",
    ])

    with tab_full:
        st.caption(
            "Stock and Sales are required. View, Hold, Offer, Wishlist and Cart are optional. "
            "The two previous Sales files are required; previous action files are optional."
        )

        st.markdown("#### 1. Latest files")
        raw_uploads = {}
        raw_specs = [
            ("Stock", "Stock file *", True),
            ("Sales", "Sales file *", True),
            ("View", "View file (optional)", False),
            ("Hold", "Hold file (optional)", False),
            ("Offer", "Offer file (optional)", False),
            ("Wishlist", "Wishlist file (optional)", False),
            ("Cart", "Cart file (optional)", False),
        ]
        raw_columns = st.columns(3)
        for index, (dataset, label, _required) in enumerate(raw_specs):
            with raw_columns[index % 3]:
                raw_uploads[dataset] = st.file_uploader(
                    label,
                    type=["xlsx"],
                    key=f"auto_pricing_full_raw_{dataset.lower().replace(' ', '_')}",
                )

        st.markdown("#### 2. Previous files")
        st.caption(
            "Use the previous Sales files to keep the required sales period. Action files can be left blank. "
            "When a previous action file and the latest action file share the previous file's last date, that whole date is refreshed from the latest file."
        )

        previous_uploads = {}
        left, right = st.columns(2)
        with left:
            previous_uploads["Sales 180"] = st.file_uploader(
                "Previous 180-day Sales file *", type=["xlsx"], key="auto_pricing_full_previous_sales_180"
            )
        with right:
            previous_uploads["Sales 120"] = st.file_uploader(
                "Previous 120-day Sales file *", type=["xlsx"], key="auto_pricing_full_previous_sales_120"
            )

        st.markdown("##### Previous action files (optional)")
        history_action_specs = [
            ("View", "View file (optional)"),
            ("Hold", "Hold file (optional)"),
            ("Offer", "Offer file (optional)"),
            ("Wishlist", "Wishlist file (optional)"),
            ("Cart", "Cart file (optional)"),
        ]
        history_columns = st.columns(3)
        for index, (dataset, label) in enumerate(history_action_specs):
            with history_columns[index % 3]:
                previous_uploads[dataset] = st.file_uploader(
                    label,
                    type=["xlsx"],
                    key=f"auto_pricing_full_previous_{dataset.lower()}",
                )

        default_remove_bytes = default_remove_party_bytes()

        with st.expander("Use a different processing date (optional)", expanded=False):
            st.caption("Normally leave this blank. Use it only when you want to recreate an older day's result.")
            test_date = st.text_input(
                "Processing date",
                value="",
                placeholder="2026-08-28",
                key="auto_pricing_full_test_date",
            ).strip()

        st.markdown("#### 3. Files ready for this run")
        required_ready = sum(
            1 for item in [
                raw_uploads.get("Stock"),
                raw_uploads.get("Sales"),
                previous_uploads.get("Sales 180"),
                previous_uploads.get("Sales 120"),
            ] if item is not None
        )
        optional_selected = sum(
            1 for dataset in ["View", "Hold", "Offer", "Wishlist", "Cart"]
            if raw_uploads.get(dataset) is not None
        ) + sum(
            1 for dataset in ["View", "Hold", "Offer", "Wishlist", "Cart"]
            if previous_uploads.get(dataset) is not None
        )
        s1, s2, s3 = st.columns(3)
        s1.metric("Required files", f"{required_ready}/4 ready")
        s2.metric("Optional action files", f"{optional_selected} selected")
        s3.metric("Removed parties list", "Ready" if default_remove_bytes is not None else "Needs attention")

        file_rows = []
        for dataset, label, required in raw_specs:
            uploaded = raw_uploads.get(dataset)
            file_rows.append({
                "Group": "Latest",
                "File": label.replace(" *", "").replace(" (optional)", ""),
                "Selected": uploaded.name if uploaded is not None else "—",
                "Status": "Ready" if uploaded is not None else ("Required" if required else "Optional — not selected"),
            })
        previous_specs = [
            ("Sales 180", "Previous 180-day Sales", True),
            ("Sales 120", "Previous 120-day Sales", True),
            ("View", "Previous View", False),
            ("Hold", "Previous Hold", False),
            ("Offer", "Previous Offer", False),
            ("Wishlist", "Previous Wishlist", False),
            ("Cart", "Previous Cart", False),
        ]
        for dataset, label, required in previous_specs:
            uploaded = previous_uploads.get(dataset)
            file_rows.append({
                "Group": "Previous",
                "File": label,
                "Selected": uploaded.name if uploaded is not None else "—",
                "Status": "Ready" if uploaded is not None else ("Required" if required else "Optional — not selected"),
            })
        file_rows.append({
            "Group": "Saved",
            "File": "Removed parties list",
            "Selected": "Loaded automatically" if default_remove_bytes is not None else "—",
            "Status": "Ready" if default_remove_bytes is not None else "Needs attention",
        })
        st.dataframe(file_rows, hide_index=True, use_container_width=True)

        missing = []
        if raw_uploads.get("Stock") is None:
            missing.append("Stock file")
        if raw_uploads.get("Sales") is None:
            missing.append("Sales file")
        if previous_uploads.get("Sales 180") is None:
            missing.append("Previous 180-day Sales file")
        if previous_uploads.get("Sales 120") is None:
            missing.append("Previous 120-day Sales file")
        if default_remove_bytes is None:
            missing.append("Removed parties list")

        if st.button(
            "Create Auto Pricing",
            type="primary",
            use_container_width=True,
            disabled=bool(missing),
            key="run_auto_pricing_full_pipeline",
        ):
            try:
                if test_date and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", test_date):
                    raise ValueError("Processing date must be YYYY-MM-DD, or left blank.")
                with st.spinner("Preparing the latest files and creating Auto Pricing..."):
                    data, filename, logs = run_auto_pricing_full_pipeline_from_files(
                        raw_files={
                            name: uploaded.getvalue() if uploaded is not None else None
                            for name, uploaded in raw_uploads.items()
                        },
                        previous_base_files={
                            name: uploaded.getvalue() if uploaded is not None else None
                            for name, uploaded in previous_uploads.items()
                        },
                        remove_party_bytes=None,
                        test_date=test_date or None,
                    )
                    _save_result("auto_pricing_full_result", data, filename, logs)
                st.rerun()
            except Exception as exc:
                _show_user_error("Unable to complete Auto Pricing", exc)

        if missing:
            st.info("Add these required files before continuing: " + "; ".join(missing))

        result = st.session_state.get("auto_pricing_full_result")
        if result:
            try:
                with zipfile.ZipFile(io.BytesIO(result["data"])) as zf:
                    created = [n for n in zf.namelist() if not n.endswith("/")]
                    run_summary = json.loads(zf.read("Run_Summary.json")) if "Run_Summary.json" in zf.namelist() else {}

                st.markdown("#### 4. Files used and results")
                summary_lookup = {str(row.get("Dataset")): row for row in (run_summary.get("summary") or [])}
                result_rows = []
                for dataset in ["Stock", "Sales 120", "Sales 180", "View", "Hold", "Offer", "Wishlist", "Cart"]:
                    if dataset == "Stock":
                        latest = raw_uploads.get("Stock")
                        previous = None
                    elif dataset.startswith("Sales"):
                        latest = raw_uploads.get("Sales")
                        previous = previous_uploads.get(dataset)
                    else:
                        latest = raw_uploads.get(dataset)
                        previous = previous_uploads.get(dataset)
                    report_row = summary_lookup.get(dataset, {})
                    if not latest and not previous and dataset in {"View", "Hold", "Offer", "Wishlist", "Cart"}:
                        work_status = "Not selected"
                    else:
                        work_status = "Completed" if str(report_row.get("Status", "PASS")).upper() == "PASS" else "Check needed"
                    result_rows.append({
                        "Data": dataset,
                        "Latest file": latest.name if latest is not None else "—",
                        "Previous file": previous.name if previous is not None else ("Not needed" if dataset == "Stock" else "—"),
                        "Rows created": report_row.get("Output Rows", 0),
                        "Result": work_status,
                    })
                st.dataframe(result_rows, hide_index=True, use_container_width=True)

                st.markdown("#### 5. Files created")
                created_rows = []
                for name in created:
                    display_name = Path(name).name
                    if display_name == "Run_Summary.json":
                        label = "Run summary"
                    elif display_name == "Processing_Check.xlsx":
                        label = "Processing check"
                    elif display_name.startswith("Updated_Files_"):
                        label = "Updated files"
                    elif display_name.startswith("Auto_Pricing_") and display_name.lower().endswith(".xlsx"):
                        label = "Final Auto Pricing workbook"
                    elif name.startswith("base_files/"):
                        label = "Updated " + display_name.replace(".xlsx", "")
                    else:
                        label = display_name
                    created_rows.append({"Output": label, "File": display_name, "Result": "Created"})
                st.dataframe(created_rows, hide_index=True, use_container_width=True)
            except Exception:
                pass

        _show_result("auto_pricing_full_result", "Download Auto Pricing files")

        with st.expander("How the files are updated", expanded=False):
            st.markdown(
                """
- If the last date in a previous action file is also present in the latest action file, the old rows for that date are replaced with the complete latest rows for that date.
- Hold always uses `SERIAL_NO` as the stone number.
- View keeps one uploaded row as one row. Blank or zero stone numbers are removed.
- Existing Sales rows are matched by `Serial No + Sale Date`; the latest Sales file updates `SPer` and `SRate`, and new Sales rows are added.
- `Amt` is calculated as `SRate × Carat`.
- Blank Lab or Flo values become `NON`.
- Parties in the saved removal list are excluded automatically.
- `SALE_PERSON` is kept as it is and is not used to decide whether a row should be replaced.
- Repeated activity rows are kept for now unless they are removed in a separately approved duplicate-cleaning step.
                """
            )

    with tab_prepared:
        st.caption(
            "Use this option when your Stock, 120-day Sales and 180-day Sales files are already prepared. "
            "Action files are optional."
        )

        left, middle, right = st.columns(3)
        with left:
            stock = st.file_uploader("Stock file *", type=["xlsx"], key="auto_pricing_stock")
        with middle:
            sales_180 = st.file_uploader("180-day Sales file *", type=["xlsx"], key="auto_pricing_sales_180")
        with right:
            sales_120 = st.file_uploader("120-day Sales file *", type=["xlsx"], key="auto_pricing_sales_120")

        st.markdown("#### Action files (optional)")
        action_uploads = {}
        columns = st.columns(3)
        for index, action_name in enumerate(AUTO_PRICING_ACTION_NAMES):
            with columns[index % 3]:
                action_uploads[action_name] = st.file_uploader(
                    f"{action_name} file (optional)",
                    type=["xlsx"],
                    key=f"auto_pricing_action_{action_name.lower()}",
                )

        ready_rows = []
        for label, uploaded, required in [
            ("Stock", stock, True),
            ("180-day Sales", sales_180, True),
            ("120-day Sales", sales_120, True),
        ]:
            ready_rows.append({
                "File": label,
                "Selected": uploaded.name if uploaded is not None else "—",
                "Status": "Ready" if uploaded is not None else "Required",
            })
        for action_name, uploaded in action_uploads.items():
            ready_rows.append({
                "File": action_name,
                "Selected": uploaded.name if uploaded is not None else "—",
                "Status": "Ready" if uploaded is not None else "Optional — not selected",
            })
        st.markdown("#### Files ready for this run")
        st.dataframe(ready_rows, hide_index=True, use_container_width=True)

        prepared_missing = []
        if stock is None:
            prepared_missing.append("Stock file")
        if sales_180 is None:
            prepared_missing.append("180-day Sales file")
        if sales_120 is None:
            prepared_missing.append("120-day Sales file")

        if st.button(
            "Create Auto Pricing from Ready Files",
            type="primary",
            use_container_width=True,
            disabled=bool(prepared_missing),
            key="run_auto_pricing",
        ):
            try:
                with st.spinner("Creating Auto Pricing..."):
                    data, filename, logs = run_auto_pricing(
                        stock_bytes=stock.getvalue(),
                        sales_180_bytes=sales_180.getvalue(),
                        sales_120_bytes=sales_120.getvalue(),
                        action_files={
                            name: uploaded.getvalue() if uploaded is not None else None
                            for name, uploaded in action_uploads.items()
                        },
                    )
                    _save_result("auto_pricing_result", data, filename, logs)
                st.rerun()
            except Exception as exc:
                _show_user_error("Unable to complete Auto Pricing", exc)

        if prepared_missing:
            st.info("Add these required files before continuing: " + "; ".join(prepared_missing))
        _show_result("auto_pricing_result", "Download Auto Pricing files")

def dealmaker_page() -> None:
    section_header(
        "Production deal allocation",
        "Dealmaker",
        "Run the supplied Dealmaker v8.2.2 production package end-to-end: preflight, preparation, stock profiling, sequential allocation, finalization, deal books, and release-gate audit.",
        ["Upload Sales / Actions / Stock", "Use bundled or updated control", "Run production stages", "Download review ZIP"],
    )

    st.caption(
        "Default behavior uses the bundled Dealmaker_Full_Control_v6.xlsx from the supplied v8.2.2 package. "
        "If the base/control workbook changes, upload the updated Dealmaker_Full_Control_v6.xlsx here for that run. "
        "The bundled file is never overwritten."
    )

    left, middle, right = st.columns(3)
    with left:
        sales = st.file_uploader("Full Sales / BID workbook *", type=["xlsx"], key="dealmaker_sales")
    with middle:
        actions = st.file_uploader("Full Actions workbook *", type=["xlsx"], key="dealmaker_actions")
    with right:
        stock = st.file_uploader("Actual current-stock workbook *", type=["xlsx"], key="dealmaker_stock")

    base_col, upload_col = st.columns(2)
    with base_col:
        st.download_button(
            "Download bundled Dealmaker_Full_Control_v6.xlsx",
            data=bundled_dealmaker_control_bytes(),
            file_name="Dealmaker_Full_Control_v6.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="download_dealmaker_bundled_control",
            on_click="ignore",
        )
    with upload_col:
        control = st.file_uploader(
            "Dealmaker_Full_Control_v6.xlsx (optional updated base/control override)",
            type=["xlsx"],
            key="dealmaker_control",
            help="Leave blank to use the verified control workbook bundled with Dealmaker v8.2.2.",
        )
    if control is None:
        st.info("Control source for this run: bundled Dealmaker_Full_Control_v6.xlsx")
    else:
        st.success("Control source for this run: uploaded control override (bundled base remains unchanged)")

    with st.expander("Advanced input detection and production settings", expanded=False):
        st.caption("Leave sheet names blank and header rows on Auto unless preflight cannot detect the source correctly.")
        chunk_size = st.number_input(
            "Parties per allocation chunk",
            min_value=1,
            max_value=100,
            value=10,
            step=1,
            key="dealmaker_chunk_size",
        )
        sheet_cols = st.columns(3)
        sales_sheet = sheet_cols[0].text_input("Sales sheet name (optional)", key="dealmaker_sales_sheet")
        actions_sheet = sheet_cols[1].text_input("Actions sheet name (optional)", key="dealmaker_actions_sheet")
        stock_sheet = sheet_cols[2].text_input("Stock sheet name (optional)", key="dealmaker_stock_sheet")
        header_cols = st.columns(3)
        sales_header_auto = header_cols[0].checkbox("Auto-detect Sales header row", value=True, key="dealmaker_sales_header_auto")
        actions_header_auto = header_cols[1].checkbox("Auto-detect Actions header row", value=True, key="dealmaker_actions_header_auto")
        stock_header_auto = header_cols[2].checkbox("Auto-detect Stock header row", value=True, key="dealmaker_stock_header_auto")
        sales_header = None if sales_header_auto else int(header_cols[0].number_input("Sales header row", min_value=1, value=1, step=1, key="dealmaker_sales_header_row"))
        actions_header = None if actions_header_auto else int(header_cols[1].number_input("Actions header row", min_value=1, value=1, step=1, key="dealmaker_actions_header_row"))
        stock_header = None if stock_header_auto else int(header_cols[2].number_input("Stock header row", min_value=1, value=1, step=1, key="dealmaker_stock_header_row"))

    missing = []
    if sales is None:
        missing.append("Sales / BID workbook")
    if actions is None:
        missing.append("Actions workbook")
    if stock is None:
        missing.append("current-stock workbook")

    if st.button(
        "Run Dealmaker production package",
        type="primary",
        use_container_width=True,
        disabled=bool(missing),
        key="run_dealmaker",
    ):
        try:
            with st.spinner("Running Dealmaker preflight, allocation, finalization, deal books, and release audit..."):
                data, filename, logs = run_dealmaker(
                    sales_bytes=sales.getvalue(),
                    actions_bytes=actions.getvalue(),
                    stock_bytes=stock.getvalue(),
                    control_bytes=control.getvalue() if control else None,
                    chunk_size=int(chunk_size),
                    sales_sheet=sales_sheet,
                    actions_sheet=actions_sheet,
                    stock_sheet=stock_sheet,
                    sales_header_row=sales_header,
                    actions_header_row=actions_header,
                    stock_header_row=stock_header,
                )
                _save_result("dealmaker_result", data, filename, logs)
            st.rerun()
        except Exception as exc:
            _show_user_error("Unable to complete the Dealmaker production run", exc)

    if missing:
        st.info("Complete the following before processing: " + "; ".join(missing))
    _show_result("dealmaker_result", "Download Dealmaker production review ZIP")


def layout_stock_matcher_page() -> None:
    section_header(
        "Layout-driven stock suggestions",
        "Layout Stock Matcher",
        "Create salesperson-wise layout suggestion workbooks from Layout Stock, Sales, and Actions using the preserved exact/tolerance matching engine.",
        ["Upload Layout / Sales / Actions", "Keep or adjust source controls", "Generate salesperson workbooks", "Download ZIP"],
    )

    st.caption(
        "Default source behavior: Shape + revised Carat group + Lab are exact keys; Color, Clarity, and Fluorescence use tolerance 1. "
        "Output retains the source's Exact / Nearby / Multi Nearby classification and workbook formatting."
    )

    left, middle, right = st.columns(3)
    with left:
        layout = st.file_uploader("Layout Stock workbook *", type=["xlsx"], key="layout_matcher_layout")
    with middle:
        sales = st.file_uploader("Sales workbook *", type=["xlsx"], key="layout_matcher_sales")
    with right:
        actions = st.file_uploader("Actions workbook *", type=["xlsx"], key="layout_matcher_actions")

    with st.expander("Layout matching controls", expanded=False):
        c1, c2 = st.columns(2)
        tolerance = c1.selectbox(
            "Color / Clarity / Fluorescence tolerance",
            options=[0, 1],
            index=1,
            format_func=lambda value: "1 - allow one rank" if value == 1 else "0 - exact only",
            key="layout_matcher_tolerance",
        )
        min_layout_pct = c2.number_input(
            "Minimum layout percentage",
            min_value=0.0,
            max_value=1.0,
            value=0.10,
            step=0.01,
            format="%.2f",
            key="layout_matcher_min_layout_pct",
        )
        c3, c4 = st.columns(2)
        min_absolute_stones = c3.number_input(
            "Minimum absolute stones",
            min_value=1,
            value=5,
            step=1,
            key="layout_matcher_min_absolute_stones",
        )
        min_action_weight = c4.number_input(
            "Minimum action weight",
            min_value=0,
            value=45,
            step=5,
            key="layout_matcher_min_action_weight",
        )
        include_match_columns = st.checkbox(
            "Include match audit columns",
            value=True,
            key="layout_matcher_include_match_columns",
        )
        create_empty_salesperson_files = st.checkbox(
            "Create empty salesperson files when no suggestions qualify",
            value=True,
            key="layout_matcher_create_empty",
        )

    missing = []
    if layout is None:
        missing.append("Layout Stock workbook")
    if sales is None:
        missing.append("Sales workbook")
    if actions is None:
        missing.append("Actions workbook")

    if st.button(
        "Run Layout Stock Matcher",
        type="primary",
        use_container_width=True,
        disabled=bool(missing),
        key="run_layout_stock_matcher",
    ):
        try:
            with st.spinner("Building party suggestions and salesperson layout workbooks..."):
                data, filename, logs = run_layout_stock_matcher(
                    layout_bytes=layout.getvalue(),
                    sales_bytes=sales.getvalue(),
                    actions_bytes=actions.getvalue(),
                    tolerance=int(tolerance),
                    min_layout_pct=float(min_layout_pct),
                    min_absolute_stones=int(min_absolute_stones),
                    min_action_weight=int(min_action_weight),
                    include_match_columns=bool(include_match_columns),
                    create_empty_salesperson_files=bool(create_empty_salesperson_files),
                )
                _save_result("layout_stock_matcher_result", data, filename, logs)
            st.rerun()
        except Exception as exc:
            _show_user_error("Unable to complete Layout Stock Matcher", exc)

    if missing:
        st.info("Complete the following before processing: " + "; ".join(missing))
    _show_result("layout_stock_matcher_result", "Download Layout Stock Matcher ZIP")


def demand_filter_page() -> None:
    section_header(
        "Rule-based stock demand",
        "Demand Filter Mapper",
        "Apply your demand rules to Stock Actual and the Filters workbook. Reports are generated for each salesperson.",
        ["Upload stock", "Upload filters", "Set tolerance", "Download reports"],
    )
    left, right = st.columns(2)
    stock = left.file_uploader("Stock Actual workbook *", type=["xlsx"], key="demand_stock")
    filters = right.file_uploader("Filters workbook *", type=["xlsx"], key="demand_filters")

    with st.expander("Report settings", expanded=True):
        c1, c2, c3 = st.columns(3)
        tolerance = c1.number_input("Tolerance level", min_value=0, max_value=3, value=1, step=1)
        flo_mapping = c2.toggle("Group VSL / SL with FNT", value=False)
        related_tag = c3.toggle("Include related Tag_No", value=True)

    if st.button("Generate demand-filter reports", type="primary", use_container_width=True, disabled=stock is None or filters is None):
        try:
            with st.spinner("Applying demand rules and generating salesperson workbooks..."):
                data, filename, logs = run_filter_demand(
                    stock_bytes=stock.getvalue(),
                    filter_bytes=filters.getvalue(),
                    runtime_overrides={
                        "TOLERANCE_LEVEL": int(tolerance),
                        "ENABLE_FLO_MAPPING": bool(flo_mapping),
                        "ENABLE_RELATED_TAG": bool(related_tag),
                    },
                )
                _save_result("demand_result", data, filename, logs)
            st.rerun()
        except Exception as exc:
            _show_user_error("Unable to complete this action", exc)
    _show_result("demand_result", "Download demand-filter reports ZIP")


def sorting_page() -> None:
    section_header(
        "Suggested-output prioritization",
        "Sort Suggested Output",
        "Prioritize suggested output using your selected match rules, limits, history, and scoring choices.",
        ["Upload suggested ZIP", "Choose match rules", "Configure scoring", "Download sorted ZIP"],
    )
    input_zip = st.file_uploader("Suggested output ZIP *", type=["zip"], key="sort_zip")

    tab_main, tab_priority, tab_detection = st.tabs(
        ["Filtering & limits", "Priority & history", "Match detection & weights"]
    )
    runtime: dict[str, object] = {}

    with tab_main:
        c1, c2, c3 = st.columns(3)
        runtime["ONLY_MATCHED_TYPE"] = c1.toggle("Only selected match types", value=True)
        runtime["COUNT_THRESHOLD"] = int(
            c2.number_input("Navigation count threshold", min_value=0, value=200, step=10)
        )
        runtime["NORMAL_PARTY_LIMIT"] = int(
            c3.number_input(
                "Normal-party limit",
                min_value=1,
                value=100,
                step=10,
                disabled=runtime["ONLY_MATCHED_TYPE"],
            )
        )
        allowed = st.multiselect(
            "Allowed match types",
            ["Exact", "Near by", "May be", "Filter"],
            default=["Exact", "Near by", "Filter"],
            disabled=not runtime["ONLY_MATCHED_TYPE"],
        )
        runtime["ALLOWED_MATCH_TYPES"] = set(allowed)

        d1, d2, d3, d4 = st.columns(4)
        runtime["HIGH_PRIORITY_PARTY_LIMIT"] = int(
            d1.number_input(
                "High-priority limit",
                min_value=1,
                value=150,
                step=10,
                disabled=runtime["ONLY_MATCHED_TYPE"],
            )
        )
        runtime["EXACT_CAP"] = int(
            d2.number_input("Exact family cap", min_value=0, value=3, step=1, disabled=runtime["ONLY_MATCHED_TYPE"])
        )
        runtime["NEAR_CAP"] = int(
            d3.number_input("Nearby family cap", min_value=0, value=2, step=1, disabled=runtime["ONLY_MATCHED_TYPE"])
        )
        runtime["MAYBE_CAP"] = int(
            d4.number_input("Maybe family cap", min_value=0, value=1, step=1, disabled=runtime["ONLY_MATCHED_TYPE"])
        )

    with tab_priority:
        priority_text = st.text_input(
            "High-priority party types",
            value=", ".join(DEFAULT_PRIORITY_TYPES),
            disabled=runtime["ONLY_MATCHED_TYPE"],
        )
        runtime["HIGH_PRIORITY_PARTY_TYPES"] = set(
            value.upper() for value in (parse_delimited_values(priority_text) or DEFAULT_PRIORITY_TYPES)
        )
        runtime["HIGH_PRIORITY_TOTAL_HISTORY_THRESHOLD"] = int(
            st.number_input(
                "High-priority total-history threshold",
                min_value=0,
                value=200,
                step=10,
                disabled=runtime["ONLY_MATCHED_TYPE"],
            )
        )
        st.caption(
            "These settings are used only when 'Only selected match types' is off and the sales-history scoring path is active."
        )

    with tab_detection:
        st.markdown("#### Worksheet fill colors used to classify rows")
        e1, e2, e3 = st.columns(3)
        green_text = e1.text_area("Exact / green RGB values", value="FFCCFFCC, 00CCFFCC")
        orange_text = e2.text_area("Nearby / orange RGB values", value="FFFFE699, 00FFE699")
        filter_text = e3.text_area(
            "Filter / yellow RGB values",
            value="FFFFF2CC, 00FFF2CC, FFFFFFCC, 00FFFFCC, FFFFFF99, 00FFFF99",
        )
        runtime["GREEN_RGB_VALUES"] = set(value.upper() for value in parse_delimited_values(green_text))
        runtime["ORANGE_RGB_VALUES"] = set(value.upper() for value in parse_delimited_values(orange_text))
        runtime["FILTER_RGB_VALUES"] = set(value.upper() for value in parse_delimited_values(filter_text))

        pair_words_text = st.text_input("Pair-detection words", value="PAIR, PAIRED, MATCH PAIR")
        runtime["PAIR_WORDS"] = tuple(parse_delimited_values(pair_words_text) or ["PAIR", "PAIRED", "MATCH PAIR"])

        default_weights = {
            "SALE": 20,
            "BID": 15,
            "OFFER": 10,
            "WISHLIST": 10,
            "CART": 10,
            "HOLD": 7,
            "INQUIRY": 2,
            "ENQUIRY": 2,
            "LIST": 2,
            "VIEW": 4,
            "SEARCH": 3,
            "DETAILS": 5,
        }
        weights_text = st.text_area(
            "Sales/action type weights",
            value=format_type_weights(default_weights),
            disabled=runtime["ONLY_MATCHED_TYPE"],
            help="Use NAME=WEIGHT entries separated by commas or new lines.",
        )
        runtime["TYPE_WEIGHTS"] = parse_type_weights(weights_text) or default_weights
        runtime["DEFAULT_TYPE_WEIGHT"] = int(
            st.number_input(
                "Default history-event weight",
                min_value=0,
                value=1,
                step=1,
                disabled=runtime["ONLY_MATCHED_TYPE"],
            )
        )

    sales_file = st.file_uploader(
        "Sales history workbook (required when 'Only selected match types' is off)",
        type=["xlsx"],
        key="sort_sales",
        disabled=runtime["ONLY_MATCHED_TYPE"],
    )

    missing: list[str] = []
    if input_zip is None:
        missing.append("Suggested output ZIP")
    if runtime["ONLY_MATCHED_TYPE"] and not runtime["ALLOWED_MATCH_TYPES"]:
        missing.append("at least one allowed match type")
    if not runtime["ONLY_MATCHED_TYPE"] and sales_file is None:
        missing.append("Sales history workbook")
    if not runtime["GREEN_RGB_VALUES"] or not runtime["ORANGE_RGB_VALUES"] or not runtime["FILTER_RGB_VALUES"]:
        missing.append("RGB match-detection values")

    if st.button("Sort suggested output", type="primary", use_container_width=True, disabled=bool(missing)):
        try:
            with st.spinner("Filtering and rewriting party sheets while preserving workbook order..."):
                data, filename, logs = run_sorting_process(
                    input_zip_bytes=input_zip.getvalue(),
                    sales_file_bytes=sales_file.getvalue() if sales_file else None,
                    runtime_overrides=runtime,
                )
                _save_result("sort_result", data, filename, logs)
            st.rerun()
        except Exception as exc:
            _show_user_error("Unable to complete this action", exc)
    if missing:
        st.info("Complete the following before processing: " + "; ".join(missing))
    _show_result("sort_result", "Download sorted suggested ZIP")

def usa_suggested_stock_mapper_page() -> None:
    section_header(
        "USA sales-history recommendation",
        "USA Suggested Stock Mapper",
        "Generate suggested stock files for the USA team from Stock Actual and one or more USA sales-history workbooks.",
        ["Upload USA inputs", "Build top Super Tags", "Match exact + nearby stock", "Download salesperson ZIP"],
    )

    st.caption(
        "This USA workflow uses each party's historical Shape / Carat Group / Color / Clarity / "
        "Fluorescence / Lab combinations. Exact matches are included first, followed by adjacent "
        "Color, Clarity, and Fluorescence suggestions without repeating Serial No values."
    )

    left, right = st.columns(2)
    with left:
        stock = st.file_uploader(
            "USA Stock Actual workbook *", type=["xlsx"], key="usa_suggest_stock"
        )
    with right:
        sales_history = st.file_uploader(
            "USA sales history workbook(s) *",
            type=["xlsx"],
            accept_multiple_files=True,
            key="usa_suggest_sales",
        )

    top_n = st.number_input(
        "Top historical Super Tags per party",
        min_value=1,
        max_value=5000,
        value=251,
        step=25,
        key="usa_suggest_top_n",
        help="The supplied USA mapper uses 251 by default.",
    )

    missing: list[str] = []
    if stock is None:
        missing.append("USA Stock Actual workbook")
    if not sales_history:
        missing.append("at least one USA sales history workbook")

    if st.button(
        "Generate USA suggested stock",
        type="primary",
        use_container_width=True,
        disabled=bool(missing),
        key="run_usa_suggested_stock",
    ):
        try:
            with st.spinner("Matching USA sales history to stock and building party reports..."):
                data, filename, logs = run_usa_suggested_stock_mapper(
                    stock_bytes=stock.getvalue(),
                    sales_history_files=[
                        (file.name, file.getvalue()) for file in (sales_history or [])
                    ],
                    top_n_supertags=int(top_n),
                )
                _save_result("usa_suggested_result", data, filename, logs)
            st.rerun()
        except Exception as exc:
            _show_user_error("Unable to generate USA suggested stock", exc)

    if missing:
        st.info("Complete the following before processing: " + "; ".join(missing))
    _show_result("usa_suggested_result", "Download USA suggested stock ZIP")


def suggested_list_splitter_page() -> None:
    section_header(
        "Stone-count distribution",
        "Suggested List Splitter",
        "Split suggested Excel files for each salesperson into Under 10 Stones and 10+ Stones packages.",
        ["Upload suggested ZIP", "Count Serial No rows", "Split by stone count", "Download final ZIP"],
    )
    st.caption("Boundary rule: 0-9 stones -> Under 10 Stones; exactly 10 and above -> 10+ Stones.")
    input_zip = st.file_uploader(
        "Suggested list ZIP *", type=["zip"], key="suggested_list_split_zip"
    )
    if st.button(
        "Split suggested lists by stone count",
        type="primary",
        use_container_width=True,
        disabled=input_zip is None,
    ):
        try:
            with st.spinner("Counting stones and packaging salesperson lists..."):
                data, filename, logs = run_suggested_list_splitter(
                    input_zip_bytes=input_zip.getvalue()
                )
                _save_result("suggested_list_split_result", data, filename, logs)
            st.rerun()
        except Exception as exc:
            _show_user_error("Unable to complete this action", exc)
    _show_result(
        "suggested_list_split_result",
        "Download split suggested-list ZIP",
    )


def bid_drop_page() -> None:
    section_header(
        "Bid conversion diagnostics",
        "BID Drop Checking Process",
        "Explain why bid stones were suggested, blocked, out of stock, or missed by matching and capacity rules.",
        ["Upload reports", "Use default rules", "Run diagnostics", "Download XLSX"],
    )

    left, right = st.columns(2)
    with left:
        suggested = st.file_uploader(
            "MASTER_CONSOLIDATED_REPORT workbook *", type=["xlsx"], key="bid_drop_suggested"
        )
        bids = st.file_uploader("BID workbook *", type=["xlsx"], key="bid_drop_bids")
        sales = st.file_uploader("Sales & Actions workbook *", type=["xlsx"], key="bid_drop_sales")
    with right:
        stock = st.file_uploader("Stock Actual workbook *", type=["xlsx"], key="bid_drop_stock")
        filters = st.file_uploader(
            "Filters workbook (optional override)", type=["xlsx"], key="bid_drop_filters"
        )
        no_need = st.file_uploader(
            "No-need workbook (optional override)", type=["xlsx"], key="bid_drop_no_need"
        )

    st.caption(
        "When omitted, Filters uses data/filters - EBID.xlsx and No-need uses data/no_need.xlsx."
    )

    missing = []
    if suggested is None:
        missing.append("MASTER_CONSOLIDATED_REPORT")
    if bids is None:
        missing.append("BID workbook")
    if sales is None:
        missing.append("Sales & Actions workbook")
    if stock is None:
        missing.append("Stock Actual workbook")
    if filters is None and not _repository_default_exists("filters"):
        missing.append("Filters upload or data/filters - EBID.xlsx")
    if no_need is None and not _repository_default_exists("no_need"):
        missing.append("No-need upload or data/no_need.xlsx")

    if st.button(
        "Run BID drop checking",
        type="primary",
        use_container_width=True,
        disabled=bool(missing),
    ):
        try:
            with st.spinner("Checking bid conversion gaps and building the diagnostic workbook..."):
                data, filename, logs = run_bid_drop_checking(
                    suggested_bytes=suggested.getvalue(),
                    bids_bytes=bids.getvalue(),
                    sales_bytes=sales.getvalue(),
                    stock_bytes=stock.getvalue(),
                    filter_bytes=filters.getvalue() if filters else None,
                    no_need_bytes=no_need.getvalue() if no_need else None,
                )
                _save_result(
                    "bid_drop_result",
                    data,
                    filename,
                    logs,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.rerun()
        except Exception as exc:
            _show_user_error("Unable to complete this action", exc)

    if missing:
        st.info("Complete the following before processing: " + "; ".join(missing))
    _show_result("bid_drop_result", "Download BID drop analysis XLSX")


def separation_page() -> None:
    section_header(
        "Workbook distribution",
        "Separate Excel Sheets",
        "Split each worksheet from every Excel workbook inside an uploaded ZIP. Each uploaded workbook becomes its own ZIP of separate sheet files.",
        ["Upload master ZIP", "Find workbooks", "Separate sheets", "Download final ZIP"],
    )
    input_zip = st.file_uploader("ZIP containing Excel workbooks *", type=["zip"], key="separate_zip")
    if st.button("Separate all workbook sheets", type="primary", use_container_width=True, disabled=input_zip is None):
        try:
            with st.spinner("Separating workbook sheets and building nested ZIP packages..."):
                data, filename, logs = run_separate_excel_files(input_zip_bytes=input_zip.getvalue())
                _save_result("separate_result", data, filename, logs)
            st.rerun()
        except Exception as exc:
            _show_user_error("Unable to complete this action", exc)
    _show_result("separate_result", "Download separated workbook ZIP")


if page == "Overview":
    overview_page()
elif page == "Suggested Stock Mapper":
    suggested_page()
elif page == "USA Suggested Stock Mapper":
    usa_suggested_stock_mapper_page()
elif page == "Auto Pricing":
    auto_pricing_page()
elif page == "Dealmaker":
    dealmaker_page()
elif page == "Layout Stock Matcher":
    layout_stock_matcher_page()
elif page == "Demand Filter Mapper":
    demand_filter_page()
elif page == "Sort Suggested Output":
    sorting_page()
elif page == "Suggested List Splitter":
    suggested_list_splitter_page()
elif page == "BID Drop Checking Process":
    bid_drop_page()
elif page == "Separate Excel Sheets":
    separation_page()

page_footer()
