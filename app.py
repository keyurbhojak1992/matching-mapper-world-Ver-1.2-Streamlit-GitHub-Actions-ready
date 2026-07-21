from __future__ import annotations

import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from adapters.filter_demand_adapter import run_filter_demand
from adapters.separate_adapter import run_separate_excel_files
from adapters.sort_adapter import run_sorting_process
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
st.sidebar.caption(f"Application version {APP_VERSION}")

PAGES = [
    "Overview",
    "Suggested Stock Mapper",
    "Demand Filter Mapper",
    "Sort Suggested Output",
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


def _save_result(key: str, data: bytes, filename: str, logs: str) -> None:
    st.session_state[key] = {"data": data, "filename": filename, "logs": logs}


def _show_result(key: str, label: str) -> None:
    result = st.session_state.get(key)
    if not result:
        return
    c1, c2, c3 = st.columns(3)
    c1.metric("Output status", "Ready")
    c2.metric("File size", f"{len(result['data']) / (1024 * 1024):,.2f} MB")
    try:
        with zipfile.ZipFile(io.BytesIO(result["data"])) as zf:
            count = len([n for n in zf.namelist() if not n.endswith("/")])
    except Exception:
        count = 1
    c3.metric("Files packaged", f"{count:,}")
    st.success("Processing completed. Your output is ready to download.")
    st.download_button(
        label,
        data=result["data"],
        file_name=result["filename"],
        mime="application/zip",
        type="primary",
        use_container_width=True,
        key=f"download_{key}",
        on_click="ignore",
    )
    with st.expander("Processing log", expanded=False):
        st.code(result["logs"] or "Process completed.", language="text")


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

    st.markdown("#### Submitted full-run job")
    settings = _github_job_settings()
    if settings is None:
        st.error("GitHub background-job settings are no longer available in Streamlit secrets.")
        return

    backend = GitHubReleaseBackend(settings)
    snapshot = None
    try:
        snapshot = backend.job_snapshot(int(job["release_id"]))
    except Exception as exc:
        st.warning(f"Unable to refresh the private GitHub job right now: {exc}")

    status_payload = snapshot.get("status", {}) if snapshot else {}
    state = str(status_payload.get("state") or job.get("status", "submitted"))
    columns = st.columns(4)
    columns[0].metric("Job status", state.replace("_", " ").title())
    columns[1].metric("Batch size", str(job.get("batch_size", 3)))
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
    columns[3].metric("Batch progress", progress_label)
    st.code(str(job["tag_name"]), language="text")
    message = str(status_payload.get("message", "Waiting for the GitHub Actions runner."))
    st.caption(message)
    worker_version = status_payload.get("worker_version")
    if worker_version:
        st.caption(f"GitHub worker version: {worker_version}")
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
                with st.spinner("Downloading the private result to the Streamlit server..."):
                    backend.download_asset_to_path(asset, target)
                st.session_state["github_suggested_download"] = str(target)
                st.session_state.pop("github_suggested_download_url", None)
            st.rerun()
        except Exception as exc:
            st.exception(exc)

    confirm_delete = cleanup.checkbox("Allow cleanup", key="confirm_github_cleanup")
    if cleanup.button(
        "Delete job files",
        use_container_width=True,
        disabled=not confirm_delete,
        key="delete_github_job",
        help="Deletes the private draft release, result files, and temporary tag after you have downloaded the result.",
    ):
        try:
            backend.delete_job(int(job["release_id"]), tag_name=str(job.get("tag_name", "")))
            download_path = st.session_state.pop("github_suggested_download", None)
            st.session_state.pop("github_suggested_download_url", None)
            if download_path:
                Path(download_path).unlink(missing_ok=True)
            st.session_state.pop("github_suggested_job", None)
            st.success("Temporary GitHub job files were deleted.")
            st.rerun()
        except Exception as exc:
            st.exception(exc)

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
        st.error(
            "The background worker reported a failure. Open the repository's Actions page or inspect "
            "failure_log.txt and worker_console.txt in the private draft release."
        )
    elif state == "partial_success":
        st.warning("The result is ready, but at least one batch failed or was skipped. Review batch_manifest.json.")


def _show_github_recovery_panel(settings: GitHubCloudConfig | None) -> None:
    if settings is None:
        return
    with st.expander("Recover a previous full-run job", expanded=False):
        backend = GitHubReleaseBackend(settings)
        manual_tag = st.text_input(
            "Job tag",
            placeholder="mmw-job-YYYYMMDD-HHMMSS-xxxxxxxx",
            key="recover_github_tag",
        )
        a, b = st.columns(2)
        if a.button("Load job tag", use_container_width=True, disabled=not manual_tag.strip()):
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
                st.exception(exc)

        if b.button("Find recent jobs", use_container_width=True):
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
                st.exception(exc)

        recent = st.session_state.get("recent_github_jobs", [])
        if recent:
            labels = [f"{item['name']} | {item['tag_name']}" for item in recent]
            selected_label = st.selectbox("Recent private jobs", labels, key="recent_github_job_select")
            selected = recent[labels.index(selected_label)]
            if st.button("Load selected recent job", use_container_width=True):
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
    metrics[0].metric("Production workflows", "4")
    metrics[1].metric("Master scripts", "4 preserved")
    metrics[2].metric("Processing model", "Streamlit + GitHub worker")
    metrics[3].metric("Primary outputs", "XLSX / ZIP")

    st.markdown("### Workflow studio")
    cards = [
        ("PROFILE MATCHING", "Suggested Stock Mapper", "Output: Suggested report ZIP"),
        ("DEMAND RULES", "Demand Filter Mapper", "Output: Salesperson report ZIP"),
        ("PRIORITIZATION", "Sort Suggested Output", "Output: Sorted suggested ZIP"),
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
        "Run the complete preserved recommendation engine with every operational control from the original USER INPUT section available in the web interface.",
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
            profile_perm = st.file_uploader("Permanent party profile", type=["parquet"], key="suggest_perm")
            profile_temp = st.file_uploader("Temporary party profile", type=["parquet"], key="suggest_temp")
        with right:
            filters = st.file_uploader("Filters workbook *", type=["xlsx"], key="suggest_filters")
            no_need = st.file_uploader("No-need / exclusion workbook", type=["xlsx"], key="suggest_noneed")
            sales_history = st.file_uploader(
                "Sales history workbook(s)", type=["xlsx"], accept_multiple_files=True, key="suggest_sales"
            )
        st.caption(
            "Required: Stock Actual and Filters. At least one party-profile file is recommended. "
            "Uploaded files are passed to the unchanged master engine through temporary paths."
        )

        discovery_identity = (
            _uploaded_file_identity(profile_perm),
            _uploaded_file_identity(profile_temp),
            _uploaded_file_identity(filters),
        )
        scan_names = st.button(
            "Scan uploaded files for salesperson and party names (optional)",
            use_container_width=True,
            disabled=not any((profile_perm, profile_temp, filters)),
            help=(
                "Use this only when you need uploaded names in the manual selection lists. "
                "A full 'All salespeople' GitHub job discovers names again on the worker, so scanning here is not required."
            ),
        )
        if scan_names:
            with st.spinner("Reading only salesperson and party columns..."):
                found_salespeople, found_parties = discover_suggested_names_from_streams(
                    profile_perm,
                    profile_temp,
                    filters,
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
                help=(
                    "This setting is passed to the preserved master engine for "
                    "intermediate batch workbooks only. The final downloaded "
                    "MASTER_CONSOLIDATED_REPORT.xlsx is always assembled into one "
                    "Master_Data worksheet when it fits Excel's hard row limit."
                ),
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
            help="Options combine the original script lists with names found in uploaded files.",
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
    if filters is None:
        missing.append("Filters")
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
        st.warning("Advanced sorting is enabled without sales-history files. The preserved master code will continue and bypass history-based limits if no history can be loaded.")

    st.markdown("#### Processing destination")
    processing_mode = st.radio(
        "Choose where the Suggested Stock Mapper runs",
        [
            "GitHub Actions full run (recommended)",
            "Run inside this Streamlit app (small test only)",
        ],
        horizontal=True,
        help="GitHub Actions keeps the heavy Excel work away from Streamlit Community Cloud's smaller app server.",
    )

    if processing_mode.startswith("GitHub Actions"):
        settings = _github_job_settings()
        c1, c2, c3 = st.columns(3)
        batch_size = int(c1.number_input("Salespeople per batch", min_value=1, max_value=20, value=1, step=1))
        continue_after_failure = c2.toggle("Continue after failed batch", value=True)
        scope_mode = c3.selectbox(
            "Batch scope",
            ["All salespeople discovered from current inputs", "Only salespeople selected above"],
        )
        requested_by = st.text_input(
            "Requested by (optional)",
            placeholder="Name or work email for the manifest",
        )

        cloud_scope = "all" if scope_mode.startswith("All") else "selected"
        cloud_salespeople = [] if cloud_scope == "all" else list(runtime.get("TEST_SALESPEOPLE") or [])

        cloud_missing = [
            item
            for item in missing
            if not (cloud_scope == "all" and item == "one or more selected salespeople")
        ]
        if cloud_scope == "selected" and not cloud_salespeople:
            cloud_missing.append("one or more selected salespeople for batching")
        if settings is None:
            cloud_missing.append("GitHub settings in Streamlit secrets")
            st.warning(
                "The background worker is not connected yet. Add the [github] secret block described in "
                "STREAMLIT_GITHUB_SETUP.md after pushing this project to your private repository."
            )

        if cloud_scope == "all":
            st.caption(
                "The worker will read the exact salesperson names from the uploaded Parquet profiles and Filters workbook, "
                f"then process them in groups of {batch_size}. All matching, filtering, styling and output settings above "
                "are copied into every batch. Only TEST_MODE and TEST_SALESPEOPLE change per batch."
            )
        else:
            estimated_batches = (len(cloud_salespeople) + batch_size - 1) // batch_size if cloud_salespeople else 0
            st.caption(
                f"This job will process {len(cloud_salespeople)} selected salesperson(s) in approximately "
                f"{estimated_batches} batch(es). All other configuration remains identical across batches."
            )
        st.info(
            "The worker uses isolated batches for memory safety, then restores the normal final output layout: "
            "one global MASTER_CONSOLIDATED_REPORT.xlsx when Merged Master is enabled, plus every salesperson "
            "workbook and separated-party ZIP together in the same final ZIP when those options are enabled."
        )

        if st.button(
            "Submit full run to GitHub Actions",
            type="primary",
            use_container_width=True,
            disabled=bool(cloud_missing),
            key="submit_github_full_run",
        ):
            try:
                assert settings is not None
                backend = GitHubReleaseBackend(settings)
                with st.spinner("Checking the private repository connection..."):
                    backend.check_connection(require_private=True)
                    active_jobs = backend.list_active_jobs(limit=30)
                    if active_jobs:
                        active_tag = str(active_jobs[0].get("tag_name", "another job"))
                        raise GitHubJobError(
                            "A Matching Mapper job is already queued or running: "
                            f"{active_tag}. Wait for it to finish or delete the stale job before submitting another."
                        )
                job_tag = new_job_tag()
                with tempfile.TemporaryDirectory(prefix="matching_mapper_submit_") as tmp:
                    bundle_path = Path(tmp) / "input_bundle.zip"
                    with st.spinner("Packaging inputs and submitting the private background job..."):
                        build_job_bundle(
                            output_path=bundle_path,
                            job_id=job_tag,
                            runtime_overrides=runtime,
                            scope=cloud_scope,
                            salespeople=cloud_salespeople,
                            batch_size=batch_size,
                            continue_after_failed_batch=continue_after_failure,
                            stock_file=(stock.name, stock),
                            filter_file=(filters.name, filters),
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
                st.success("Full run submitted. GitHub Actions will continue even after this browser page is closed.")
                st.rerun()
            except GitHubJobError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.exception(exc)

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
        st.warning(
            "Local processing is deliberately limited to at most 3 selected salespeople. "
            "Use GitHub Actions for an all-salespeople or full production run."
        )
        if st.button(
            "Generate suggested stock reports in Streamlit",
            type="primary",
            use_container_width=True,
            disabled=bool(local_missing),
            key="run_suggested_locally",
        ):
            try:
                with st.spinner("Running the preserved master matching engine. Large profiles may take several minutes..."):
                    data, filename, logs = run_suggested_process(
                        stock_bytes=stock.getvalue(),
                        profile_perm_bytes=profile_perm.getvalue() if profile_perm else None,
                        profile_temp_bytes=profile_temp.getvalue() if profile_temp else None,
                        filter_bytes=filters.getvalue(),
                        clean_bytes=no_need.getvalue() if no_need else None,
                        sales_history_files=[(file.name, file.getvalue()) for file in (sales_history or [])],
                        runtime_overrides=runtime,
                    )
                    _save_result("suggested_result", data, filename, logs)
                st.rerun()
            except Exception as exc:
                st.exception(exc)

        if local_missing:
            st.info(
                "Complete the following before processing: "
                + "; ".join(dict.fromkeys(local_missing))
            )
        _show_result("suggested_result", "Download suggested reports ZIP")

def demand_filter_page() -> None:
    section_header(
        "Rule-based stock demand",
        "Demand Filter Mapper",
        "Apply the original filter-demand master code to Stock Actual and the Filters workbook. Reports are generated by salesperson.",
        ["Upload stock", "Upload filters", "Set tolerance", "Download reports"],
    )
    left, right = st.columns(2)
    stock = left.file_uploader("Stock Actual workbook *", type=["xlsx"], key="demand_stock")
    filters = right.file_uploader("Filters workbook *", type=["xlsx"], key="demand_filters")

    with st.expander("Runtime controls", expanded=True):
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
            st.exception(exc)
    _show_result("demand_result", "Download demand-filter reports ZIP")


def sorting_page() -> None:
    section_header(
        "Suggested-output prioritization",
        "Sort Suggested Output",
        "Run the preserved sorting engine with all business controls from its CONFIG section.",
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
            st.exception(exc)
    if missing:
        st.info("Complete the following before processing: " + "; ".join(missing))
    _show_result("sort_result", "Download sorted suggested ZIP")

def separation_page() -> None:
    section_header(
        "Workbook distribution",
        "Separate Excel Sheets",
        "Split each worksheet from every Excel workbook inside an uploaded ZIP. Each source workbook becomes its own nested ZIP of separate sheet files.",
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
            st.exception(exc)
    _show_result("separate_result", "Download separated workbook ZIP")


if page == "Overview":
    overview_page()
elif page == "Suggested Stock Mapper":
    suggested_page()
elif page == "Demand Filter Mapper":
    demand_filter_page()
elif page == "Sort Suggested Output":
    sorting_page()
elif page == "Separate Excel Sheets":
    separation_page()

page_footer()
