from __future__ import annotations

import html
import re
import shutil
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Iterable, Sequence
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MASTER_REPORT_NAME = "MASTER_CONSOLIDATED_REPORT.xlsx"
ASSEMBLER_VERSION = "1.4.3"
FINAL_MASTER_SHEET_POLICY = "strict_single_sheet"
DIAGNOSTICS_FOLDER = "_batch_diagnostics"
EXCEL_MAX_DATA_ROWS = 1_048_575  # Row 1 is reserved for the header.
MAX_HYPERLINKS_PER_SHEET = 0  # 0 disables native-hyperlink sheet splitting.
MASTER_SINGLE_SHEET_NAME = "Master_Data"
HYPERLINK_FORMULA_LIMIT = 7_900

_FIXED_WIDTH_COLUMNS = {"DETAIL", "PHYGITAL", "IMAGE", "VIDEO", "CERTIFICATE", "LAYOUT"}
_DECIMAL_COLUMNS = {
    "Carat",
    "RapRate",
    "Reserve%",
    "Rate",
    "Amount",
    "Ratio",
    "Dep",
    "Tab",
    "Dia",
    "Girdle",
    "C/A",
    "C/H",
    "P/A",
    "P/H",
    "Disc%",
    "Price/carat",
    "Total Amount",
    "Depth",
    "Table",
}

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_HYPERLINK_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)

# Self-closing alternatives must come first. Otherwise a self-closing cell can
# be mistaken for the opening tag of a later normal cell in the same row.
_ROW_RE = re.compile(rb"<row\b([^>]*)/>|<row\b([^>]*)>(.*?)</row>", re.S)
_CELL_RE = re.compile(rb"<c\b([^>]*)/>|<c\b([^>]*)>(.*?)</c>", re.S)
_COLUMN_RE = re.compile(rb"<col\b([^>]*)/?>")
_HYPERLINK_RE = re.compile(rb"<hyperlink\b([^>]*)/>")
_ATTRIBUTE_PATTERNS: dict[bytes, re.Pattern[bytes]] = {}


@dataclass
class BatchCollection:
    """Files collected from one preserved-master batch output."""

    master_parts: list[Path] = field(default_factory=list)
    packaged_files: list[str] = field(default_factory=list)
    empty: bool = True


@dataclass(frozen=True)
class MasterMergeSummary:
    """Metadata describing the globally merged master workbook."""

    source_workbooks: int
    total_rows: int
    total_columns: int
    worksheets: int
    hyperlinks: int
    output_path: Path


@dataclass(frozen=True)
class _SourceSheet:
    workbook_path: Path
    sheet_path: str
    headers: tuple[str, ...]
    widths: dict[int, float]
    row_numbers: tuple[int, ...]
    hyperlink_counts: dict[int, int]


def _member_basename(name: str) -> str:
    pure = PurePosixPath(str(name))
    if pure.is_absolute() or ".." in pure.parts or not pure.name:
        raise ValueError(f"Unsafe file name in preserved-master result: {name!r}")
    return pure.name


def _unique_root_name(name: str, used_names: set[str], batch_name: str) -> str:
    """Keep the normal master layout while preventing accidental ZIP duplicates."""
    clean = _member_basename(name)
    key = clean.casefold()
    if key not in used_names:
        used_names.add(key)
        return clean

    path = Path(clean)
    candidate = f"{path.stem}__{batch_name}{path.suffix}"
    counter = 2
    while candidate.casefold() in used_names:
        candidate = f"{path.stem}__{batch_name}_{counter}{path.suffix}"
        counter += 1
    used_names.add(candidate.casefold())
    return candidate


def _stream_member_to_zip(
    source_archive: zipfile.ZipFile,
    source_info: zipfile.ZipInfo,
    target_archive: zipfile.ZipFile,
    target_name: str,
) -> None:
    target_info = zipfile.ZipInfo(target_name, date_time=source_info.date_time)
    # XLSX files and nested ZIPs are already compressed. Storing them avoids
    # wasting GitHub-runner CPU and keeps the packaging step predictable.
    target_info.compress_type = zipfile.ZIP_STORED
    target_info.external_attr = source_info.external_attr or (0o600 << 16)
    with source_archive.open(source_info, "r") as source, target_archive.open(
        target_info,
        "w",
        force_zip64=True,
    ) as destination:
        shutil.copyfileobj(source, destination, length=1024 * 1024)


def collect_preserved_batch_output(
    *,
    batch_zip_path: Path,
    batch_name: str,
    master_parts_directory: Path,
    final_archive: zipfile.ZipFile,
    used_root_names: set[str],
) -> BatchCollection:
    """
    Flatten one preserved-master result into the final user-facing ZIP.

    Batch master workbooks are staged for one global merge. Salesperson
    workbooks and ``*_Separated_Parties.zip`` files are copied to the final ZIP
    root, which matches the layout of a normal unbatched master-code run.
    """
    batch_zip_path = Path(batch_zip_path)
    master_parts_directory.mkdir(parents=True, exist_ok=True)
    collection = BatchCollection()

    with zipfile.ZipFile(batch_zip_path, "r") as source_archive:
        for member_index, info in enumerate(source_archive.infolist(), start=1):
            if info.is_dir():
                continue
            base_name = _member_basename(info.filename)
            collection.empty = False

            if base_name.casefold() == MASTER_REPORT_NAME.casefold():
                target = (
                    master_parts_directory
                    / f"{batch_name}_{member_index:03d}_{MASTER_REPORT_NAME}"
                )
                with source_archive.open(info, "r") as source, target.open("wb") as destination:
                    shutil.copyfileobj(source, destination, length=1024 * 1024)
                collection.master_parts.append(target)
                continue

            target_name = _unique_root_name(base_name, used_root_names, batch_name)
            _stream_member_to_zip(source_archive, info, final_archive, target_name)
            collection.packaged_files.append(target_name)

    return collection


def add_diagnostic_file(
    archive: zipfile.ZipFile,
    *,
    source_path: Path,
    batch_name: str,
    filename: str,
) -> None:
    source_path = Path(source_path)
    if source_path.is_file():
        archive.write(
            source_path,
            f"{DIAGNOSTICS_FOLDER}/{batch_name}/{filename}",
            compress_type=zipfile.ZIP_DEFLATED,
        )


def iter_nested_batch_zips(
    source_archive: zipfile.ZipFile,
) -> Iterable[tuple[str, zipfile.ZipInfo]]:
    """Yield legacy ``batch_XX/suggested_reports.zip`` members for conversion."""
    for info in source_archive.infolist():
        pure = PurePosixPath(info.filename)
        if info.is_dir() or pure.name != "suggested_reports.zip" or len(pure.parts) < 2:
            continue
        yield pure.parts[-2], info


def _attribute(attrs: bytes, name: bytes) -> bytes | None:
    pattern = _ATTRIBUTE_PATTERNS.get(name)
    if pattern is None:
        pattern = re.compile(rb"(?:^|\s)" + re.escape(name) + rb'="([^"]*)"')
        _ATTRIBUTE_PATTERNS[name] = pattern
    match = pattern.search(attrs)
    return match.group(1) if match else None


def _row_parts(match: re.Match[bytes]) -> tuple[bytes, bytes]:
    if match.group(1) is not None:
        return match.group(1), b""
    return match.group(2) or b"", match.group(3) or b""


def _cell_parts(match: re.Match[bytes]) -> tuple[bytes, bytes, bool]:
    if match.group(1) is not None:
        return match.group(1), b"", True
    return match.group(2) or b"", match.group(3) or b"", False


def _column_index(reference: bytes | str) -> int:
    raw = reference.encode("ascii", "ignore") if isinstance(reference, str) else reference
    value = 0
    for byte in raw:
        if 65 <= byte <= 90:
            value = value * 26 + byte - 64
        elif 97 <= byte <= 122:
            value = value * 26 + byte - 96
        else:
            break
    return value


def _row_number(reference: bytes | str) -> int:
    raw = reference.decode("ascii", "ignore") if isinstance(reference, bytes) else reference
    match = re.search(r"\d+", raw)
    return int(match.group(0)) if match else 0


def _cell_value(cell_type: bytes | None, inner: bytes) -> object | None:
    if not inner:
        return None
    if cell_type == b"inlineStr":
        values = re.findall(rb"<t(?:\s[^>]*)?>(.*?)</t>", inner, flags=re.S)
        if not values:
            return None
        return "".join(
            html.unescape(value.decode("utf-8", "replace"))
            for value in values
        )

    value_match = re.search(rb"<v(?:\s[^>]*)?>(.*?)</v>", inner, flags=re.S)
    if not value_match:
        return None
    text = html.unescape(value_match.group(1).decode("utf-8", "replace"))
    if cell_type == b"b":
        return text == "1"
    return text


def _parse_header(row_inner: bytes) -> tuple[str, ...]:
    values: dict[int, str] = {}
    maximum_column = 0
    for match in _CELL_RE.finditer(row_inner):
        attrs, inner, _self_closing = _cell_parts(match)
        reference = _attribute(attrs, b"r")
        if not reference:
            continue
        column = _column_index(reference)
        if not column:
            continue
        maximum_column = max(maximum_column, column)
        value = _cell_value(_attribute(attrs, b"t"), inner)
        values[column] = str(value).strip() if value is not None else ""
    return tuple(values.get(index, "") for index in range(1, maximum_column + 1))


def _parse_widths(sheet_xml: bytes) -> dict[int, float]:
    widths: dict[int, float] = {}
    prefix = sheet_xml.split(b"<sheetData", 1)[0]
    for match in _COLUMN_RE.finditer(prefix):
        attrs = match.group(1)
        minimum = _attribute(attrs, b"min")
        maximum = _attribute(attrs, b"max")
        width = _attribute(attrs, b"width")
        if not minimum or not maximum or not width:
            continue
        try:
            parsed_width = float(width)
            for index in range(int(minimum), int(maximum) + 1):
                widths[index] = parsed_width
        except ValueError:
            continue
    return widths


def _sheet_paths(archive: zipfile.ZipFile) -> list[str]:
    paths = [
        name
        for name in archive.namelist()
        if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)
    ]
    return sorted(
        paths,
        key=lambda name: int(re.search(r"(\d+)", Path(name).stem).group(1)),
    )


def _relationship_targets(
    archive: zipfile.ZipFile,
    sheet_path: str,
) -> dict[str, str]:
    relationship_path = (
        f"{str(Path(sheet_path).parent)}/_rels/{Path(sheet_path).name}.rels"
    )
    try:
        root = ET.fromstring(archive.read(relationship_path))
    except KeyError:
        return {}
    return {
        element.attrib["Id"]: element.attrib["Target"]
        for element in root
        if element.attrib.get("Id") and element.attrib.get("Target")
    }


def _hyperlinks(
    archive: zipfile.ZipFile,
    sheet_path: str,
    sheet_xml: bytes,
) -> dict[str, str]:
    relationships = _relationship_targets(archive, sheet_path)
    result: dict[str, str] = {}
    for match in _HYPERLINK_RE.finditer(sheet_xml):
        attrs = match.group(1)
        reference = _attribute(attrs, b"ref")
        relationship_id = _attribute(attrs, b"r:id")
        location = _attribute(attrs, b"location")
        if not reference:
            continue
        target = (
            relationships.get(relationship_id.decode("utf-8", "replace"))
            if relationship_id
            else None
        )
        if target is None and location:
            target = "#" + location.decode("utf-8", "replace")
        if target:
            result[reference.decode("ascii", "ignore")] = target
    return result


def _scan_sources(
    workbook_paths: Sequence[Path],
    *,
    max_rows_per_sheet: int,
    max_hyperlinks_per_sheet: int,
) -> tuple[list[_SourceSheet], list[str], dict[str, float], list[tuple[int, int]]]:
    descriptors: list[_SourceSheet] = []
    union_headers: list[str] = []
    seen_headers: set[str] = set()
    widths_by_header: dict[str, float] = {}
    all_link_counts: list[int] = []

    for workbook_path in workbook_paths:
        with zipfile.ZipFile(workbook_path, "r") as archive:
            for sheet_path in _sheet_paths(archive):
                sheet_xml = archive.read(sheet_path)
                header_row: bytes | None = None
                row_numbers: list[int] = []
                for row_match in _ROW_RE.finditer(sheet_xml):
                    attrs, inner = _row_parts(row_match)
                    row_attr = _attribute(attrs, b"r")
                    if not row_attr:
                        continue
                    row_number = int(row_attr)
                    if row_number == 1:
                        header_row = inner
                    elif row_number > 1:
                        row_numbers.append(row_number)
                if header_row is None or not row_numbers:
                    continue

                headers = _parse_header(header_row)
                widths = _parse_widths(sheet_xml)
                for index, header in enumerate(headers, start=1):
                    if not header:
                        continue
                    if header not in seen_headers:
                        seen_headers.add(header)
                        union_headers.append(header)
                    if index in widths:
                        widths_by_header[header] = max(
                            widths_by_header.get(header, 0.0),
                            widths[index],
                        )

                links = _hyperlinks(archive, sheet_path, sheet_xml)
                link_counts = Counter(_row_number(reference) for reference in links)
                all_link_counts.extend(link_counts.get(row_number, 0) for row_number in row_numbers)
                descriptors.append(
                    _SourceSheet(
                        workbook_path=Path(workbook_path),
                        sheet_path=sheet_path,
                        headers=headers,
                        widths=widths,
                        row_numbers=tuple(row_numbers),
                        hyperlink_counts=dict(link_counts),
                    )
                )

    if not descriptors or not union_headers:
        return descriptors, union_headers, widths_by_header, []

    sheet_plan: list[tuple[int, int]] = []
    current_rows = 0
    current_links = 0
    for row_links in all_link_counts:
        hyperlink_limit_reached = (
            max_hyperlinks_per_sheet > 0
            and current_links + row_links > max_hyperlinks_per_sheet
        )
        if current_rows and (
            current_rows >= max_rows_per_sheet
            or hyperlink_limit_reached
        ):
            sheet_plan.append((current_rows, current_links))
            current_rows = 0
            current_links = 0
        current_rows += 1
        current_links += row_links
    if current_rows:
        sheet_plan.append((current_rows, current_links))
    return descriptors, union_headers, widths_by_header, sheet_plan


def _make_template(
    output_path: Path,
    *,
    sheet_count: int,
    enable_styling: bool,
    enable_new_fill_logic: bool,
) -> tuple[int, dict[tuple[str, str], int]]:
    workbook = Workbook()
    workbook.active.title = MASTER_SINGLE_SHEET_NAME if sheet_count == 1 else "Master_Data_Part_1"
    for index in range(2, sheet_count + 1):
        workbook.create_sheet(f"Master_Data_Part_{index}")

    worksheet = workbook.worksheets[0]
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
    fills: dict[str, PatternFill | None] = {
        "none": None,
        "exact": PatternFill(
            start_color="CCFFCC",
            end_color="CCFFCC",
            fill_type="solid",
        ),
        "filter": PatternFill(
            start_color="FFF2CC",
            end_color="FFF2CC",
            fill_type="solid",
        ),
        "search": PatternFill(
            start_color="E0F7FA",
            end_color="E0F7FA",
            fill_type="solid",
        ),
        "nearby": PatternFill(
            start_color="FFE699" if enable_new_fill_logic else "FFFFFF",
            end_color="FFE699" if enable_new_fill_logic else "FFFFFF",
            fill_type="solid",
        ),
        "maybe": PatternFill(
            start_color="FFFFFF",
            end_color="FFFFFF",
            fill_type="solid",
        ),
    }

    header = worksheet.cell(1, 1, "header")
    if enable_styling:
        header.fill = PatternFill(
            start_color="D3D3D3",
            end_color="D3D3D3",
            fill_type="solid",
        )
        header.font = Font(bold=True)
        header.border = border
        header.alignment = alignment
    header_style_id = header.style_id

    style_ids: dict[tuple[str, str], int] = {}
    column = 2
    for fill_key, fill in fills.items():
        for kind in ("general", "decimal", "hyperlink", "serial"):
            cell = worksheet.cell(1, column, "style")
            column += 1
            if enable_styling:
                cell.border = border
                cell.alignment = alignment
                if fill is not None:
                    cell.fill = fill
                if kind == "decimal":
                    cell.number_format = "0.00"
                if kind in {"hyperlink", "serial"}:
                    cell.font = Font(color="0000FF", underline="single")
            elif kind == "hyperlink":
                # Native hyperlinks are styled by the original master even
                # when general worksheet styling is disabled.
                cell.font = Font(color="0000FF", underline="single")
            style_ids[(fill_key, kind)] = cell.style_id

    workbook.save(output_path)
    workbook.close()
    return header_style_id, style_ids


def _excel_formula_string(value: object) -> str:
    return str(value).replace('"', '""')


def _hyperlink_formula(target: object, display: object) -> str | None:
    formula = (
        f'HYPERLINK("{_excel_formula_string(target)}",'
        f'"{_excel_formula_string(display)}")'
    )
    if len(formula) > HYPERLINK_FORMULA_LIMIT:
        return None
    return formula


def _row_fill(values: dict[str, object | None], enable_new_fill_logic: bool) -> str:
    is_nearby = values.get("Is_Nearby")
    is_filter_match = values.get("Is_Filter_Match")
    match_type = values.get("Match_Type")
    tolerance_type = values.get("Tolerance_Type")

    if is_nearby is False:
        return "exact"
    if is_filter_match is True:
        return "filter"
    if match_type == "SEARCH":
        return "search"
    if tolerance_type == "NEARBY":
        return "nearby"
    if tolerance_type == "MAYBE":
        return "maybe"
    if is_nearby is True and not enable_new_fill_logic:
        return "exact"
    return "none"


class _SheetWriter:
    def __init__(
        self,
        root: Path,
        *,
        sheet_index: int,
        planned_rows: int,
        headers: Sequence[str],
        widths_by_header: dict[str, float],
        header_style_id: int,
        style_ids: dict[tuple[str, str], int],
        enable_styling: bool,
        hyperlink_mode: str = "formula",
    ) -> None:
        self.sheet_index = sheet_index
        self.headers = list(headers)
        self.enable_styling = enable_styling
        self.hyperlink_mode = hyperlink_mode
        self.data_rows = 0
        self.hyperlinks = 0
        self.native_hyperlinks = 0
        self._column_letters = [
            get_column_letter(index)
            for index in range(1, len(self.headers) + 1)
        ]
        self._base_kinds = [
            "serial"
            if header == "Sr.No"
            else "decimal"
            if header in _DECIMAL_COLUMNS
            else "general"
            for header in self.headers
        ]
        self._style_attributes = {
            key: (f' s="{style_id}"'.encode("ascii") if style_id else b"")
            for key, style_id in style_ids.items()
        }

        self.sheet_path = root / f"sheet{sheet_index}.xml"
        self.links_path = root / f"sheet{sheet_index}.links"
        self.rels_path = root / f"sheet{sheet_index}.rels"
        self._sheet = self.sheet_path.open("wb", buffering=1024 * 1024)
        self._links = self.links_path.open("wb", buffering=1024 * 1024)
        self._rels = self.rels_path.open("wb", buffering=1024 * 1024)

        last_column = get_column_letter(len(self.headers))
        self._sheet.write(
            (
                f'<worksheet xmlns="{_MAIN_NS}" xmlns:r="{_DOC_REL_NS}">'
                '<sheetPr><outlinePr summaryBelow="1" summaryRight="1"/>'
                '<pageSetUpPr/></sheetPr>'
                f'<dimension ref="A1:{last_column}{planned_rows + 1}"/>'
                '<sheetViews><sheetView workbookViewId="0">'
                '<selection activeCell="A1" sqref="A1"/>'
                '</sheetView></sheetViews>'
                '<sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>'
            ).encode("utf-8")
        )

        if enable_styling:
            column_xml: list[bytes] = [b"<cols>"]
            for index, header in enumerate(self.headers, start=1):
                width = widths_by_header.get(header)
                if width is None and header in _FIXED_WIDTH_COLUMNS:
                    width = 12.0
                if width is not None:
                    column_xml.append(
                        f'<col width="{width}" customWidth="1" min="{index}" max="{index}"/>'.encode(
                            "utf-8"
                        )
                    )
            column_xml.append(b"</cols>")
            self._sheet.write(b"".join(column_xml))

        header_style = f' s="{header_style_id}"' if header_style_id else ""
        header_parts: list[bytes] = [b'<sheetData><row r="1">']
        for index, header in enumerate(self.headers, start=1):
            coordinate = f"{self._column_letters[index - 1]}1"
            header_parts.append(
                f'<c r="{coordinate}"{header_style} t="inlineStr"><is><t>{escape(header)}</t></is></c>'.encode(
                    "utf-8"
                )
            )
        header_parts.append(b"</row>")
        self._sheet.write(b"".join(header_parts))
        self._rels.write(
            f'<Relationships xmlns="{_PACKAGE_REL_NS}">'.encode("utf-8")
        )

    def add_row(
        self,
        *,
        cells_by_output_column: dict[
            int,
            tuple[bytes | None, bytes, bool, str | None],
        ],
        fill_key: str,
    ) -> None:
        self.data_rows += 1
        row_number = self.data_rows + 1
        row_parts: list[bytes] = [f'<row r="{row_number}">'.encode("ascii")]
        link_parts: list[bytes] = []
        relationship_parts: list[bytes] = []

        for zero_index, _header in enumerate(self.headers):
            output_column = zero_index + 1
            item = cells_by_output_column.get(output_column)
            hyperlink = item[3] if item else None
            kind = "hyperlink" if hyperlink else self._base_kinds[zero_index]
            style_attribute = self._style_attributes[(fill_key, kind)]
            coordinate = f"{self._column_letters[zero_index]}{row_number}"
            coordinate_bytes = coordinate.encode("ascii")

            if hyperlink and item and self.hyperlink_mode == "formula":
                cell_type, inner, _self_closing, _target = item
                display_value = _cell_value(cell_type, inner)
                display_text = str(display_value).strip() if display_value is not None else ""
                if not display_text:
                    display_text = str(hyperlink)
                formula = _hyperlink_formula(hyperlink, display_text)
                self.hyperlinks += 1
                if formula is not None:
                    row_parts.extend(
                        [
                            b'<c r="',
                            coordinate_bytes,
                            b'"',
                            style_attribute,
                            b' t="str"><f>',
                            escape(formula).encode("utf-8"),
                            b"</f><v>",
                            escape(display_text).encode("utf-8"),
                            b"</v></c>",
                        ]
                    )
                    continue
                row_parts.extend(
                    [
                        b'<c r="',
                        coordinate_bytes,
                        b'"',
                        style_attribute,
                        b' t="inlineStr"><is><t>',
                        escape(display_text).encode("utf-8"),
                        b"</t></is></c>",
                    ]
                )
                continue

            if item and item[1]:
                cell_type, inner, _self_closing, _target = item
                type_attribute = b' t="' + cell_type + b'"' if cell_type else b""
                row_parts.extend(
                    [
                        b'<c r="',
                        coordinate_bytes,
                        b'"',
                        style_attribute,
                        type_attribute,
                        b'>',
                        inner,
                        b"</c>",
                    ]
                )
            elif self.enable_styling or style_attribute:
                row_parts.extend(
                    [b'<c r="', coordinate_bytes, b'"', style_attribute, b"/>"]
                )

            if hyperlink:
                self.hyperlinks += 1
                self.native_hyperlinks += 1
                relationship_id = f"rId{self.native_hyperlinks}"
                link_parts.append(
                    f'<hyperlink ref="{coordinate}" r:id="{relationship_id}"/>'.encode(
                        "utf-8"
                    )
                )
                relationship_parts.append(
                    (
                        f'<Relationship Type="{_HYPERLINK_REL_TYPE}" '
                        f'Target={quoteattr(str(hyperlink))} TargetMode="External" '
                        f'Id="{relationship_id}"/>'
                    ).encode("utf-8")
                )

        row_parts.append(b"</row>")
        self._sheet.write(b"".join(row_parts))
        if link_parts:
            self._links.write(b"".join(link_parts))
            self._rels.write(b"".join(relationship_parts))

    def close(self) -> None:
        self._sheet.write(b"</sheetData>")
        self._links.close()
        if self.native_hyperlinks:
            self._sheet.write(b"<hyperlinks>")
            with self.links_path.open("rb") as source:
                shutil.copyfileobj(source, self._sheet, length=1024 * 1024)
            self._sheet.write(b"</hyperlinks>")
        self._sheet.write(
            b'<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>'
            b"</worksheet>"
        )
        self._sheet.close()
        self._rels.write(b"</Relationships>")
        self._rels.close()
        self.links_path.unlink(missing_ok=True)
        if not self.native_hyperlinks:
            self.rels_path.unlink(missing_ok=True)


def merge_master_workbooks(
    *,
    master_parts: Sequence[Path],
    output_path: Path,
    max_rows_per_sheet: int,
    enable_styling: bool,
    enable_new_fill_logic: bool = True,
) -> MasterMergeSummary | None:
    """
    Merge batch-level master workbooks into one low-memory workbook.

    The preserved engine's data cells are streamed directly from worksheet
    OOXML. Styles are recreated according to the original master-code rules.
    Hyperlinks are written as HYPERLINK formulas so all data remains in one
    worksheet without using Excel's native hyperlink collection limit. The
    final assembly deliberately ignores the intermediate batch row-limit
    setting and fails clearly only when Excel's hard one-sheet row limit is
    exceeded. This
    avoids loading millions of cells into pandas or a normal openpyxl workbook
    during the final assembly step.
    """
    clean_parts = [Path(path) for path in master_parts if Path(path).is_file()]
    if not clean_parts:
        return None

    # ``MAX_ROWS_PER_SHEET`` belongs to the preserved engine's intermediate
    # batch workbooks. It must not split the final user-facing master. The
    # consolidated workbook is therefore forced into one worksheet whenever
    # the data fits Excel's hard worksheet limit (1,048,576 rows including the
    # header). ``max_rows_per_sheet`` is retained in the public signature for
    # backward compatibility with older callers but is intentionally ignored
    # for final assembly.
    _requested_intermediate_limit = max(
        1,
        min(int(max_rows_per_sheet), EXCEL_MAX_DATA_ROWS),
    )
    descriptors, headers, widths, sheet_plan = _scan_sources(
        clean_parts,
        max_rows_per_sheet=EXCEL_MAX_DATA_ROWS,
        max_hyperlinks_per_sheet=MAX_HYPERLINKS_PER_SHEET,
    )
    if not descriptors or not headers or not sheet_plan:
        return None

    total_planned_rows = sum(rows for rows, _links in sheet_plan)
    if total_planned_rows > EXCEL_MAX_DATA_ROWS or len(sheet_plan) != 1:
        raise ValueError(
            "The final consolidated master contains "
            f"{total_planned_rows:,} data rows, which cannot fit in one Excel "
            "worksheet. Excel allows at most 1,048,575 data rows plus one "
            "header row. Reduce the full-run scope or use a non-Excel export."
        )

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.unlink(missing_ok=True)

    with tempfile.TemporaryDirectory(
        prefix="matching_mapper_master_merge_"
    ) as temporary_directory:
        root = Path(temporary_directory)
        template_path = root / "template.xlsx"
        header_style_id, style_ids = _make_template(
            template_path,
            sheet_count=len(sheet_plan),
            enable_styling=enable_styling,
            enable_new_fill_logic=enable_new_fill_logic,
        )

        sheet_index = 1
        writer = _SheetWriter(
            root,
            sheet_index=sheet_index,
            planned_rows=sheet_plan[0][0],
            headers=headers,
            widths_by_header=widths,
            header_style_id=header_style_id,
            style_ids=style_ids,
            enable_styling=enable_styling,
            hyperlink_mode="formula",
        )
        total_rows = 0
        total_hyperlinks = 0
        output_lookup = {
            header: index
            for index, header in enumerate(headers, start=1)
        }

        for descriptor in descriptors:
            with zipfile.ZipFile(descriptor.workbook_path, "r") as archive:
                sheet_xml = archive.read(descriptor.sheet_path)
                links = _hyperlinks(archive, descriptor.sheet_path, sheet_xml)

            source_to_output = {
                source_index: output_lookup[header]
                for source_index, header in enumerate(
                    descriptor.headers,
                    start=1,
                )
                if header and header in output_lookup
            }
            status_source_indexes = {
                header: source_index
                for source_index, header in enumerate(
                    descriptor.headers,
                    start=1,
                )
                if header
                in {
                    "Is_Nearby",
                    "Is_Filter_Match",
                    "Match_Type",
                    "Tolerance_Type",
                }
            }

            for row_match in _ROW_RE.finditer(sheet_xml):
                row_attrs, row_inner = _row_parts(row_match)
                row_attr = _attribute(row_attrs, b"r")
                if not row_attr or int(row_attr) <= 1:
                    continue

                if writer.data_rows >= sheet_plan[sheet_index - 1][0]:
                    total_hyperlinks += writer.hyperlinks
                    writer.close()
                    sheet_index += 1
                    writer = _SheetWriter(
                        root,
                        sheet_index=sheet_index,
                        planned_rows=sheet_plan[sheet_index - 1][0],
                        headers=headers,
                        widths_by_header=widths,
                        header_style_id=header_style_id,
                        style_ids=style_ids,
                        enable_styling=enable_styling,
                        hyperlink_mode="formula",
                    )

                cells_by_output_column: dict[
                    int,
                    tuple[bytes | None, bytes, bool, str | None],
                ] = {}
                status_values: dict[str, object | None] = {}
                for cell_match in _CELL_RE.finditer(row_inner):
                    attrs, inner, self_closing = _cell_parts(cell_match)
                    reference = _attribute(attrs, b"r")
                    if not reference:
                        continue
                    source_column = _column_index(reference)
                    output_column = source_to_output.get(source_column)
                    if output_column is None:
                        continue
                    cell_type = _attribute(attrs, b"t")
                    source_reference = reference.decode("ascii", "ignore")
                    cells_by_output_column[output_column] = (
                        cell_type,
                        inner,
                        self_closing,
                        links.get(source_reference),
                    )
                    for status_header, status_index in status_source_indexes.items():
                        if status_index == source_column:
                            status_values[status_header] = _cell_value(
                                cell_type,
                                inner,
                            )
                            break

                fill_key = (
                    _row_fill(status_values, enable_new_fill_logic)
                    if enable_styling
                    else "none"
                )
                writer.add_row(
                    cells_by_output_column=cells_by_output_column,
                    fill_key=fill_key,
                )
                total_rows += 1

        total_hyperlinks += writer.hyperlinks
        writer.close()

        package_root = root / "package"
        package_root.mkdir()
        with zipfile.ZipFile(template_path, "r") as template_archive:
            template_archive.extractall(package_root)

        for index in range(1, len(sheet_plan) + 1):
            target_sheet = (
                package_root
                / "xl"
                / "worksheets"
                / f"sheet{index}.xml"
            )
            target_sheet.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(root / f"sheet{index}.xml", target_sheet)

            target_relationships = (
                package_root
                / "xl"
                / "worksheets"
                / "_rels"
                / f"sheet{index}.xml.rels"
            )
            source_relationships = root / f"sheet{index}.rels"
            if source_relationships.is_file():
                target_relationships.parent.mkdir(
                    parents=True,
                    exist_ok=True,
                )
                shutil.copy2(source_relationships, target_relationships)
            else:
                target_relationships.unlink(missing_ok=True)

        # Compression level 1 reduces final-assembly time substantially while
        # still producing a compact XLSX. The batch workbooks themselves remain
        # untouched and the merged result stays well below GitHub's asset limit.
        with zipfile.ZipFile(
            output_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=1,
            allowZip64=True,
        ) as output_archive:
            package_files = sorted(
                path
                for path in package_root.rglob("*")
                if path.is_file()
            )
            for package_file in package_files:
                output_archive.write(
                    package_file,
                    package_file.relative_to(package_root).as_posix(),
                )

    return MasterMergeSummary(
        source_workbooks=len(clean_parts),
        total_rows=total_rows,
        total_columns=len(headers),
        worksheets=len(sheet_plan),
        hyperlinks=total_hyperlinks,
        output_path=output_path,
    )
