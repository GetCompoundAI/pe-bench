import io
import logging
from datetime import datetime

import lxml.etree

from pe_bench.office.workbook.excel.io.constants import Namespaces
from pe_bench.office.workbook.excel.io.xml_functions import Element, is_true, safe_float
from pe_bench.office.workbook.excel.type_coerce import datetime_to_serial_number
from pe_bench.office.workbook.types import (
    Cell,
    CellDataType,
    CellValueUnion,
    ExcelErrorType,
)
from pe_bench.office.workbook.utils import CellRangeMapper, cell_coord_to_indices

logger = logging.getLogger(__name__)

TYPE_SHARED_STRING = "s"
TYPE_FORMULA_CACHE_STRING = "str"
TYPE_BOOLEAN = "b"
TYPE_ERROR = "e"
TYPE_NUMBER = "n"
TYPE_DATE = "d"

SHEET_MAIN_NS = Namespaces.SHEET_MAIN
CELL_TAG = "{%s}c" % SHEET_MAIN_NS
VALUE_TAG = "{%s}v" % SHEET_MAIN_NS
FORMULA_TAG = "{%s}f" % SHEET_MAIN_NS
INLINE_STRING_TAG = "{%s}is" % SHEET_MAIN_NS
RUN_TAG = "{%s}r" % SHEET_MAIN_NS
TEXT_TAG = "{%s}t" % SHEET_MAIN_NS
ROW_TAG = "{%s}row" % SHEET_MAIN_NS
DATA_TAG = "{%s}sheetData" % SHEET_MAIN_NS


def _parse_formula(
    formula_elem: Element,
    cell_ref: str,
    cell: Cell,
    group_id_to_master_cell_coord: dict[str, str],
):
    formula_type = formula_elem.get("t", None)
    formula_text = formula_elem.text
    if formula_type == "shared":
        group_id = formula_elem.get("si", None)
        if not group_id:
            return
        if formula_text:
            range_ref = formula_elem.get("ref", None)
            cell.f = f"={formula_text}"
            cell.fr = range_ref
            group_id_to_master_cell_coord[group_id] = cell_ref
            return
        master_cell_coord = group_id_to_master_cell_coord.get(group_id)
        if master_cell_coord is None:
            return
        cell.fm = master_cell_coord
        return
    elif formula_type == "array":
        if formula_text:
            cell.f = f"={formula_text}"
            cell.fa = True
            cell.fr = formula_elem.get("ref", None)
        return
    elif formula_type == "normal" or formula_text is not None:
        cell.f = f"={formula_text}"
        return
    elif formula_type == "dataTable":
        return
    else:
        cell.fa = True


def _parse_inline_string(inline_string_elem: Element) -> str | None:
    text_elem: Element | None = inline_string_elem.find(TEXT_TAG, None)
    if text_elem is not None and text_elem.text is not None:
        return text_elem.text
    runs: list[Element] = inline_string_elem.findall(RUN_TAG, None)
    parts = [
        run_text.text
        for run in runs
        if (run_text := run.find(TEXT_TAG, None)) is not None
        and run_text.text is not None
    ]
    return "".join(parts) if parts else None


def _parse_cell_value(
    value_elem: Element, data_type: str, shared_strings: list[str] | None
) -> tuple[CellValueUnion, CellDataType]:
    value_text: str = value_elem.text

    if data_type == TYPE_SHARED_STRING:
        if shared_strings:
            try:
                string_index = int(value_text)
                return shared_strings[string_index], CellDataType.TEXT
            except (ValueError, IndexError):
                return value_text, CellDataType.TEXT
        else:
            return value_text, CellDataType.TEXT
    elif data_type == TYPE_FORMULA_CACHE_STRING:
        return value_text, CellDataType.TEXT
    elif data_type == TYPE_BOOLEAN:
        return is_true(value_text), CellDataType.BOOLEAN
    elif data_type == TYPE_ERROR:
        try:
            return ExcelErrorType(value_text), CellDataType.ERROR
        except ValueError:
            logger.warning(
                {"type": "failed_to_parse_excel_error_type", "value_text": value_text}
            )
            return value_text, CellDataType.TEXT
    elif data_type == TYPE_DATE:
        try:
            date = datetime_to_serial_number(datetime.fromisoformat(value_text))
            return date, (
                CellDataType.NUMBER if isinstance(date, float) else CellDataType.ERROR
            )
        except ValueError:
            logger.warning(
                {"type": "failed_to_parse_excel_date", "value_text": value_text}
            )
            return value_text, CellDataType.TEXT
    elif data_type == TYPE_NUMBER:
        value = safe_float(value_text)
        if value.is_integer():
            value = int(value)
        return value, CellDataType.NUMBER
    else:
        raise AssertionError(f"Unknown data type: {data_type}")


def _parse_cell(
    cell_elem: Element,
    cell_ref: str,
    group_id_to_master_cell_coord: dict[str, str],
    shared_strings: list[str] | None = None,
) -> Cell | None:
    cell = Cell()

    style_id = cell_elem.get("s", None)
    if style_id:
        cell.s = int(style_id)

    data_type = cell_elem.get("t", CellDataType.NUMBER)

    value_elem: Element | None = None
    for child in cell_elem.iter(None):
        tag = child.tag
        if tag == FORMULA_TAG:
            _parse_formula(child, cell_ref, cell, group_id_to_master_cell_coord)
        elif tag == INLINE_STRING_TAG and cell.v is None:
            inline = _parse_inline_string(child)
            if inline:
                cell.v = inline
                cell.t = CellDataType.TEXT
        elif tag == VALUE_TAG:
            value_elem = child

    if (
        cell.v is None
        and value_elem is not None
        and (data_type != TYPE_NUMBER or (value_elem.text or "").strip())
    ):
        cell.v, cell.t = _parse_cell_value(value_elem, data_type, shared_strings)

    return cell


def _populate_array_formula_child_refs(cells: dict[str, Cell]) -> None:
    for master_coord, master_cell in cells.items():
        if not (master_cell.f is not None and master_cell.fa and master_cell.fr):
            continue

        try:
            child_coords, _ = CellRangeMapper(master_cell.fr).to_coords()
        except ValueError:
            logger.warning(
                {"type": "failed_to_parse_array_formula_range"},
                extra={"cell": master_coord, "range": master_cell.fr},
            )
            continue

        for child_coord in child_coords:
            if child_coord == master_coord:
                continue

            child_cell = cells.get(child_coord)
            if (
                child_cell is None
                or child_cell.f is not None
                or child_cell.fm is not None
            ):
                continue
            child_cell.fm = master_coord
            child_cell.fa = True


def parse_worksheet_cells(
    xml_content: bytes,
    shared_strings: list[str] | None = None,
    max_rows_and_columns: tuple[int, int] | None = None,
) -> tuple[Element, dict[str, Cell]]:
    cells: dict[str, Cell] = {}
    group_id_to_master_cell_coord: dict[str, str] = {}
    is_past_max_rows = False
    streamed = lxml.etree.iterparse(
        io.BytesIO(xml_content),
        events=("end",),
        tag=CELL_TAG,
        resolve_entities=False,
    )
    for _event, cell_elem in streamed:
        cell_ref = cell_elem.get("r")
        if cell_ref and not is_past_max_rows:
            is_within_limits = True
            if max_rows_and_columns:
                max_rows, max_columns = max_rows_and_columns
                row_idx, col_idx = cell_coord_to_indices(cell_ref)
                is_past_max_rows = row_idx >= max_rows
                is_within_limits = not is_past_max_rows and col_idx < max_columns
            if is_within_limits:
                cell = _parse_cell(
                    cell_elem, cell_ref, group_id_to_master_cell_coord, shared_strings
                )
                if cell is not None:
                    cells[cell_ref] = cell
        cell_elem.clear(keep_tail=True)
        row_elem = cell_elem.getparent()
        while row_elem is not None and cell_elem.getprevious() is not None:
            del row_elem[0]

    _populate_array_formula_child_refs(cells)

    return streamed.root, cells
