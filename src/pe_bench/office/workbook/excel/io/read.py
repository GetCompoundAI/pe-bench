import io
import logging

from pe_bench.office.workbook.excel.io.archive import ExcelArchiveReader
from pe_bench.office.workbook.excel.io.parsers.packaging import (
    parse_core_properties,
    parse_extended_properties,
)
from pe_bench.office.workbook.excel.io.parsers.relationships import parse_relationships
from pe_bench.office.workbook.excel.io.parsers.shared_strings import (
    parse_shared_strings,
)
from pe_bench.office.workbook.excel.io.parsers.sheet import parse_worksheet
from pe_bench.office.workbook.excel.io.parsers.styles import parse_stylesheet
from pe_bench.office.workbook.excel.io.parsers.workbook import parse_workbook
from pe_bench.office.workbook.types import Sheet, Workbook

logger = logging.getLogger(__name__)


def excel_bytes_to_workbook(
    file_bytes: bytes,
    lean: bool = False,
    max_rows_and_columns: tuple[int, int] | None = None,
) -> Workbook:

    with ExcelArchiveReader(io.BytesIO(file_bytes)) as archive:
        workbook_rels_xml = archive.read_workbook_rels()
        workbook_rels = parse_relationships(workbook_rels_xml)
        sheet_paths: dict[str, str] = {
            rel.id: rel.target
            for rel in workbook_rels
            if "worksheet" in rel.type.lower()
        }

        workbook_xml = archive.read_workbook()
        workbook, sheet_metas = parse_workbook(workbook_xml)

        shared_strings_xml = archive.read_shared_strings()
        shared_strings: list[str] = (
            parse_shared_strings(shared_strings_xml) if shared_strings_xml else []
        )

        if not lean:
            styles_xml = archive.read_styles()
            workbook.style_sheet = parse_stylesheet(styles_xml) if styles_xml else None

            theme_xml = archive.read_theme()
            workbook.theme = theme_xml.decode() if theme_xml else None

            core_props_xml = archive.read_core_properties()
            if core_props_xml:
                workbook.core_properties = parse_core_properties(core_props_xml)

            app_props_xml = archive.read_app_properties()
            if app_props_xml:
                workbook.extended_properties = parse_extended_properties(app_props_xml)

        sheets: dict[str, Sheet] = {}
        for i, sheet_meta in enumerate(sheet_metas):
            sheet_path = sheet_paths.get(sheet_meta.rid)
            if not sheet_path:
                logger.warning(
                    {"type": "sheet_path_not_found", "sheet_meta": sheet_meta}
                )
                continue

            if sheet_path.startswith("xl/"):
                pass
            elif sheet_path.startswith("/xl/"):
                sheet_path = sheet_path[1:]
            else:
                sheet_path = f"xl/{sheet_path}"

            worksheet_rels_xml = archive.read_worksheet_rels(i + 1)
            if worksheet_rels_xml:
                worksheet_rels = parse_relationships(worksheet_rels_xml)
            else:
                worksheet_rels = None

            worksheet_xml = archive.read_worksheet(sheet_path)
            sheet = parse_worksheet(
                worksheet_xml, shared_strings, worksheet_rels, max_rows_and_columns
            )
            sheet.index = i

            if sheet_meta.state:
                sheet.state = sheet_meta.state

            sheets[sheet_meta.name] = sheet

        workbook.sheets = sheets
        if lean:
            workbook.is_truncated = True
        return workbook
