import re
from dataclasses import dataclass
from enum import StrEnum
from typing import Any

from pydantic import BaseModel, field_serializer, field_validator, model_validator

from pe_bench.office.workbook.types_document import (
    DocumentCoreProperties,
    ExtendedProperties,
)
from pe_bench.office.workbook.types_format import ConditionalFormatting
from pe_bench.office.workbook.types_sheet import (
    ColumnDimension,
    HeaderFooter,
    Hyperlink,
    PageMargins,
    PageSetup,
    PrintOptions,
    RowDimension,
    SheetFormatProperties,
    SheetProperties,
    SheetState,
    SheetView,
)
from pe_bench.office.workbook.types_style import Stylesheet
from pe_bench.office.workbook.types_workbook import (
    CalcProperties,
    DefinedName,
    FileSharing,
    FileVersion,
    WorkbookProperties,
    WorkbookProtection,
    WorkbookView,
)


def clean_docx_source_spans(spans: list[str] | None) -> list[str] | None:
    if not spans:
        return spans
    cleaned: list[str] = []
    for span in spans:
        span = re.sub(r"^#+\s*", "", span, flags=re.MULTILINE)
        span = re.sub(r"\n{2,}", "\n", span)
        span = span.strip()
        cleaned.append(span)
    return cleaned


class Citation(BaseModel):
    file_id: str | None = None
    filename: str | None = None
    title: str | None = None
    source_spans: list[str] | None = None
    source_spans_pages: list[int] | None = None
    cell_range: str | None = None
    sheet_name: str | None = None
    slide_number: int | None = None
    line_range: str | None = None
    source_label: str | None = None
    url: str | None = None
    mcp_id: str | None = None
    as_of: str | None = None

    @model_validator(mode="before")
    @classmethod
    def migrate_old_format(cls, data: Any) -> Any:
        if not isinstance(data, dict):
            return data
        source_spans_pages = data.get("source_spans_pages")
        if source_spans_pages and isinstance(source_spans_pages[0], list):
            data["source_spans_pages"] = [p[0] if p else 0 for p in source_spans_pages]
        return data

    @model_validator(mode="after")
    def require_file_xor_mcp(self) -> "Citation":
        has_file = self.file_id is not None
        has_mcp = self.source_label is not None
        if has_file == has_mcp:
            raise ValueError(
                "exactly one of file_id (file citation) or source_label "
                "(MCP citation) must be set"
            )
        return self

    @model_validator(mode="after")
    def clean_docx_spans(self) -> "Citation":
        if (
            self.filename is not None
            and self.filename.lower().endswith(".docx")
            and self.source_spans
        ):
            object.__setattr__(
                self, "source_spans", clean_docx_source_spans(self.source_spans)
            )
        return self


class CellDataType(StrEnum):
    TEXT = "t"
    NUMBER = "n"
    BOOLEAN = "b"
    ERROR = "e"


class ExcelErrorType(StrEnum):
    NULL = "#NULL!"
    DIV_ZERO = "#DIV/0!"
    VALUE = "#VALUE!"
    REF = "#REF!"
    NAME = "#NAME?"
    NUM = "#NUM!"
    NA = "#N/A"
    GETTING_DATA = "#GETTING_DATA"
    CALC = "#CALC!"
    SPILL = "#SPILL!"
    FIELD = "#FIELD!"
    BLOCK = "#BLOCK!"


CellValueUnion = ExcelErrorType | str | int | float | bool


@dataclass(slots=True)
class Cell:
    v: CellValueUnion | None = None
    f: str | None = None
    t: CellDataType | None = None
    c: Citation | None = None
    s: int | None = None

    fm: str | None = None
    fr: str | None = None
    fa: bool | None = None

    def is_formula(self) -> bool:
        return self.f is not None or self.fm is not None

    def set_value(self, data: CellValueUnion) -> None:
        if isinstance(data, ExcelErrorType):
            self.v = data
            self.t = CellDataType.ERROR
        elif isinstance(data, str):
            self.v = data
            self.t = CellDataType.TEXT
        elif isinstance(data, bool):
            self.v = data
            self.t = CellDataType.BOOLEAN
        elif isinstance(data, (int, float)):
            self.v = data
            self.t = CellDataType.NUMBER
        else:
            raise ValueError(f"Invalid cell value: {data}")

    def set_shared_formula_master(self, master_coord: str) -> None:
        self.fm = master_coord
        self.f = None
        self.fr = None
        self.v = None
        self.t = None
        self.c = None
        self.fa = None

    def clear_data(self) -> None:
        self.v = None
        self.f = None
        self.fm = None
        self.fr = None
        self.t = None
        self.c = None
        self.fa = None

    @classmethod
    def from_value(cls, data: CellValueUnion, formula: str | None = None) -> "Cell":
        if isinstance(data, ExcelErrorType):
            return cls(v=data, t=CellDataType.ERROR, f=formula)
        elif isinstance(data, str):
            return cls(v=data, t=CellDataType.TEXT, f=formula)
        if isinstance(data, bool):
            return cls(v=data, t=CellDataType.BOOLEAN, f=formula)
        if isinstance(data, (int, float)):
            return cls(v=data, t=CellDataType.NUMBER, f=formula)
        raise ValueError(f"Invalid cell value: {data}")

    @classmethod
    def from_error(
        cls, error_type: ExcelErrorType, formula: str | None = None
    ) -> "Cell":
        return cls(v=error_type, t=CellDataType.ERROR, f=formula)


class SheetReference(BaseModel):
    file_id: str
    sheet_name: str


class Sheet(BaseModel):
    cells: dict[str, Cell]
    merged_cell_ranges: list[str] | None = None

    row_dims_ids: dict[int, int] | None = None
    row_dims: list[RowDimension] | None = None
    columns_dims: dict[str, ColumnDimension] | None = None
    format_properties: SheetFormatProperties | None = None

    state: SheetState | None = None
    properties: SheetProperties | None = None
    conditional_formattings: list[ConditionalFormatting] | None = None
    views: list[SheetView] | None = None

    print_options: PrintOptions | None = None
    page_margins: PageMargins | None = None
    page_setup: PageSetup | None = None
    header_footer: HeaderFooter | None = None

    hyperlinks: dict[str, Hyperlink] | None = None

    reference: SheetReference | None = None
    index: int | None = None
    is_truncated: bool | None = None
    max_rows: int | None = None
    max_columns: int | None = None

    @field_serializer("row_dims_ids")
    def serialize_row_dims_ids(
        self, row_dims_ids: dict[int, int] | None
    ) -> dict[str, int] | None:
        if row_dims_ids is None:
            return None
        return {str(k): v for k, v in row_dims_ids.items()}

    @field_validator("row_dims_ids", mode="before")
    @classmethod
    def validate_row_dims_ids(
        cls, v: dict[str, int] | dict[int, int] | None
    ) -> dict[int, int] | None:
        if v is None:
            return None
        if isinstance(v, dict):
            return {int(k) if isinstance(k, str) else k: v for k, v in v.items()}
        return v

    @field_validator("cells", mode="before")
    @classmethod
    def validate_cell_data(
        cls, data: dict[str, Cell | dict[str, object]]
    ) -> dict[str, Cell | dict[str, object]]:
        for cell_data in data.values():
            if isinstance(cell_data, dict):
                if cell_data.get("t") == CellDataType.ERROR:
                    if not isinstance(cell_data.get("v"), ExcelErrorType):
                        cell_data["v"] = ExcelErrorType(cell_data["v"])
        return data


class Workbook(BaseModel):
    sheets: dict[str, Sheet]
    active_sheet: str | None = None

    style_sheet: Stylesheet | None = None
    theme: str | None = None
    properties: WorkbookProperties | None = None
    calculation: CalcProperties | None = None
    views: list[WorkbookView] | None = None
    defined_names: list[DefinedName] | None = None
    file_version: FileVersion | None = None
    file_sharing: FileSharing | None = None
    protection: WorkbookProtection | None = None

    core_properties: DocumentCoreProperties | None = None
    extended_properties: ExtendedProperties | None = None

    is_truncated: bool | None = None
    has_circular_deps: bool | None = None

    @property
    def date_1904(self) -> bool:
        return bool(self.properties is not None and self.properties.date_1904)
