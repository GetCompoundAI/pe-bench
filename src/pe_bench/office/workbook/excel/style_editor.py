from copy import deepcopy
from typing import Generic, TypeVar

from pydantic import BaseModel

from pe_bench.office.workbook.excel.styles.number_format_builtin import (
    BUILTIN_NUM_FORMATS_ID2STR,
    BUILTIN_NUM_FORMATS_STR2ID,
    number_format_is_datetime,
)
from pe_bench.office.workbook.types import Cell, Workbook
from pe_bench.office.workbook.types_style import (
    Alignment,
    Border,
    CellStyle,
    Font,
    GradientFill,
    PatternFill,
    Protection,
    Stylesheet,
)


class ExcelCellStyle(BaseModel):
    number_format: str | None = None
    font: Font | None = None
    fill: PatternFill | GradientFill | None = None
    border: Border | None = None
    alignment: Alignment | None = None
    protection: Protection | None = None

    def is_default(self) -> bool:
        return (
            self.number_format is None
            and self.font is None
            and self.fill is None
            and self.border is None
            and self.alignment is None
            and self.protection is None
        )


def merge_font(base_font: Font | None, other_font: Font | None) -> Font | None:
    if base_font and other_font:
        return Font.model_validate(
            base_font.model_dump(exclude_none=True)
            | other_font.model_dump(exclude_none=True)
        )
    return other_font or base_font


def merge_border(
    base_border: Border | None, other_border: Border | None
) -> Border | None:
    if base_border and other_border:
        return Border.model_validate(
            base_border.model_dump(exclude_none=True)
            | other_border.model_dump(exclude_none=True)
        )
    return other_border or base_border


def merge_alignment(
    base_alignment: Alignment | None, other_alignment: Alignment | None
) -> Alignment | None:
    if base_alignment and other_alignment:
        return Alignment.model_validate(
            base_alignment.model_dump(exclude_none=True)
            | other_alignment.model_dump(exclude_none=True)
        )
    return other_alignment or base_alignment


def merge_protection(
    base_protection: Protection | None, other_protection: Protection | None
) -> Protection | None:
    if base_protection and other_protection:
        return Protection.model_validate(
            base_protection.model_dump(exclude_none=True)
            | other_protection.model_dump(exclude_none=True)
        )
    return other_protection or base_protection


def merge_style(
    base_style: ExcelCellStyle | None, other_style: ExcelCellStyle | None
) -> ExcelCellStyle | None:
    if base_style and other_style:
        return ExcelCellStyle(
            number_format=other_style.number_format or base_style.number_format,
            font=merge_font(base_style.font, other_style.font),
            fill=other_style.fill or base_style.fill,
            border=merge_border(base_style.border, other_style.border),
            alignment=merge_alignment(base_style.alignment, other_style.alignment),
            protection=merge_protection(base_style.protection, other_style.protection),
        )
    return other_style or base_style


T = TypeVar("T", bound=BaseModel)


def _get_hash(element: BaseModel) -> str:
    return element.model_dump_json(exclude_none=True)


class IndexedList(Generic[T]):
    def __init__(self, elements: list[T]):
        self._elements = elements
        self._index_map: dict[str, int] = {
            _get_hash(element): i for i, element in enumerate(elements)
        }

    def get_or_add(self, element: T) -> int:
        if _get_hash(element) not in self._index_map:
            self._index_map[_get_hash(element)] = len(self._elements)
            self._elements.append(element)
        return self._index_map[_get_hash(element)]

    def get(self, index: int) -> T | None:
        if index < 0 or index >= len(self._elements):
            return None
        return self._elements[index]

    def __len__(self) -> int:
        return len(self._elements)

    def clone(self, elements: list[T]) -> "IndexedList[T]":
        clone = self.__class__.__new__(self.__class__)
        clone._elements = elements
        clone._index_map = deepcopy(self._index_map)
        return clone


class StylesheetEditor:
    def __init__(self, workbook: Workbook):
        self._workbook = workbook
        if self._workbook.style_sheet is None:
            self._workbook.style_sheet = Stylesheet()
        stylesheet = self._workbook.style_sheet

        if stylesheet.num_fmts is None:
            stylesheet.num_fmts = {}
        if stylesheet.fonts is None:
            stylesheet.fonts = [Font()]
        if stylesheet.fills is None:
            stylesheet.fills = [PatternFill()]
        if stylesheet.borders is None:
            stylesheet.borders = [Border()]
        if stylesheet.alignments is None:
            stylesheet.alignments = [Alignment()]
        if stylesheet.protections is None:
            stylesheet.protections = [Protection()]
        if stylesheet.cell_styles is None:
            stylesheet.cell_styles = [CellStyle()]

        self._id_2_num_fmt: dict[int, str] = stylesheet.num_fmts
        self._num_fmt_2_id: dict[str, int] = {
            v: k for k, v in self._id_2_num_fmt.items()
        }
        self._max_num_fmt_id = max(self._num_fmt_2_id.values(), default=163)

        self._fonts: IndexedList[Font] = IndexedList(stylesheet.fonts)
        self._fills: IndexedList[PatternFill | GradientFill] = IndexedList(
            stylesheet.fills
        )
        self._borders: IndexedList[Border] = IndexedList(stylesheet.borders)
        self._alignments: IndexedList[Alignment] = IndexedList(stylesheet.alignments)
        self._protections: IndexedList[Protection] = IndexedList(stylesheet.protections)
        self._cell_styles: IndexedList[CellStyle] = IndexedList(stylesheet.cell_styles)

    def clone(self, workbook: Workbook) -> "StylesheetEditor":
        clone = self.__class__.__new__(self.__class__)
        clone._workbook = workbook

        stylesheet = workbook.style_sheet
        assert stylesheet is not None
        assert stylesheet.num_fmts is not None
        assert stylesheet.fonts is not None
        assert stylesheet.fills is not None
        assert stylesheet.borders is not None
        assert stylesheet.alignments is not None
        assert stylesheet.protections is not None
        assert stylesheet.cell_styles is not None

        clone._id_2_num_fmt = stylesheet.num_fmts
        clone._num_fmt_2_id = deepcopy(self._num_fmt_2_id)
        clone._max_num_fmt_id = self._max_num_fmt_id

        clone._fonts = self._fonts.clone(stylesheet.fonts)
        clone._fills = self._fills.clone(stylesheet.fills)
        clone._borders = self._borders.clone(stylesheet.borders)
        clone._alignments = self._alignments.clone(stylesheet.alignments)
        clone._protections = self._protections.clone(stylesheet.protections)
        clone._cell_styles = self._cell_styles.clone(stylesheet.cell_styles)

        return clone

    def _get_num_fmt_id(self, num_fmt: str) -> int:
        if num_fmt in BUILTIN_NUM_FORMATS_STR2ID:
            return BUILTIN_NUM_FORMATS_STR2ID[num_fmt]
        if num_fmt in self._num_fmt_2_id:
            return self._num_fmt_2_id[num_fmt]
        num_fmt_id = self._max_num_fmt_id + 1
        self._num_fmt_2_id[num_fmt] = num_fmt_id
        self._id_2_num_fmt[num_fmt_id] = num_fmt
        self._max_num_fmt_id = num_fmt_id
        return num_fmt_id

    def _get_or_add_style_id(self, style: ExcelCellStyle) -> int:
        num_fmt_id = (
            self._get_num_fmt_id(style.number_format) if style.number_format else None
        )
        font_id = self._fonts.get_or_add(style.font) if style.font else None
        fill_id = self._fills.get_or_add(style.fill) if style.fill else None
        border_id = self._borders.get_or_add(style.border) if style.border else None
        alignment_id = (
            self._alignments.get_or_add(style.alignment) if style.alignment else None
        )
        protection_id = (
            self._protections.get_or_add(style.protection) if style.protection else None
        )

        cell_style = CellStyle(
            num_fmt_id=num_fmt_id,
            font_id=font_id,
            fill_id=fill_id,
            border_id=border_id,
            alignment_id=alignment_id,
            protection_id=protection_id,
        )
        style_id = self._cell_styles.get_or_add(cell_style)
        return style_id

    def set_cells_style(
        self, sheet_name: str, coords: list[str], style: ExcelCellStyle
    ) -> None:
        sheet = self._workbook.sheets[sheet_name]

        if style.is_default():
            self.delete_cell_styles(sheet_name, coords)
            return

        style_id = self._get_or_add_style_id(style)
        for cell_coord in coords:
            if cell_coord not in sheet.cells:
                sheet.cells[cell_coord] = Cell(s=style_id)
            else:
                sheet.cells[cell_coord].s = style_id

    def merge_cells_style(
        self, sheet_name: str, coords: list[str], style: ExcelCellStyle
    ) -> None:
        sheet = self._workbook.sheets[sheet_name]
        for cell_coord in coords:
            if cell_coord not in sheet.cells:
                sheet.cells[cell_coord] = Cell()
            cell = sheet.cells[cell_coord]
            if cell.s and (cell_style := self.get_cell_style(cell.s)):
                cell.s = self._get_or_add_style_id(
                    merge_style(cell_style, style) or ExcelCellStyle()
                )
            else:
                cell.s = self._get_or_add_style_id(style)

    def set_cell_number_formats(
        self, sheet_name: str, coords: list[str], num_fmt: str
    ) -> None:
        sheet = self._workbook.sheets[sheet_name]
        for cell_coord in coords:
            if cell_coord not in sheet.cells:
                sheet.cells[cell_coord] = Cell()
            cell = sheet.cells[cell_coord]
            if cell.s and (style := self.get_cell_style(cell.s)):
                style.number_format = num_fmt
            else:
                style = ExcelCellStyle(number_format=num_fmt)
            cell.s = self._get_or_add_style_id(style)

    def num_fmt_is_datetime(self, style_id: int) -> bool:
        cell_style = self._cell_styles.get(style_id)
        if cell_style is None or cell_style.num_fmt_id is None:
            return False
        if cell_style.num_fmt_id in self._id_2_num_fmt:
            return number_format_is_datetime(self._id_2_num_fmt[cell_style.num_fmt_id])
        elif cell_style.num_fmt_id in BUILTIN_NUM_FORMATS_ID2STR:
            return number_format_is_datetime(
                BUILTIN_NUM_FORMATS_ID2STR[cell_style.num_fmt_id]
            )
        return False

    def get_cell_style(self, style_id: int) -> ExcelCellStyle | None:
        cell_style = self._cell_styles.get(style_id)
        if cell_style is None:
            return None

        number_format = None
        if (num_fmt_id := cell_style.num_fmt_id) is not None:
            if num_fmt_id in self._id_2_num_fmt:
                number_format = self._id_2_num_fmt[num_fmt_id]
            elif num_fmt_id in BUILTIN_NUM_FORMATS_ID2STR:
                number_format = BUILTIN_NUM_FORMATS_ID2STR[num_fmt_id]
            else:
                number_format = "General"

        font = (
            self._fonts.get(cell_style.font_id)
            if cell_style.font_id is not None
            else None
        )
        fill = (
            self._fills.get(cell_style.fill_id)
            if cell_style.fill_id is not None
            else None
        )
        border = (
            self._borders.get(cell_style.border_id)
            if cell_style.border_id is not None
            else None
        )
        alignment = (
            self._alignments.get(cell_style.alignment_id)
            if cell_style.alignment_id is not None
            else None
        )
        protection = (
            self._protections.get(cell_style.protection_id)
            if cell_style.protection_id is not None
            else None
        )

        return ExcelCellStyle(
            number_format=number_format,
            font=font,
            fill=fill,
            border=border,
            alignment=alignment,
            protection=protection,
        )

    def get_all_styles(self) -> dict[int, ExcelCellStyle]:
        return {
            style_id: self.get_cell_style(style_id)
            for style_id in range(len(self._cell_styles))
        }  # type: ignore

    def delete_cell_styles(self, sheet_name: str, coords: list[str]) -> set[str]:
        sheet = self._workbook.sheets.get(sheet_name)
        if not sheet:
            return set()

        deleted_coords = set()
        for cell_coord in coords:
            if cell_coord in sheet.cells:
                sheet.cells[cell_coord].s = None
                deleted_coords.add(cell_coord)
        return deleted_coords
