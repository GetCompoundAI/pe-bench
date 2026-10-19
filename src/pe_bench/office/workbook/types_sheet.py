from enum import StrEnum

from pydantic import BaseModel

from pe_bench.office.workbook.types_style import Color


class Hyperlink(BaseModel):
    url: str | None = None
    cell_coord: str | None = None
    display: str | None = None
    tooltip: str | None = None


class Dimension(BaseModel):
    hidden: bool | None = None
    outline_level: int | None = None
    collapsed: bool | None = None
    style_id: int | None = None


class RowDimension(Dimension):
    height: float | None = None
    thick_bottom: bool | None = None
    thick_top: bool | None = None

    def __hash__(self) -> int:
        return hash(
            (
                self.height,
                self.thick_bottom,
                self.thick_top,
                self.outline_level,
                self.collapsed,
                self.style_id,
                self.hidden,
            )
        )

    def is_empty(self) -> bool:
        return (
            self.height is None
            and self.thick_bottom is None
            and self.thick_top is None
            and self.outline_level is None
            and self.collapsed is None
            and self.style_id is None
            and self.hidden is None
        )


class ColumnDimension(Dimension):
    width: float | None = None
    best_fit: bool | None = None
    min: int | None = None
    max: int | None = None


class SheetFormatProperties(BaseModel):
    base_col_width: int | None = None
    default_col_width: float | None = None
    default_row_height: float | None = None
    custom_height: bool | None = None
    zero_height: bool | None = None
    thick_top: bool | None = None
    thick_bottom: bool | None = None
    outline_level_row: int | None = None
    outline_level_col: int | None = None


class SheetState(StrEnum):
    VISIBLE = "visible"
    HIDDEN = "hidden"
    VERY_HIDDEN = "veryHidden"


class SheetViewType(StrEnum):
    NORMAL = "normal"
    PAGE_BREAK_PREVIEW = "pageBreakPreview"
    PAGE_LAYOUT = "pageLayout"


class SheetViewPaneType(StrEnum):
    BOTTOM_RIGHT = "bottomRight"
    TOP_RIGHT = "topRight"
    BOTTOM_LEFT = "bottomLeft"
    TOP_LEFT = "topLeft"


class SheetViewSelection(BaseModel):
    pane: SheetViewPaneType | None = None
    active_cell: str | None = None
    active_cell_id: int | None = None
    sqref: str | None = None


class SheetViewPaneStateType(StrEnum):
    SPLIT = "split"
    FROZEN = "frozen"
    FROZEN_SPLIT = "frozenSplit"


class SheetViewPane(BaseModel):
    x_split: float | None = None
    y_split: float | None = None
    top_left_cell: str | None = None
    active_pane: SheetViewPaneType
    state: SheetViewPaneStateType


class SheetView(BaseModel):
    window_protection: bool | None = None
    show_formulas: bool | None = None
    show_grid_lines: bool | None = None
    show_row_col_headers: bool | None = None
    show_zeros: bool | None = None
    right_to_left: bool | None = None
    tab_selected: bool | None = None
    show_ruler: bool | None = None
    show_outline_symbols: bool | None = None
    default_grid_color: bool | None = None
    show_white_space: bool | None = None
    view: SheetViewType | None = None
    top_left_cell: str | None = None
    color_id: int | None = None
    zoom_scale: int | None = None
    zoom_scale_normal: int | None = None
    zoom_scale_sheet_layout_view: int | None = None
    zoom_scale_page_layout_view: int | None = None
    zoom_to_fit: bool | None = None
    workbook_view_id: int | None = None
    selection: list[SheetViewSelection] | None = None
    pane: SheetViewPane | None = None


class SheetProperties(BaseModel):
    tab_color: Color | None = None


class PrintOptions(BaseModel):
    horizontal_centered: bool | None = None
    vertical_centered: bool | None = None
    headings: bool | None = None
    grid_lines: bool | None = None
    grid_lines_set: bool | None = None


class PageMargins(BaseModel):
    left: float | None = None
    right: float | None = None
    top: float | None = None
    bottom: float | None = None
    header: float | None = None
    footer: float | None = None


class PageSetup(BaseModel):
    paper_size: int | None = None
    scale: int | None = None
    first_page_number: int | None = None
    fit_to_width: int | None = None
    fit_to_height: int | None = None
    orientation: str | None = None
    use_printer_defaults: bool | None = None
    black_and_white: bool | None = None
    draft: bool | None = None
    horizontal_dpi: int | None = None
    vertical_dpi: int | None = None
    copies: int | None = None


class HeaderFooter(BaseModel):
    different_first: bool | None = None
    different_odd_even: bool | None = None
    scale_with_doc: bool | None = None
    align_with_margins: bool | None = None

    odd_header: str | None = None
    odd_footer: str | None = None
    even_header: str | None = None
    even_footer: str | None = None
    first_header: str | None = None
    first_footer: str | None = None
