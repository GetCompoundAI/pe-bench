import logging

from pe_bench.office.workbook.excel.io.constants import Namespaces
from pe_bench.office.workbook.excel.io.parsers.cell import (
    DATA_TAG,
    ROW_TAG,
    parse_worksheet_cells,
)
from pe_bench.office.workbook.excel.io.parsers.hyperlink import parse_hyperlinks
from pe_bench.office.workbook.excel.io.parsers.relationships import Relationship
from pe_bench.office.workbook.excel.io.parsers.styles import parse_color
from pe_bench.office.workbook.excel.io.xml_functions import (
    Element,
    is_true,
    safe_float,
    safe_int,
)
from pe_bench.office.workbook.excel.sheet_property_editor import SheetPropertyEditor
from pe_bench.office.workbook.types import Sheet
from pe_bench.office.workbook.types_format import (
    ColorScale,
    ConditionalFormatting,
    DataBar,
    FormatObject,
    FormatObjectType,
    IconSet,
    IconSetType,
    Rule,
    RuleOperator,
    RuleType,
    TimePeriodType,
)
from pe_bench.office.workbook.types_sheet import (
    ColumnDimension,
    HeaderFooter,
    PageMargins,
    PageSetup,
    PrintOptions,
    RowDimension,
    SheetFormatProperties,
    SheetProperties,
    SheetView,
    SheetViewPane,
    SheetViewPaneStateType,
    SheetViewPaneType,
    SheetViewSelection,
    SheetViewType,
)
from pe_bench.office.workbook.utils import (
    cell_coord_to_indices,
    get_range_start_and_end,
    idx_to_column_letter,
)

logger = logging.getLogger(__name__)


SHEET_MAIN_NS = Namespaces.SHEET_MAIN

MERGE_TAG = "{%s}mergeCells" % SHEET_MAIN_NS
MERGE_CELL_TAG = "{%s}mergeCell" % SHEET_MAIN_NS
COL_TAG = "{%s}col" % SHEET_MAIN_NS
COLS_TAG = "{%s}cols" % SHEET_MAIN_NS
CF_TAG = "{%s}conditionalFormatting" % SHEET_MAIN_NS
PRINT_TAG = "{%s}printOptions" % SHEET_MAIN_NS
MARGINS_TAG = "{%s}pageMargins" % SHEET_MAIN_NS
PAGE_TAG = "{%s}pageSetup" % SHEET_MAIN_NS
HEADER_TAG = "{%s}headerFooter" % SHEET_MAIN_NS
ODD_HEADER_TAG = "{%s}oddHeader" % SHEET_MAIN_NS
ODD_FOOTER_TAG = "{%s}oddFooter" % SHEET_MAIN_NS
EVEN_HEADER_TAG = "{%s}evenHeader" % SHEET_MAIN_NS
EVEN_FOOTER_TAG = "{%s}evenFooter" % SHEET_MAIN_NS
FIRST_HEADER_TAG = "{%s}firstHeader" % SHEET_MAIN_NS
FIRST_FOOTER_TAG = "{%s}firstFooter" % SHEET_MAIN_NS
PROPERTIES_TAG = "{%s}sheetPr" % SHEET_MAIN_NS
VIEWS_TAG = "{%s}sheetViews" % SHEET_MAIN_NS
VIEW_TAG = "{%s}sheetView" % SHEET_MAIN_NS
PANE_TAG = "{%s}pane" % SHEET_MAIN_NS
SELECTION_TAG = "{%s}selection" % SHEET_MAIN_NS
FORMAT_TAG = "{%s}sheetFormatPr" % SHEET_MAIN_NS
DIMENSION_TAG = "{%s}dimension" % SHEET_MAIN_NS
TAB_COLOR_TAG = "{%s}tabColor" % SHEET_MAIN_NS

CF_RULE_TAG = "{%s}cfRule" % SHEET_MAIN_NS
CFVO_TAG = "{%s}cfvo" % SHEET_MAIN_NS
COLOR_SCALE_TAG = "{%s}colorScale" % SHEET_MAIN_NS
DATA_BAR_TAG = "{%s}dataBar" % SHEET_MAIN_NS
ICON_SET_TAG = "{%s}iconSet" % SHEET_MAIN_NS
FORMULA_TAG = "{%s}formula" % SHEET_MAIN_NS
COLOR_TAG = "{%s}color" % SHEET_MAIN_NS


def _parse_sheet_properties(sheet_elem: Element) -> SheetProperties | None:
    sheet_pr_elem: Element | None = sheet_elem.find(PROPERTIES_TAG, None)
    if sheet_pr_elem is None:
        return None

    properties = SheetProperties()

    tab_color_elem: Element | None = sheet_pr_elem.find(TAB_COLOR_TAG, None)
    if tab_color_elem is not None:
        properties.tab_color = parse_color(tab_color_elem)

    return properties


def _parse_sheet_view_selection(selection_elem: Element) -> SheetViewSelection:
    selection_data = SheetViewSelection()

    if pane := selection_elem.get("pane", None):
        selection_data.pane = SheetViewPaneType(pane)

    if active_cell := selection_elem.get("activeCell", None):
        selection_data.active_cell = active_cell

    if active_cell_id := selection_elem.get("activeCellId", None):
        selection_data.active_cell_id = safe_int(active_cell_id)

    if sqref := selection_elem.get("sqref", None):
        selection_data.sqref = sqref

    return selection_data


def _parse_sheet_view_pane(pane_elem: Element) -> SheetViewPane:
    active_pane = pane_elem.get("activePane", "bottomRight")
    active_pane = SheetViewPaneType(active_pane)

    state = pane_elem.get("state", "split")
    state = SheetViewPaneStateType(state)

    pane_data = SheetViewPane(active_pane=active_pane, state=state)

    if x_split := pane_elem.get("xSplit", None):
        pane_data.x_split = safe_float(x_split)

    if y_split := pane_elem.get("ySplit", None):
        pane_data.y_split = safe_float(y_split)

    if top_left_cell := pane_elem.get("topLeftCell", None):
        pane_data.top_left_cell = top_left_cell

    return pane_data


def _parse_sheet_view(view_elem: Element) -> SheetView:
    view_data = SheetView()

    if window_protection := view_elem.get("windowProtection", None):
        view_data.window_protection = is_true(window_protection)

    if show_formulas := view_elem.get("showFormulas", None):
        view_data.show_formulas = is_true(show_formulas)

    if show_grid_lines := view_elem.get("showGridLines", None):
        view_data.show_grid_lines = is_true(show_grid_lines)

    if show_row_col_headers := view_elem.get("showRowColHeaders", None):
        view_data.show_row_col_headers = is_true(show_row_col_headers)

    if show_zeros := view_elem.get("showZeros", None):
        view_data.show_zeros = is_true(show_zeros)

    if right_to_left := view_elem.get("rightToLeft", None):
        view_data.right_to_left = is_true(right_to_left)

    if tab_selected := view_elem.get("tabSelected", None):
        view_data.tab_selected = is_true(tab_selected)

    if show_ruler := view_elem.get("showRuler", None):
        view_data.show_ruler = is_true(show_ruler)

    if show_outline_symbols := view_elem.get("showOutlineSymbols", None):
        view_data.show_outline_symbols = is_true(show_outline_symbols)

    if default_grid_color := view_elem.get("defaultGridColor", None):
        view_data.default_grid_color = is_true(default_grid_color)

    if show_white_space := view_elem.get("showWhiteSpace", None):
        view_data.show_white_space = is_true(show_white_space)

    if zoom_to_fit := view_elem.get("zoomToFit", None):
        view_data.zoom_to_fit = is_true(zoom_to_fit)

    if view_type := view_elem.get("view", None):
        view_data.view = SheetViewType(view_type)

    if top_left_cell := view_elem.get("topLeftCell", None):
        view_data.top_left_cell = top_left_cell

    if color_id := view_elem.get("colorId", None):
        view_data.color_id = safe_int(color_id)

    if zoom_scale := view_elem.get("zoomScale", None):
        view_data.zoom_scale = safe_int(zoom_scale)

    if zoom_scale_normal := view_elem.get("zoomScaleNormal", None):
        view_data.zoom_scale_normal = safe_int(zoom_scale_normal)

    if zoom_scale_sheet_layout_view := view_elem.get("zoomScaleSheetLayoutView", None):
        view_data.zoom_scale_sheet_layout_view = safe_int(zoom_scale_sheet_layout_view)

    if zoom_scale_page_layout_view := view_elem.get("zoomScalePageLayoutView", None):
        view_data.zoom_scale_page_layout_view = safe_int(zoom_scale_page_layout_view)

    if workbook_view_id := view_elem.get("workbookViewId", None):
        view_data.workbook_view_id = safe_int(workbook_view_id)

    pane_elem: Element | None = view_elem.find(PANE_TAG, None)
    if pane_elem is not None:
        view_data.pane = _parse_sheet_view_pane(pane_elem)

    selections: list[SheetViewSelection] = [
        _parse_sheet_view_selection(selection_elem)
        for selection_elem in view_elem.findall(SELECTION_TAG, None)
    ]
    if selections:
        view_data.selection = selections

    return view_data


def _parse_sheet_views(sheet_elem: Element) -> list[SheetView] | None:
    views_elem: Element | None = sheet_elem.find(VIEWS_TAG, None)
    if views_elem is None:
        return None

    views: list[SheetView] = [
        _parse_sheet_view(view_elem) for view_elem in views_elem.findall(VIEW_TAG)
    ]  # type: ignore
    return views if views else None


def _parse_sheet_format_properties(sheet_elem: Element) -> SheetFormatProperties | None:
    format_elem: Element | None = sheet_elem.find(FORMAT_TAG, None)
    if format_elem is None:
        return None

    format_data = SheetFormatProperties()

    if base_col_width := format_elem.get("baseColWidth", None):
        format_data.base_col_width = safe_int(base_col_width)

    if default_col_width := format_elem.get("defaultColWidth", None):
        format_data.default_col_width = safe_float(default_col_width)

    if default_row_height := format_elem.get("defaultRowHeight", None):
        format_data.default_row_height = safe_float(default_row_height)

    if custom_height := format_elem.get("customHeight", None):
        format_data.custom_height = is_true(custom_height)

    if zero_height := format_elem.get("zeroHeight", None):
        format_data.zero_height = is_true(zero_height)

    if thick_top := format_elem.get("thickTop", None):
        format_data.thick_top = is_true(thick_top)

    if thick_bottom := format_elem.get("thickBottom", None):
        format_data.thick_bottom = is_true(thick_bottom)

    if outline_level_row := format_elem.get("outlineLevelRow", None):
        format_data.outline_level_row = safe_int(outline_level_row)

    if outline_level_col := format_elem.get("outlineLevelCol", None):
        format_data.outline_level_col = safe_int(outline_level_col)

    return format_data


def _parse_column_dimension(col_elem: Element) -> ColumnDimension:
    column_dim = ColumnDimension()

    if width := col_elem.get("width", None):
        column_dim.width = safe_float(width)

    if min_col := col_elem.get("min", None):
        column_dim.min = safe_int(min_col)

    if max_col := col_elem.get("max", None):
        column_dim.max = safe_int(max_col)

    if style_id := col_elem.get("style", None):
        column_dim.style_id = safe_int(style_id)

    if hidden := col_elem.get("hidden", None):
        column_dim.hidden = is_true(hidden)

    if best_fit := col_elem.get("bestFit", None):
        column_dim.best_fit = is_true(best_fit)

    if collapsed := col_elem.get("collapsed", None):
        column_dim.collapsed = is_true(collapsed)

    if outline_level := col_elem.get("outlineLevel", None):
        column_dim.outline_level = safe_int(outline_level)

    return column_dim


def _parse_column_dimensions(sheet_elem: Element) -> dict[str, ColumnDimension] | None:
    cols_elems: list[Element] = sheet_elem.findall(COLS_TAG, None)
    if not cols_elems:
        return None

    col_dimensions: dict[str, ColumnDimension] = {}
    for cols_elem in cols_elems:
        for col_elem in cols_elem.findall(COL_TAG, None):
            start_idx = safe_int(col_elem.get("min", None)) - 1
            end_idx = safe_int(col_elem.get("max", None)) - 1
            if start_idx < 0 or end_idx < 0:
                continue
            key = f"{idx_to_column_letter(start_idx)}:{idx_to_column_letter(end_idx)}"
            col_dimensions[key] = _parse_column_dimension(col_elem)

    return col_dimensions


def _parse_row_dimension(row_elem: Element) -> RowDimension:
    row_dim = RowDimension()

    if height := row_elem.get("ht", None):
        row_dim.height = safe_float(height)

    if style_id := row_elem.get("s", None):
        row_dim.style_id = safe_int(style_id)

    if hidden := row_elem.get("hidden", None):
        row_dim.hidden = is_true(hidden)

    if collapsed := row_elem.get("collapsed", None):
        row_dim.collapsed = is_true(collapsed)

    if thick_top := row_elem.get("thickTop", None):
        row_dim.thick_top = is_true(thick_top)

    if thick_bottom := row_elem.get("thickBot", None):
        row_dim.thick_bottom = is_true(thick_bottom)

    if outline_level := row_elem.get("outlineLevel", None):
        row_dim.outline_level = safe_int(outline_level)

    return row_dim


def _parse_row_dimensions(
    sheet_elem: Element,
) -> tuple[list[RowDimension] | None, dict[int, int] | None]:
    data_elem: Element | None = sheet_elem.find(DATA_TAG, None)
    if data_elem is None:
        return None, None

    row_dims: list[RowDimension] = []
    row_dims_ids: dict[int, int] = {}
    dimension_map: dict[RowDimension, int] = {}

    for row_elem in data_elem.findall(ROW_TAG):  # type: ignore
        row_r: str | None = row_elem.get("r")
        if not row_r:
            continue

        row_number = safe_int(row_r)
        if row_number is None:
            continue

        row_dim = _parse_row_dimension(row_elem)
        if row_dim.is_empty():
            continue

        dim_index = dimension_map.get(row_dim, None)
        if dim_index is None:
            dim_index = len(row_dims)
            row_dims.append(row_dim)
            dimension_map[row_dim] = dim_index

        row_dims_ids[row_number] = dim_index

    if not row_dims:
        return None, None

    return row_dims, row_dims_ids


def _parse_merged_ranges(sheet_elem: Element) -> list[str]:
    merged_ranges: list[str] = []

    merge_cells_elem: Element | None = sheet_elem.find(MERGE_TAG, None)
    if merge_cells_elem is not None:
        for merge_cell_elem in merge_cells_elem.findall(MERGE_CELL_TAG):  # type: ignore
            if ref := merge_cell_elem.get("ref"):
                merged_ranges.append(ref)

    return merged_ranges


def _parse_format_object(cfvo_elem: Element) -> FormatObject:
    type_val = cfvo_elem.get("type", "num")
    format_obj = FormatObject(type=FormatObjectType(type_val))

    if val := cfvo_elem.get("val", None):
        try:
            format_obj.val = float(val)
        except ValueError:
            format_obj.val = val

    if gte := cfvo_elem.get("gte", None):
        format_obj.gte = is_true(gte)

    return format_obj


def _parse_color_scale(color_scale_elem: Element) -> ColorScale:
    color_scale = ColorScale()

    cfvo_elems = color_scale_elem.findall(CFVO_TAG, None)
    if cfvo_elems:
        color_scale.cfvo = [_parse_format_object(cfvo_elem) for cfvo_elem in cfvo_elems]

    color_elems = color_scale_elem.findall(COLOR_TAG, None)
    if color_elems:
        colors = [parse_color(color_elem) for color_elem in color_elems]
        colors = [color for color in colors if color is not None]
        color_scale.color = colors

    return color_scale


def _parse_data_bar(data_bar_elem: Element) -> DataBar:
    data_bar = DataBar()

    if min_length := data_bar_elem.get("minLength", None):
        data_bar.min_length = safe_int(min_length) or 0

    if max_length := data_bar_elem.get("maxLength", None):
        data_bar.max_length = safe_int(max_length) or 100

    if show_value := data_bar_elem.get("showValue", None):
        data_bar.show_value = is_true(show_value)

    cfvo_elems = data_bar_elem.findall(CFVO_TAG, None)
    if cfvo_elems:
        data_bar.cfvo = [_parse_format_object(cfvo_elem) for cfvo_elem in cfvo_elems]

    color_elem = data_bar_elem.find(COLOR_TAG, None)
    if color_elem is not None:
        data_bar.color = parse_color(color_elem)

    return data_bar


def _parse_icon_set(icon_set_elem: Element) -> IconSet:
    icon_set = IconSet()

    if icon_set_value := icon_set_elem.get("iconSet", None):
        icon_set.icon_set = IconSetType(icon_set_value)

    if show_value := icon_set_elem.get("showValue", None):
        icon_set.show_value = is_true(show_value)

    if percent := icon_set_elem.get("percent", None):
        icon_set.percent = is_true(percent)

    if reverse := icon_set_elem.get("reverse", None):
        icon_set.reverse = is_true(reverse)

    cfvo_elems = icon_set_elem.findall(CFVO_TAG, None)
    if cfvo_elems:
        icon_set.cfvo = [_parse_format_object(cfvo_elem) for cfvo_elem in cfvo_elems]

    return icon_set


def _parse_cf_rule(rule_elem: Element) -> Rule:
    rule_type = rule_elem.get("type", "expression")
    priority = rule_elem.get("priority", "1")

    rule = Rule(type=RuleType(rule_type), priority=safe_int(priority) or 1)

    if dxf_id := rule_elem.get("dxfId", None):
        rule.dxf_id = safe_int(dxf_id)

    if stop_if_true := rule_elem.get("stopIfTrue", None):
        rule.stop_if_true = is_true(stop_if_true)

    if above_average := rule_elem.get("aboveAverage", None):
        rule.above_average = is_true(above_average)

    if percent := rule_elem.get("percent", None):
        rule.percent = is_true(percent)

    if bottom := rule_elem.get("bottom", None):
        rule.bottom = is_true(bottom)

    if operator := rule_elem.get("operator", None):
        rule.operator = RuleOperator(operator)

    if text := rule_elem.get("text", None):
        rule.text = text

    if time_period := rule_elem.get("timePeriod", None):
        rule.time_period = TimePeriodType(time_period)

    if rank := rule_elem.get("rank", None):
        rule.rank = safe_int(rank)

    if std_dev := rule_elem.get("stdDev", None):
        rule.std_dev = safe_int(std_dev) or 1

    if equal_average := rule_elem.get("equalAverage", None):
        rule.equal_average = is_true(equal_average)

    formula_elems = rule_elem.findall(FORMULA_TAG, None)
    if formula_elems:
        rule.formula = [formula_elem.text or "" for formula_elem in formula_elems]

    color_scale_elem = rule_elem.find(COLOR_SCALE_TAG, None)
    if color_scale_elem is not None:
        rule.color_scale = _parse_color_scale(color_scale_elem)

    data_bar_elem = rule_elem.find(DATA_BAR_TAG, None)
    if data_bar_elem is not None:
        rule.data_bar = _parse_data_bar(data_bar_elem)

    icon_set_elem = rule_elem.find(ICON_SET_TAG, None)
    if icon_set_elem is not None:
        rule.icon_set = _parse_icon_set(icon_set_elem)

    return rule


def _parse_conditional_formattings(
    sheet_elem: Element,
) -> list[ConditionalFormatting] | None:
    cf_elems = sheet_elem.findall(CF_TAG, None)
    if not cf_elems:
        return None

    conditional_formattings: list[ConditionalFormatting] = []

    for cf_elem in cf_elems:
        sqref = cf_elem.get("sqref", "")
        cell_ranges = sqref.split() if sqref else []

        pivot = cf_elem.get("pivot")
        pivot_bool = is_true(pivot) if pivot else None

        rules: list[Rule] = []
        for rule_elem in cf_elem.findall(CF_RULE_TAG):
            try:
                rule = _parse_cf_rule(rule_elem)
                rules.append(rule)
            except (ValueError, KeyError) as e:
                logger.warning({"type": "failed_to_parse_cf_rule", "error": e})
                continue

        if rules:
            conditional_formattings.append(
                ConditionalFormatting(
                    cell_ranges=cell_ranges, cf_rule=rules, pivot=pivot_bool
                )
            )

    return conditional_formattings if conditional_formattings else None


def _parse_worksheet_dimensions(root: Element) -> tuple[int | None, int | None]:
    dimension_elem: Element | None = root.find(DIMENSION_TAG, None)
    if dimension_elem is not None:
        ref: str | None = dimension_elem.get("ref", None)
        if ref:
            _, end_ref = get_range_start_and_end(ref)
            end_row, end_col = cell_coord_to_indices(end_ref)
            return (end_row + 1, end_col + 1)
    return None, None


def _parse_print_options(sheet_elem: Element) -> PrintOptions | None:

    print_elem: Element | None = sheet_elem.find(PRINT_TAG, None)
    if print_elem is None:
        return None

    options = PrintOptions()

    if val := print_elem.get("horizontalCentered", None):
        options.horizontal_centered = is_true(val)

    if val := print_elem.get("verticalCentered", None):
        options.vertical_centered = is_true(val)

    if val := print_elem.get("headings", None):
        options.headings = is_true(val)

    if val := print_elem.get("gridLines", None):
        options.grid_lines = is_true(val)

    if val := print_elem.get("gridLinesSet", None):
        options.grid_lines_set = is_true(val)

    return options


def _parse_page_margins(sheet_elem: Element) -> PageMargins | None:
    margins_elem: Element | None = sheet_elem.find(MARGINS_TAG, None)
    if margins_elem is None:
        return None

    margins = PageMargins()
    for attr in ["left", "right", "top", "bottom", "header", "footer"]:
        if val := margins_elem.get(attr, None):
            setattr(margins, attr if attr != "" else attr, safe_float(val))

    return margins


def _parse_page_setup(sheet_elem: Element) -> PageSetup | None:
    setup_elem: Element | None = sheet_elem.find(PAGE_TAG, None)
    if setup_elem is None:
        return None

    setup = PageSetup()

    if val := setup_elem.get("paperSize", None):
        setup.paper_size = safe_int(val)
    if val := setup_elem.get("scale", None):
        setup.scale = safe_int(val)
    if val := setup_elem.get("firstPageNumber", None):
        setup.first_page_number = safe_int(val)
    if val := setup_elem.get("fitToWidth", None):
        setup.fit_to_width = safe_int(val)
    if val := setup_elem.get("fitToHeight", None):
        setup.fit_to_height = safe_int(val)
    if val := setup_elem.get("horizontalDpi", None):
        setup.horizontal_dpi = safe_int(val)
    if val := setup_elem.get("verticalDpi", None):
        setup.vertical_dpi = safe_int(val)
    if val := setup_elem.get("copies", None):
        setup.copies = safe_int(val)

    if val := setup_elem.get("orientation", None):
        setup.orientation = val

    if val := setup_elem.get("usePrinterDefaults", None):
        setup.use_printer_defaults = is_true(val)
    if val := setup_elem.get("blackAndWhite", None):
        setup.black_and_white = is_true(val)
    if val := setup_elem.get("draft", None):
        setup.draft = is_true(val)

    return setup


def _parse_header_footer(sheet_elem: Element) -> HeaderFooter | None:
    hf_elem: Element | None = sheet_elem.find(HEADER_TAG, None)
    if hf_elem is None:
        return None

    hf = HeaderFooter()

    if val := hf_elem.get("differentFirst", None):
        hf.different_first = is_true(val)
    if val := hf_elem.get("differentOddEven", None):
        hf.different_odd_even = is_true(val)
    if val := hf_elem.get("scaleWithDoc", None):
        hf.scale_with_doc = is_true(val)
    if val := hf_elem.get("alignWithMargins", None):
        hf.align_with_margins = is_true(val)

    tag_map = [
        (ODD_HEADER_TAG, "odd_header"),
        (ODD_FOOTER_TAG, "odd_footer"),
        (EVEN_HEADER_TAG, "even_header"),
        (EVEN_FOOTER_TAG, "even_footer"),
        (FIRST_HEADER_TAG, "first_header"),
        (FIRST_FOOTER_TAG, "first_footer"),
    ]
    for tag, field in tag_map:
        elem: Element | None = hf_elem.find(tag, None)
        if elem is not None and elem.text is not None:
            setattr(hf, field, elem.text)

    return hf


def parse_worksheet(
    xml_content: bytes,
    shared_strings: list[str] | None = None,
    relationships: list[Relationship] | None = None,
    max_rows_and_columns: tuple[int, int] | None = None,
) -> Sheet:
    root, cells = parse_worksheet_cells(
        xml_content, shared_strings, max_rows_and_columns
    )
    sheet = Sheet(cells=cells)
    SheetPropertyEditor(sheet)

    sheet.properties = _parse_sheet_properties(root)

    sheet.views = _parse_sheet_views(root)

    sheet.format_properties = _parse_sheet_format_properties(root)

    sheet.columns_dims = _parse_column_dimensions(root)

    sheet.row_dims, sheet.row_dims_ids = _parse_row_dimensions(root)

    sheet.merged_cell_ranges = _parse_merged_ranges(root)

    sheet.conditional_formattings = _parse_conditional_formattings(root)

    sheet.hyperlinks = parse_hyperlinks(root, relationships or [])

    sheet.print_options = _parse_print_options(root)

    sheet.page_margins = _parse_page_margins(root)

    sheet.page_setup = _parse_page_setup(root)

    sheet.header_footer = _parse_header_footer(root)

    sheet.max_rows, sheet.max_columns = _parse_worksheet_dimensions(root)

    return sheet
