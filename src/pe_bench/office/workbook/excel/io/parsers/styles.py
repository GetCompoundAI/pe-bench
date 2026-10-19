from pydantic import BaseModel

from pe_bench.office.workbook.excel.io.constants import Namespaces
from pe_bench.office.workbook.excel.io.parsers.indexed_list import IndexedList
from pe_bench.office.workbook.excel.io.xml_functions import (
    Element,
    from_string,
    is_true,
    safe_float,
    safe_int,
)
from pe_bench.office.workbook.types_style import (
    Alignment,
    Border,
    BorderStyleType,
    CellStyle,
    Color,
    ColorStop,
    ColorType,
    DifferentialStyle,
    Font,
    FontAlignmentType,
    FontUnderlineType,
    GradientFill,
    GradientFillType,
    HorizontalAlignmentType,
    NamedStyle,
    PatternFill,
    PatternFillType,
    Protection,
    Side,
    Stylesheet,
    ThemeColorType,
    VerticalAlignmentType,
)

SHEET_MAIN_NS = Namespaces.SHEET_MAIN

NAME_TAG = "{%s}name" % SHEET_MAIN_NS
SIZE_TAG = "{%s}sz" % SHEET_MAIN_NS
CHARSET_TAG = "{%s}charset" % SHEET_MAIN_NS
FAMILY_TAG = "{%s}family" % SHEET_MAIN_NS
SCHEME_TAG = "{%s}scheme" % SHEET_MAIN_NS
BOLD_TAG = "{%s}b" % SHEET_MAIN_NS
ITALIC_TAG = "{%s}i" % SHEET_MAIN_NS
STRIKE_TAG = "{%s}strike" % SHEET_MAIN_NS
OUTLINE_TAG = "{%s}outline" % SHEET_MAIN_NS
SHADOW_TAG = "{%s}shadow" % SHEET_MAIN_NS
CONDENSE_TAG = "{%s}condense" % SHEET_MAIN_NS
EXTEND_TAG = "{%s}extend" % SHEET_MAIN_NS
UNDERLINE_TAG = "{%s}u" % SHEET_MAIN_NS
VERT_ALIGN_TAG = "{%s}vertAlign" % SHEET_MAIN_NS
COLOR_TAG = "{%s}color" % SHEET_MAIN_NS
PATTERN_FILL_TAG = "{%s}patternFill" % SHEET_MAIN_NS
FG_COLOR_TAG = "{%s}fgColor" % SHEET_MAIN_NS
BG_COLOR_TAG = "{%s}bgColor" % SHEET_MAIN_NS
GRADIENT_FILL_TAG = "{%s}gradientFill" % SHEET_MAIN_NS
STOP_TAG = "{%s}stop" % SHEET_MAIN_NS
LEFT_TAG = "{%s}left" % SHEET_MAIN_NS
RIGHT_TAG = "{%s}right" % SHEET_MAIN_NS
TOP_TAG = "{%s}top" % SHEET_MAIN_NS
BOTTOM_TAG = "{%s}bottom" % SHEET_MAIN_NS
DIAGONAL_TAG = "{%s}diagonal" % SHEET_MAIN_NS
START_TAG = "{%s}start" % SHEET_MAIN_NS
END_TAG = "{%s}end" % SHEET_MAIN_NS
HORIZONTAL_TAG = "{%s}horizontal" % SHEET_MAIN_NS
VERTICAL_TAG = "{%s}vertical" % SHEET_MAIN_NS
ALIGNMENT_TAG = "{%s}alignment" % SHEET_MAIN_NS
PROTECTION_TAG = "{%s}protection" % SHEET_MAIN_NS
FONT_TAG = "{%s}font" % SHEET_MAIN_NS
FILL_TAG = "{%s}fill" % SHEET_MAIN_NS
BORDER_TAG = "{%s}border" % SHEET_MAIN_NS
NUM_FMT_TAG = "{%s}numFmt" % SHEET_MAIN_NS
NUM_FMTS_TAG = "{%s}numFmts" % SHEET_MAIN_NS
FONTS_TAG = "{%s}fonts" % SHEET_MAIN_NS
FILLS_TAG = "{%s}fills" % SHEET_MAIN_NS
BORDERS_TAG = "{%s}borders" % SHEET_MAIN_NS
CELL_STYLE_XFS_TAG = "{%s}cellStyleXfs" % SHEET_MAIN_NS
CELL_XFS_TAG = "{%s}cellXfs" % SHEET_MAIN_NS
CELL_STYLES_TAG = "{%s}cellStyles" % SHEET_MAIN_NS
CELL_STYLE_TAG = "{%s}cellStyle" % SHEET_MAIN_NS
DXFS_TAG = "{%s}dxfs" % SHEET_MAIN_NS
DXF_TAG = "{%s}dxf" % SHEET_MAIN_NS
XF_TAG = "{%s}xf" % SHEET_MAIN_NS


class _CellStyleXf(BaseModel):
    num_fmt_id: int
    font_id: int
    fill_id: int
    border_id: int
    xf_id: int | None = None

    quote_prefix: bool | None = None
    pivot_button: bool | None = None

    apply_number_format: bool | None = None
    apply_font: bool | None = None
    apply_fill: bool | None = None
    apply_border: bool | None = None
    apply_alignment: bool | None = None
    apply_protection: bool | None = None

    alignment: Alignment | None = None
    protection: Protection | None = None


class _NamedCellStyle(BaseModel):
    name: str
    xf_id: int
    builtin_id: int | None = None
    i_level: int | None = None
    hidden: bool | None = None
    custom_builtin: bool | None = None


def _parse_num_fmt(num_fmt_elem: Element) -> tuple[int, str]:
    num_fmt_id = safe_int(num_fmt_elem.get("numFmtId", None))
    format_code = num_fmt_elem.get("formatCode", "General")
    return num_fmt_id, format_code


def parse_color(color_elem: Element) -> Color | None:
    tint = color_elem.get("tint", None)
    tint = safe_float(tint) if tint else None

    if rgb := color_elem.get("rgb", None):
        return Color(type=ColorType.RGB, rgb=rgb, tint=tint)

    if theme := color_elem.get("theme", None):
        return Color(
            type=ColorType.THEME, theme=ThemeColorType(safe_int(theme)), tint=tint
        )

    if indexed := color_elem.get("indexed", None):
        return Color(type=ColorType.INDEXED, indexed=safe_int(indexed), tint=tint)

    if is_true(color_elem.get("auto", None)):
        return Color(type=ColorType.AUTO, auto=True)

    return None


def _parse_font(font_elem: Element) -> Font:
    font = Font()

    name_elem: Element | None = font_elem.find(NAME_TAG, None)
    if name_elem is not None:
        if val := name_elem.get("val", None):
            font.name = val

    size_elem: Element | None = font_elem.find(SIZE_TAG, None)
    if size_elem is not None:
        if val := size_elem.get("val", None):
            font.size = safe_float(val)

    charset_elem: Element | None = font_elem.find(CHARSET_TAG, None)
    if charset_elem is not None:
        if val := charset_elem.get("val", None):
            font.charset = safe_int(val)

    family_elem: Element | None = font_elem.find(FAMILY_TAG, None)
    if family_elem is not None:
        if val := family_elem.get("val", None):
            font.family = safe_int(val)

    scheme_elem: Element | None = font_elem.find(SCHEME_TAG, None)
    if scheme_elem is not None:
        font.scheme = scheme_elem.get("val", None)

    bold_elem: Element | None = font_elem.find(BOLD_TAG, None)
    if bold_elem is not None:
        font.bold = True

    italic_elem: Element | None = font_elem.find(ITALIC_TAG, None)
    if italic_elem is not None:
        font.italic = True

    strike_elem: Element | None = font_elem.find(STRIKE_TAG, None)
    if strike_elem is not None:
        font.strike = True

    outline_elem: Element | None = font_elem.find(OUTLINE_TAG, None)
    if outline_elem is not None:
        font.outline = True

    shadow_elem: Element | None = font_elem.find(SHADOW_TAG, None)
    if shadow_elem is not None:
        font.shadow = True

    condense_elem: Element | None = font_elem.find(CONDENSE_TAG, None)
    if condense_elem is not None:
        font.condense = True

    extend_elem: Element | None = font_elem.find(EXTEND_TAG, None)
    if extend_elem is not None:
        font.extend = True

    u_elem: Element | None = font_elem.find(UNDERLINE_TAG, None)
    if u_elem is not None:
        if val := u_elem.get("val", None):
            font.underline = FontUnderlineType(val)
        else:
            font.underline = FontUnderlineType.SINGLE

    vert_align_elem: Element | None = font_elem.find(VERT_ALIGN_TAG, None)
    if vert_align_elem is not None:
        if val := vert_align_elem.get("val", None):
            font.vert_align = FontAlignmentType(val)
        else:
            font.vert_align = FontAlignmentType.BASELINE

    color_elem: Element | None = font_elem.find(COLOR_TAG, None)
    if color_elem is not None:
        font.color = parse_color(color_elem)

    return font


def _parse_fill(fill_elem: Element) -> PatternFill | GradientFill:
    pattern_fill_elem: Element | None = fill_elem.find(PATTERN_FILL_TAG, None)
    if pattern_fill_elem is not None:
        fill = PatternFill()

        if pattern_type := pattern_fill_elem.get("patternType", None):
            fill.pattern_type = PatternFillType(pattern_type)

        fg_color_elem: Element | None = pattern_fill_elem.find(FG_COLOR_TAG, None)
        if fg_color_elem is not None:
            fill.fg_color = parse_color(fg_color_elem)

        bg_color_elem: Element | None = pattern_fill_elem.find(BG_COLOR_TAG, None)
        if bg_color_elem is not None:
            fill.bg_color = parse_color(bg_color_elem)

        return fill

    gradient_fill_elem: Element | None = fill_elem.find(GRADIENT_FILL_TAG, None)
    if gradient_fill_elem is not None:
        grad_fill = GradientFill()

        if grad_type := gradient_fill_elem.get("type", None):
            grad_fill.gradient_type = GradientFillType(grad_type)

        if degree := gradient_fill_elem.get("degree", None):
            grad_fill.degree = safe_float(degree)

        for attr in ("left", "right", "top", "bottom"):
            if val := gradient_fill_elem.get(attr, None):
                setattr(grad_fill, attr, safe_float(val))

        stops: list[ColorStop] = []
        for stop_elem in gradient_fill_elem.findall(STOP_TAG, None):
            pos_attr = stop_elem.get("position", None)
            if pos_attr is None:
                continue
            position = safe_float(pos_attr)
            color_elem: Element | None = stop_elem.find(COLOR_TAG, None)
            color_obj = parse_color(color_elem) if color_elem is not None else None
            if color_obj is None:
                continue
            stops.append(ColorStop(position=position, color=color_obj))

        if stops:
            grad_fill.stops = stops

        return grad_fill

    return PatternFill()


def _parse_side(side_elem: Element) -> Side | None:

    style_val = side_elem.get("style", None)
    if not style_val or style_val == "none":
        return None
    style = BorderStyleType(style_val)

    color_elem = side_elem.find(COLOR_TAG, None)
    color = parse_color(color_elem) if color_elem is not None else None

    return Side(style=style, color=color)


def _parse_border(border_elem: Element) -> Border:
    border = Border()

    left_elem = border_elem.find(LEFT_TAG, None)
    if left_elem is not None:
        border.left = _parse_side(left_elem)

    right_elem = border_elem.find(RIGHT_TAG, None)
    if right_elem is not None:
        border.right = _parse_side(right_elem)

    top_elem = border_elem.find(TOP_TAG, None)
    if top_elem is not None:
        border.top = _parse_side(top_elem)

    bottom_elem = border_elem.find(BOTTOM_TAG, None)
    if bottom_elem is not None:
        border.bottom = _parse_side(bottom_elem)

    diagonal_elem = border_elem.find(DIAGONAL_TAG, None)
    if diagonal_elem is not None:
        border.diagonal = _parse_side(diagonal_elem)

    start_elem = border_elem.find(START_TAG, None)
    if start_elem is not None:
        border.start = _parse_side(start_elem)

    end_elem = border_elem.find(END_TAG, None)
    if end_elem is not None:
        border.end = _parse_side(end_elem)

    horizontal_elem = border_elem.find(HORIZONTAL_TAG, None)
    if horizontal_elem is not None:
        border.horizontal = _parse_side(horizontal_elem)

    vertical_elem = border_elem.find(VERTICAL_TAG, None)
    if vertical_elem is not None:
        border.vertical = _parse_side(vertical_elem)

    if outline := border_elem.get("outline", None):
        border.outline = is_true(outline)

    if diagonal_up := border_elem.get("diagonalUp", None):
        border.diagonal_up = is_true(diagonal_up)

    if diagonal_down := border_elem.get("diagonalDown", None):
        border.diagonal_down = is_true(diagonal_down)

    return border


def _parse_alignment(alignment_elem: Element) -> Alignment:
    alignment = Alignment()

    if horizontal := alignment_elem.get("horizontal", None):
        alignment.horizontal = HorizontalAlignmentType(horizontal)

    if vertical := alignment_elem.get("vertical", None):
        alignment.vertical = VerticalAlignmentType(vertical)

    if text_rotation := alignment_elem.get("textRotation", None):
        alignment.text_rotation = safe_int(text_rotation)

    if wrap_text := alignment_elem.get("wrapText", None):
        alignment.wrap_text = is_true(wrap_text)

    if shrink_to_fit := alignment_elem.get("shrinkToFit", None):
        alignment.shrink_to_fit = is_true(shrink_to_fit)

    if justify_last_line := alignment_elem.get("justifyLastLine", None):
        alignment.justify_last_line = is_true(justify_last_line)

    if indent := alignment_elem.get("indent", None):
        alignment.indent = safe_int(indent)

    if relative_indent := alignment_elem.get("relativeIndent", None):
        alignment.relative_indent = safe_int(relative_indent)

    if reading_order := alignment_elem.get("readingOrder", None):
        alignment.reading_order = safe_int(reading_order)

    return alignment


def _parse_protection(protection_elem: Element) -> Protection:
    protection = Protection()

    if locked := protection_elem.get("locked", None):
        protection.locked = is_true(locked)

    if hidden := protection_elem.get("hidden", None):
        protection.hidden = is_true(hidden)

    return protection


def _parse_cell_style_xf(xf_elem: Element) -> _CellStyleXf:
    xf = _CellStyleXf(
        num_fmt_id=safe_int(xf_elem.get("numFmtId", None)),
        font_id=safe_int(xf_elem.get("fontId", None)),
        fill_id=safe_int(xf_elem.get("fillId", None)),
        border_id=safe_int(xf_elem.get("borderId", None)),
    )

    if xf_id := xf_elem.get("xfId", None):
        xf.xf_id = safe_int(xf_id)

    if pivot_button := xf_elem.get("pivotButton", None):
        xf.pivot_button = is_true(pivot_button)

    if quote_prefix := xf_elem.get("quotePrefix", None):
        xf.quote_prefix = is_true(quote_prefix)

    if apply_number_format := xf_elem.get("applyNumberFormat", None):
        xf.apply_number_format = is_true(apply_number_format)

    if apply_font := xf_elem.get("applyFont", None):
        xf.apply_font = is_true(apply_font)

    if apply_fill := xf_elem.get("applyFill", None):
        xf.apply_fill = is_true(apply_fill)

    if apply_border := xf_elem.get("applyBorder", None):
        xf.apply_border = is_true(apply_border)

    if apply_alignment := xf_elem.get("applyAlignment", None):
        xf.apply_alignment = is_true(apply_alignment)

    if apply_protection := xf_elem.get("applyProtection", None):
        xf.apply_protection = is_true(apply_protection)

    alignment_elem = xf_elem.find(ALIGNMENT_TAG, None)
    if alignment_elem is not None:
        xf.alignment = _parse_alignment(alignment_elem)

    protection_elem = xf_elem.find(PROTECTION_TAG, None)
    if protection_elem is not None:
        xf.protection = _parse_protection(protection_elem)

    return xf


def _parse_named_cell_style(cell_style_elem: Element) -> _NamedCellStyle:
    cell_style = _NamedCellStyle(
        name=cell_style_elem.get("name", ""),
        xf_id=safe_int(cell_style_elem.get("xfId", None)),
    )

    if builtin_id := cell_style_elem.get("builtinId", None):
        cell_style.builtin_id = safe_int(builtin_id)

    if i_level := cell_style_elem.get("iLevel", None):
        cell_style.i_level = safe_int(i_level)

    if hidden := cell_style_elem.get("hidden", None):
        cell_style.hidden = is_true(hidden)

    return cell_style


def _merge_xfs_into_cell_style(
    cell_xf: _CellStyleXf,
    cell_style_xf: _CellStyleXf,
    alignments: IndexedList[Alignment],
    protections: IndexedList[Protection],
) -> CellStyle:

    num_fmt_id = None
    if cell_xf.num_fmt_id is not None and (cell_xf.apply_number_format is not False):
        num_fmt_id = cell_xf.num_fmt_id
    elif cell_style_xf.num_fmt_id is not None and (
        cell_style_xf.apply_number_format is not False
    ):
        num_fmt_id = cell_style_xf.num_fmt_id

    font_id = None
    if cell_xf.font_id is not None and (cell_xf.apply_font is not False):
        font_id = cell_xf.font_id
    elif cell_style_xf.font_id is not None and (cell_style_xf.apply_font is not False):
        font_id = cell_style_xf.font_id

    fill_id = None
    if cell_xf.fill_id is not None and (cell_xf.apply_fill is not False):
        fill_id = cell_xf.fill_id
    elif cell_style_xf.fill_id is not None and (cell_style_xf.apply_fill is not False):
        fill_id = cell_style_xf.fill_id

    border_id = None
    if cell_xf.border_id is not None and (cell_xf.apply_border is not False):
        border_id = cell_xf.border_id
    elif cell_style_xf.border_id is not None and (
        cell_style_xf.apply_border is not False
    ):
        border_id = cell_style_xf.border_id

    alignment = None
    if cell_xf.alignment and (cell_xf.apply_alignment is not False):
        alignment = cell_xf.alignment
    elif cell_style_xf.alignment and (cell_style_xf.apply_alignment is not False):
        alignment = cell_style_xf.alignment

    protection = None
    if cell_xf.protection and (cell_xf.apply_protection is not False):
        protection = cell_xf.protection
    elif cell_style_xf.protection and (cell_style_xf.apply_protection is not False):
        protection = cell_style_xf.protection

    pivot_button = cell_xf.pivot_button
    quote_prefix = cell_xf.quote_prefix

    return CellStyle(
        num_fmt_id=num_fmt_id,
        font_id=font_id,
        fill_id=fill_id,
        border_id=border_id,
        quote_prefix=quote_prefix,
        pivot_button=pivot_button,
        alignment_id=alignments.get_index(alignment) if alignment is not None else None,
        protection_id=(
            protections.get_index(protection) if protection is not None else None
        ),
    )


def _add_or_find_empty_xfs_id(cell_style_xfs: list[_CellStyleXf]) -> int:
    for i, cell_style_xf in enumerate(cell_style_xfs):
        if cell_style_xf == _CellStyleXf(
            num_fmt_id=0, font_id=0, fill_id=0, border_id=0
        ):
            return i
    cell_style_xfs.append(_CellStyleXf(num_fmt_id=0, font_id=0, fill_id=0, border_id=0))
    return len(cell_style_xfs) - 1


def _from_xl_arrays(
    cell_style_xfs: list[_CellStyleXf],
    cell_xfs: list[_CellStyleXf],
    names: list[_NamedCellStyle],
) -> tuple[list[CellStyle], list[NamedStyle], list[Alignment], list[Protection]]:
    empty_xfs_id = _add_or_find_empty_xfs_id(cell_style_xfs)

    style_refs: list[CellStyle] = []
    named_styles: list[NamedStyle] = []
    alignments: IndexedList[Alignment] = IndexedList()
    protections: IndexedList[Protection] = IndexedList()

    alignments.get_index(Alignment())
    protections.get_index(Protection())

    for cell_xf in cell_xfs:
        if cell_xf.xf_id is None or cell_xf.xf_id >= len(cell_style_xfs):
            cell_xf.xf_id = empty_xfs_id

        cell_style_xf = cell_style_xfs[cell_xf.xf_id]
        style_refs.append(
            _merge_xfs_into_cell_style(cell_xf, cell_style_xf, alignments, protections)
        )

    for named_style in names:
        cell_style_xf = cell_style_xfs[named_style.xf_id]

        named_styles.append(
            NamedStyle(
                name=named_style.name,
                builtin_id=named_style.builtin_id,
                i_level=named_style.i_level,
                hidden=named_style.hidden,
                custom_builtin=named_style.custom_builtin,
                xf_id=named_style.xf_id,
                num_fmt_id=(
                    None
                    if cell_style_xf.apply_number_format is False
                    else cell_style_xf.num_fmt_id
                ),
                font_id=(
                    None if cell_style_xf.apply_font is False else cell_style_xf.font_id
                ),
                fill_id=(
                    None if cell_style_xf.apply_fill is False else cell_style_xf.fill_id
                ),
                border_id=(
                    None
                    if cell_style_xf.apply_border is False
                    else cell_style_xf.border_id
                ),
                alignment_id=(
                    alignments.get_index(cell_style_xf.alignment)
                    if cell_style_xf.apply_alignment
                    and cell_style_xf.alignment is not None
                    else None
                ),
                protection_id=(
                    protections.get_index(cell_style_xf.protection)
                    if cell_style_xf.apply_protection
                    and cell_style_xf.protection is not None
                    else None
                ),
            )
        )
    named_styles.sort(key=lambda x: x.xf_id)

    return style_refs, named_styles, alignments, protections


def _parse_differential_style(dxf_elem: Element) -> DifferentialStyle:
    style = DifferentialStyle()

    font_elem = dxf_elem.find(FONT_TAG, None)
    if font_elem is not None:
        style.font = _parse_font(font_elem)

    fill_elem = dxf_elem.find(FILL_TAG, None)
    if fill_elem is not None:
        style.fill = _parse_fill(fill_elem)

    border_elem = dxf_elem.find(BORDER_TAG, None)
    if border_elem is not None:
        style.border = _parse_border(border_elem)

    alignment_elem = dxf_elem.find(ALIGNMENT_TAG, None)
    if alignment_elem is not None:
        style.alignment = _parse_alignment(alignment_elem)

    protection_elem = dxf_elem.find(PROTECTION_TAG, None)
    if protection_elem is not None:
        style.protection = _parse_protection(protection_elem)

    num_fmt_elem = dxf_elem.find(NUM_FMT_TAG, None)
    if num_fmt_elem is not None:
        format_code = num_fmt_elem.get("formatCode", "")
        style.number_format = format_code

    return style


def parse_stylesheet(xml_content: bytes) -> Stylesheet:
    root = from_string(xml_content)

    num_fmts_elem: Element | None = root.find(NUM_FMTS_TAG, None)
    num_fmts: dict[int, str] = {}
    if num_fmts_elem is not None:
        num_fmts = dict(
            _parse_num_fmt(num_fmt_elem)
            for num_fmt_elem in num_fmts_elem.iter(NUM_FMT_TAG)
        )

    fonts_elem: Element | None = root.find(FONTS_TAG, None)
    fonts: list[Font] = []
    if fonts_elem is not None:
        fonts = [_parse_font(font_elem) for font_elem in fonts_elem.iter(FONT_TAG)]

    fills_elem: Element | None = root.find(FILLS_TAG, None)
    fills: list[PatternFill | GradientFill] = []
    if fills_elem is not None:
        fills = [_parse_fill(fill_elem) for fill_elem in fills_elem.iter(FILL_TAG)]

    borders_elem: Element | None = root.find(BORDERS_TAG, None)
    borders: list[Border] = []
    if borders_elem is not None:
        borders = [
            _parse_border(border_elem) for border_elem in borders_elem.iter(BORDER_TAG)
        ]

    cell_style_xfs_elem: Element | None = root.find(CELL_STYLE_XFS_TAG, None)
    cell_style_xfs: list[_CellStyleXf] = []
    if cell_style_xfs_elem is not None:
        cell_style_xfs = [
            _parse_cell_style_xf(xf_elem)
            for xf_elem in cell_style_xfs_elem.iter(XF_TAG)
        ]

    cell_xfs_elem: Element | None = root.find(CELL_XFS_TAG, None)
    cell_xfs: list[_CellStyleXf] = []
    if cell_xfs_elem is not None:
        cell_xfs = [
            _parse_cell_style_xf(xf_elem) for xf_elem in cell_xfs_elem.iter(XF_TAG)
        ]

    cell_styles_elem: Element | None = root.find(CELL_STYLES_TAG, None)
    cell_styles: list[_NamedCellStyle] = []
    if cell_styles_elem is not None:
        cell_styles = [
            _parse_named_cell_style(cell_style_elem)
            for cell_style_elem in cell_styles_elem.iter(CELL_STYLE_TAG)
        ]

    dxfs_elem: Element | None = root.find(DXFS_TAG, None)
    dxfs: list[DifferentialStyle] = []
    if dxfs_elem is not None:
        dxfs = [
            _parse_differential_style(dxf_elem) for dxf_elem in dxfs_elem.iter(DXF_TAG)
        ]

    cell_references, named_styles, alignments, protections = _from_xl_arrays(
        cell_style_xfs, cell_xfs, cell_styles
    )

    return Stylesheet(
        num_fmts=num_fmts,
        fonts=fonts,
        fills=fills,
        borders=borders,
        cell_styles=cell_references,
        named_styles=named_styles,
        differential_styles=dxfs,
        alignments=alignments,
        protections=protections,
    )
