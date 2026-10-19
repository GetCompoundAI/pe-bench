import json
import re
from collections.abc import Mapping

from defusedxml.ElementTree import ParseError, fromstring
from openpyxl.styles.colors import COLOR_INDEX

from pe_bench.office.deliverable_checks.cells import style_of as style_of
from pe_bench.office.workbook.excel.style_editor import ExcelCellStyle, StylesheetEditor
from pe_bench.office.workbook.types import Workbook
from pe_bench.office.workbook.types_style import Color, ColorType, PatternFillType

_THEME_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

_THEME_ELEMENTS = {
    0: "lt1",
    1: "dk1",
    2: "lt2",
    3: "dk2",
    4: "accent1",
    5: "accent2",
    6: "accent3",
    7: "accent4",
    8: "accent5",
    9: "accent6",
    10: "hlink",
    11: "folHlink",
}

_INDEXED_PALETTE: dict[int, str] = {
    index: f"FF{argb[-6:].upper()}" for index, argb in enumerate(COLOR_INDEX)
} | {64: "FF000000", 65: "FFFFFFFF"}


def theme_rgbs(book: Workbook) -> dict[int, str]:
    if not book.theme:
        return {}
    try:
        root = fromstring(book.theme)
    except ParseError:
        return {}
    found: dict[int, str] = {}
    for index, element_name in _THEME_ELEMENTS.items():
        element = root.find(f".//{{{_THEME_NS}}}{element_name}")
        if element is None:
            continue
        srgb = element.find(f"{{{_THEME_NS}}}srgbClr")
        if srgb is not None and srgb.get("val"):
            found[index] = f"FF{srgb.get('val', '').upper()}"
            continue
        system = element.find(f"{{{_THEME_NS}}}sysClr")
        if system is not None:
            value = system.get("lastClr") or system.get("val") or ""
            if value:
                found[index] = f"FF{value.upper()}"
    return found


def apply_tint(rgb: str, tint: float) -> str:
    channels = (int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16))
    tinted = [
        max(
            0,
            min(
                255,
                round(
                    channel * (1 + tint)
                    if tint < 0
                    else channel + (255 - channel) * tint
                ),
            ),
        )
        for channel in channels
    ]
    return rgb[:2] + "".join(f"{channel:02X}" for channel in tinted)


def resolved_rgb(color: Color | None, palette: Mapping[int, str]) -> str | None:
    if color is None:
        return None
    rgb: str | None = None
    if color.rgb is not None:
        rgb = color.rgb
    elif color.theme is not None:
        rgb = palette.get(int(color.theme))
    elif color.indexed is not None:
        rgb = _INDEXED_PALETTE.get(color.indexed)
    elif color.auto or color.type == ColorType.AUTO:
        rgb = "FF000000"
    if rgb is None:
        return None
    rgb = f"FF{rgb[-6:].upper()}"
    return apply_tint(rgb, color.tint) if color.tint else rgb


_DEFAULT_SPELLINGS: tuple[object, ...] = (
    None,
    "",
    {},
    [],
    "none",
    "General",
    "general",
    "bottom",
    False,
)
_FONT_METADATA = ("family", "scheme")
_PROTECTION_DEFAULTS = {"locked": True, "hidden": False}
_COLOR_TYPES = tuple(ColorType)
_AUTOMATIC_TEXT_RGB = "FF000000"
_LITERAL_IN_FORMAT = re.compile(r"\\([$+\-/():!^&'~{}<>= ])")


def visible_style(
    style: ExcelCellStyle | None,
    *,
    palette: Mapping[int, str],
    normal: ExcelCellStyle | None = None,
) -> dict[str, object]:

    def shown(value: object) -> object:
        if isinstance(value, Mapping):
            if value.get("type") in _COLOR_TYPES:
                rgb = resolved_rgb(Color.model_validate(dict(value)), palette)
                return rgb if rgb is not None else json.dumps(value, sort_keys=True)
            kept = {key: shown(item) for key, item in value.items()}
            return {
                key: item
                for key, item in kept.items()
                if item not in _DEFAULT_SPELLINGS
            }
        if isinstance(value, list):
            return [shown(item) for item in value]
        if isinstance(value, str):
            return value.strip()
        return value

    normal_color = None if normal is None or normal.font is None else normal.font.color
    normal_font = None if normal is None else normal.font
    normal_name = None if normal_font is None else normal_font.name
    normal_size = None if normal_font is None else normal_font.size
    automatic = (
        _AUTOMATIC_TEXT_RGB
        if normal_color is None
        else shown(normal_color.model_dump(exclude_none=True))
    )

    def attributes(one: ExcelCellStyle | None) -> dict[str, object]:
        found: dict[str, object] = {}
        for name, value in (one.model_dump(exclude_none=True) if one else {}).items():
            if name == "font":
                value = {
                    "name": normal_name,
                    "size": normal_size,
                    **{k: v for k, v in value.items() if k not in _FONT_METADATA},
                }
            elif name == "number_format" and isinstance(value, str):
                value = _LITERAL_IN_FORMAT.sub(r"\1", value)
            elif name == "fill" and value.get("pattern_type") == PatternFillType.SOLID:
                value = {k: v for k, v in value.items() if k != "bg_color"}
            seen = (
                {**_PROTECTION_DEFAULTS, **value}
                if name == "protection" and isinstance(value, Mapping)
                else shown(value)
            )
            if (
                name == "font"
                and isinstance(seen, dict)
                and seen.get("color") == automatic
            ):
                seen = {k: v for k, v in seen.items() if k != "color"}
            if seen not in _DEFAULT_SPELLINGS:
                found[name] = seen
        return found

    plain = attributes(normal)
    plain.setdefault("protection", _PROTECTION_DEFAULTS)
    return {
        name: value
        for name, value in attributes(style).items()
        if value != plain.get(name)
    }


def sheet_style_signatures(
    book: Workbook, sheet_name: str, *, normal: ExcelCellStyle | None
) -> dict[str, dict[str, object]]:
    styles = StylesheetEditor(book)
    palette = theme_rgbs(book)
    own = styles.get_cell_style(0)
    found: dict[str, dict[str, object]] = {}
    for coord, cell in book.sheets[sheet_name].cells.items():
        style = own if cell.s is None else style_of(styles, cell)
        visible = visible_style(style, palette=palette, normal=normal)
        if visible:
            found[coord] = visible
    return found


def rgb_channels(rgb: str) -> tuple[int, int, int]:
    raw = rgb[2:] if len(rgb) == 8 else rgb
    return int(raw[0:2], 16), int(raw[2:4], 16), int(raw[4:6], 16)
