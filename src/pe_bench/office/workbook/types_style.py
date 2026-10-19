import re
from enum import IntEnum, StrEnum

from pydantic import BaseModel, Field, field_serializer, field_validator

aRGB_REGEX = re.compile("^([A-Fa-f0-9]{8}|[A-Fa-f0-9]{6})$")


class ColorType(StrEnum):
    RGB = "rgb"
    INDEXED = "indexed"
    AUTO = "auto"
    THEME = "theme"


class ThemeColorType(IntEnum):
    DARK1 = 0
    LIGHT1 = 1
    DARK2 = 2
    LIGHT2 = 3
    ACCENT1 = 4
    ACCENT2 = 5
    ACCENT3 = 6
    ACCENT4 = 7
    ACCENT5 = 8
    ACCENT6 = 9
    HYPERLINK = 10
    FOLLOWED_HYPERLINK = 11


class Color(BaseModel):
    type: ColorType
    rgb: str | None = None
    indexed: int | None = None
    auto: bool | None = None
    theme: ThemeColorType | None = None
    tint: float | None = Field(default=None, ge=-1.0, le=1.0)

    @field_validator("rgb", mode="before")
    @classmethod
    def validate_rgb(cls, v: str | None) -> str | None:
        if v is None:
            return v
        normalized = v.strip().lstrip("#")
        if aRGB_REGEX.match(normalized) is None:
            return None
        if len(normalized) == 6:
            normalized = "00" + normalized
        return normalized


class PatternFillType(StrEnum):
    NONE = "none"
    SOLID = "solid"
    DARK_DOWN = "darkDown"
    DARK_GRAY = "darkGray"
    DARK_GRID = "darkGrid"
    DARK_HORIZONTAL = "darkHorizontal"
    DARK_TRELLIS = "darkTrellis"
    DARK_UP = "darkUp"
    DARK_VERTICAL = "darkVertical"
    GRAY_0625 = "gray0625"
    GRAY_125 = "gray125"
    LIGHT_DOWN = "lightDown"
    LIGHT_GRAY = "lightGray"
    LIGHT_GRID = "lightGrid"
    LIGHT_HORIZONTAL = "lightHorizontal"
    LIGHT_TRELLIS = "lightTrellis"
    LIGHT_UP = "lightUp"
    LIGHT_VERTICAL = "lightVertical"
    MEDIUM_GRAY = "mediumGray"


class PatternFill(BaseModel):
    pattern_type: PatternFillType | None = None
    fg_color: Color | None = None
    bg_color: Color | None = None


class GradientFillType(StrEnum):
    LINEAR = "linear"
    PATH = "path"


class ColorStop(BaseModel):
    position: float = Field(ge=0.0, le=1.0)
    color: Color


class GradientFill(BaseModel):
    gradient_type: GradientFillType | None = None
    degree: float | None = Field(default=None, ge=0.0, le=360.0)

    left: float | None = Field(default=None, ge=0.0, le=1.0)
    right: float | None = Field(default=None, ge=0.0, le=1.0)
    top: float | None = Field(default=None, ge=0.0, le=1.0)
    bottom: float | None = Field(default=None, ge=0.0, le=1.0)

    stops: list[ColorStop] | None = None


class FontUnderlineType(StrEnum):
    NONE = "none"
    SINGLE = "single"
    DOUBLE = "double"
    SINGLE_ACCOUNTING = "singleAccounting"
    DOUBLE_ACCOUNTING = "doubleAccounting"


class FontAlignmentType(StrEnum):
    SUPERSCRIPT = "superscript"
    SUBSCRIPT = "subscript"
    BASELINE = "baseline"


class Font(BaseModel):
    name: str | None = None
    charset: int | None = None
    family: int | None = Field(default=None, ge=0, le=14)
    size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    strike: bool | None = None
    outline: bool | None = None
    shadow: bool | None = None
    condense: bool | None = None
    extend: bool | None = None
    underline: FontUnderlineType | None = None
    vert_align: FontAlignmentType | None = None
    color: Color | None = None
    scheme: str | None = None


class HorizontalAlignmentType(StrEnum):
    GENERAL = "general"
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    FILL = "fill"
    JUSTIFY = "justify"
    CENTER_CONTINUOUS = "centerContinuous"
    DISTRIBUTED = "distributed"


class VerticalAlignmentType(StrEnum):
    TOP = "top"
    CENTER = "center"
    BOTTOM = "bottom"
    JUSTIFY = "justify"
    DISTRIBUTED = "distributed"


class Alignment(BaseModel):
    horizontal: HorizontalAlignmentType | None = None
    vertical: VerticalAlignmentType | None = None
    text_rotation: int | None = Field(default=None, ge=0, le=255)
    wrap_text: bool | None = None
    shrink_to_fit: bool | None = None
    indent: int | None = Field(default=None, ge=0, le=255)
    relative_indent: int | None = Field(default=None, ge=-255, le=255)
    justify_last_line: bool | None = None
    reading_order: int | None = Field(default=None, ge=0)

    def __hash__(self) -> int:
        return hash(
            (
                self.horizontal,
                self.vertical,
                self.text_rotation,
                self.wrap_text,
                self.shrink_to_fit,
                self.indent,
                self.relative_indent,
                self.justify_last_line,
                self.reading_order,
            )
        )


class BorderStyleType(StrEnum):
    DASH_DOT = "dashDot"
    DASH_DOT_DOT = "dashDotDot"
    DASHED = "dashed"
    DOTTED = "dotted"
    DOUBLE = "double"
    HAIR = "hair"
    MEDIUM = "medium"
    MEDIUM_DASH_DOT = "mediumDashDot"
    MEDIUM_DASH_DOT_DOT = "mediumDashDotDot"
    MEDIUM_DASHED = "mediumDashed"
    SLANT_DASH_DOT = "slantDashDot"
    THICK = "thick"
    THIN = "thin"


class Side(BaseModel):
    style: BorderStyleType
    color: Color | None = None


class Border(BaseModel):
    start: Side | None = None
    end: Side | None = None
    left: Side | None = None
    right: Side | None = None
    top: Side | None = None
    bottom: Side | None = None
    diagonal: Side | None = None
    vertical: Side | None = None
    horizontal: Side | None = None

    outline: bool | None = None
    diagonal_up: bool | None = None
    diagonal_down: bool | None = None


class Protection(BaseModel):
    locked: bool | None = None
    hidden: bool | None = None

    def __hash__(self) -> int:
        return hash((self.locked, self.hidden))


class CellStyle(BaseModel):
    num_fmt_id: int | None = None
    font_id: int | None = None
    fill_id: int | None = None
    border_id: int | None = None
    alignment_id: int | None = None
    protection_id: int | None = None

    pivot_button: bool | None = None
    quote_prefix: bool | None = None


class NamedStyle(BaseModel):
    name: str
    xf_id: int
    builtin_id: int | None = None
    i_level: int | None = None
    hidden: bool | None = None
    custom_builtin: bool | None = None

    num_fmt_id: int | None = None
    font_id: int | None = None
    fill_id: int | None = None
    border_id: int | None = None
    alignment_id: int | None = None
    protection_id: int | None = None


class DifferentialStyle(BaseModel):
    font: Font | None = None
    fill: PatternFill | GradientFill | None = None
    border: Border | None = None
    alignment: Alignment | None = None
    number_format: str | None = None
    protection: Protection | None = None


class Stylesheet(BaseModel):
    num_fmts: dict[int, str] | None = None
    fonts: list[Font] | None = None
    fills: list[PatternFill | GradientFill] | None = None
    borders: list[Border] | None = None
    alignments: list[Alignment] | None = None
    protections: list[Protection] | None = None
    cell_styles: list[CellStyle] | None = None
    named_styles: list[NamedStyle] | None = None
    differential_styles: list[DifferentialStyle] | None = None

    @field_serializer("num_fmts")
    def serialize_num_fmts(
        self, num_fmts: dict[int, str] | None
    ) -> dict[str, str] | None:
        if num_fmts is None:
            return None
        return {str(k): v for k, v in num_fmts.items()}

    @field_validator("num_fmts", mode="before")
    @classmethod
    def validate_num_fmts(
        cls, v: dict[str, str] | dict[int, str] | list[str] | None
    ) -> dict[int, str] | None:
        if v is None:
            return None
        if isinstance(v, dict):
            return {int(k) if isinstance(k, str) else k: v for k, v in v.items()}
        if isinstance(v, list):
            return {i + 164: v for i, v in enumerate(v)}
        return v
