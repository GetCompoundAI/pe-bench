from enum import StrEnum

from pydantic import BaseModel

from pe_bench.office.workbook.types_style import Color


class FormatObjectType(StrEnum):
    NUM = "num"
    PERCENT = "percent"
    MAX = "max"
    MIN = "min"
    FORMULA = "formula"
    PERCENTILE = "percentile"


class IconSetType(StrEnum):
    THREE_ARROWS = "3Arrows"
    THREE_ARROWS_GRAY = "3ArrowsGray"
    THREE_FLAGS = "3Flags"
    THREE_TRAFFIC_LIGHTS_1 = "3TrafficLights1"
    THREE_TRAFFIC_LIGHTS_2 = "3TrafficLights2"
    THREE_SIGNS = "3Signs"
    THREE_SYMBOLS = "3Symbols"
    THREE_SYMBOLS_2 = "3Symbols2"
    FOUR_ARROWS = "4Arrows"
    FOUR_ARROWS_GRAY = "4ArrowsGray"
    FOUR_RED_TO_BLACK = "4RedToBlack"
    FOUR_RATING = "4Rating"
    FOUR_TRAFFIC_LIGHTS = "4TrafficLights"
    FIVE_ARROWS = "5Arrows"
    FIVE_ARROWS_GRAY = "5ArrowsGray"
    FIVE_RATING = "5Rating"
    FIVE_QUARTERS = "5Quarters"


class RuleType(StrEnum):
    EXPRESSION = "expression"
    CELL_IS = "cellIs"
    COLOR_SCALE = "colorScale"
    DATA_BAR = "dataBar"
    ICON_SET = "iconSet"
    TOP_10 = "top10"
    UNIQUE_VALUES = "uniqueValues"
    DUPLICATE_VALUES = "duplicateValues"
    CONTAINS_TEXT = "containsText"
    NOT_CONTAINS_TEXT = "notContainsText"
    BEGINS_WITH = "beginsWith"
    ENDS_WITH = "endsWith"
    CONTAINS_BLANKS = "containsBlanks"
    NOT_CONTAINS_BLANKS = "notContainsBlanks"
    CONTAINS_ERRORS = "containsErrors"
    NOT_CONTAINS_ERRORS = "notContainsErrors"
    TIME_PERIOD = "timePeriod"
    ABOVE_AVERAGE = "aboveAverage"


class RuleOperator(StrEnum):
    LESS_THAN = "lessThan"
    LESS_THAN_OR_EQUAL = "lessThanOrEqual"
    EQUAL = "equal"
    NOT_EQUAL = "notEqual"
    GREATER_THAN_OR_EQUAL = "greaterThanOrEqual"
    GREATER_THAN = "greaterThan"
    BETWEEN = "between"
    NOT_BETWEEN = "notBetween"
    CONTAINS_TEXT = "containsText"
    NOT_CONTAINS = "notContains"
    BEGINS_WITH = "beginsWith"
    ENDS_WITH = "endsWith"


class TimePeriodType(StrEnum):
    TODAY = "today"
    YESTERDAY = "yesterday"
    TOMORROW = "tomorrow"
    LAST_7_DAYS = "last7Days"
    THIS_MONTH = "thisMonth"
    LAST_MONTH = "lastMonth"
    NEXT_MONTH = "nextMonth"
    THIS_WEEK = "thisWeek"
    LAST_WEEK = "lastWeek"
    NEXT_WEEK = "nextWeek"


class FormatObject(BaseModel):
    type: FormatObjectType
    val: str | float | None = None
    gte: bool | None = None


class ColorScale(BaseModel):
    cfvo: list[FormatObject] | None = None
    color: list[Color] | None = None


class DataBar(BaseModel):
    min_length: int | None = None
    max_length: int | None = None
    show_value: bool | None = None
    cfvo: list[FormatObject] | None = None
    color: Color | None = None


class IconSet(BaseModel):
    icon_set: IconSetType | None = None
    show_value: bool | None = None
    percent: bool | None = None
    reverse: bool | None = None
    cfvo: list[FormatObject] | None = None


class Rule(BaseModel):
    type: RuleType
    dxf_id: int | None = None
    priority: int
    stop_if_true: bool | None = None
    above_average: bool | None = None
    percent: bool | None = None
    bottom: bool | None = None
    operator: RuleOperator | None = None
    text: str | None = None
    time_period: TimePeriodType | None = None
    rank: int | None = None
    std_dev: int | None = None
    equal_average: bool | None = None
    formula: list[str] | None = None
    color_scale: ColorScale | None = None
    data_bar: DataBar | None = None
    icon_set: IconSet | None = None


class ConditionalFormatting(BaseModel):
    cell_ranges: list[str]
    cf_rule: list[Rule]
    pivot: bool | None = None
