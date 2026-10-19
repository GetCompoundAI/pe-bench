import calendar
import re
from dataclasses import dataclass
from datetime import date, datetime
from typing import Literal

from pe_bench.grading.workbook.xlsx import column_number, parse_coord, same_text
from pe_bench.office.workbook.excel.type_coerce import serial_number_to_datetime
from pe_bench.office.workbook.types import Sheet

_YEAR_TEXT = re.compile(r"^(?:fy|cy)?((?:19|20)\d{2})([ae])?$")
_SHORT_YEAR_TEXT = re.compile(r"^(?:fy|cy)(\d{2})([ae])?$")
_MONTH_YEAR_TEXT = re.compile(r"^([a-z]{3,9})((?:19|20)\d{2}|\d{2})$")
_US_DATE_TEXT = re.compile(r"^\d{1,2}/\d{1,2}/((?:19|20)\d{2}|\d{2})$")
_EARLIEST_YEAR = 1900
_LATEST_YEAR = 2199
_DATE_SERIALS = (20000, 80000)

YearHeader = tuple[int, str, int, str | None]

_MONTH_NUMBER = {
    name: number
    for number, name in enumerate(
        (
            "january",
            "february",
            "march",
            "april",
            "may",
            "june",
            "july",
            "august",
            "september",
            "october",
            "november",
            "december",
        ),
        start=1,
    )
}
_ISO_DATE_TEXT = re.compile(r"^((?:19|20)\d{2})[-/](\d{1,2})[-/](\d{1,2})$")
_SLASH_DATE_TEXT = re.compile(r"^(\d{1,2})[/-](\d{1,2})[/-]((?:19|20)\d{2}|\d{2})$")
_MONTH_FIRST_TEXT = re.compile(
    r"^([a-z]{3,9})[\s\-.,']*(?:(\d{1,2})[\s\-.,']+)?((?:19|20)\d{2}|\d{2})$"
)
_DAY_FIRST_TEXT = re.compile(
    r"^(\d{1,2})[\s\-.]([a-z]{3,9})[\s\-.,']+((?:19|20)\d{2}|\d{2})$"
)


def _month_number(name: str) -> int | None:
    if name == "sept":
        return 9
    matched = {
        number for month, number in _MONTH_NUMBER.items() if month.startswith(name)
    }
    return matched.pop() if len(matched) == 1 else None


def _full_year(digits: str) -> int:
    year = int(digits)
    return year if year >= 100 else 2000 + year


def _named_month_date(text: str) -> tuple[int, int, int] | None:
    named = _MONTH_FIRST_TEXT.match(text)
    if named is not None:
        name, day_digits, year_digits = named.groups()
    else:
        day_first = _DAY_FIRST_TEXT.match(text)
        if day_first is None:
            return None
        day_digits, name, year_digits = day_first.groups()
    month = _month_number(name)
    if month is None:
        return None
    return (
        _full_year(year_digits),
        month,
        1 if day_digits is None else int(day_digits),
    )


def text_date(value: str) -> date | None:
    text = re.sub(r"\s+", " ", same_text(value)).strip()
    iso = _ISO_DATE_TEXT.match(text)
    slashed = _SLASH_DATE_TEXT.match(text)
    if iso is not None:
        parts = (int(iso.group(1)), int(iso.group(2)), int(iso.group(3)))
    elif slashed is not None:
        parts = (
            _full_year(slashed.group(3)),
            int(slashed.group(1)),
            int(slashed.group(2)),
        )
    else:
        found = _named_month_date(text)
        if found is None:
            return None
        parts = found
    year, month, day = parts
    if not (_EARLIEST_YEAR <= year <= _LATEST_YEAR and 1 <= month <= 12):
        return None
    try:
        return date(year, month, day)
    except ValueError:
        return None


def year_header(value: object | None) -> tuple[int, str | None] | None:
    if isinstance(value, (date, datetime)):
        return (value.year, None)
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        if number.is_integer() and _EARLIEST_YEAR <= number <= _LATEST_YEAR:
            return (int(number), None)
        if _DATE_SERIALS[0] <= number <= _DATE_SERIALS[1]:
            converted = serial_number_to_datetime(number)
            return (
                None if not isinstance(converted, datetime) else (converted.year, None)
            )
        return None
    raw = same_text(str(value))
    text = re.sub(r"[^a-z0-9]+", "", raw)
    found = _YEAR_TEXT.match(text)
    if found is not None:
        return (int(found.group(1)), found.group(2))
    short = _SHORT_YEAR_TEXT.match(text)
    if short is not None:
        return (2000 + int(short.group(1)), short.group(2))
    dated = _MONTH_YEAR_TEXT.match(text)
    if dated is not None and _month_number(dated.group(1)) is not None:
        return (_full_year(dated.group(2)), None)
    us_dated = _US_DATE_TEXT.match(raw.replace(" ", ""))
    if us_dated is not None:
        return (_full_year(us_dated.group(1)), None)
    spelled = text_date(raw)
    return None if spelled is None else (spelled.year, None)


def year_header_row(
    sheet: Sheet, *, least: int = 4, years: range = range(1990, 2101)
) -> tuple[int, list[YearHeader]] | None:
    by_row: dict[int, list[YearHeader]] = {}
    for coord, cell in sheet.cells.items():
        parsed = year_header(cell.v)
        if parsed is None or parsed[0] not in years:
            continue
        column, row = parse_coord(coord)
        by_row.setdefault(row, []).append((column_number(column), column, *parsed))
    best: tuple[int, list[YearHeader]] | None = None
    for row, headers in sorted(by_row.items()):
        distinct = len({year for _number, _letters, year, _marker in headers})
        if distinct >= least and (
            best is None or distinct > len({year for _n, _l, year, _m in best[1]})
        ):
            best = (row, sorted(headers))
    return best


def header_day(value: object | None) -> int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, str):
        spelled = text_date(value)
        return None if spelled is None else spelled.toordinal()
    if isinstance(value, datetime):
        return value.date().toordinal()
    if isinstance(value, date):
        return value.toordinal()
    if (
        isinstance(value, (int, float))
        and _DATE_SERIALS[0] <= value <= _DATE_SERIALS[1]
    ):
        converted = serial_number_to_datetime(float(value))
        return converted.date().toordinal() if isinstance(converted, datetime) else None
    return None


_QUARTER_NAMED = re.compile(
    r"(?<![a-z0-9])(?:q([1-4])|([1-4])q)[\s'\u2019/-]*(?:fy)?[\s'\u2019/-]*((?:20)?\d{2})(?!\d)",
    re.IGNORECASE,
)


def quarter_named(value: object | None) -> tuple[int, int] | None:
    if not isinstance(value, str):
        return None
    named = _QUARTER_NAMED.search(value)
    if named is None:
        return None
    return (_full_year(named.group(3)), int(named.group(1) or named.group(2)))


@dataclass(frozen=True)
class Period:
    kind: Literal["year", "quarter", "half", "ltm", "date"]
    year: int | None
    part: int | None = None
    day: date | None = None


_QUARTER_ALONE = re.compile(
    r"(?:q([1-4])|([1-4])q)[\s'\u2019/-]*(?:fy|cy)?[\s'\u2019/-]*((?:20)?\d{2})[ae]?"
)
_YEAR_FIRST_QUARTER = re.compile(
    r"(?:(?:fy|cy)\s*'?(\d{2})|(?:fy|cy)?\s*((?:19|20)\d{2}))\s*[-/]?\s*q([1-4])[ae]?"
)
_YEAR_FIRST_QUARTER_IN = re.compile(r"(?<![0-9])((?:19|20)\d{2}) ?-?q([1-4])(?![0-9])")
_DATED_SERIAL_YEARS = range(2000, 2100)
_LTM = re.compile(
    r"(?<![a-z0-9_])(?:ltm|ttm)(?![a-z_])|\b(?:last|trailing) twelve months\b"
)
_HALF = re.compile(
    r"(?<![a-z0-9])(?:h([12])|([12])h|(6m)|(first half)|(second half))(?![a-z0-9])"
)
_HALF_FIRST = re.compile(r"([12])h'?((?:19|20)\d{2}|\d{2})[ae]?")
_SIX_MONTHS_OF_A_YEAR = re.compile(r"\(\s*6m\s*\)")
_LTM_ATTACHED_YEAR = re.compile(r"(?<![a-z0-9_])(?:ltm|ttm) ?(\d{4})(?![0-9])")
_LAYERED = re.compile(r"[\n()]")
_COMPARED = re.compile(r"(?<![a-z])(?:vs\.?|versus)(?![a-z])", re.IGNORECASE)
_FISCAL = re.compile(r"(?<![a-z])fy")
_ENDED = re.compile(
    r"^(?:(three|six|twelve|3|6|12) ?(?:months?|m|mo)(?: ended| ending)?|"
    r"(?:quarter|qtr|qe|q/e|period|year|fiscal year)(?: ended| ending)?)[\s:,-]*(.+)$"
)
_QUALIFIERS = re.compile(
    r"\b(?:actuals?|act|audited|unaudited|reported|restated|estimated?|est|"
    r"forecast|budget|plan|projected|projection|pro ?forma|pf)\b|[()]"
)
_YEAR_SUFFIX = re.compile(r"(\d{2})[aefbp]$")
_YEAR_ANYWHERE = re.compile(r"(?<![0-9])((?:19|20)\d{2})(?![0-9])")
_SHORT_FY = re.compile(r"(?<![a-z0-9])fy ?'?(\d{2})(?![0-9])")


_NAMED_PERIOD = re.compile(
    r"^(?:FY(?P<year>\d{4})|(?P<quarter_year>\d{4})Q(?P<quarter>[1-4])"
    r"|(?P<half_year>\d{4})H(?P<half>[12])"
    r"|LTM(?:(?P<ltm_year>\d{4})(?:Q(?P<ltm_quarter>[1-4]))?)?"
    r"|(?P<day>\d{4}-\d{2}-\d{2})|(?P<month_year>\d{4})-(?P<month>\d{2}))$"
)
NAMED_PERIOD_NOTATION = (
    "FY2024 for a fiscal year, 2023Q1 for a quarter, 2024H1 for a half year, "
    "LTM or LTM2024 for the twelve months to a period end, 2024-12-31 for a "
    "period-end date"
)


def period_named(text: str) -> Period | None:
    matched = _NAMED_PERIOD.fullmatch(text.strip())
    if matched is None:
        return None
    named_year = next(
        (
            int(digits[:4])
            for digits in (
                matched["year"],
                matched["quarter_year"],
                matched["half_year"],
                matched["day"],
                matched["month_year"],
                matched["ltm_year"],
            )
            if digits
        ),
        None,
    )
    if named_year is not None and not _EARLIEST_YEAR <= named_year <= _LATEST_YEAR:
        return None
    if matched["year"]:
        return Period("year", int(matched["year"]))
    if matched["quarter"]:
        return Period("quarter", int(matched["quarter_year"]), int(matched["quarter"]))
    if matched["half"]:
        return Period("half", int(matched["half_year"]), int(matched["half"]))
    if matched["day"]:
        try:
            day = date.fromisoformat(matched["day"])
        except ValueError:
            return None
        return Period("date", day.year, day=day)
    if matched["month"]:
        year, month = int(matched["month_year"]), int(matched["month"])
        if not 1 <= month <= 12:
            return None
        return Period(
            "date", year, day=date(year, month, calendar.monthrange(year, month)[1])
        )
    if matched["ltm_quarter"]:
        year, quarter = int(matched["ltm_year"]), int(matched["ltm_quarter"])
        return Period("ltm", year, day=_quarter_end(year, quarter))
    return Period("ltm", int(matched["ltm_year"]) if matched["ltm_year"] else None)


def period_notation(period: Period) -> str:
    if period.kind == "date" and period.day is not None:
        return period.day.isoformat()
    if period.kind == "ltm":
        return f"LTM{period.year or ''}"
    if period.kind == "quarter":
        return f"{period.year}Q{period.part}"
    if period.kind == "half":
        return f"{period.year}H{period.part}"
    return f"FY{period.year}"


def _quarter_end(year: int, quarter: int) -> date:
    month = 3 * quarter
    return date(year, month, calendar.monthrange(year, month)[1])


def _year_first_quarter(text: str) -> tuple[int, int] | None:
    named = _YEAR_FIRST_QUARTER.fullmatch(text)
    if named is None:
        return None
    return (_full_year(named.group(1) or named.group(2)), int(named.group(3)))


def _named_parts(raw: str, separator: re.Pattern[str]) -> list[Period] | None:
    parts = [
        part
        for part in separator.split(raw)
        if re.search(r"[a-z0-9]", part, re.IGNORECASE)
    ]
    if len(parts) < 2:
        return None
    named: list[Period] = []
    for part in parts:
        period = parse_period(part)
        if period is None or period.year is None:
            return None
        named.append(period)
    return named


def _first_of_layered(raw: str) -> Period | None:
    named = _named_parts(raw, _LAYERED)
    if named is None:
        return None
    first = named[0]
    if any(
        one.kind != first.kind or (one.year or 0) > (first.year or 0) for one in named
    ):
        return None
    return first


def _ltm_end(text: str) -> date | None:
    year_first = _YEAR_FIRST_QUARTER_IN.search(text)
    quarter = quarter_named(text) or (
        None
        if year_first is None
        else (int(year_first.group(1)), int(year_first.group(2)))
    )
    if quarter is not None:
        return None if _FISCAL.search(text) else _quarter_end(*quarter)
    rest = re.sub(r"\s+", " ", _LTM.sub(" ", _YEAR_SUFFIX.sub(r"\1", text))).strip()
    day = _dated(rest)
    month_only = _MONTH_FIRST_TEXT.match(rest)
    if day is None or month_only is None or month_only.group(2) is not None:
        return day
    return date(day.year, day.month, calendar.monthrange(day.year, day.month)[1])


def _dated(text: str) -> date | None:
    spelled = text_date(text)
    if spelled is not None:
        return spelled
    cleaned = re.sub(r"(\d)(?:st|nd|rd|th)\b", r"\1", text.replace(",", " "))
    return text_date(re.sub(r"\s+", " ", cleaned).strip())


def parse_period(value: object | None) -> Period | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, datetime):
        return Period("date", value.year, day=value.date())
    if isinstance(value, date):
        return Period("date", value.year, day=value)
    if isinstance(value, (int, float)):
        number = float(value)
        if number.is_integer() and _EARLIEST_YEAR <= number <= _LATEST_YEAR:
            return Period("year", int(number))
        day = header_day(value)
        if day is None or date.fromordinal(day).year not in _DATED_SERIAL_YEARS:
            return None
        return Period("date", date.fromordinal(day).year, day=date.fromordinal(day))
    text = same_text(str(value))
    if not text:
        return None
    if _named_parts(str(value), _COMPARED) is not None:
        return None
    layered = _first_of_layered(str(value))
    if layered is not None:
        return layered
    if _LTM.search(text):
        attached = _LTM_ATTACHED_YEAR.search(text)
        if attached is not None and not (
            _EARLIEST_YEAR <= int(attached.group(1)) <= _LATEST_YEAR
        ):
            return None
        end = _ltm_end(text)
        quarter = quarter_named(text)
        years = _YEAR_ANYWHERE.findall(text)
        if end is not None:
            return Period("ltm", end.year, day=end)
        if quarter is not None:
            return Period("ltm", quarter[0])
        return Period("ltm", int(years[-1]) if years else None)
    ended = _ENDED.match(text)
    if ended is not None:
        day = _dated(ended.group(2))
        span = ended.group(1)
        if day is not None:
            if (
                span in ("three", "3")
                or span is None
                and text.startswith(("quarter", "qtr", "qe", "q/e"))
            ):
                return Period("quarter", day.year, (day.month - 1) // 3 + 1, day)
            if span in ("six", "6"):
                return Period("half", day.year, 1 if day.month <= 6 else 2, day)
            return Period("year", day.year, day=day)
    bare = re.sub(r"\s+", " ", _QUALIFIERS.sub(" ", text)).strip()
    quarter = _QUARTER_ALONE.fullmatch(bare)
    if quarter is not None:
        return Period(
            "quarter",
            _full_year(quarter.group(3)),
            int(quarter.group(1) or quarter.group(2)),
        )
    year_first = _year_first_quarter(bare)
    if year_first is not None:
        return Period("quarter", *year_first)
    half_first = _HALF_FIRST.fullmatch(bare)
    if half_first is not None:
        return Period("half", _full_year(half_first.group(2)), int(half_first.group(1)))
    half = _HALF.search(text)
    if half is not None:
        if half.group(3) and _SIX_MONTHS_OF_A_YEAR.search(text):
            ended = parse_period(_SIX_MONTHS_OF_A_YEAR.sub(" ", text))
            if ended is not None and ended.kind == "date" and ended.day is not None:
                return Period("half", ended.day.year, 1 if ended.day.month <= 6 else 2)
            if re.search(r"(?<![a-z])fy", text):
                return None
        years = _YEAR_ANYWHERE.findall(text)
        short = re.search(r"'(\d{2})(?![0-9])|(?<![0-9])(\d{2})$", text)
        year = (
            int(years[-1])
            if years
            else (_full_year(short.group(1) or short.group(2)) if short else None)
        )
        part = 2 if half.group(5) else int(half.group(1) or half.group(2) or 1)
        return Period("half", year, part)
    for candidate in (text, bare, _YEAR_SUFFIX.sub(r"\1", bare)):
        if not candidate:
            continue
        day = _dated(candidate)
        if day is not None:
            return Period("date", day.year, day=day)
        read = year_header(candidate)
        if read is not None:
            return Period("year", read[0])
        short = _SHORT_FY.fullmatch(candidate)
        if short is not None:
            return Period("year", _full_year(short.group(1)))
        if candidate != text:
            years = _YEAR_ANYWHERE.findall(candidate)
            rest = _YEAR_ANYWHERE.sub("", candidate)
            if len(years) == 1 and re.fullmatch(r"[\s\-/.fy']*", rest):
                return Period("year", int(years[0]))
    return None
