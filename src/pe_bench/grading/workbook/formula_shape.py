import re
from dataclasses import dataclass

from pe_bench.grading.workbook.xlsx import column_number, parse_coord

_ENDPOINT = r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}"
FORMULA_TOKEN = re.compile(
    r'(?P<string>"(?:[^"]|"")*")'
    r"|(?P<func>[A-Za-z_][A-Za-z0-9_.]*\s*\()"
    r"|(?P<ref>(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    rf"(?:{_ENDPOINT}(?::{_ENDPOINT})?"
    r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
    r"|\$?[0-9]{1,7}:\$?[0-9]{1,7})"
    r"(?![A-Za-z0-9_$]))"
    r"|(?P<name>[A-Za-z_][A-Za-z0-9_.]*)"
    r"|(?P<number>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)"
    r"|(?P<space>\s+)"
    r"|(?P<other><=|>=|<>|.)"
)
_SHEET_PREFIX = re.compile(
    r"^(?:'(?P<quoted>(?:[^']|'')+)'|(?P<bare>[A-Za-z_][A-Za-z0-9_.]*))!"
)


@dataclass(frozen=True)
class Reference:
    sheet: str | None
    first_column: str | None
    first_row: int | None
    last_column: str | None
    last_row: int | None
    first_column_anchored: bool = False
    first_row_anchored: bool = False
    last_column_anchored: bool = False
    last_row_anchored: bool = False
    counter: bool = False

    @property
    def is_range(self) -> bool:
        return self.last_column is not None or self.last_row is not None

    @property
    def on_own_sheet(self) -> bool:
        return self.sheet is None

    @property
    def is_whole_column(self) -> bool:
        return self.is_range and self.first_row is None and self.last_row is None

    @property
    def is_whole_row(self) -> bool:
        return self.is_range and self.first_column is None and self.last_column is None


def formula_body(formula: str) -> str:
    return formula.strip().lstrip("=+@").strip()


def _endpoint(text: str) -> tuple[str | None, bool, int | None, bool]:
    leading = text.startswith("$")
    rest = text[1:] if leading else text
    letters = ""
    while rest and rest[0].isalpha():
        letters, rest = letters + rest[0], rest[1:]
    inner = rest.startswith("$")
    if inner:
        rest = rest[1:]
    row = int(rest) if rest else None
    if letters:
        return letters.upper(), leading, row, inner
    return None, False, row, leading


def parse_reference(text: str, own_sheet: str, *, counter: bool = False) -> Reference:
    prefix = _SHEET_PREFIX.match(text)
    sheet: str | None = None
    if prefix is not None:
        quoted = prefix.group("quoted")
        named = quoted.replace("''", "'") if quoted else prefix.group("bare")
        text = text[prefix.end() :]
        if named.strip().lower() != own_sheet.strip().lower():
            sheet = named
    start, _, end = text.partition(":")
    first_column, first_column_anchored, first_row, first_row_anchored = _endpoint(
        start
    )
    last_column, last_column_anchored, last_row, last_row_anchored = (
        _endpoint(end) if end else (None, False, None, False)
    )
    return Reference(
        sheet=sheet,
        first_column=first_column,
        first_row=first_row,
        last_column=last_column,
        last_row=last_row,
        first_column_anchored=first_column_anchored,
        first_row_anchored=first_row_anchored,
        last_column_anchored=last_column_anchored,
        last_row_anchored=last_row_anchored,
        counter=counter,
    )


_COUNTER_FUNCTIONS = frozenset({"ROWS", "COLUMNS"})


def formula_references(formula: str, own_sheet: str) -> list[Reference]:
    found: list[Reference] = []
    calls: list[str] = []
    for token in FORMULA_TOKEN.finditer(formula_body(formula)):
        if token.lastgroup == "func":
            calls.append(token.group().rstrip("(").strip().upper())
        elif token.group() == "(":
            calls.append("")
        elif token.group() == ")":
            if calls:
                calls.pop()
        elif token.lastgroup == "ref":
            found.append(
                parse_reference(
                    token.group(),
                    own_sheet,
                    counter=bool(calls) and calls[-1] in _COUNTER_FUNCTIONS,
                )
            )
    return found


def _offset(distance: int) -> str:
    return "0" if distance == 0 else f"{distance:+d}"


def _endpoint_shape(
    column: str | None,
    column_anchored: bool,
    row: int | None,
    row_anchored: bool,
    *,
    moving: bool,
    host_column: int,
    host_row: int,
) -> str:
    parts = []
    if column is not None:
        if column_anchored:
            parts.append(f"${column}")
        else:
            parts.append(
                "C?" if moving else f"C[{_offset(column_number(column) - host_column)}]"
            )
    if row is not None:
        if row_anchored:
            parts.append(f"${row}")
        else:
            parts.append("R?" if moving else f"R[{_offset(row - host_row)}]")
    return " ".join(parts)


def _reference_shape(one: Reference, host_column: int, host_row: int) -> str:
    moving = not one.on_own_sheet
    start = _endpoint_shape(
        one.first_column,
        one.first_column_anchored,
        one.first_row,
        one.first_row_anchored,
        moving=moving,
        host_column=host_column,
        host_row=host_row,
    )
    end = (
        _endpoint_shape(
            one.last_column,
            one.last_column_anchored,
            one.last_row,
            one.last_row_anchored,
            moving=moving,
            host_column=host_column,
            host_row=host_row,
        )
        if one.is_range
        else ""
    )
    prefix = "" if one.sheet is None else f"{one.sheet.strip().lower()}!"
    return f"{prefix}{start}:{end}" if one.is_range else f"{prefix}{start}"


def formula_shape(formula: str, host: str, own_sheet: str) -> str:
    host_letters, host_row = parse_coord(host)
    host_column = column_number(host_letters)
    out = []
    for found in FORMULA_TOKEN.finditer(formula_body(formula)):
        kind, text = found.lastgroup, found.group()
        if kind == "space":
            continue
        if kind == "ref":
            out.append(
                _reference_shape(
                    parse_reference(text, own_sheet), host_column, host_row
                )
            )
        elif kind == "func":
            out.append(text.rstrip("(").strip().upper() + "(")
        elif kind == "name":
            out.append(text.upper())
        else:
            out.append(text)
    return "".join(out)
