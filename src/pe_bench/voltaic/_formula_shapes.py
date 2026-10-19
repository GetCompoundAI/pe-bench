import re

from pe_bench.grading.workbook.formula_shape import FORMULA_TOKEN, formula_body

_STRING = re.compile(r'"[^"]*"')
_SHEET_PREFIX = re.compile(r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!")
_FUNCTION = re.compile(r"([A-Z][A-Z0-9.]*)\(", re.IGNORECASE)
_BRANCHES = frozenset({"IF", "IFS", "IFERROR", "IFNA", "CHOOSE", "SWITCH"})
_OPERATORS = frozenset("+-*/^&")
_COMPARISON = re.compile(r"[<>]|<>|(?<![<>!])=(?!=)")
_REFERENCE_SHEET = re.compile(r"^(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!")
_ROW_NUMBER = re.compile(r"\d+")
LONG_FORMULA = 100


def _after_call(body: str, start: int) -> str:
    depth = 1
    index = start
    while index < len(body) and depth:
        depth += {"(": 1, ")": -1}.get(body[index], 0)
        index += 1
    return body[index:].lstrip()[:1]


def is_complex(formula: str) -> bool:
    body = _STRING.sub('""', formula)
    if len(_SHEET_PREFIX.sub("", body).replace("$", "")) > LONG_FORMULA:
        return True
    stack: list[str] = []
    index = 0
    while index < len(body):
        found = _FUNCTION.match(body, index)
        if found and (index == 0 or not body[index - 1].isalnum()):
            name = found.group(1).upper()
            before = body[:index].rstrip()[-1:]
            enclosing = [one for one in stack if one]
            if name in _BRANCHES and (
                before in _OPERATORS
                or (enclosing and enclosing[-1] not in _BRANCHES)
                or _after_call(body, found.end()) in _OPERATORS
            ):
                return True
            calculations = [one for one in enclosing if one not in _BRANCHES]
            if name not in _BRANCHES and calculations:
                return True
            stack.append(name)
            index = found.end()
            continue
        if body[index] == "(":
            stack.append("")
        elif body[index] == ")" and stack:
            stack.pop()
        index += 1
    return False


def is_a_check(value: object, formula: str) -> bool:
    if isinstance(value, str):
        return True
    return (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and value in (0, 1)
        and bool(_COMPARISON.search(_STRING.sub('""', formula)[1:]))
    )


def moved_form(formula: str) -> str:
    parts: list[str] = []
    for found in FORMULA_TOKEN.finditer(formula_body(formula)):
        if found.lastgroup == "space":
            continue
        text = found.group().upper()
        if found.lastgroup == "ref":
            text = _ROW_NUMBER.sub("#", _REFERENCE_SHEET.sub("", text))
        parts.append(text)
    return "".join(parts)
