import re
from typing import Final, Pattern

_ABSTIME: Final[Pattern[str]] = re.compile("\\[[HhMmSs\u0e0a\u0e19\u0e17]+\\]")


def _is_general(s: str, i: int = 0) -> bool:
    return s[i : i + 7].lower() == "general"


def fmt_is_date(fmt: str) -> bool:
    fmtt = fmt.title()
    if fmtt in ("Date", "Short Date", "Long Date", "Time"):
        return True
    if fmtt in (
        "General",
        "Number",
        "Currency",
        "Accounting",
        "Percentage",
        "Fraction",
        "Scientific",
        "Text",
    ):
        return False
    i = 0
    while i < len(fmt):
        c = fmt[i]
        if c == "G":
            if _is_general(fmt, i):
                i += 7
                continue
        elif c == '"':
            j = fmt.find('"', i + 1)
            i = j + 1 if j > i else i + 1
            continue
        elif c == "\\":
            i += 2
            continue
        elif c == "_":
            i += 2
            continue
        elif c == "@":
            i += 1
            continue
        elif c in ("B", "b"):
            if fmt[i + 1 : i + 2] in ("1", "2"):
                return True
        if c in (
            "B",
            "b",
            "M",
            "D",
            "Y",
            "H",
            "S",
            "E",
            "m",
            "d",
            "y",
            "h",
            "s",
            "e",
            "g",
        ):
            return True
        if c in ("A", "a", "上"):
            if fmt[i : i + 3].upper() == "A/P":
                return True
            if fmt[i : i + 5].upper() in ("AM/PM", "上午/下午"):
                return True
            i += 1
        elif c == "[":
            j = fmt.find("]", i + 1)
            if j < 0:
                return False
            if _ABSTIME.match(fmt[i : j + 1]):
                return True
            i = j + 1
        elif c in (".", "0", "#"):
            while i < len(fmt):
                i += 1
                c = fmt[i : i + 1]
                if "0#?.,E+-%".find(c) > -1 or (
                    c == "\\"
                    and fmt[i + 1 : i + 2] == "-"
                    and "0#".find(fmt[i + 2 : i + 3]) > -1
                ):
                    continue
        elif c == "?":
            i += 1
            while fmt[i : i + 1] == c:
                i += 1
        elif c == "*":
            i += 1
            if fmt[i : i + 1] in (" ", "*"):
                i += 1
        elif c in ("(", ")"):
            i += 1
        elif c in ("1", "2", "3", "4", "5", "6", "7", "8", "9"):
            i += 1
            while fmt[i : i + 1].isdigit():
                i += 1
        elif c == " ":
            i += 1
        else:
            i += 1
    return False
