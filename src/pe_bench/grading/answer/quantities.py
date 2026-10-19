import re

_SPELLED = {
    "zero": 0.0,
    "one": 1.0,
    "two": 2.0,
    "three": 3.0,
    "four": 4.0,
    "five": 5.0,
    "six": 6.0,
    "seven": 7.0,
    "eight": 8.0,
    "nine": 9.0,
    "ten": 10.0,
    "eleven": 11.0,
    "twelve": 12.0,
    "thirteen": 13.0,
    "fourteen": 14.0,
    "fifteen": 15.0,
    "sixteen": 16.0,
    "seventeen": 17.0,
    "eighteen": 18.0,
    "nineteen": 19.0,
    "twenty": 20.0,
}

_NUMBER = r"\d[\d,]*(?:\.\d+)?|\b(?:{})\b".format(
    "|".join(sorted(_SPELLED, key=len, reverse=True))
)

_START = r"(?<![\w.,/])"

_RANGE = re.compile(
    rf"{_START}(?:{_NUMBER})[^\S\n]*(?:[-–—]|\bto\b|\bor\b)[^\S\n]*(?:{_NUMBER})"
)

_QUANTITY = rf"{_START}-?(?:{_NUMBER})"

_GAP = r"[^\w\n%/]"

_BETWEEN = rf"{_GAP}{{0,4}}(?:[^\W\d_]+{_GAP}{{0,4}}){{0,2}}"


_PART_OF = (
    rf"({_QUANTITY}){_GAP}{{0,4}}(?:out{_GAP}{{1,4}})?of{_GAP}{{1,4}}"
    rf"(?:(?:the|these|those|all){_GAP}{{1,4}}){{0,2}}{_QUANTITY}"
)
_FRONT = (_PART_OF, f"({_QUANTITY})")


def counts_of(text: str, *units: str) -> list[float]:
    lowered = _RANGE.sub("\n", text.lower().replace("*", ""))
    found = {
        one.start(): one.group(1)
        for unit in units
        for front in _FRONT
        for one in re.finditer(
            rf"{front}{_BETWEEN}{re.escape(unit.lower())}\w*", lowered
        )
    }
    return [
        _SPELLED[word] if word in _SPELLED else float(word.replace(",", ""))
        for _, word in sorted(found.items())
    ]
