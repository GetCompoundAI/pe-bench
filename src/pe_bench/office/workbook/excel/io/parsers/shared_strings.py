from pe_bench.office.workbook.excel.io.constants import Namespaces
from pe_bench.office.workbook.excel.io.xml_functions import Element, from_string

SHEET_MAIN_NS = Namespaces.SHEET_MAIN
STRING_TAG = "{%s}si" % SHEET_MAIN_NS
TEXT_TAG = "{%s}t" % SHEET_MAIN_NS
RICH_TEXT_TAG = "{%s}r" % SHEET_MAIN_NS


def parse_shared_strings(xml_content: bytes) -> list[str]:
    strings: list[str] = []

    root = from_string(xml_content)

    si_elems: list[Element] = root.findall(STRING_TAG, None)
    for si_elem in si_elems:
        t_elem: Element | None = si_elem.find(TEXT_TAG, None)
        if t_elem is not None:
            text = t_elem.text or ""
            strings.append(text)
        else:
            text_parts: list[str] = []
            r_elems: list[Element] = si_elem.findall(RICH_TEXT_TAG, None)
            for r_elem in r_elems:
                t_elem: Element | None = r_elem.find(TEXT_TAG, None)
                if t_elem is not None:
                    text_parts.append(t_elem.text or "")

            if text_parts:
                strings.append("".join(text_parts))
            else:
                strings.append("")

    return strings
