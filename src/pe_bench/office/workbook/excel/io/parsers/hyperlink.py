from pe_bench.office.workbook.excel.io.constants import Namespaces
from pe_bench.office.workbook.excel.io.parsers.relationships import Relationship
from pe_bench.office.workbook.excel.io.xml_functions import Element
from pe_bench.office.workbook.types_sheet import Hyperlink

SHEET_MAIN_NS = Namespaces.SHEET_MAIN
HYPERLINK_TAG = "{%s}hyperlinks" % SHEET_MAIN_NS
HYPERLINK_CHILD_TAG = "{%s}hyperlink" % SHEET_MAIN_NS


def parse_hyperlinks(
    sheet_elem: Element, relationships: list[Relationship]
) -> dict[str, Hyperlink] | None:

    links_parent: Element | None = sheet_elem.find(HYPERLINK_TAG, None)
    if links_parent is None:
        return None

    relationships_by_id: dict[str, Relationship] = {r.id: r for r in relationships}
    hyperlinks: dict[str, Hyperlink] = {}
    link_elems: list[Element] = links_parent.findall(HYPERLINK_CHILD_TAG, None)

    for link_elem in link_elems:
        ref = link_elem.get("ref", None)
        if not ref:
            continue

        rid_attr_name = f"{{{Namespaces.DOC_REL}}}id"
        target_rid = link_elem.get(rid_attr_name, None)
        target_url = None

        if target_rid and target_rid in relationships_by_id:
            target_url = relationships_by_id[target_rid].target

        hyperlink = Hyperlink(
            url=target_url,
            cell_coord=link_elem.get("location", None),
            display=link_elem.get("display", None),
            tooltip=link_elem.get("tooltip", None),
        )

        hyperlinks[ref] = hyperlink

    return hyperlinks if hyperlinks else None
