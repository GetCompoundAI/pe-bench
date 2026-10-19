class _Constants:
    def __setattr__(self, key: str, value: object) -> None:
        raise AttributeError("Cannot modify constant values.")

    def __init__(self):
        raise NotImplementedError("This class should not be instantiated.")


class ArchiveFilePaths(_Constants):
    CONTENT_TYPES: str = "[Content_Types].xml"
    ROOT_RELS: str = "_rels/.rels"
    WORKBOOK_RELS: str = "xl/_rels/workbook.xml.rels"
    CORE: str = "docProps/core.xml"
    APP: str = "docProps/app.xml"
    CUSTOM: str = "docProps/custom.xml"
    WORKBOOK: str = "xl/workbook.xml"
    STYLES: str = "xl/styles.xml"
    THEME: str = "xl/theme/theme1.xml"
    SHARED_STRINGS: str = "xl/sharedStrings.xml"
    CUSTOM_UI: str = "customUI/customUI.xml"
    WORKSHEETS: str = "xl/worksheets/"
    WORKSHEET_TEMPLATE: str = "xl/worksheets/sheet{}.xml"
    WORKSHEET_RELS_TEMPLATE: str = "xl/worksheets/_rels/sheet{}.xml.rels"
    WORKSHEET_RELS: str = "xl/worksheets/_rels"


_DOC_NS = "http://schemas.openxmlformats.org/officeDocument/2006/"
_PKG_NS = "http://schemas.openxmlformats.org/package/2006/"


class Namespaces(_Constants):
    XSI: str = "http://www.w3.org/2001/XMLSchema-instance"

    XML: str = "http://www.w3.org/XML/1998/namespace"

    DCORE: str = "http://purl.org/dc/elements/1.1/"
    DCTERMS: str = "http://purl.org/dc/terms/"
    DCMITYPE: str = "http://purl.org/dc/dcmitype/"

    DOC_REL: str = f"{_DOC_NS}relationships"
    DOC_VTYPES: str = f"{_DOC_NS}docPropsVTypes"
    DOC_XPROPS: str = f"{_DOC_NS}extended-properties"
    DOC_CUSTPROPS: str = f"{_DOC_NS}custom-properties"

    PKG_REL: str = f"{_PKG_NS}relationships"
    PKG_COREPROPS: str = f"{_PKG_NS}metadata/core-properties"
    PKG_CONTYPES: str = f"{_PKG_NS}content-types"

    SHEET_MAIN: str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    MARKING_COMPATIBILITY: str = (
        "http://schemas.openxmlformats.org/markup-compatibility/2006"
    )

    CHART: str = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    DRAWING: str = "http://schemas.openxmlformats.org/drawingml/2006/main"
