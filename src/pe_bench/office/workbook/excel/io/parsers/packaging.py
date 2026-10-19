from pe_bench.office.workbook.excel.io.constants import Namespaces
from pe_bench.office.workbook.excel.io.xml_functions import (
    Element,
    from_string,
    safe_datetime,
    safe_int,
)
from pe_bench.office.workbook.types_document import (
    DocumentCoreProperties,
    ExtendedProperties,
)

COREPROPS_NS = Namespaces.PKG_COREPROPS
DCORE_NS = Namespaces.DCORE
DCTERMS_NS = Namespaces.DCTERMS
XPROPS_NS = Namespaces.DOC_XPROPS

CATEGORY_TAG = f"{{{COREPROPS_NS}}}category"
CONTENT_STATUS_TAG = f"{{{COREPROPS_NS}}}contentStatus"
KEYWORDS_TAG = f"{{{COREPROPS_NS}}}keywords"
LAST_MODIFIED_BY_TAG = f"{{{COREPROPS_NS}}}lastModifiedBy"
LAST_PRINTED_TAG = f"{{{COREPROPS_NS}}}lastPrinted"
REVISION_TAG = f"{{{COREPROPS_NS}}}revision"
VERSION_TAG = f"{{{COREPROPS_NS}}}version"

CREATOR_TAG = f"{{{DCORE_NS}}}creator"
TITLE_TAG = f"{{{DCORE_NS}}}title"
SUBJECT_TAG = f"{{{DCORE_NS}}}subject"
DESCRIPTION_TAG = f"{{{DCORE_NS}}}description"
IDENTIFIER_TAG = f"{{{DCORE_NS}}}identifier"
LANGUAGE_TAG = f"{{{DCORE_NS}}}language"

CREATED_TAG = f"{{{DCTERMS_NS}}}created"
MODIFIED_TAG = f"{{{DCTERMS_NS}}}modified"

TEMPLATE_TAG = f"{{{XPROPS_NS}}}Template"
MANAGER_TAG = f"{{{XPROPS_NS}}}Manager"
COMPANY_TAG = f"{{{XPROPS_NS}}}Company"
APPLICATION_TAG = f"{{{XPROPS_NS}}}Application"
APP_VERSION_TAG = f"{{{XPROPS_NS}}}AppVersion"
DOC_SECURITY_TAG = f"{{{XPROPS_NS}}}DocSecurity"


def parse_core_properties(xml_content: bytes) -> DocumentCoreProperties:
    root = from_string(xml_content)

    props = DocumentCoreProperties()

    category_elem: Element | None = root.find(CATEGORY_TAG)  # type: ignore
    if category_elem is not None and category_elem.text:
        props.category = category_elem.text

    content_status_elem: Element | None = root.find(CONTENT_STATUS_TAG)  # type: ignore
    if content_status_elem is not None and content_status_elem.text:
        props.content_status = content_status_elem.text

    keywords_elem: Element | None = root.find(KEYWORDS_TAG)  # type: ignore
    if keywords_elem is not None and keywords_elem.text:
        props.keywords = keywords_elem.text

    last_modified_by_elem: Element | None = root.find(LAST_MODIFIED_BY_TAG)  # type: ignore
    if last_modified_by_elem is not None and last_modified_by_elem.text:
        props.last_modified_by = last_modified_by_elem.text

    last_printed_elem: Element | None = root.find(LAST_PRINTED_TAG)  # type: ignore
    if last_printed_elem is not None and last_printed_elem.text:
        props.last_printed = safe_datetime(last_printed_elem.text)

    revision_elem: Element | None = root.find(REVISION_TAG)  # type: ignore
    if revision_elem is not None and revision_elem.text:
        props.revision = revision_elem.text

    version_elem: Element | None = root.find(VERSION_TAG)  # type: ignore
    if version_elem is not None and version_elem.text:
        props.version = version_elem.text

    creator_elem: Element | None = root.find(CREATOR_TAG)  # type: ignore
    if creator_elem is not None and creator_elem.text:
        props.creator = creator_elem.text

    title_elem: Element | None = root.find(TITLE_TAG)  # type: ignore
    if title_elem is not None and title_elem.text:
        props.title = title_elem.text

    subject_elem: Element | None = root.find(SUBJECT_TAG)  # type: ignore
    if subject_elem is not None and subject_elem.text:
        props.subject = subject_elem.text

    description_elem: Element | None = root.find(DESCRIPTION_TAG)  # type: ignore
    if description_elem is not None and description_elem.text:
        props.description = description_elem.text

    identifier_elem: Element | None = root.find(IDENTIFIER_TAG)  # type: ignore
    if identifier_elem is not None and identifier_elem.text:
        props.identifier = identifier_elem.text

    language_elem: Element | None = root.find(LANGUAGE_TAG)  # type: ignore
    if language_elem is not None and language_elem.text:
        props.language = language_elem.text

    created_elem: Element | None = root.find(CREATED_TAG)  # type: ignore
    if created_elem is not None and created_elem.text:
        props.created = safe_datetime(created_elem.text)

    modified_elem: Element | None = root.find(MODIFIED_TAG)  # type: ignore
    if modified_elem is not None and modified_elem.text:
        props.modified = safe_datetime(modified_elem.text)

    return props


def parse_extended_properties(xml_content: bytes) -> ExtendedProperties:
    root = from_string(xml_content)

    props = ExtendedProperties()

    template_elem: Element | None = root.find(TEMPLATE_TAG)  # type: ignore
    if template_elem is not None and template_elem.text:
        props.template = template_elem.text

    manager_elem: Element | None = root.find(MANAGER_TAG)  # type: ignore
    if manager_elem is not None and manager_elem.text:
        props.manager = manager_elem.text

    company_elem: Element | None = root.find(COMPANY_TAG)  # type: ignore
    if company_elem is not None and company_elem.text:
        props.company = company_elem.text

    application_elem: Element | None = root.find(APPLICATION_TAG)  # type: ignore
    if application_elem is not None and application_elem.text:
        props.application = application_elem.text

    app_version_elem: Element | None = root.find(APP_VERSION_TAG)  # type: ignore
    if app_version_elem is not None and app_version_elem.text:
        props.app_version = app_version_elem.text

    doc_security_elem: Element | None = root.find(DOC_SECURITY_TAG)  # type: ignore
    if doc_security_elem is not None and doc_security_elem.text:
        props.doc_security = safe_int(doc_security_elem.text)

    return props
