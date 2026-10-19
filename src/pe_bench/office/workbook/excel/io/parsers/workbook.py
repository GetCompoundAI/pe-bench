import re

from pydantic import BaseModel

from pe_bench.office.workbook.excel.io.constants import Namespaces
from pe_bench.office.workbook.excel.io.xml_functions import (
    Element,
    decode_base64_binary,
    decode_hex_binary,
    from_string,
    is_true,
    safe_float,
    safe_int,
)
from pe_bench.office.workbook.types import Workbook
from pe_bench.office.workbook.types_sheet import SheetState
from pe_bench.office.workbook.types_workbook import (
    CalcModeType,
    CalcProperties,
    DefinedName,
    FileSharing,
    FileVersion,
    RefModeType,
    ShowObjectsType,
    UpdateLinksType,
    VisibilityType,
    WorkbookProperties,
    WorkbookProtection,
    WorkbookView,
)

EXTERNAL_WORKBOOK_PATTERN = re.compile(r"\[(\d+)\]")

SHEET_MAIN_NS = Namespaces.SHEET_MAIN

SHEETS_TAG = "{%s}sheets" % SHEET_MAIN_NS
SHEET_TAG = "{%s}sheet" % SHEET_MAIN_NS
BOOK_VIEWS_TAG = "{%s}bookViews" % SHEET_MAIN_NS
WORKBOOK_VIEW_TAG = "{%s}workbookView" % SHEET_MAIN_NS
DEFINED_NAMES_TAG = "{%s}definedNames" % SHEET_MAIN_NS
DEFINED_NAME_TAG = "{%s}definedName" % SHEET_MAIN_NS
CALC_PR_TAG = "{%s}calcPr" % SHEET_MAIN_NS
WORKBOOK_PR_TAG = "{%s}workbookPr" % SHEET_MAIN_NS
WORKBOOK_PROTECTION_TAG = "{%s}workbookProtection" % SHEET_MAIN_NS
FILE_VERSION_TAG = "{%s}fileVersion" % SHEET_MAIN_NS
FILE_SHARING_TAG = "{%s}fileSharing" % SHEET_MAIN_NS

REL_NS = Namespaces.DOC_REL
REL_ID_ATTR = "{%s}id" % REL_NS


def _parse_workbook_properties(workbook_elem: Element) -> WorkbookProperties | None:
    workbook_pr_elem: Element | None = workbook_elem.find(WORKBOOK_PR_TAG, None)
    if workbook_pr_elem is None:
        return None

    properties = WorkbookProperties()

    if date_1904 := workbook_pr_elem.get("date1904", None):
        properties.date_1904 = is_true(date_1904)

    if date_compatibility := workbook_pr_elem.get("dateCompatibility", None):
        properties.date_compatibility = is_true(date_compatibility)

    if show_border_unselected_tables := workbook_pr_elem.get(
        "showBorderUnselectedTables", None
    ):
        properties.show_border_unselected_tables = is_true(
            show_border_unselected_tables
        )

    if filter_privacy := workbook_pr_elem.get("filterPrivacy", None):
        properties.filter_privacy = is_true(filter_privacy)

    if prompted_solutions := workbook_pr_elem.get("promptedSolutions", None):
        properties.prompted_solutions = is_true(prompted_solutions)

    if show_ink_annotation := workbook_pr_elem.get("showInkAnnotation", None):
        properties.show_ink_annotation = is_true(show_ink_annotation)

    if backup_file := workbook_pr_elem.get("backupFile", None):
        properties.backup_file = is_true(backup_file)

    if save_external_link_values := workbook_pr_elem.get(
        "saveExternalLinkValues", None
    ):
        properties.save_external_link_values = is_true(save_external_link_values)

    if hide_pivot_field_list := workbook_pr_elem.get("hidePivotFieldList", None):
        properties.hide_pivot_field_list = is_true(hide_pivot_field_list)

    if show_pivot_chart_filter := workbook_pr_elem.get("showPivotChartFilter", None):
        properties.show_pivot_chart_filter = is_true(show_pivot_chart_filter)

    if allow_refresh_query := workbook_pr_elem.get("allowRefreshQuery", None):
        properties.allow_refresh_query = is_true(allow_refresh_query)

    if publish_items := workbook_pr_elem.get("publishItems", None):
        properties.publish_items = is_true(publish_items)

    if check_compatibility := workbook_pr_elem.get("checkCompatibility", None):
        properties.check_compatibility = is_true(check_compatibility)

    if auto_compress_pictures := workbook_pr_elem.get("autoCompressPictures", None):
        properties.auto_compress_pictures = is_true(auto_compress_pictures)

    if refresh_all_connections := workbook_pr_elem.get("refreshAllConnections", None):
        properties.refresh_all_connections = is_true(refresh_all_connections)

    if show_objects := workbook_pr_elem.get("showObjects", None):
        properties.show_objects = ShowObjectsType(show_objects)

    if update_links := workbook_pr_elem.get("updateLinks", None):
        properties.update_links = UpdateLinksType(update_links)

    if code_name := workbook_pr_elem.get("codeName", None):
        properties.code_name = code_name

    if default_theme_version := workbook_pr_elem.get("defaultThemeVersion", None):
        properties.default_theme_version = safe_int(default_theme_version, 0)

    return properties


def _parse_calc_properties(workbook_elem: Element) -> CalcProperties | None:
    calc_pr_elem: Element | None = workbook_elem.find(CALC_PR_TAG, None)
    if calc_pr_elem is None:
        return None

    calc_id = safe_int(calc_pr_elem.get("calcId", "0"))
    calc_mode = CalcModeType(calc_pr_elem.get("calcMode", "auto"))

    properties = CalcProperties(calc_id=calc_id, calc_mode=calc_mode)

    if full_calc_on_load := calc_pr_elem.get("fullCalcOnLoad", None):
        properties.full_calc_on_load = is_true(full_calc_on_load)

    if iterate := calc_pr_elem.get("iterate", None):
        properties.iterate = is_true(iterate)

    if full_precision := calc_pr_elem.get("fullPrecision", None):
        properties.full_precision = is_true(full_precision)

    if calc_completed := calc_pr_elem.get("calcCompleted", None):
        properties.calc_completed = is_true(calc_completed)

    if calc_on_save := calc_pr_elem.get("calcOnSave", None):
        properties.calc_on_save = is_true(calc_on_save)

    if concurrent_calc := calc_pr_elem.get("concurrentCalc", None):
        properties.concurrent_calc = is_true(concurrent_calc)

    if force_full_calc := calc_pr_elem.get("forceFullCalc", None):
        properties.force_full_calc = is_true(force_full_calc)

    if ref_mode := calc_pr_elem.get("refMode", None):
        properties.ref_mode = RefModeType(ref_mode)

    if iterate_count := calc_pr_elem.get("iterateCount", None):
        properties.iterate_count = safe_int(iterate_count, 0)

    if concurrent_manual_count := calc_pr_elem.get("concurrentManualCount", None):
        properties.concurrent_manual_count = safe_int(concurrent_manual_count, 0)

    if iterate_delta := calc_pr_elem.get("iterateDelta", None):
        properties.iterate_delta = safe_float(iterate_delta)

    return properties


def _parse_workbook_view(view_elem: Element) -> WorkbookView:
    view_data = WorkbookView()

    if minimized := view_elem.get("minimized", None):
        view_data.minimized = is_true(minimized)

    if show_horizontal_scroll := view_elem.get("showHorizontalScroll", None):
        view_data.show_horizontal_scroll = is_true(show_horizontal_scroll)

    if show_vertical_scroll := view_elem.get("showVerticalScroll", None):
        view_data.show_vertical_scroll = is_true(show_vertical_scroll)

    if show_sheet_tabs := view_elem.get("showSheetTabs", None):
        view_data.show_sheet_tabs = is_true(show_sheet_tabs)

    if auto_filter_date_grouping := view_elem.get("autoFilterDateGrouping", None):
        view_data.auto_filter_date_grouping = is_true(auto_filter_date_grouping)

    if visibility := view_elem.get("visibility", None):
        view_data.visibility = VisibilityType(visibility)

    if x_window := view_elem.get("xWindow", None):
        view_data.x_window = safe_int(x_window, 0)

    if y_window := view_elem.get("yWindow", None):
        view_data.y_window = safe_int(y_window, 0)

    if window_width := view_elem.get("windowWidth", None):
        view_data.window_width = safe_int(window_width, 0)

    if window_height := view_elem.get("windowHeight", None):
        view_data.window_height = safe_int(window_height, 0)

    if tab_ratio := view_elem.get("tabRatio", None):
        view_data.tab_ratio = safe_int(tab_ratio, 0)

    if first_sheet := view_elem.get("firstSheet", None):
        view_data.first_sheet = safe_int(first_sheet, 0)

    if active_tab := view_elem.get("activeTab", None):
        view_data.active_tab = safe_int(active_tab, 0)

    return view_data


def _parse_workbook_views(workbook_elem: Element) -> list[WorkbookView] | None:
    book_views_elem: Element | None = workbook_elem.find(BOOK_VIEWS_TAG, None)
    if book_views_elem is None:
        return None

    views: list[WorkbookView] = [
        _parse_workbook_view(view_elem)
        for view_elem in book_views_elem.findall(WORKBOOK_VIEW_TAG, None)
    ]
    return views if views else None


def _parse_defined_name(name_elem: Element) -> DefinedName:
    name = name_elem.get("name", "")
    value = name_elem.text or ""

    defined_name = DefinedName(name=name, value=value)

    if comment := name_elem.get("comment", None):
        defined_name.comment = comment

    if custom_menu := name_elem.get("customMenu", None):
        defined_name.custom_menu = custom_menu

    if description := name_elem.get("description", None):
        defined_name.description = description

    if help_text := name_elem.get("help", None):
        defined_name.help = help_text

    if status_bar := name_elem.get("statusBar", None):
        defined_name.status_bar = status_bar

    if shortcut_key := name_elem.get("shortcutKey", None):
        defined_name.shortcut_key = shortcut_key

    if local_sheet_id := name_elem.get("localSheetId", None):
        defined_name.local_sheet_id = safe_int(local_sheet_id, 0)

    if function_group_id := name_elem.get("functionGroupId", None):
        defined_name.function_group_id = safe_int(function_group_id, 0)

    if hidden := name_elem.get("hidden", None):
        defined_name.hidden = is_true(hidden)

    if function := name_elem.get("function", None):
        defined_name.function = is_true(function)

    if vb_procedure := name_elem.get("vbProcedure", None):
        defined_name.vb_procedure = is_true(vb_procedure)

    if xlm := name_elem.get("xlm", None):
        defined_name.xlm = is_true(xlm)

    if publish_to_server := name_elem.get("publishToServer", None):
        defined_name.publish_to_server = is_true(publish_to_server)

    if workbook_parameter := name_elem.get("workbookParameter", None):
        defined_name.workbook_parameter = is_true(workbook_parameter)

    return defined_name


def _skip_defined_name(defined_name: DefinedName) -> bool:
    return defined_name.hidden or bool(
        EXTERNAL_WORKBOOK_PATTERN.search(defined_name.value)
    )


def _parse_defined_names(workbook_elem: Element) -> list[DefinedName] | None:
    defined_names_elem: Element | None = workbook_elem.find(DEFINED_NAMES_TAG, None)
    if defined_names_elem is None:
        return None

    defined_names = [
        _parse_defined_name(name_elem)
        for name_elem in defined_names_elem.findall(DEFINED_NAME_TAG, None)
    ]
    defined_names = [name for name in defined_names if not _skip_defined_name(name)]
    return defined_names if defined_names else None


def _parse_file_version(workbook_elem: Element) -> FileVersion | None:
    file_version_elem: Element | None = workbook_elem.find(FILE_VERSION_TAG, None)
    if file_version_elem is None:
        return None

    file_version = FileVersion()

    if app_name := file_version_elem.get("appName", None):
        file_version.app_name = app_name

    if last_edited := file_version_elem.get("lastEdited", None):
        file_version.last_edited = last_edited

    if lowest_edited := file_version_elem.get("lowestEdited", None):
        file_version.lowest_edited = lowest_edited

    if rup_build := file_version_elem.get("rupBuild", None):
        file_version.rup_build = rup_build

    if code_name := file_version_elem.get("codeName", None):
        file_version.code_name = code_name

    return file_version


def _parse_file_sharing(workbook_elem: Element) -> FileSharing | None:
    file_sharing_elem: Element | None = workbook_elem.find(FILE_SHARING_TAG, None)
    if file_sharing_elem is None:
        return None

    file_sharing = FileSharing()

    if read_only_recommended := file_sharing_elem.get("readOnlyRecommended", None):
        file_sharing.read_only_recommended = is_true(read_only_recommended)

    if user_name := file_sharing_elem.get("userName", None):
        file_sharing.user_name = user_name

    if algorithm_name := file_sharing_elem.get("algorithmName", None):
        file_sharing.algorithm_name = algorithm_name

    if spin_count := file_sharing_elem.get("spinCount", None):
        file_sharing.spin_count = safe_int(spin_count, 0)

    if reservation_password := file_sharing_elem.get("reservationPassword", None):
        file_sharing.reservation_password = decode_hex_binary(reservation_password)

    if hash_value := file_sharing_elem.get("hashValue", None):
        file_sharing.hash_value = decode_base64_binary(hash_value)

    if salt_value := file_sharing_elem.get("saltValue", None):
        file_sharing.salt_value = decode_base64_binary(salt_value)

    return file_sharing


def _parse_workbook_protection(workbook_elem: Element) -> WorkbookProtection | None:
    protection_elem: Element | None = workbook_elem.find(WORKBOOK_PROTECTION_TAG, None)
    if protection_elem is None:
        return None

    protection = WorkbookProtection()

    if workbook_password := protection_elem.get("workbookPassword", None):
        protection.workbook_password = workbook_password

    if workbook_password_character_set := protection_elem.get(
        "workbookPasswordCharacterSet", None
    ):
        protection.workbook_password_character_set = workbook_password_character_set

    if revision_password := protection_elem.get("revisionsPassword", None):
        protection.revision_password = revision_password

    if revisions_password_character_set := protection_elem.get(
        "revisionsPasswordCharacterSet", None
    ):
        protection.revisions_password_character_set = revisions_password_character_set

    if revisions_algorithm_name := protection_elem.get("revisionsAlgorithmName", None):
        protection.revisions_algorithm_name = revisions_algorithm_name

    if workbook_algorithm_name := protection_elem.get("workbookAlgorithmName", None):
        protection.workbook_algorithm_name = workbook_algorithm_name

    if lock_structure := protection_elem.get("lockStructure", None):
        protection.lock_structure = is_true(lock_structure)

    if lock_windows := protection_elem.get("lockWindows", None):
        protection.lock_windows = is_true(lock_windows)

    if lock_revision := protection_elem.get("lockRevision", None):
        protection.lock_revision = is_true(lock_revision)

    if revisions_spin_count := protection_elem.get("revisionsSpinCount", None):
        protection.revisions_spin_count = safe_int(revisions_spin_count, 0)

    if workbook_spin_count := protection_elem.get("workbookSpinCount", None):
        protection.workbook_spin_count = safe_int(workbook_spin_count, 0)

    if revisions_hash_value := protection_elem.get("revisionsHashValue", None):
        protection.revisions_hash_value = decode_base64_binary(revisions_hash_value)

    if revisions_salt_value := protection_elem.get("revisionsSaltValue", None):
        protection.revisions_salt_value = decode_base64_binary(revisions_salt_value)

    if workbook_hash_value := protection_elem.get("workbookHashValue", None):
        protection.workbook_hash_value = decode_base64_binary(workbook_hash_value)

    if workbook_salt_value := protection_elem.get("workbookSaltValue", None):
        protection.workbook_salt_value = decode_base64_binary(workbook_salt_value)

    return protection


class _SheetMeta(BaseModel):
    name: str
    rid: str
    type: str
    state: SheetState | None = None


def _parse_workbook_sheet_metas(xml_content: bytes) -> list[_SheetMeta]:
    root = from_string(xml_content)
    sheets: list[_SheetMeta] = []

    sheets_elem: Element | None = root.find(SHEETS_TAG, None)
    if sheets_elem is not None:
        sheet_elems: list[Element] = sheets_elem.findall(SHEET_TAG, None)
        for sheet_elem in sheet_elems:
            sheet_name: str = sheet_elem.get("name", "")
            sheet_rid = sheet_elem.get(REL_ID_ATTR, "")
            sheet_type = "worksheet"

            sheet_state = None
            if state_attr := sheet_elem.get("state", None):
                sheet_state = SheetState(state_attr)

            if "chart" in sheet_name.lower():
                sheet_type = "chartsheet"
            elif "dialog" in sheet_name.lower():
                sheet_type = "dialogsheet"

            sheets.append(
                _SheetMeta(
                    name=sheet_name, rid=sheet_rid, type=sheet_type, state=sheet_state
                )
            )

    return sheets


def _parse_active_sheet_index(xml_content: bytes) -> int | None:
    root = from_string(xml_content)

    book_views_elem: Element | None = root.find(BOOK_VIEWS_TAG, None)
    if book_views_elem is not None:
        workbook_view_elem: Element | None = book_views_elem.find(
            WORKBOOK_VIEW_TAG, None
        )
        if workbook_view_elem is not None:
            active_tab: str | None = workbook_view_elem.get("activeTab", None)
            if active_tab is not None:
                return safe_int(active_tab, 0)
    return None


def parse_workbook(xml_content: bytes) -> tuple[Workbook, list[_SheetMeta]]:
    root = from_string(xml_content)
    workbook = Workbook(sheets={})

    if workbook_properties := _parse_workbook_properties(root):
        workbook.properties = workbook_properties

    if calc_properties := _parse_calc_properties(root):
        workbook.calculation = calc_properties

    if workbook_views := _parse_workbook_views(root):
        workbook.views = workbook_views

    if defined_names := _parse_defined_names(root):
        workbook.defined_names = defined_names

    if file_version := _parse_file_version(root):
        workbook.file_version = file_version

    if file_sharing := _parse_file_sharing(root):
        workbook.file_sharing = file_sharing

    if workbook_protection := _parse_workbook_protection(root):
        workbook.protection = workbook_protection

    sheet_metas = _parse_workbook_sheet_metas(xml_content)

    active_sheet_index = _parse_active_sheet_index(xml_content)
    if active_sheet_index is not None:
        workbook.active_sheet = sheet_metas[active_sheet_index].name

    return workbook, sheet_metas
