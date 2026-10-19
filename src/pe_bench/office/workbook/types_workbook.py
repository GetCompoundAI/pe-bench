from enum import StrEnum

from pydantic import BaseModel


class DefinedName(BaseModel):
    name: str
    value: str
    comment: str | None = None
    custom_menu: str | None = None
    description: str | None = None
    help: str | None = None
    status_bar: str | None = None
    local_sheet_id: int | None = None
    hidden: bool | None = None
    function: bool | None = None
    vb_procedure: bool | None = None
    xlm: bool | None = None
    function_group_id: int | None = None
    shortcut_key: str | None = None
    publish_to_server: bool | None = None
    workbook_parameter: bool | None = None


class ShowObjectsType(StrEnum):
    ALL = "all"
    PLACEHOLDERS = "placeholders"


class UpdateLinksType(StrEnum):
    USER_SET = "userSet"
    NEVER = "never"
    ALWAYS = "always"


class WorkbookProperties(BaseModel):
    date_1904: bool | None = None
    date_compatibility: bool | None = None
    show_objects: ShowObjectsType | None = None
    show_border_unselected_tables: bool | None = None
    filter_privacy: bool | None = None
    prompted_solutions: bool | None = None
    show_ink_annotation: bool | None = None
    backup_file: bool | None = None
    save_external_link_values: bool | None = None
    update_links: UpdateLinksType | None = None
    code_name: str | None = None
    hide_pivot_field_list: bool | None = None
    show_pivot_chart_filter: bool | None = None
    allow_refresh_query: bool | None = None
    publish_items: bool | None = None
    check_compatibility: bool | None = None
    auto_compress_pictures: bool | None = None
    refresh_all_connections: bool | None = None
    default_theme_version: int | None = None


class CalcModeType(StrEnum):
    MANUAL = "manual"
    AUTO = "auto"
    AUTO_NO_TABLE = "autoNoTable"


class RefModeType(StrEnum):
    A1 = "A1"
    R1C1 = "R1C1"


class CalcProperties(BaseModel):
    calc_id: int
    calc_mode: CalcModeType
    full_calc_on_load: bool | None = None
    ref_mode: RefModeType | None = None
    iterate: bool | None = None
    iterate_count: int | None = None
    iterate_delta: float | None = None
    full_precision: bool | None = None
    calc_completed: bool | None = None
    calc_on_save: bool | None = None
    concurrent_calc: bool | None = None
    concurrent_manual_count: int | None = None
    force_full_calc: bool | None = None


class FileVersion(BaseModel):
    app_name: str | None = None
    last_edited: str | None = None
    lowest_edited: str | None = None
    rup_build: str | None = None
    code_name: str | None = None


class WorkbookProtection(BaseModel):
    workbook_password: str | None = None
    workbook_password_character_set: str | None = None
    revision_password: str | None = None
    revisions_password_character_set: str | None = None
    lock_structure: bool | None = None
    lock_windows: bool | None = None
    lock_revision: bool | None = None
    revisions_algorithm_name: str | None = None
    revisions_hash_value: bytes | None = None
    revisions_salt_value: bytes | None = None
    revisions_spin_count: int | None = None
    workbook_algorithm_name: str | None = None
    workbook_hash_value: bytes | None = None
    workbook_salt_value: bytes | None = None
    workbook_spin_count: int | None = None


class FileSharing(BaseModel):
    read_only_recommended: bool | None = None
    user_name: str | None = None
    reservation_password: bytes | None = None
    algorithm_name: str | None = None
    hash_value: bytes | None = None
    salt_value: bytes | None = None
    spin_count: int | None = None


class VisibilityType(StrEnum):
    VISIBLE = "visible"
    HIDDEN = "hidden"
    VERY_HIDDEN = "veryHidden"


class WorkbookView(BaseModel):
    visibility: VisibilityType | None = None
    minimized: bool | None = None
    show_horizontal_scroll: bool | None = None
    show_vertical_scroll: bool | None = None
    show_sheet_tabs: bool | None = None
    x_window: int | None = None
    y_window: int | None = None
    window_width: int | None = None
    window_height: int | None = None
    tab_ratio: int | None = None
    first_sheet: int | None = None
    active_tab: int | None = None
    auto_filter_date_grouping: bool | None = None
