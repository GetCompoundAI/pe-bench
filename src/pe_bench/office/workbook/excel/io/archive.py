import io
import zipfile
from types import TracebackType

from pe_bench.office.workbook.excel.io.constants import ArchiveFilePaths


class ExcelArchiveReader:
    def __init__(self, file_bytes: io.BytesIO):
        self._zip_file = zipfile.ZipFile(file_bytes, "r")

    def __enter__(self):
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ) -> None:
        self.close()

    def close(self):
        if hasattr(self, "_zip_file"):
            self._zip_file.close()

    def _list_files(self) -> list[str]:
        return self._zip_file.namelist()

    def _has_file(self, path: str) -> bool:
        return path in self._zip_file.namelist()

    def _read_bytes(self, path: str) -> bytes:
        try:
            return self._zip_file.read(path)
        except KeyError as e:
            raise KeyError(f"File '{path}' not found in archive") from e

    def read_content_types(self) -> bytes:
        return self._read_bytes(ArchiveFilePaths.CONTENT_TYPES)

    def read_package_rels(self) -> bytes:
        return self._read_bytes(ArchiveFilePaths.ROOT_RELS)

    def read_workbook(self) -> bytes:
        return self._read_bytes(ArchiveFilePaths.WORKBOOK)

    def read_workbook_rels(self) -> bytes:
        return self._read_bytes(ArchiveFilePaths.WORKBOOK_RELS)

    def read_worksheet(self, sheet_path: str) -> bytes:
        return self._read_bytes(sheet_path)

    def read_shared_strings(self) -> bytes | None:
        if self._has_file(ArchiveFilePaths.SHARED_STRINGS):
            return self._read_bytes(ArchiveFilePaths.SHARED_STRINGS)
        return None

    def read_styles(self) -> bytes | None:
        if self._has_file(ArchiveFilePaths.STYLES):
            return self._read_bytes(ArchiveFilePaths.STYLES)
        return None

    def read_theme(self) -> bytes | None:
        if self._has_file(ArchiveFilePaths.THEME):
            return self._read_bytes(ArchiveFilePaths.THEME)
        return None

    def read_core_properties(self) -> bytes | None:
        if self._has_file(ArchiveFilePaths.CORE):
            return self._read_bytes(ArchiveFilePaths.CORE)
        return None

    def read_app_properties(self) -> bytes | None:
        if self._has_file(ArchiveFilePaths.APP):
            return self._read_bytes(ArchiveFilePaths.APP)
        return None

    def get_worksheet_paths(self) -> list[str]:
        return [
            f
            for f in self._list_files()
            if f.startswith(ArchiveFilePaths.WORKSHEETS) and f.endswith(".xml")
        ]

    def get_relationship_paths(self) -> list[str]:
        return [f for f in self._list_files() if f.endswith(".rels")]

    def read_worksheet_rels(self, sheet_number: int) -> bytes | None:
        path = ArchiveFilePaths.WORKSHEET_RELS_TEMPLATE.format(sheet_number)
        if self._has_file(path):
            return self._read_bytes(path)
        return None

    def get_worksheet_rels_paths(self) -> list[str]:
        return [
            f
            for f in self._list_files()
            if f.startswith(ArchiveFilePaths.WORKSHEET_RELS)
        ]
