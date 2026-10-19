import functools
import re

from pe_bench.office.workbook.types import Cell

_CELL_COORD_REGEX = re.compile(r"^([A-Z]+)([0-9]+)$")


@functools.lru_cache(maxsize=1000)
def column_letter_to_idx(column_letter: str) -> int:
    return (
        sum(
            (ord(char) - ord("A") + 1) * (26**i)
            for i, char in enumerate(reversed(column_letter))
        )
        - 1
    )


@functools.lru_cache(maxsize=1000)
def idx_to_column_letter(idx: int) -> str:
    result = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


@functools.lru_cache(maxsize=100000)
def cell_coord_to_indices(cell_coord: str) -> tuple[int, int]:
    match = _CELL_COORD_REGEX.match(cell_coord)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_coord}")
    col_str, row_str = match.groups()

    col_idx = column_letter_to_idx(col_str)
    row_idx = int(row_str) - 1
    return row_idx, col_idx


@functools.lru_cache(maxsize=100000)
def indices_to_cell_coord(row_idx: int, col_idx: int) -> str:
    col_str = idx_to_column_letter(col_idx)
    row_str = str(row_idx + 1)
    return col_str + row_str


def get_range_start_and_end(cell_range: str) -> tuple[str, str]:
    segs = cell_range.split(":")
    if len(segs) == 1:
        return segs[0], segs[0]
    elif len(segs) == 2:
        return segs[0], segs[1]
    else:
        raise ValueError(f"Invalid cell range: {cell_range}")


class CellRangeMapper:
    def __init__(self, cell_range: str, dimensions: tuple[int, int] | None = None):
        self._range = cell_range
        self._start_ref, self._end_ref = get_range_start_and_end(cell_range)

        if self._start_ref.isalpha() and self._end_ref.isalpha():
            if dimensions is None:
                raise ValueError("Column range is not supported without dimensions")
            self._start_col_idx = column_letter_to_idx(self._start_ref)
            self._end_col_idx = column_letter_to_idx(self._end_ref)
            self._start_row_idx = 0
            self._end_row_idx = dimensions[0] - 1
        else:
            self._start_row_idx, self._start_col_idx = cell_coord_to_indices(
                self._start_ref
            )
            self._end_row_idx, self._end_col_idx = cell_coord_to_indices(self._end_ref)

    @property
    def range(self) -> str:
        return self._range

    @property
    def start_coord(self) -> str:
        return self._start_ref

    @property
    def end_coord(self) -> str:
        return self._end_ref

    def coord_is_in_range(self, cell_ref: str) -> bool:
        row_idx, col_idx = cell_coord_to_indices(cell_ref)
        return (
            self._start_row_idx <= row_idx <= self._end_row_idx
            and self._start_col_idx <= col_idx <= self._end_col_idx
        )

    def slice_cells(
        self,
        cells: dict[str, Cell],
        max_rows: int | None = None,
        max_cols: int | None = None,
    ) -> tuple[dict[str, Cell], bool]:
        start_row, start_col = self._start_row_idx, self._start_col_idx
        end_row, end_col = self._end_row_idx, self._end_col_idx
        is_truncated = False
        if max_rows is not None and end_row - start_row + 1 > max_rows:
            end_row = start_row + max_rows - 1
            is_truncated = True
        if max_cols is not None and end_col - start_col + 1 > max_cols:
            end_col = start_col + max_cols - 1
            is_truncated = True
        cell_within_range: dict[str, Cell] = {}
        for cell_ref, cell in cells.items():
            row_idx, col_idx = cell_coord_to_indices(cell_ref)
            if start_row <= row_idx <= end_row and start_col <= col_idx <= end_col:
                cell_within_range[cell_ref] = cell
        return cell_within_range, is_truncated

    def to_coords(self) -> tuple[list[str], tuple[int, int]]:
        result: list[str] = []
        for row_idx in range(self._start_row_idx, self._end_row_idx + 1):
            for col_idx in range(self._start_col_idx, self._end_col_idx + 1):
                result.append(indices_to_cell_coord(row_idx, col_idx))
        return result, (
            self._end_row_idx - self._start_row_idx + 1,
            self._end_col_idx - self._start_col_idx + 1,
        )

    def collide_with_range(self, other_range: "CellRangeMapper") -> bool:
        return (
            self._start_col_idx <= other_range._end_col_idx
            and self._end_col_idx >= other_range._start_col_idx
            and self._start_row_idx <= other_range._end_row_idx
            and self._end_row_idx >= other_range._start_row_idx
        )
