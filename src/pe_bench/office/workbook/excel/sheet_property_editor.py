from pe_bench.office.workbook.types import Sheet
from pe_bench.office.workbook.types_format import ConditionalFormatting
from pe_bench.office.workbook.types_sheet import ColumnDimension, RowDimension
from pe_bench.office.workbook.utils import CellRangeMapper


class SheetPropertyEditor:
    def __init__(self, sheet: Sheet):
        self._sheet = sheet

        self._merged_mappers: list[CellRangeMapper] = []

        for cell_range in sheet.merged_cell_ranges or []:
            mapper = CellRangeMapper(cell_range)
            if not any(mapper.collide_with_range(r) for r in self._merged_mappers):
                self._merged_mappers.append(mapper)
        self._sheet.merged_cell_ranges = [
            range.range for range in self._merged_mappers
        ] or None

    @property
    def sheet(self) -> Sheet:
        return self._sheet

    def merge_cells(self, cell_range: str, clear_children_cells: bool = True) -> bool:
        if self._sheet.merged_cell_ranges is None:
            self._sheet.merged_cell_ranges = []

        mapper = CellRangeMapper(cell_range)
        if any(mapper.collide_with_range(r) for r in self._merged_mappers):
            return False

        coords, _ = mapper.to_coords()
        if coords:
            for coord in coords[1:]:
                if cell := self._sheet.cells.get(coord):
                    if clear_children_cells:
                        cell.clear_data()
                    else:
                        if cell.v is not None or cell.is_formula():
                            return False

        self._merged_mappers.append(mapper)
        self._sheet.merged_cell_ranges.append(cell_range)
        return True

    def unmerge_cells(self, cell_range: str):
        if self._sheet.merged_cell_ranges is None:
            return
        mapper = CellRangeMapper(cell_range)
        normalized_range = mapper.range
        self._merged_mappers = [
            m for m in self._merged_mappers if m.range != normalized_range
        ]
        self._sheet.merged_cell_ranges = [m.range for m in self._merged_mappers]

    def set_row_dimensions(self, row_ids: list[int], row_dim: RowDimension):
        if row_dim.is_empty():
            self.delete_row_dimensions(row_ids)
            return

        if self._sheet.row_dims_ids is None:
            self._sheet.row_dims_ids = {}
        if self._sheet.row_dims is None:
            self._sheet.row_dims = []

        new_row_id = None
        for i, existing_row_dim in enumerate(self._sheet.row_dims):
            if existing_row_dim == row_dim:
                new_row_id = i
                break
        if new_row_id is None:
            new_row_id = len(self._sheet.row_dims)
            self._sheet.row_dims.append(row_dim)
        for row_id in row_ids:
            self._sheet.row_dims_ids[row_id] = new_row_id

    def delete_row_dimensions(self, row_ids: list[int]):
        if self._sheet.row_dims_ids is None:
            return
        self._sheet.row_dims_ids = {
            k: v for k, v in self._sheet.row_dims_ids.items() if k not in row_ids
        }

    def set_column_dimensions(self, col_group: str, col_dim: ColumnDimension):
        if self._sheet.columns_dims is None:
            self._sheet.columns_dims = {}
        self._sheet.columns_dims[col_group] = col_dim

    def delete_column_dimensions(self, col_group: str):
        if self._sheet.columns_dims is None:
            return
        self._sheet.columns_dims = {
            k: v for k, v in self._sheet.columns_dims.items() if k != col_group
        }

    def set_conditional_formatting(self, conditional_formatting: ConditionalFormatting):
        if self._sheet.conditional_formattings is None:
            self._sheet.conditional_formattings = []
        self._sheet.conditional_formattings.append(conditional_formatting)

    def delete_conditional_formatting(self, conditional_formatting_id: int):
        if self._sheet.conditional_formattings is None:
            return
        if conditional_formatting_id < len(self._sheet.conditional_formattings):
            self._sheet.conditional_formattings.pop(conditional_formatting_id)
