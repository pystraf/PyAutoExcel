"""
A class representing a range of cells in an Excel-style spreadsheet.
"""
from typing import Optional, Union

from openpyxl.utils.cell import coordinate_to_tuple

from .Cell import Cell
from .Utils import pos2string


class CellRange:
    """
    A class representing a range of cells in an Excel-style spreadsheet.
    """

    def __init__(
        self,
        row_start: Union[
            Cell,
            str,
            int,
            range,
            slice,
            tuple[int, int],
            tuple[int, int, int, int],
            tuple[tuple[int, int], tuple[int, int]],
        ],
        col_start: Union[Cell, str, int, tuple, range, slice, None] = None,
        row_end: Optional[int] = None,
        col_end: Optional[int] = None,
    ):
        # CellRange(0, 0, 1, 1)
        if all(
            map(lambda x: isinstance(x, int), [row_start, col_start, row_end, col_end])
        ):
            self.row_start = row_start
            self.col_start = col_start
            self.row_end = row_end
            self.col_end = col_end
        # CellRange('A1', 'B2')
        elif all(map(lambda x: isinstance(x, str), [row_start, col_start])):
            self.row_start, self.col_start = map(
                lambda x: x - 1, coordinate_to_tuple(coordinate=row_start)
            )
            self.row_end, self.col_end = map(
                lambda x: x - 1, coordinate_to_tuple(coordinate=col_start)
            )
        # CellRange(Cell(0, 0), Cell(1, 1))
        elif all(map(lambda x: isinstance(x, Cell), [row_start, col_start])):
            self.row_start, self.col_start = row_start.row, row_start.col
            self.row_end, self.col_end = col_start.row, col_start.col
        # CellRange((0, 0), (1, 1))
        elif all(map(lambda x: isinstance(x, (tuple, list)), [row_start, col_start])):
            self.row_start, self.col_start = row_start
            self.row_end, self.col_end = col_start
        # CellRange(range(0, 2), range(0, 2))
        # CellRange(slice(0, 2), slice(0, 2))
        elif all(map(lambda x: isinstance(x, (range, slice)), [row_start, col_start])):
            self.row_start = row_start.start
            self.row_end = row_start.stop - 1
            self.col_start = col_start.start
            self.col_end = col_start.stop - 1
        # CellRange(1, 1)
        elif all(map(lambda x: isinstance(x, int), [row_start, col_start])):
            self.row_start = 0
            self.col_start = 0
            self.row_end = row_start
            self.col_end = col_start
        elif isinstance(row_start, str):
            # CellRange('A1:B2')
            if ":" in row_start:
                start, end = row_start.split(":")
                row1, col1 = coordinate_to_tuple(coordinate=start)
                row2, col2 = coordinate_to_tuple(coordinate=end)
                self.row_start = row1 - 1
                self.col_start = col1 - 1
                self.row_end = row2 - 1
                self.col_end = col2 - 1
            # CellRange('B2')
            else:
                row, col = coordinate_to_tuple(coordinate=row_start)
                self.row_start = 0
                self.col_start = 0
                self.row_end = row - 1
                self.col_end = col - 1
        # CellRange(Cell(1, 1))
        elif isinstance(row_start, Cell):
            self.row_start = 0
            self.col_start = 0
            self.row_end = row_start.row
            self.col_end = row_start.col
        # CellRange((1, 1))
        # CellRange((0, 0, 1, 1))
        elif isinstance(row_start, tuple):
            if len(row_start) == 2:
                data = (0, 0, row_start[0], row_start[1])
            else:
                data = row_start
            self.row_start, self.col_start, self.row_end, self.col_end = data
        else:
            raise ValueError("Cannot initinitalizing CellRange object.")

    def __repr__(self):
        return (
            "PyAutoExcel.CellRange(row_start=%d, col_start=%d, row_end=%d, col_end=%d)"
            % (self.row_start, self.col_start, self.row_end, self.col_end)
        )

    def __hash__(self):
        return hash((self.row_start, self.col_start, self.row_end, self.col_end))

    def __eq__(self, other):
        return all(
            [
                self.row_start == other.row_start,
                self.col_start == other.col_start,
                self.row_end == other.row_end,
                self.col_end == other.col_end,
            ]
        )

    @property
    def area(self):
        """
        The number of cells in the range.
        """
        return self.row_count * self.col_count

    @property
    def single_cell(self):
        """
        True if the range contains a single cell.
        """
        return self.row_start == self.row_end and self.col_start == self.col_end

    @property
    def row_count(self):
        """
        The number of rows in the range.
        """
        return self.row_end - self.row_start + 1

    @property
    def col_count(self):
        """
        The number of columns in the range.
        """
        return self.col_end - self.col_start + 1

    def to_string(self):
        """
        Returns a string representation of the Range.

        :return: A string representation of the Range.
        :rtype: str
        """
        return (
            pos2string(self.row_start, self.col_start)
            + ":"
            + pos2string(self.row_end, self.col_end)
        )
