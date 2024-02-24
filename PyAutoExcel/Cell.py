"""
Create a class named "Cell" to represent a cell in an Excel-style spreadsheet.
The class should have attributes to store the row and column indices of the cell,
as well as methods to set and get the value of the cell.
Additionally, implement a method to convert the cell coordinates to a tuple format.
"""
from typing import Optional, Union

from openpyxl.utils.cell import coordinate_to_tuple

from .Utils import pos2string


class Cell:
    """
    A class representing a cell in an Excel-style spreadsheet.
    """

    def __init__(self, row: Union[int, str, tuple], col: Optional[int] = None):
        if isinstance(row, str):
            row, col = coordinate_to_tuple(coordinate=row)
            row -= 1
            col -= 1
        if isinstance(row, tuple):
            row, col = row
        self.row = row
        self.col = col

    def __repr__(self):
        return "PyAutoExcel.Cell(row=%d, col=%d)" % (self.row, self.col)

    def __eq__(self, other):
        return other.row == self.row and other.col == self.col

    def __ne__(self, other):
        return not self == other

    def __hash__(self):
        return hash((self.row, self.col))

    def to_string(self):
        return pos2string(self.row, self.col)
