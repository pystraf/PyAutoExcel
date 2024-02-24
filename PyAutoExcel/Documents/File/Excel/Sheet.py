from PyAutoExcel.BaseReader import ReadSheet
from PyAutoExcel.BaseWriter import WriteSheet
from PyAutoExcel.CellRange import CellRange
from PyAutoExcel.Grid import ListGrid
from PyAutoExcel.Utils import FinalMeta


class Sheet(metaclass=FinalMeta):
    """
    Represents a sheet within an Excel file.

    :param name: The name of the sheet.
    :type name: str
    """

    def __init__(self, name: str):
        """
        Initializes a new instance of the Sheet class.

        :param name: The name of the sheet.
        :type name: str
        """
        self.name = name
        self.grid = ListGrid()
        self.data = [[]]

    def _update(self):
        """Updates the internal data representation from the grid."""
        self.data = self.grid.get().copy()

    def _clean(self):
        """Cleans the cache of internal grid."""
        self.grid._grid = None

    def set_cell(self, row: int, col: int, value):
        """
        Sets the value of a specific cell.

        :param row: The row index of the cell.
        :param col: The column index of the cell.
        :param value: The value to set in the cell.
        """
        self.grid.cell(row, col, value)
        self._update()
        self._clean()

    def set_row(self, row: int, values: list):
        """
        Sets the values of a specific row.

        :param row: The row index.
        :param values: A list of values to set in the row.
        """
        self.grid.row(row, values)
        self._update()
        self._clean()

    def set_col(self, col: int, values: list):
        """
        Sets the values of a specific column.

        :param col: The column index.
        :param values: A list of values to set in the column.
        """
        self.grid.column(col, values)
        self._update()
        self._clean()

    def get_cell(self, row: int, col: int):
        """
        Gets the value of a specific cell.

        :param row: The row index of the cell.
        :param col: The column index of the cell.
        :return: The value of the cell.
        """
        return self.data[row][col]

    def get_row(self, row: int):
        """
        Gets the values of a specific row.

        :param row: The row index.
        :return: A list of values in the row.
        """
        return self.data[row]

    def get_col(self, col: int):
        """
        Gets the values of a specific column.

        :param col: The column index.
        :return: A list of values in the column.
        """
        return [self.get_row(r)[col] for r in range(self.nrows())]

    def get_range(self, rng: CellRange):
        """
        Gets the values in a specific range.

        :param rng: A CellRange object representing the range to get.
        :type rng: CellRange
        :return: A list of lists, where each inner list contains the values in a specific row.
        :rtype: list[list[Any]]
        """
        res = []
        for r in range(rng.row_start, rng.row_end + 1):
            res.append(self.data[r][rng.col_start : rng.col_end + 1])
        return res

    def set_range(self, rng: CellRange, content: list[list]):
        """
        Sets the values in a specific range.

        :param rng: A CellRange object representing the range to set.
        :type rng: CellRange
        :param content: A list of lists,
                        where each inner list contains the values to set in a specific row.
        :type content: list[list[Any]]
        """
        for row in range(rng.row_start, rng.row_end + 1):
            for col in range(rng.col_start, rng.col_end + 1):
                self.grid.cell(
                    row, col, content[row - rng.row_start][col - rng.col_start]
                )
        self._update()
        self._clean()

    def nrows(self):
        """
        Gets the number of rows in the sheet.

        :return: The number of rows.
        """
        return len(self.data)

    def ncols(self):
        """
        Gets the number of columns in the sheet.

        :return: The number of columns.
        """
        return len(self.data[0])

    def _dump(self, s: WriteSheet):
        """
        Dumps the sheet data to a writer.

        :param s: The writer to dump the data to.
        :type s: WriteSheet
        """
        for i, row in enumerate(self.data):
            s.write_row(i, row)

    def _load(self, r: ReadSheet):
        """
        Loads the sheet data from a reader.

        :param r: The reader to load the data from.
        :type r: ReadSheet
        """
        for i, row in enumerate(r.data):
            self.set_row(i, row)

    def __repr__(self):
        return f"PyAutoExcel.Documents.File.Excel.Sheet(name={self.name!r})"
