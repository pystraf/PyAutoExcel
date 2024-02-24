from abc import ABCMeta, abstractmethod
from typing import Any


class Grid(metaclass=ABCMeta):
    def __init__(self):
        self.records = []
        self._grid = None  # Cache of the grid.

    def get(self):
        """
        Returns a copy of the grid.
        If the grid has not been calculated yet, it will be calculated and then returned.

        :return: A copy of the grid.
        """
        if self._grid:
            return self._grid
        else:
            grid = self._calc_grid()
            self._grid = grid
            return grid.copy()

    @abstractmethod
    def _calc_grid(self):
        """
        Calculates the grid.
        Subclasses must implement this method.
        """
        pass

    def cell(self, row: int, col: int, value):
        """
        Modify the value of a cell.

        :param row: The row of the cell.
        :type row: int
        :param col: The column of the cell.
        :type col: int
        :param value: The new value of the cell.
        :type value: Any
        """
        self.records.append((row, col, value))

    def row(self, row: int, values: list):
        """
        Modify the values of a row.

        :param row: The row to modify.
        :type row: int
        :param values: The new values of the row.
        :type values: list
        """
        for i, v in enumerate(values):
            self.cell(row=row, col=i, value=v)

    def column(self, col: int, values: list):
        """
        Modify the values of a column.

        :param col: The column to modify.
        :type col: int
        :param values: The new values of the column.
        :type values: list
        """
        for i, v in enumerate(values):
            self.cell(row=i, col=col, value=v)

    def __repr__(self):
        return "%s.%s()" % (self.__class__.__module__, self.__class__.__qualname__)


def calc_list_grid_size(records: list[tuple[int, int, Any]]) -> tuple[int, int]:
    """
    Calculate the size of the grid based on the given records.

    :param records: A list of records, where each record is a tuple of (row, column, value).
    :type records: list[tuple[int, int, Any]]
    :return: A tuple of (rows, columns), where rows is the maximum row number
             and columns is the maximum column number.
    :rtype: tuple[int, int]
    """
    max_row = max(records, key=lambda x: x[0])[0]
    max_col = max(records, key=lambda x: x[1])[1]
    return max_row + 1, max_col + 1


class ListGrid(Grid):
    def _calc_grid(self):
        rows, columns = calc_list_grid_size(records=self.records)
        g = []
        for _ in range(rows):
            temp = []
            for _ in range(columns):
                temp.append("")
            g.append(temp)
        for r in self.records:
            g[r[0]][r[1]] = r[2]
        return g

    def __repr__(self):
        return "%s.%s()" % (self.__class__.__module__, self.__class__.__qualname__)


class ColumnDictGrid(ListGrid):
    def _calc_grid(self):
        li = ListGrid._calc_grid(self)
        header = li[0]
        data = list(zip(*li[1:]))
        ret = {}
        for h, d in zip(header, data):
            ret[h] = list(d)
        return ret

    def __repr__(self):
        return "%s.%s()" % (self.__class__.__module__, self.__class__.__qualname__)


class RecordGrid(ListGrid):
    def _calc_grid(self):
        li = ListGrid._calc_grid(self)
        header = li[0]
        data = li[1:]
        ret = []
        for d in data:
            ret.append(dict(zip(header, d)))
        return ret.copy()

    def __repr__(self):
        return "%s.%s()" % (self.__class__.__module__, self.__class__.__qualname__)


def list_to_column_dict(li: list[list]) -> dict:
    """
    Convert a list of rows into a dict of columns. The first row is the header.

    :param li: list of rows
    :type li: list[list]
    :return: dict of columns
    :rtype: dict
    """
    grid = ColumnDictGrid()
    for i, row in enumerate(li):
        grid.row(row=i, values=row)
    return grid.get().copy()


def list_to_records(li: list[list]):
    """
    Convert a list of rows into a list of records (Each item is a dict).
    The first row is the header.

    :param li: list of rows
    :type li: list[list]
    :return: list of records
    :rtype: list[dict]
    """
    grid = RecordGrid()
    for i, row in enumerate(li):
        grid.row(row=i, values=row)
    return grid.get().copy()
