"""
Base classes for write engines.
"""
import io
from abc import ABCMeta, abstractmethod
from os.path import join
from typing import Union

from . import CellRange, Deprecated, Grid


class WriteSheet:
    """
    A class representing a single sheet in a workbook.

    :param sheet_name: The name of the sheet.
    :param grid_class: The grid class to be used for the sheet.
    :ivar name: The name of the sheet.
    :ivar records: A list to store records if no grid is provided.
    """

    def __init__(self, sheet_name: str, grid_class=None):
        """
        Initialize the WriteSheet object with the given sheet name and grid class.

        :param sheet_name: The name of the sheet.
        :param grid_class: The grid class to be used for the sheet.
        """
        self.name = sheet_name
        self.records = []
        if grid_class:
            self.__grid = grid_class()
        else:
            self.__grid = None

    def write_cell(self, row: int, col: int, value):
        """
        Write a value to a cell in the sheet.

        :param row: The row index (0-based) of the cell.
        :param col: The column index (0-based) of the cell.
        :param value: The value to be written to the cell.
        """
        if self.__grid:
            self.__grid.cell(row=row, col=col, value=value)
        else:
            self.records.append((row, col, value))

    def write_row(self, row: int, values: list):
        """
        Write a list of values to a row in the sheet.

        :param row: The row index (0-based) to write the values.
        :param values: A list of values to be written to the row.
        """
        if self.__grid:
            self.__grid.row(row=row, values=values)
        else:
            for i, v in enumerate(values):
                self.write_cell(row=row, col=i, value=v)

    def write_col(self, col: int, values: list):
        """
        Write a list of values to a column in the sheet.

        :param col: The column index (0-based) to write the values.
        :param values: A list of values to be written to the column.
        """
        if self.__grid:
            self.__grid.column(col=col, values=values)
        else:
            for i, v in enumerate(values):
                self.write_cell(row=i, col=col, value=v)

    def write_range(self, rng: CellRange.CellRange, content: list[list]):
        """
        Write a 2D list of values to a range of cells in the sheet.

        :param rng: The CellRange object representing the range of cells to write to.
        :param content: A 2D list of values to be written to the range.
        """
        if self.__grid:
            for row in range(rng.row_start, rng.row_end + 1):
                for col in range(rng.col_start, rng.col_end + 1):
                    self.__grid.cell(
                        row, col, content[row - rng.row_start][col - rng.col_start]
                    )
        else:
            for row in range(rng.row_start, rng.row_end + 1):
                for col in range(rng.col_start, rng.col_end + 1):
                    self.write_cell(
                        row=row,
                        col=col,
                        value=content[row - rng.row_start][col - rng.col_start],
                    )

    def get_grid(self):
        """
        Get the grid object associated with the sheet.

        :return: The grid object associated with the sheet.
        :raises: AttributeError if the sheet has no grid associated with it.
        """
        if self.__grid:
            return self.__grid
        raise AttributeError("'WriteSheet' object has no attribute 'get_grid'")

    def __repr__(self):
        show_grid_class = repr(self.__grid.__class__) if self.__grid else "None"
        return (
            f"{self.__class__.__module__}."
            f"{self.__class__.__qualname__}"
            f"(sheet_name={self.name!r}, grid_class={show_grid_class})"
        )

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass


class WriteBook(metaclass=ABCMeta):
    """
    Abstract base class for writing workbook data.

    :ivar __engine__: Identifier for the engine used to write the workbook.
    :ivar __deprecated__: Information about deprecated features.
    :ivar gcls: The grid class to be used for the workbook.
    """

    __engine__ = ""
    __deprecated__ = Deprecated.DeprecatedInfo()
    gcls = None

    def __init__(self):
        """
        Initialize the WriteBook object.
        """
        self.__sheets = []
        self.__sheet_names = []

    def add_sheet(self, sheet_name: str, index: int = -1):
        """
        Add a new sheet to the workbook.

        :param sheet_name: The name of the sheet to be added.
        :param index: The index at which the sheet should be added.
        :return: The WriteSheet object representing the added sheet.
        """
        ws = WriteSheet(sheet_name=sheet_name, grid_class=self.gcls)
        if index == -1:
            self.__sheet_names.append(sheet_name)
            self.__sheets.append(ws)
        else:
            self.__sheet_names.insert(index, sheet_name)
            self.__sheets.insert(index, ws)
        return ws

    def get_sheet(self, name_or_idx: Union[int, str]) -> WriteSheet:
        """
        Get a sheet from the workbook by name or index.

        :param name_or_idx: The name or index of the sheet to retrieve.
        :return: The WriteSheet object representing the retrieved sheet.
        """
        if isinstance(name_or_idx, int):
            return self.sheets[name_or_idx]
        ret = dict(zip(self.sheet_names, self.sheets)).get(name_or_idx)
        if ret is None:
            raise NameError("Not worksheet named %s" % name_or_idx)

    @property
    def sheets(self) -> list["WriteSheet"]:
        """
        Get the list of sheets in the workbook.

        :return: A list of WriteSheet objects representing the sheets in the workbook.
        """
        return self.__sheets

    @property
    def sheet_names(self) -> list[str]:
        """
        Get the list of sheet names in the workbook.

        :return: A list of sheet names in the workbook.
        """
        return self.__sheet_names

    @abstractmethod
    def save_file(self, file_name: str):
        """
        Abstract method to save the workbook data to a file.

        :param file_name: The name of the file to save the workbook data to.
        """
        pass

    def save_virtual(self):
        """
        Save the workbook data to a virtual file and return the data as bytes.

        :return: The workbook data as bytes.
        """
        import os

        from .ExcelTempFile import TEMPDIR

        file_name = join(TEMPDIR, "excel.tmp")
        self.save_file(file_name)
        with open(file_name, mode="rb") as fp:
            data = fp.read()
        os.remove(file_name)
        return data

    def save_io(self, file: io.IOBase):
        """
        Save the workbook data to an IO stream.

        :param file: The IO stream to save the workbook data to.
        """
        file.write(self.save_virtual())

    def save(self, saver: Union[None, str, io.BytesIO] = None):
        """
        Save the workbook data.

        :param saver: The destination to save the workbook data to.
        :return: The saved workbook data.
        """
        if not saver:
            return self.save_virtual()
        else:
            if isinstance(saver, str):
                self.save_file(file_name=saver)
            else:
                self.save_io(file=saver)

    def __repr__(self):
        return "%s.%s()" % (self.__class__.__module__, self.__class__.__qualname__)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass


class ListGridWriteBook(WriteBook, metaclass=ABCMeta):
    """
    A class representing a workbook with a list grid.

    This class inherits from WriteBook and uses a list grid for storing data.
    """

    gcls = Grid.ListGrid


class ColumnDictGridWriteBook(WriteBook, metaclass=ABCMeta):
    """
    A class representing a workbook with a column dictionary grid.

    This class inherits from WriteBook and uses a column dictionary grid for storing data.
    """

    gcls = Grid.ColumnDictGrid


class RecordGridWriteBook(WriteBook, metaclass=ABCMeta):
    """
    A class representing a workbook with a record grid.

    This class inherits from WriteBook and uses a record grid for storing data.
    """

    gcls = Grid.RecordGrid
