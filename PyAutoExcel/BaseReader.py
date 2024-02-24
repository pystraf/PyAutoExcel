"""
Base classes for read engines.
"""
import io
import os
from abc import ABCMeta, abstractmethod
from os.path import abspath, dirname, join
from typing import Union

from . import CellRange, Deprecated

CUR = dirname(abspath(__file__))
TEMPDIR = join(CUR, "TempFiles")


class ReadSheet:
    """
    A class to wrap parsed worksheet data.

    This class is designed to provide a convenient interface to access data from a worksheet.
    Instances of this class are not meant to be created directly by the user; instead, they are
    returned by methods of the ReadBook class, such as `sheet_by_index` and `sheet_by_name`.

    :param data: The raw data of the worksheet as a list of lists.
    :param name: The name of the worksheet.
    :ivar _data: The raw data of the worksheet stored internally.
    :ivar _translated: The transposed data of the worksheet stored internally.
    :ivar name: The name of the worksheet.
    """

    def __init__(self, data: list[list], name: str):
        """
        Init the worksheet
        :param data: list of list
        :param name: the name of the worksheet
        """
        self._data = data
        self._translated = list(zip(*data))
        self.name = name

    @property
    def nrows(self):
        """
        Get the number of rows in the worksheet.

        :return: The number of rows.
        """
        return len(self._data)

    @property
    def ncolumns(self):
        """
        Get the number of columns in the worksheet.

        :return: The number of columns.
        """
        return len(self._data[0])

    def read_cell(self, row: int, column: int):
        """
        Retrieve the value of a cell at a specified row and column.

        :param row: The row number (0-based index) of the cell.
        :param column: The column number (0-based index) of the cell.
        :return: The value of the cell.
        """
        return self._data[row][column]

    def read_row(self, row: int):
        """
        Retrieve all values from a specified row.

        :param row: The row number (0-based index) to retrieve.
        :return: A list of values from the specified row.
        """
        return list(self._data[row]).copy()

    def read_column(self, column: int):
        """Retrieve all values from a specified column.

        :param column: The column number (0-based index) to retrieve.
        :return: A list of values from the specified column.
        """
        return list(self._translated[column])

    def read_range(self, rng: CellRange.CellRange):
        """
        Retrieve all values from a specified range.

        :param rng: The range to retrieve.
        :return: A list of lists of values from the specified range.
        """
        return [
            self._data[r][rng.col_start : rng.col_end + 1]
            for r in range(rng.row_start, rng.row_end + 1)
        ]

    @property
    def data(self):
        """
        Get a copy of the worksheet data.

        :return: A copy of the worksheet data as a list of lists.
        """
        return self._data.copy()

    def __repr__(self):
        return "PyAutoExcel.BaseReader.ReadSheet(data=%s, name=%s)" % (
            repr(self._data),
            repr(self.name),
        )


class ReadBook(metaclass=ABCMeta):
    """
    Abstract base class for reading workbook data.

    This class defines the structure and required methods for reading data from different types
    of workbook files. Concrete implementations should provide the specific parsing logic.

    :ivar __engine__: Identifier for the engine used to parse the workbook.
    :ivar __deprecated__: Information about deprecated features.
    :ivar _datas: A dictionary mapping sheet names to their data.
    :ivar _datas_list: A list of sheet data, indexed by sheet order in the workbook.
    :ivar file_name: The path to the workbook file.
    """

    __engine__ = ""
    __deprecated__ = Deprecated.DeprecatedInfo()

    def __init__(self, file: Union[bytes, str, io.IOBase]):
        """
        Initialize the ReadBook object with a file.

        :param file: The file path as a string, a bytes object containing the file data,
                     or an IO stream.
        """
        self._datas = {}
        self._datas_list = []
        if isinstance(file, str):
            self.file_name = file
            self._parse()
            self._indexing()
        elif isinstance(file, bytes):
            self._dump_bytes(stream=file)
        else:
            self._dump_io(stream=file)

    def _dump_io(self, stream: io.IOBase):
        """
        Reads data from an IO stream and writes it to a temporary file for parsing.

        :param stream: An IO stream that contains the workbook data.
        """
        stream.seek(0)
        data = stream.read()
        path = join(TEMPDIR, "streaming_cache.tmp")
        with open(path, "wb") as fp:
            fp.write(data)
        self.file_name = path
        self._parse()
        self._indexing()
        os.remove(path)

    def _dump_bytes(self, stream: bytes):
        """
        Writes byte stream data to a temporary file for parsing.

        :param stream: A bytes object that contains the workbook data.
        """
        path = join(TEMPDIR, "streaming_cache.tmp")
        with open(path, "wb") as fp:
            fp.write(stream)
        self.file_name = path
        self._parse()
        self._indexing()
        os.remove(path)

    @abstractmethod
    def _parse(self):
        """
        Abstract method to parse the workbook data.

        This method must be implemented by subclasses to parse the workbook file.
        """

    def _indexing(self):
        """
        Index the parsed data for quick access to sheets by index or name.
        """
        self._datas_list = list(self._datas.values())

    def sheet_by_index(self, idx: int) -> ReadSheet:
        """
        Return a ReadSheet object for the sheet at the given index.

        :param idx: The index of the sheet (0-based).
        :return: A ReadSheet object corresponding to the sheet at the given index.
        """
        return ReadSheet(data=self._datas_list[idx], name=self.sheet_names()[idx])

    def sheet_by_name(self, name: str) -> ReadSheet:
        """
        Return a ReadSheet object for the sheet with the given name.

        :param name: The name of the sheet.
        :return: A ReadSheet object corresponding to the sheet with the given name.
        """
        return ReadSheet(data=self._datas[name], name=name)

    def sheets(self):
        """
        Return a list of ReadSheet objects for all sheets in the workbook.

        :return: A list of ReadSheet objects.
        """
        return [self.sheet_by_index(idx=s) for s in range(len(self._datas))]

    def nsheets(self):
        """
        Return the number of sheets in the workbook.

        :return: The number of sheets.
        """
        return len(self._datas_list)

    def sheet_names(self):
        """
        Return a list of all sheet names in the workbook.

        :return: A list of sheet names.
        """
        return list(self._datas.keys())

    def __repr__(self):
        return (
            f"{self.__class__.__module__}."
            f"{self.__class__.__qualname__}"
            f"(file_name={self.file_name!r})"
        )

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass
