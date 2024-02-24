"""
Base class of excel workbook

History (most recent first):
3.0.1  2024/2/7     Add docstrings.
2.0.1  2023/3/11    Created.

"""
# Importing modules

# bulitin
from abc import ABC
from typing import Optional

# self
from ..File.Excel.Sheet import Sheet as Worksheet  # Worksheet is same as Sheet
from .Engine import EngineBase, ReaderStream, WriterStream  # Base engine class


# define BaseWorkbook class.
class BaseWorkbook(ABC):
    engine: Optional[type[EngineBase]] = None

    def __init__(self):
        self._sheets = []

    def load(self, file: ReaderStream):
        """
        Load excel workbook from file.

        :param file: file stream to load.
        :type file: ReaderStream
        :return: None
        """
        self._sheets = self.engine.read(file)

    def save(self, file: WriterStream):
        """
        Save excel workbook to file.

        :param file: The target file. Can be a string or a file-like object with a write method.
                     If it's None, The file will be saved in memory.
        :type file: WriterStream
        :return: If param 'file' is None, return the file content as bytes.
                 Otherwise, return None.
        :rtype: Optional[bytes]
        """
        res = self.engine.save(file, self._sheets)
        return res

    # from ExcelDocument class.
    def add_sheet(self, s: Worksheet, index: int = -1):
        """
        Add a sheet to workbook.
        If index is -1, append to the end.
        Otherwise, insert at the given index.

        :param s: The sheet to add.
        :type s: Worksheet
        :param index: The index to insert the sheet.
        :type index: int
        """
        if index == -1:
            self._sheets.append(s)
        else:
            self._sheets.insert(index, s)

    def __repr__(self):
        return f"{self.__class__.__module__}.{self.__class__.__qualname__}()"

    @property
    def sheets(self):
        """
        Return all sheets in workbook.

        :return: List of all sheets in workbook.
        :rtype: list[Worksheet]
        """
        return self._sheets

    @property
    def nsheets(self):
        """
        Return the number of sheets in workbook.

        :return: The number of sheets in workbook.
        :rtype: int
        """
        return len(self._sheets)

    def sheet_by_name(self, name: str):
        """
        Return the sheet with the given name.

        :param name: The name of the sheet.
        :type name: str
        :return: The sheet with the given name.
        :rtype: Worksheet
        :raise KeyError: If no sheet with the given name exists.
        """

        if name not in [sheet.name for sheet in self._sheets]:
            raise KeyError(f"No sheet named '{name}'.")
        return [sheet for sheet in self._sheets if sheet.name == name][0]

    def sheet_by_index(self, idx: int):
        """
        Return the sheet with the given index.

        :param idx: The index of the sheet.
        :type idx: int
        :return: The sheet with the given index.
        :rtype: Worksheet
        :raises IndexError: If the index is out of range.
        """
        return self._sheets[idx]
