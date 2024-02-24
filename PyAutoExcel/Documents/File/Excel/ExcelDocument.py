"""
The Document class represents an Excel file and provides methods to read and write its content.
"""
import io
from typing import Union

from .Reader.Excel import ExcelReader
from .Sheet import Sheet
from .Writer.Excel import ExcelWriter


class Document:
    def __init__(self):
        self._sheets = []

    def load(
        self,
        file: Union[str, io.BytesIO, bytes],
        engine: str = "",
        fmt: str = "",
    ):
        """
        Loads an Excel document from a file.

        :param file: The file to load. Can be a file path, a file-like object, or bytes.
        :type file: Union[str, io.BytesIO, bytes]
        :param engine: The engine to use for loading the document. Default, it's auto-detected.
        :type engine: str
        :param fmt: The format of the file. Default, it's auto-detected.
        :type fmt: str
        """
        reader = ExcelReader(file, engine, fmt)
        self._sheets = reader.sheets().copy()

    def save(
        self,
        saver: Union[None, str, io.BytesIO] = None,
        engine: str = "",
        fmt: str = "xlsx",
    ):
        """
        Saves the current document to a file.

        :param saver: The target file. If not provided, it will be saved to memory.
        :type saver: Union[None, str, io.BytesIO]
        :param engine: The engine to use for saving the document. Default, it's auto-detected.
        :type engine: str
        :param fmt: The format of the file. Default, it's "xlsx".
        :type fmt: str
        :return: If param 'saver' is None, return the content as bytes.
                 Otherwise, return None.
        """
        writer = ExcelWriter(engine, fmt)
        for s in self._sheets:
            writer.add_sheet(s)
        return writer.save(saver)

    def add_sheet(self, s: Sheet, index: int = -1):
        """
        Add a sheet to the document.
        If index is -1, the sheet will be added to the end of the list of sheets.

        :param s: The sheet to add.
        :type s: str
        :param index: Index at which to add the sheet (default is -1).
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
        Return a list of Sheet objects for all sheets in the workbook.

        :return: A list of Sheet objects.
        """
        return self._sheets

    @property
    def nsheets(self):
        """
        Return the number of sheets in the workbook.

        :return: The number of sheets in the workbook.
        """
        return len(self._sheets)

    def sheet_by_name(self, name: str):
        """
        Return a Sheet object for the sheet with the given name.

        :param name: The name of the sheet.
        :return: A Sheet object corresponding to the sheet with the given name.
        :raise KeyError: If no sheet with the given name exists.
        """
        if name not in [sheet.name for sheet in self._sheets]:
            raise KeyError(f"No sheet named '{name}'.")
        return [sheet for sheet in self._sheets if sheet.name == name][0]

    def sheet_by_index(self, idx: int):
        """
        Return a Sheet object for the sheet at the given index.

        :param idx: The index of the sheet (0-based).
        :return: A Sheet object corresponding to the sheet at the given index.
        :raise OverflowError: If the index is out of range.
        """
        if idx < 0 or idx >= len(self._sheets):
            raise OverflowError(f"Sheet index out of range: {idx}")
        return self._sheets[idx]
