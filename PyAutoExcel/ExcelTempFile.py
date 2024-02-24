"""
Create temporary excel files.
"""
import os
import random
from os.path import abspath, dirname, join
from typing import Optional

from .Documents.File.Excel.ExcelDocument import Document
from .Documents.File.Excel.Sheet import Sheet

CUR = dirname(abspath(__file__))
TEMPDIR = join(CUR, "TempFiles")


# From tempfile.py
class _RandomNameSequence:
    """An instance of _RandomNameSequence generates an endless
    sequence of unpredictable strings which can safely be incorporated
    into file names.  Each string is eight characters long.  Multiple
    threads can safely use the same instance at the same time.

    _RandomNameSequence is an iterator."""

    characters = "abcdefghijklmnopqrstuvwxyz0123456789_"
    _rng: random.Random
    _rng_pid: int

    @property
    def rng(self):
        cur_pid = os.getpid()
        if cur_pid != getattr(self, "_rng_pid", None):
            self._rng = random.Random()
            self._rng_pid = cur_pid
        return self._rng

    def __iter__(self):
        return self

    def __next__(self):
        c = self.characters
        choose = self.rng.choice
        letters = [choose(c) for _ in range(8)]
        return "".join(letters)


_name_seq = _RandomNameSequence()


def get_tempdir() -> str:
    """
    Returns the path to the temporary directory.

    :return: The path to the temporary directory.
    :rtype: str
    """
    return TEMPDIR


class TemporanyExcelFile:
    """
    Represents a temporary Excel file.

    :param file_format: The format of the file (e.g., 'xls', 'xlsx').
    :type file_format: str
    :param sheets: The worksheet contained in the file,
                   can also be added using add_sheet() after creating objects.
    :type sheets: list[Sheet], optional
    :param tempdir: The directory to store this file. Default to TEMPDIR.
    :type tempdir: str
    :param delete_after_use: Whether to delete the file after use.
    :type delete_after_use: bool
    """

    def __init__(
        self,
        file_format: str = "xlsx",
        sheets: Optional[list[Sheet]] = None,
        tempdir: str = "",
        delete_after_use: bool = True,
    ):
        self._file_format = file_format
        self._sheets = sheets if sheets is not None else []
        self._tempdir = get_tempdir() if not tempdir else tempdir
        self._delete_after_use = delete_after_use
        self._gen_filename()

        self._doc = Document()
        for sheet in self._sheets:
            self._doc.add_sheet(sheet)

    def _gen_filename(self):
        """
        Generate a filename for this temporany file.
        """
        self._filename = f"{next(_name_seq)}.{self._file_format}"
        self._path = join(self._tempdir, self._filename)

    def update(self):
        """
        Update the temporany file in the disk.
        """
        self._doc.save(self._path, fmt=self._file_format)

    def add_sheet(self, s: Sheet, index: int = -1):
        """
        Add a sheet to the document.
        If index is -1, the sheet will be added to the end of the list of sheets.

        :param s: The sheet to add.
        :type s: Sheet
        :param index: Index at which to add the sheet (default is -1).
        :type index: int
        """
        self._doc.add_sheet(s, index)

    @property
    def sheets(self):
        """
        Return a list of Sheet objects for all sheets in the workbook.

        :return: A list of Sheet objects.
        """
        return self._doc.sheets

    @property
    def nsheets(self):
        """
        Return the number of sheets in the workbook.

        :return: The number of sheets in the workbook.
        """
        return self._doc.nsheets

    def sheet_by_name(self, name: str):
        """
        Return a Sheet object for the sheet with the given name.

        :param name: The name of the sheet.
        :return: A Sheet object corresponding to the sheet with the given name.
        """
        return self._doc.sheet_by_name(name)

    def sheet_by_index(self, idx: int):
        """
        Return a Sheet object for the sheet at the given index.

        :param idx: The index of the sheet (0-based).
        :return: A Sheet object corresponding to the sheet at the given index.
        """
        return self._doc.sheet_by_index(idx)

    def close(self):
        """
        Close the temporany file and destory the document instance.
        If delete_after_use is True, the file will be deleted.
        """
        if self._delete_after_use:
            os.remove(self._path)
        self._doc = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
