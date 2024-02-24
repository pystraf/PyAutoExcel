import io
from typing import Type, Union

import proglog

from PyAutoExcel.BaseReader import ReadBook
from PyAutoExcel.Deprecated import process_deprecated as _process_deprecated

from ..Register import Register
from ..Sheet import Sheet

readers = Register()


def add_reader(engine: Type[ReadBook]):
    """
    Registers a new reader engine.

    :param engine: The engine class to be registered, which must be a subclass of ReadBook.
    """
    global readers
    readers.add(engine_name=engine.__engine__, engine=engine)


def remove_reader(name: str):
    """
    Removes a reader engine from the registry by name.

    :param name: The name of the engine to be removed.
    """
    global readers
    try:
        del readers.engines[name]
    except KeyError:
        pass


def install_builtin_readers():
    """
    Registers all built-in reader engines found in the Readers module.
    This function searches for subclasses of ReadBook in the Readers module
         and registers them using their ``__engine__`` attribute.
    """
    global readers
    from .....Engines import Readers

    readers.add_from_module(
        module=Readers,
        sub_class_filters=[ReadBook],
        field_name="__engine__",
    )


def auto_engine(fmt: str):
    """
    Determines the appropriate engine to use based on the file format.

    :param fmt: The format of the file (e.g., 'xls', 'xlsx').
    :return: The name of the engine to use ('xlrd' for 'xls' files, 'openpyxl' for 'xlsx' files).
    """
    return "xlrd" if fmt == "xls" else "openpyxl"


class ExcelReader:
    """
    A class for reading Excel files, supporting multiple engines and formats.

    :param file: The path to the file or a file-like object.
    :param engine: The name of the engine to use for reading.
           If not specified, it is auto-detected based on the file format.
    :param fmt: The format of the file (e.g., 'xls', 'xlsx').
           If not specified, it is inferred from the file name.
    """

    _engine: ReadBook

    def __init__(
        self,
        file: Union[str, io.BytesIO, bytes],
        engine: str = "",
        fmt: str = "",
    ):
        self._params = f"(file={file!r}, engine={engine!r}, fmt={fmt!r})"
        if not engine:
            if fmt:
                engine = auto_engine(fmt)
            else:
                if isinstance(file, str):
                    fmt = file.split(".")[-1].lower()
                    engine = auto_engine(fmt)
        logger = proglog.default_bar_logger("bar")
        logger(message=f"PyAutoExcel - Reading {file}.")
        self._engine = readers.get(engine)(file)
        _process_deprecated(self._engine.__deprecated__, engine)
        self._sheets = []
        self._sheets_dic = {}
        for s in self._engine.sheets():
            st = Sheet(s.name)
            st._load(s)
            self._sheets.append(st)
            self._sheets_dic[s.name] = st
        logger(message="PyAutoExcel - Done.")

    def sheet_by_index(self, idx: int) -> Sheet:
        """
        Return a Sheet object for the sheet at the given index.

        :param idx: The index of the sheet (0-based).
        :return: A ReadSheet object corresponding to the sheet at the given index.
        :rtype: Sheet
        :raise IndexError: If the index is out of range.
        """
        if idx < 0 or idx >= self.nsheets():
            raise IndexError(f"Sheet index out of range: {idx}")
        return self._sheets[idx]

    def sheet_by_name(self, name: str) -> Sheet:
        """
        Return a Sheet object for the sheet with the given name.

        :param name: The name of the sheet.
        :return: A ReadSheet object corresponding to the sheet with the given name.
        :rtype: Sheet
        :raise KeyError: If no sheet with the given name exists.
        """
        if name not in self._sheets_dic:
            raise KeyError(f"No sheet named '{name}'.")
        return self._sheets_dic[name]

    def sheets(self):
        """
        Return a list of ReadSheet objects for all sheets in the workbook.

        :return: A list of ReadSheet objects.
        :rtype: list[ReadSheet]
        """
        return self._sheets.copy()

    def nsheets(self):
        """
        Return the number of sheets in the workbook.

        :return: The number of sheets.
        :rtype: int
        """
        return len(self._sheets)

    def sheet_names(self):
        """
        Return a list of all sheet names in the workbook.

        :return: A list of sheet names.
        :rtype: list[str]
        """
        return list(self._sheets_dic.keys())

    def __repr__(self):
        return (
            f"{self.__class__.__module__}.{self.__class__.__qualname__}{self._params}"
        )

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass
