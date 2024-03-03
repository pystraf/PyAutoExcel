import io
from typing import Type, Union

import proglog

from PyAutoExcel.Deprecated import process_deprecated as _process_deprecated
from PyAutoExcel.Engines.WriterBase import BaseWriter
from ..Reader.Excel import ExcelReader
from ..Register import Register
from ..Sheet import Sheet

writers = Register()


def add_writer(engine: Type[BaseWriter]):
    """
    Registers a new writer engine.

    :param engine: The engine class to be registered, which must be a subclass of WriteBook.
    """
    global writers
    writers.add(engine_name=engine.__engine__, engine=engine)


def remove_writer(name: str):
    """
    Removes a writer engine from the registry by name.

    :param name: The name of the engine to be removed.
    """
    global writers
    try:
        del writers.engines[name]
    except KeyError:
        pass


def install_builtin_writers():
    """
    Registers all built-in writer engines found in the Readers module.
    This function searches for subclasses of WriteBook in the Readers module
         and registers them using their ``__engine__`` attribute.
    """
    global writers
    from .....Engines import Writers

    writers.add_from_module(
        module=Writers,
        sub_class_filters=[BaseWriter],
        field_name="__engine__",
    )


def auto_engine(fmt: str):
    """
    Determines the appropriate engine to use based on the file format.

    :param fmt: The format of the file (e.g., 'xls', 'xlsx').
    :return: The name of the engine to use ('xlwt' for 'xls' files, 'xlsxwriter' for 'xlsx' files).
    """
    return "xlwt" if fmt == "xls" else "xlsxwriter"


class ExcelWriter:
    """
    A class for writing Excel files, supporting multiple engines and formats.

    :param engine: The name of the engine to use for writing.
           If not specified, it is auto-detected based on the file format.
    :param fmt: The format of the file (e.g., 'xls', 'xlsx'). Default to 'xlsx'.
    """

    _engine: BaseWriter

    def __init__(self, engine: str = "", fmt: str = "xlsx"):
        self._params = f"(engine={engine!r}, fmt={fmt!r})"
        engine = engine or auto_engine(fmt)
        self._engine = writers.get(engine)()
        _process_deprecated(self._engine.__deprecated__, engine)

    def add_sheet(self, s: Sheet, index: int = -1):
        """
        Adds a sheet to the Excel file at the specified index.

        :param s: The sheet to add.
        :type s: Sheet
        :param index: The index at which to insert the sheet.
                      Defaults to -1, which appends the sheet.
        :type index: int
        """
        if index == -1:
            self._engine.sheets.append(s)
        else:
            self._engine.sheets.insert(index, s)

    def get_sheet(self, name_or_idx: Union[int, str]):
        """
        Retrieves a sheet by its name or index.

        :param name_or_idx: The name or index of the sheet.
        :type name_or_idx: Union[int, str]
        :return: The requested sheet.
        :rtype: Sheet
        :raises LookupError: If the sheet cannot be found.
        """
        if isinstance(name_or_idx, int):
            return self._engine.sheets[name_or_idx]
        for s in self._engine.sheets:
            if s.name == name_or_idx:
                return s
        raise LookupError(f"Cannot find sheet {name_or_idx!r}.")

    @property
    def sheets(self):
        """
        Returns a copy of the list of sheets.

        :return: A list of sheets.
        :rtype: list[Sheet]
        """
        return self._engine.sheets

    @property
    def sheet_names(self):
        """
        Returns a list of the names of all sheets.

        :return: A list of sheet names.
        :rtype: list[str]
        """
        return [s.name for s in self._engine.sheets]

    def save(self, saver: Union[None, str, io.BytesIO] = None):
        """
        Saves the Excel file to the specified location or stream.

        :param saver: The file path or stream to save to.
                      If None, the file will be saved to memory.
        :type saver: Union[None, str, io.BytesIO]
        :return: If param 'saver' is None, return the content as bytes.
                 Otherwise, return None.
        """
        logger = proglog.default_bar_logger("bar")
        logger(message=f"PyAutoExcel - Writing {saver if saver is not None else 'into memory'}.")
        res = self._engine.save(saver)
        logger(message="PyAutoExcel - Done.")
        return res


    @classmethod
    def from_reader(cls, reader: ExcelReader, use_engine: str = "", fmt: str = "xlsx"):
        """
        Create a copy from the ExcelReader for writing.

        :param reader: The ExcelReader to copy from.
        :type reader: ExcelReader
        :param use_engine: The engine to use for writing.
        :type use_engine: str
        :param fmt: The file format to use for writing.
        :type fmt: str

        :return: A new ExcelWriter instance.
        :rtype: ExcelWriter
        """
        writer = cls(use_engine, fmt)
        writer._engine._sheets = reader._engine.sheets
        return writer

    def __repr__(self):
        return (
            f"{self.__class__.__module__}.{self.__class__.__qualname__}{self._params}"
        )

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass
