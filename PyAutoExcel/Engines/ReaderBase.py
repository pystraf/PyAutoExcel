import abc
import io
from typing import Union

from PyAutoExcel.Deprecated import DeprecatedInfo
from PyAutoExcel.Documents.File.Excel.Sheet import Sheet


class BaseReader(abc.ABC):
    """
    Base class for Excel readers.

    This class is abstract and should not be instantiated directly.

    :param file: The file to read.

    Subclasses must implement the following methods:

    - `_setup()`: Set up the reader.
    - `_parse()`: Read the sheets from the file.

    Subclasses must also set the following class variables:

    - `__engine__`: The name of the engine used by the reader.

    Subclasses may also set the following class variable:

    - `__deprecated__`: A DeprecatedInfo object containing information about the reader's deprecation.
    """
    __engine__ = ""
    __deprecated__ = DeprecatedInfo()
    _sheets: list[Sheet]
    _sheet_names: list[str]

    def __init__(self, file: Union[bytes, str, io.IOBase]):
        self._file = file
        self._sheets = []
        self._sheet_names = []
        self._workbook = None
        self._setup()
        self._parse()

    @abc.abstractmethod
    def _setup(self):
        """
        Open the file and prepare it for reading.
        """
        raise NotImplementedError

    @abc.abstractmethod
    def _parse(self):
        """
        Parse the file and create the sheets.
        """
        raise NotImplementedError

    @property
    def sheets(self) -> list[Sheet]:
        """
        Return the list of sheets.
        """
        return self._sheets

    @property
    def nsheets(self) -> int:
        """
        Return the number of sheets.
        """
        return len(self._sheets)

    @property
    def sheet_names(self) -> list[str]:
        """
        Return the list of sheet names.
        """
        return self._sheet_names

    def sheet_by_name(self, name: str):
        """
        Return the sheet with the given name.
        """
        return self._sheets[self._sheet_names.index(name)]

    def sheet_by_index(self, index: int):
        """
        Return the sheet at the given index.
        """
        return self._sheets[index]

