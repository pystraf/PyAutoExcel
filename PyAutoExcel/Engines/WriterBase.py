import abc
import io
from typing import Union, Optional

from PyAutoExcel.Deprecated import DeprecatedInfo
from PyAutoExcel.Documents.File.Excel.Sheet import Sheet


class BaseWriter(abc.ABC):
    """
    Base class for Excel writers.


    Subclasses must implement the following methods:

    - `_setup()`: Set up the writer.
    - `_write()`: Writing data.
    - `_output()`: Export the file.

    Subclasses must also set the following class variables:

    - `__engine__`: The name of the engine used by the reader.

    Subclasses may also set the following class variable:

    - `__deprecated__`: A DeprecatedInfo object containing information about the reader's deprecation.
    """
    __engine__ = ""
    __deprecated__ = DeprecatedInfo()

    def __init__(self):
        self._sheets: list[Sheet] = []
        self._workbook = None
        self._setup()

    @abc.abstractmethod
    def _setup(self):
        """
        Set up the writer.
        """
        raise NotImplementedError

    @abc.abstractmethod
    def _write(self):
        """
        Writing data.
        """
        raise NotImplementedError

    @abc.abstractmethod
    def _output(self, file: Union[str, io.IOBase]):
        """
        Export the file.
        """
        raise NotImplementedError

    def save(self, saver: Union[str, io.IOBase, None]) -> Optional[bytes]:
        """
        Saves the Excel file to the specified location or stream.

        :param saver: The file path or stream to save to.
                      If None, the file will be saved to memory.
        :type saver: Union[None, str, io.BytesIO]
        :return: If param 'saver' is None, return the content as bytes.
                 Otherwise, return None.
        """
        self._write()
        if saver is None:
            stream = io.BytesIO()
            self._output(stream)
            return stream.getvalue()
        else:
            self._output(saver)

    @property
    def sheets(self) -> list[Sheet]:
        """
        Gets the list of sheets.

        :return: The list of sheets.
        """
        return self._sheets

