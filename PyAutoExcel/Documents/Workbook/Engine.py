"""
Base class of excel engine.

History (most recent first):
3.0.1  2024/2/7     Add docstrings.
2.0.1  2023/3/11    Created.
"""

# Importing modules

# bulitin
import io
from abc import ABC, abstractmethod
from typing import Union

# self
from ..File.Excel.Sheet import Sheet

ReaderStream = Union[str, io.IOBase, bytes]  # Type of stream to read from.
WriterStream = Union[str, io.IOBase, None]  # Type of stream to write to.


# define EngineBase class
class EngineBase(ABC):
    """
    Base class of excel engine.
    Abstract methods should be implemented by subclasses.
    """

    @classmethod
    @abstractmethod
    def read(cls, file: ReaderStream) -> list[Sheet]:
        """
        Read excel file and return a list of sheets.
        Subclasses should implement this method.

        :param file: The file to read.
        :type file: Union[str, io.IOBase, bytes]
        :return: A list of sheets.
        :rtype: list[Sheet]
        """
        pass

    @classmethod
    @abstractmethod
    def save(cls, file: WriterStream, sheets: list[Sheet]):
        """
        Save a list of sheets to an excel file.
        Subclasses should implement this method.

        :param file: The file to save to.
        :type file: Union[str, io.IOBase, None]
        :param sheets: A list of sheets to save.
        :type sheets: list[Sheet]
        """
        pass
