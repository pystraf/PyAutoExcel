"""
Write table into .html files.
"""
import io
from typing import Union

from .. import HTMLFile
from ..TableGenerator import BasicTableGenerator


class TBL_HTMLWriter:
    """
    This class represents a writer for generating HTML tables from data.

    :param table_name: The name of the table.
    :type table_name: str
    :param data: The data to be written into the table.
    :type data: list[list]
    """

    def __init__(
        self,
        table_name: str,
        data: list[list],
        generator: type = BasicTableGenerator,
    ):
        self._sheet = HTMLFile.HTMLSheet(table_name)
        for i, row in enumerate(data):
            self._sheet.write_row(i, row)
        self._generator = generator

    @property
    def contents(self):
        """
        Get the HTML contents of the table.

        :return: The HTML contents of the table.
        :rtype: str
        """
        stream = io.StringIO()
        HTMLFile.save_html(self._sheet, stream, self._generator)
        return stream.getvalue()

    def write(self, file: Union[str, io.FileIO]):
        """
        Write the HTML contents of the table to a file.

        :param file: The file to write the HTML contents to.
        :type file: Union[str, io.FileIO]
        """
        HTMLFile.save_html(self._sheet, file, self._generator)
