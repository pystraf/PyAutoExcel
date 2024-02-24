from typing import Optional

from odswriter import ODSWriter, Sheet

from PyAutoExcel.Utils import FinalMeta


class OpenSpreadSheetWriter(metaclass=FinalMeta):
    """
    A class for writing data to an OpenDocument Spreadsheet (ODS) file.

    :param file_name: The name of the ODS file.
    :type file_name: str
    :param stream: The stream to write to.
                   If not provided, a file with the given name will be created, defaults to None.
    :type stream: file-like object, optional

    """

    def __init__(self, file_name="", stream=None):
        if not file_name and not stream:
            raise ValueError("Either file_name or stream must be provided.")

        self.__params = f"(file_name={file_name!r}, stream={stream!r})"
        if stream:
            self.__stream = stream
            self.__close = False
        else:
            self.__stream = open(file_name, "wb")
            self.__close = True
        self.__book = ODSWriter(odsfile=self.__stream)
        self.__sheets = {}

    def close(self):
        """
        Closes the ODS file and the associated stream if it was created internally.
        """
        self.__book.close()
        if self.__close:
            self.__stream.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def create_sheet(self, sheet_name: str, ncolumns: Optional[int] = None):
        """
        Creates a new sheet in the ODS file.

        :param sheet_name: The name of the new sheet.
        :type sheet_name: str
        :param ncolumns: The number of columns in the new sheet, defaults to None.
        :type ncolumns: int, optional
        """
        self.__sheets[sheet_name] = self.__book.new_sheet(
            name=sheet_name, cols=ncolumns
        )

    def _get_sheet(self, sheet_name: str) -> Sheet:
        """
        Retrieves a sheet by name.

        :param sheet_name: The name of the sheet to retrieve.
        :type sheet_name: str
        :return: The sheet with the specified name.
        :rtype: Sheet
        """
        return self.__sheets[sheet_name]

    def write_row(self, sheet_name: str, row: list):
        """
        Writes a single row to the specified sheet.

        :param sheet_name: The name of the sheet to write to.
        :type sheet_name: str
        :param row: The row data to write.
        :type row: list
        """
        self._get_sheet(sheet_name=sheet_name).writerow(cells=row)

    def write(self, sheet_name: str, data: list[list]):
        """
        Writes multiple rows of data to the specified sheet.

        :param sheet_name: The name of the sheet to write to.
        :type sheet_name: str
        :param data: The data to write, where each element is a list representing a row.
        :type data: list[list]
        """
        self._get_sheet(sheet_name=sheet_name).writerows(rows=data)

    def __repr__(self):
        return f"PyAutoExcel.OpenSpreadSheet.OpenSpreadSheetWriter{self.__params}"
