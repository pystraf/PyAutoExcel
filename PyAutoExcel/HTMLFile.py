"""
Write data into HTML web pages like in Excel.
"""
import io
from typing import Union

from .BaseWriter import WriteSheet
from .Grid import ListGrid
from .TableGenerator import BasicTableGenerator


class HtmlRecord:
    """
    Base class for creating HTML record strings.
    """

    def __init__(self):
        """
        Initializes a new instance of the HtmlRecord class.
        """
        self.record = ""

    def get(self) -> str:
        """
        Retrieves the HTML record.

        :return: The HTML record string with a newline character appended.
        :rtype: str
        """
        return self.record + "\n"


class DocTypeRecord(HtmlRecord):
    """
    Represents a DOCTYPE HTML record.
    """

    def __init__(self):
        """
        Initializes a new instance of the DocTypeRecord class with a DOCTYPE HTML declaration.
        """
        super().__init__()
        self.record = "<!DOCTYPE html>"


class HeaderRecord(HtmlRecord):
    """
    Represents an HTML header record.
    """

    def __init__(self, eof: bool = False):
        """
        Initializes a new instance of the HeaderRecord class.

        :param eof: Indicates whether this is the end of the header section.
        :type eof: bool
        """
        super().__init__()
        self.record = "</head>" if eof else "<head>"


class HtmlTagRecord(HtmlRecord):
    """
    Represents an HTML tag record.
    """

    def __init__(self, eof: bool = False):
        """
        Initializes a new instance of the HtmlTagRecord class.

        :param eof: Indicates whether this is the end of the HTML document.
        :type eof: bool
        """
        super().__init__()
        self.record = "</html>" if eof else "<html>"


class TitleRecord(HtmlRecord):
    """
    Represents an HTML title record.
    """

    def __init__(self, title: str):
        """
        Initializes a new instance of the TitleRecord class.

        :param title: The title of the HTML document.
        :type title: str
        """
        super().__init__()
        self.record = f"<title>{title}</title>"


class BodyRecord(HtmlRecord):
    """
    Represents an HTML body record.
    """

    def __init__(self, eof: bool = False):
        """
        Initializes a new instance of the BodyRecord class.

        :param eof: Indicates whether this is the end of the body section.
        :type eof: bool
        """
        super().__init__()
        self.record = "</body>" if eof else "<body>"


class TableRecord(HtmlRecord):
    """
    Represents an HTML table record.
    """

    def __init__(self, tbl: list[list], generator: type = BasicTableGenerator):
        """
        Initializes a new instance of the TableRecord class.

        :param tbl: A 2D list representing the table data.
        :type tbl: list[list]
        """
        super().__init__()
        self.record = generator(tbl).generate().content


class HtmlDumper:
    """
    Facilitates the dumping of HTML content to a file or stream.
    """

    def __init__(self, table_code: str, title: str):
        """
        Initializes a new instance of the HtmlDumper class.

        :param table_code: The HTML code for the table.
        :param title: The title of the HTML document.
        """
        self.title = title
        self.contents = ""
        self.table_code = table_code

    def pre_to_export(self):
        """
        Prepares the HTML content for export.
        """
        self.contents += DocTypeRecord().get()
        self.contents += HtmlTagRecord().get()

        self.contents += HeaderRecord().get()
        self.contents += TitleRecord(title=self.title).get()
        self.contents += HeaderRecord(True).get()

        self.contents += BodyRecord().get()
        self.contents += self.table_code
        self.contents += BodyRecord(True).get()

        self.contents += HtmlTagRecord(True).get()

    def save(self, file: Union[str, io.FileIO, io.StringIO]):
        """
        Saves the prepared HTML content to a file or stream.

        :param file: The file path or file/stream object to save the HTML content to.
        :type file: Union[str, io.FileIO, io.StringIO]
        """
        if isinstance(file, str):
            close = True
            file = open(file, "w")
        else:
            close = False
        file.write(self.contents)
        if close:
            file.close()


def save_html(
    sheet: WriteSheet,
    file: Union[str, io.FileIO, io.StringIO],
    table_generator: type = BasicTableGenerator,
):
    """
    Saves the HTML representation of a sheet to a file or stream.

    :param sheet: The WriteSheet object to be converted into HTML.
    :type sheet: HTMLSheet
    :param file: The file path or file/stream object where the HTML representation will be saved.
    :type file: Union[str, io.FileIO, io.StringIO]
    :param table_generator: The table generator class to use for generating the HTML table.
    :type table_generator: type
    :return: None
    """
    table = sheet.get_grid().get()
    html_table = TableRecord(table, table_generator).get()
    dumper = HtmlDumper(table_code=html_table, title=sheet.name)
    dumper.pre_to_export()
    dumper.save(file=file)


def HTMLSheet(title: str) -> WriteSheet:
    """
    Creates a sheet to writing to .html files.

    :param title: The title of the sheet.
    :return: A sheet to writing to .html files.
    :rtype: WriteSheet
    """
    sheet = WriteSheet(sheet_name=title, grid_class=ListGrid)
    return sheet
