"""
Write table into .pdf files.
Need 'pdfkit' and 'wkhtmltox' to run.

:raises ModuleNotFoundError: If 'pdfkit' package or 'wkhtmltox' is not found.
"""
import shutil

import proglog

from ..TableGenerator import BasicTableGenerator
from .HtmlWriter import TBL_HTMLWriter

try:
    import pdfkit
except:
    raise ModuleNotFoundError(
        "PyAutoExcel.Tiny.PdfWriter requires 'pdfkit' package but not found."
        "Please install it and try again."
    )

WK_PATH = shutil.which("wkhtmltopdf")

if not WK_PATH:
    raise ModuleNotFoundError(
        "PyAutoExcel.Tiny.PdfWriter requires 'wkhtmltox' but not found."
        "Please checkout https://wkhtmltopdf.org/downloads.html."
    )

config = pdfkit.configuration(wkhtmltopdf=WK_PATH)


class PDFKIT_PdfWriter:
    """
    PDF writer class using pdfkit to generate PDF files.

    :ivar _options: Dictionary containing PDF generation options.
    :ivar _html_content: HTML content to be converted to PDF.
    """

    def __init__(self):
        self._options = {"quiet": ""}
        self._html_content = ""

    def pdf_option(self, opt_name: str, opt):
        """
        Set a PDF generation option.

        :param opt_name: The name of the option.
        :type opt_name: str
        :param opt: The value of the option.
        :type opt: any
        :return: self
        """
        self._options[opt_name] = opt
        return self

    def table(
        self, name: str, data: list[list], html_generator: type = BasicTableGenerator
    ):
        """
        Set up data tables.

        :param name: The name of the table.
        :type name: str
        :param data: The data to be displayed in the table.
        :type data: list[list]
        :param html_generator: The HTML generator class to be used.
        :type html_generator: type
        :return: self
        """
        writer = TBL_HTMLWriter(name, data, html_generator)
        self._html_content = writer.contents
        return self

    def write(self, file: str):
        """
        Write the table to a PDF file.

        :param file: The path to save the PDF file.
        :type file: str
        """
        logger = proglog.default_bar_logger("bar")
        logger(message=f'PyAutoExcel - Writing "{file}"')
        pdfkit.from_string(self._html_content, file, options=self._options.copy())
        logger(message="PyAutoExcel - Done.")
