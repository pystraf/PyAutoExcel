"""
Write table into pictures.
Need 'pdfkit' and 'wkhtmltox' to run.

:raises ModuleNotFoundError: If 'imgkit' package or 'wkhtmltox' is not found.
"""
import shutil

import proglog

from ..TableGenerator import BasicTableGenerator
from .HtmlWriter import TBL_HTMLWriter

try:
    import imgkit
except:
    raise ModuleNotFoundError(
        "PyAutoExcel.Tiny.ImageWriter requires 'imgkit' package but not found."
        "Please install it and try again."
    )

WK_PATH = shutil.which("wkhtmltopdf")

if not WK_PATH:
    raise ModuleNotFoundError(
        "PyAutoExcel.Tiny.ImageWriter requires 'wkhtmltox' but not found."
        "Please checkout https://wkhtmltopdf.org/downloads.html."
    )

config = imgkit.config(wkhtmltoimage=WK_PATH)


class IMGKIT_ImageWriter:
    """
    Image writer class using imgkit to generate pictures.

    :ivar _options: Dictionary containing image generation options.
    :ivar _html_content: HTML content to be converted to pictures.
    """

    def __init__(self):
        self._options = {"quiet": ""}
        self._html_content = ""

    def image_option(self, opt_name: str, opt):
        """
        Set a image generation option.

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
        writer = TBL_HTMLWriter(name, data)
        self._html_content = writer.contents
        return self

    def write(self, file: str):
        """
        Write the table to an image file.

        :param file: The path to save the image file.
        :type file: str
        """
        logger = proglog.default_bar_logger("bar")
        logger(message=f'PyAutoExcel - Writing "{file}"')
        imgkit.from_string(self._html_content, file, options=self._options.copy())
        logger(message="PyAutoExcel - Done.")
