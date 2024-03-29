"""
An operation toolkit of MS Excel files.
Developed by pystraf (pystraf@163.com)
"""

import sys

# Check Python Version
if sys.version_info < (3, 9, 0):
    raise RuntimeError("Sorry, Python 3.9.0 or later required.")

from xlrd import inspect_format


# Migrate Bridge
from .Bridge import migrate_style

# Newer I/O Port
from .Documents.File.Excel.ExcelDocument import Document as ExcelDocument
from .Documents.File.Excel.Reader.Excel import ExcelReader  # reader
from .Documents.File.Excel.Reader.Excel import (
    add_reader,
    install_builtin_readers,
    remove_reader,
)
from .Documents.File.Excel.Sheet import Sheet  # worksheet API
from .Documents.File.Excel.Writer.Excel import ExcelWriter  # writer
from .Documents.File.Excel.Writer.Excel import (
    add_writer,
    install_builtin_writers,
    remove_writer,
)

# Workbook API
from .Documents.Workbook.BookType import WorkbookXLS, WorkbookXLSX
from .ExtractVBA import extract_vba_project

# HTML Exporter
from .HTMLFile import HTMLSheet, save_html

# HTML Table Generator
from .TableGenerator import (
    BasicTableGenerator as BasicHTMLTable,
    CustomTableGenerator as HTMLTable,
)

# XF Style API
from .XFStyles import (
    XFAlignment,
    XFAlignmentConst,
    XFBorders,
    XFBordersConst,
    XFFont,
    XFFontConst,
    XFPattern,
    XFPatternConst,
    XFProtection,
    XFStyle,
)

from PyAutoExcel.Cell import Cell
from PyAutoExcel.CellRange import CellRange

__version__ = "3.0.2"

install_builtin_readers()
install_builtin_writers()

__all__ = [
    "inspect_format",
    "ExcelDocument",
    "add_reader",
    "remove_reader",
    "add_writer",
    "remove_writer",
    "Sheet",
    "ExcelReader",
    "ExcelWriter",
    "WorkbookXLS",
    "WorkbookXLSX",
    "extract_vba_project",
    "HTMLSheet",
    "save_html",
    "BasicHTMLTable",
    "HTMLTable",
    "XFAlignment",
    "XFAlignmentConst",
    "XFBorders",
    "XFBordersConst",
    "XFFont",
    "XFFontConst",
    "XFPattern",
    "XFPatternConst",
    "XFProtection",
    "XFStyle",
    "Cell",
    "CellRange",
    "__version__",
]
