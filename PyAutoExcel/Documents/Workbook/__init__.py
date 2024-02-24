"""
This package contains the book classes.


Classes:
    WorkbookXLS - XLS Workbook (Binary format)
    WorkbookXLSX - XLSX Workbook (OpenXML format)
    BaseWorkbook - Base class for Workbook classes

"""
from .BookBase import BaseWorkbook
from .BookType import WorkbookXLS, WorkbookXLSX
