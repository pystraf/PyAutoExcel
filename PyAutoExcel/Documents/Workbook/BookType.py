"""
Excel workbook implement.

History (most recent first):
3.0.1  2024/2/7     Add docstrings.
2.0.1  2023/3/11    Created.

"""
# Importing modules

# self
from .BookBase import BaseWorkbook
from .BookImpl.Binary import EngineXLS
from .BookImpl.XML import EngineXLSX


# define WorkbookXLS
class WorkbookXLS(BaseWorkbook):
    """
    WorkbookXLS class.
    Support Excel 97-2003 Document.
    """

    engine = EngineXLS


# define WorkbookXLSX
class WorkbookXLSX(BaseWorkbook):
    """
    WorkbookXLSX class.
    Support Excel 2007+ Document.
    """

    engine = EngineXLSX
