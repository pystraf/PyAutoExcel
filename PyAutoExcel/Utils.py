"""
Helper functions for PyAutoExcel.
"""
from collections import defaultdict
from typing import Union

import pandas as pd
import xlwt
from openpyxl.utils.cell import get_column_letter

inited = False  # Flag to check if the colorama has been initialized.


def column_dict_to_list(cdict: dict) -> list[list]:
    """
    Convert a dictionary of column names to column data into a list of lists.

    :param cdict: The dictionary to convert.
    :return: An 2D list of the data.
    """
    header = list(cdict.keys())
    body = list(cdict.values())
    data = list(zip(*body))
    data = list(map(list, data))
    data.insert(0, header)
    return data.copy()


def records_to_list(rlist: list[dict]) -> list[list]:
    """
    Convert a list of dictionaries into a list of lists.

    :param rlist: The list of dictionaries to convert.
    :return: An 2D list of the data.
    """
    ret = []
    header = list(rlist[0].keys())
    ret.append(header)
    for record in rlist:
        row = list(record.values())
        ret.append(row)
    return ret.copy()


# pos2string(0, 0) -> 'A1'
def pos2string(row: int, col: int) -> str:
    """
    Convert a row and column position to an Excel-style cell reference string.

    :param row: The row number. (0-based)
    :param col: The column number. (0-based)
    :return: The cell reference string. (e.g. 'A1')
    """
    return get_column_letter(col + 1) + str(row + 1)


def neither_nor(a: str, b: str):
    raise Exception("neither %s nor %s found" % (a, b))


def AutoActiveDict():
    """
    This function returns an empty defaultdict object
        that can be used to automatically create keys when they are accessed.

    :return: An empty defaultdict object that can be used to automatically
             create keys when they are accessed.
    """
    return defaultdict(AutoActiveDict)


def humanize_items(
    items: Union[list[str], tuple[str, ...], str], using_or: bool = False
) -> str:
    """
    It takes a list of strings and returns a string that is a human-readable version of the list.

    :param items: The list of items to humanize.
    :param using_or: If True, the last separator will be 'or' instead of 'and', defaults to False.
    :return: A string.
    """
    last_seg = "or" if using_or else "and"
    if isinstance(items, str):
        return items
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} {last_seg} {items[1]}"
    exclude_last = items[:-1]
    exclude_last_string = ", ".join(exclude_last)
    last = items[-1]
    return f"{exclude_last_string} {last_seg} {last}"


def object_class_name(obj: object) -> str:
    """
    Return the name of the class of the given object.
    :param obj: object
    :return: The name of the class of the object.
    """
    return obj.__class__.__name__


def objects_class_name(objs: list[object]) -> list[str]:
    """
    returns a list of the class names of the objects in the list `objs`

    :param objs: list of objects
    :return: A list of strings.
    """
    return [object_class_name(obj=i) for i in objs]


def get_cls_name(objs: Union[object, list]) -> Union[list[str], str]:
    """
    returns the class name of an object or a list of objects

    :param objs: The object or list of objects to get the class name of.
    :return: A list of strings or a string.
    """
    if isinstance(objs, list):
        return objects_class_name(objs=objs).copy()
    return object_class_name(obj=objs)


class FinalMeta(type):
    """
    A metaclass that prevents a class from being inherited.
    """

    def __new__(mcs, name, bases, dict):
        for base in bases:
            if isinstance(base, FinalMeta):
                raise TypeError(
                    "type '{0}' is not an acceptable base type".format(base.__name__)
                )
        cls = super().__new__(mcs, name, bases, dict)
        return cls


def init_colorama():
    """
    Initialize the colorama library to enable ANSI color codes translation on Windows.

    :return: None
    """
    global inited
    if not inited:
        import colorama

        colorama.init(autoreset=True)
        inited = True


def to_excel_autowidth_and_border(
    writer: pd.ExcelWriter,
    df: pd.DataFrame,
    sheetname: str,
    startrow: int,
    startcol: int,
):
    """
    Write a DataFrame to an Excel sheet and adjust columns' width based on their content.
    Additionally, apply a border to all cells.

    :param writer: The ExcelWriter object to write the DataFrame into.
    :type writer: pd.ExcelWriter
    :param df: The DataFrame to write to the Excel sheet.
    :type df: pd.DataFrame
    :param sheetname: The name of the sheet where the DataFrame will be written.
    :type sheetname: str
    :param startrow: The starting row index to write the DataFrame.
    :type startrow: int
    :param startcol: The starting column index to write the DataFrame.
    :type startcol: int

    This function writes the given DataFrame `df` to an Excel sheet named `sheetname` using the provided `writer`.
    It automatically adjusts the width of each column based on the maximum length of its content.
    Additionally, it applies a border to all cells in the DataFrame range.

    :return: Nothing.
    """
    df.to_excel(
        writer,
        sheet_name=sheetname,
        index=False,
        startrow=startrow,
        startcol=startcol,
    )  # send df to writer
    workbook = writer.book
    worksheet = writer.sheets[sheetname]  # pull worksheet object
    formater = workbook.add_format({"border": 1})
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = (
            max(
                (
                    series.astype(str).map(len).max(),  # len of largest item
                    len(str(series.name)),  # len of column name/header
                )
            )
            * 3
            + 1
        )  # adding a little extra space
        # print(max_len)
        worksheet.set_column(
            idx + startcol, idx + startcol, max_len
        )  # set column width
    first_row = startrow
    first_col = startcol
    last_row = startrow + len(df.index)
    last_col = startcol + len(df.columns)
    worksheet.conditional_format(
        first_row,
        first_col,
        last_row,
        last_col - 1,
        options={"type": "formula", "criteria": "True", "format": formater},
    )


def make_df_from_list(data: list[list]) -> pd.DataFrame:
    """
    Create a pandas DataFrame from a list of lists.

    :param data: The input data in the form of a list of lists,
                 where the first list contains column names and
                 the subsequent lists contain the data.
    :type data: list[list]

    :return: A pandas DataFrame created from the input data.
    :rtype: pd.DataFrame
    """
    return pd.DataFrame(data[1:], columns=data[0])


def set_out_cell(outSheet: xlwt.Worksheet, row: int, col: int, value):
    """
    Change cell value without changing formatting.

    :param outSheet: The xlwt.Worksheet object representing the sheet to write to.
    :type outSheet: xlwt.Worksheet
    :param row: The row index of the cell to write to.
    :type row: int
    :param col: The column index of the cell to write to.
    :type col: int
    :param value: The value to write into the cell.
    :type value: Any
    """

    def _getOutCell(outSheet: xlwt.Worksheet, rowIndex: int, colIndex: int):
        """HACK: Extract the internal xlwt cell representation."""
        row = outSheet._Worksheet__rows.get(rowIndex)
        if not row:
            return None

        cell = row._Row__cells.get(colIndex)
        return cell

    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, row, col)
    # END HACK, PART I

    outSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, row, col)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx
