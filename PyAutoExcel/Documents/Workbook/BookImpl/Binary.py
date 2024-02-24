"""
.xls workbook's engine class.

History (most recent first):
2.0.1  2023/3/11    Created.

"""

# Importing modules

# bulitins
import io
import os

# third party
from xlrd import open_workbook
from xlrd.sheet import Sheet as RDSheet
from xlwt import Row, Workbook, Worksheet

# self
from ...File.Excel.Sheet import Sheet
from ..Engine import EngineBase, ReaderStream, WriterStream


# Define EngineXLS class
class EngineXLS(EngineBase):
    @classmethod
    def save(cls, file: WriterStream, sheets: list[Sheet]):
        # Write to workbook
        workbook = Workbook(encoding="utf-8")
        for s in sheets:
            # Create a worksheet
            worksheet: Worksheet = workbook.add_sheet(s.name)
            # Writing data to the worksheet
            for rindex, rows in enumerate(s.data):
                # Get a row
                row: Row = worksheet.row(rindex)

                # Writing cells
                for cindex, data in enumerate(rows):
                    row.write(cindex, data)
        if file is not None:
            workbook.save(file)
        else:
            stream = io.BytesIO()
            workbook.save(stream)
            return stream.getvalue()

    @classmethod
    def read(cls, file: ReaderStream) -> list[Sheet]:
        # Disable log of xlrd.
        null_device = open(os.devnull, mode="w")

        # Open workbook
        if isinstance(file, io.IOBase):
            file.seek(0)
            content = file.read()
            workbook = open_workbook(None, logfile=null_device, file_contents=content)
        elif isinstance(file, bytes):
            workbook = open_workbook(None, logfile=null_device, file_contents=file)
        else:
            workbook = open_workbook(file, logfile=null_device)

        # Reading sheets.
        sheets = []
        for s in workbook.sheets():
            s: RDSheet
            target = Sheet(s.name)
            rows = s.get_rows()
            # Writing to target.
            for num, row in enumerate(rows):
                target.set_row(num, row)
            sheets.append(target)

        return sheets.copy()
