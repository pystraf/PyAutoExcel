import io

import openpyxl
import xlsxwriter
import xlwt
from xlsxcessive import workbook, xlsx
from xlsxlite.writer import XLSXBook

from ..BaseWriter import ListGridWriteBook, WriteBook


class XlsxLiteWriter(ListGridWriteBook):
    __engine__ = "xlsxlite"

    def save_file(self, file_name: str):
        wb = XLSXBook()
        for s in self.sheets:
            ws = wb.add_sheet(name=s.name)
            for row in s.get_grid().get():
                ws.append_row(*row)
        wb.finalize(to_file=file_name, remove_dir=True)


class XlsxCessiveWriter(WriteBook):
    __engine__ = "xlsxcessive"

    def save_file(self, file_name: str):
        wb = workbook.Workbook()
        for s in self.sheets:
            ws = wb.new_sheet(name=s.name)
            for d in s.records:
                row, col, value = d
                ws.cell(coords=(row, col), value=value)
        if isinstance(file_name, io.BytesIO):
            xlsx.save(workbook=wb, filename="", stream=file_name)
        else:
            xlsx.save(workbook=wb, filename=file_name)

    save_io = save_file


class XlwtWriter(WriteBook):
    __engine__ = "xlwt"

    def save_file(self, file_name: str):
        wb = xlwt.Workbook(encoding="utf-8")
        for s in self.sheets:
            ws: xlwt.Worksheet = wb.add_sheet(sheetname=s.name, cell_overwrite_ok=True)
            for d in s.records:
                row, col, value = d
                ws.write(r=row, c=col, label=value)
        wb.save(filename_or_stream=file_name)

    save_io = save_file


class OpenpyxlWriter(ListGridWriteBook):
    __engine__ = "openpyxl"

    def save_file(self, file_name: str):
        wb = openpyxl.Workbook(write_only=True)
        for s in self.sheets:
            ws = wb.create_sheet(title=s.name)
            for row in s.get_grid().get():
                ws.append(row=row)
        wb.save(filename=file_name)

    save_io = save_file


class XlsxWriter(WriteBook):
    __engine__ = "xlsxwriter"

    def save_file(self, file_name: str):
        wb = xlsxwriter.Workbook(
            filename=file_name, options=dict(allowZip64=True, in_memory=True)
        )
        for s in self.sheets:
            ws = wb.add_worksheet(name=s.name)
            for d in s.records:
                row, col, value = d
                ws.write(row, col, value)
        wb.close()

    save_io = save_file
