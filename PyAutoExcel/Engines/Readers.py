import sxl
import xlrd
import xlsxio
from openpyxl import load_workbook

from ..BaseReader import ReadBook


class OpenpyxlReader(ReadBook):
    __engine__ = "openpyxl"

    def _parse(self):
        wb = load_workbook(self.file_name, read_only=True)
        for ws in wb.worksheets:
            datas = []
            for row in range(ws.max_row):
                tmp = []
                for col in range(ws.max_column):
                    tmp.append(ws.cell(row=row + 1, column=col + 1).value)
                datas.append(tmp.copy())
            self._datas[ws.title] = datas.copy()


class XlrdReader(ReadBook):
    __engine__ = "xlrd"

    def _parse(self):
        wb = xlrd.open_workbook(filename=self.file_name)
        for s in wb.sheets():
            s: xlrd.sheet.Sheet
            self._datas[s.name] = [
                s.row_values(rowx=ridx).copy() for ridx in range(s.nrows)
            ]


class XlsxioReader(ReadBook):
    __engine__ = "python-xlsxio"

    def _parse(self):
        wb = xlsxio.XlsxioReader(filename=self.file_name)
        sheet_nams = wb.get_sheet_names()
        for sn in sheet_nams:
            ws = wb.get_sheet(sheetname=sn)
            self._datas[sn] = ws.read_data().copy()
            ws.close()
        wb.close()


class SxlReader(ReadBook):
    __engine__ = "sxl"

    def _parse(self):
        wb = sxl.Workbook(file_obj=self.file_name)
        ws_count = len(wb.sheets)
        for i in range(ws_count):
            ws = wb.sheets[i]
            self._datas[ws.name] = list(ws.rows)
