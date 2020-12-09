# coding=utf-8
from openpyxl import load_workbook


class ExcelHandle:

    def __init__(self, file_name):
        self.file_name = file_name
        self.wb = load_workbook(file_name)

    def choose_sheet(self, sheet_name):
        if isinstance(sheet_name, int):
            return self.wb.worksheets[sheet_name]
        return self.wb.get_sheet_by_name(sheet_name)
