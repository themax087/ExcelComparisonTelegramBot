from functools import lru_cache

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf


class XlsFile:
    def __init__(self, filename):
        self.filename = filename
        self.book = open_workbook(filename, formatting_info=True)
        self.sheets = {}

    @property
    @lru_cache
    def sheets_names(self):
        return set(self.book.sheet_names())

    @lru_cache
    def sheet_by_name(self, name):
        return self.book.sheet_by_name(name)

    @lru_cache
    def rows(self, name):
        return self.sheet_by_name(name).nrows

    @lru_cache
    def cols(self, name):
        return self.sheet_by_name(name).ncols

    @lru_cache
    def cell_value(self, name, row, col):
        sheet = self.sheet_by_name(name)
        return sheet.cell_value(rowx=row, colx=col)

    def copy_with_highlights(self, path, highlights_by_sheet):
        style = easyxf('pattern: pattern solid, fore_colour red')
        new_wb = copy(self.book)
        for name, highlights in highlights_by_sheet.items():
            for row, col in highlights:
                sheet = new_wb.get_sheet(name)
                sheet.write(row, col, self.cell_value(name, row, col), style)
        new_wb.save(path)


class XlsxFile:
    def __init__(self, filename):
        self.filename = filename
        self.book = load_workbook(filename)
        self.sheets = {}

    @property
    @lru_cache
    def sheets_names(self):
        return set(self.book.sheetnames)

    @lru_cache
    def sheet_by_name(self, name):
        return self.book[name]

    @lru_cache
    def rows(self, name):
        return self.sheet_by_name(name).max_row

    @lru_cache
    def cols(self, name):
        return self.sheet_by_name(name).max_column

    @lru_cache
    def cell_value(self, name, row, col):
        sheet = self.sheet_by_name(name)
        value = sheet.cell(row + 1, col + 1).value
        if value:
            return value
        return ''

    def copy_with_highlights(self, path, highlights_by_sheet):
        red_fill = PatternFill(start_color="FF0000", fill_type="solid")
        for name, highlights in highlights_by_sheet.items():
            for row, col in highlights:
                self.book[name].cell(row + 1, col + 1).fill = red_fill
        self.book.save(path)


def open_excel_file(filename: str):
    if filename.endswith('.xls'):
        return XlsFile(filename)
    elif filename.endswith('.xlsx'):
        return XlsxFile(filename)
    else:
        raise NotImplementedError


def find_diff(current, original):
    highlights_by_sheet = {}
    sheets_names = current.sheets_names.intersection(original.sheets_names)
    for name in sheets_names:
        sheet_highlights = set()

        current_rows = current.rows(name)
        original_rows = original.rows(name)

        current_cols = current.cols(name)
        original_cols = original.cols(name)

        if current_rows > original_rows:
            for row in range(original_rows, current_rows):
                for col in range(current_cols):
                    sheet_highlights.add((row, col))

        if current_cols > original_cols:
            for col in range(original_cols, current_cols):
                for row in range(current_rows):
                    sheet_highlights.add((row, col))

        for row in range(min(current_rows, original_rows)):
            for col in range(min(current_cols, original_cols)):
                if current.cell_value(name, row, col) != original.cell_value(name, row, col):
                    sheet_highlights.add((row, col))

        highlights_by_sheet[name] = sheet_highlights

    return highlights_by_sheet
