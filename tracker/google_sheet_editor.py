import logging
from datetime import date
from pygsheets.cell import Cell
from pygsheets.worksheet import Worksheet
from tracker.google_sheet_client import GoogleSheetClient


class GoogleSheetEditor:
    EXPENSE_COLUMN = "C"
    TYPE_COLUMN = "F"

    def __init__(self, spreadsheet: str, client: GoogleSheetClient):
        self.logger = logging.getLogger(__name__)
        self.spreadsheet = spreadsheet
        self.client = client

    @staticmethod
    def get_worksheet_name(date) -> str:
        return date.strftime('%b %y').lower()

    def open_worksheet(self, worksheet_title) -> Worksheet:
        return self.client.open(self.spreadsheet).worksheet_by_title(worksheet_title)

    def add_expense(self, worksheet: Worksheet, expense):
        cell = self.find_cell_by_date(worksheet, expense.spent_at)
        print(cell)
        row = cell.row
        expense_values = expense.to_values()
        range_to_edit = self.cell_range(
            self.EXPENSE_COLUMN + str(row),
            self.end_column(expense_values) + str(row)
        )
        print(range_to_edit)
        cell_list = worksheet.range(range_to_edit)[0]
        print(cell_list)
        self.logger.debug("Processing values: {}".format(expense))
        if self.is_row_empty(cell_list):
            worksheet.update_values(self.EXPENSE_COLUMN + str(row), [expense_values])
        else:
            expense_values.insert(0, "")
            expense_values.insert(0, "")
            worksheet.insert_rows(row, number=1, values=expense_values)
        self.logger.debug("Added value to row with {} cell".format(expense.spent_at))

    def end_column(self, expense_values):
        return chr(ord(self.EXPENSE_COLUMN) + len(expense_values) - 1)

    def find_cell_by_date(self, worksheet: Worksheet, cell_date: date) -> Cell:
        fm_date = self.formated_date(cell_date)
        print(fm_date)
        return worksheet.find(fm_date)[0]

    def get_cells(self, expense_date: date, worksheet: Worksheet) -> list:
        cell = self.find_cell_by_date(worksheet, expense_date)
        cell_matrix = worksheet.get_values(
            start='A2',
            end=self.TYPE_COLUMN + str(cell.row),
            include_tailing_empty=False,
            include_tailing_empty_rows=False,
        )
        return cell_matrix

    @staticmethod
    def formated_date(expense_date: date):
        day = expense_date.day
        month = expense_date.month
        day = str(day)
        month = str(month)
        if len(day) == 1:
            day = '0' + day
        if len(month) == 1:
            month = '0' + month
        return f'{day}-{month}-{str(expense_date.year)}'

    @staticmethod
    def cell_range(start: str, end: str):
        return start + ":" + end

    @staticmethod
    def is_row_empty(cell_list):
        for cell in cell_list:
            if cell.value != '':
                return False
        return True
