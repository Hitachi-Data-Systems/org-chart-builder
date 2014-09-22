from unittest import TestCase
from spreadsheet_parser import SpreadsheetParser, ColumnDoesNotExist

__author__ = 'doreper'


class TestSpreadsheetParser(TestCase):
    def setUp(self):
        self.spreadsheetParser = SpreadsheetParser("./testSpreadsheet.xlsx", "TestSheet")

    def test_GetCellValue(self):
        self.assertEqual("a", self.spreadsheetParser.getColValueByName(next(self.spreadsheetParser.dataRows()), "col1"))

    def test_GetEmptyCellValue(self):
        self.assertEqual("", self.spreadsheetParser.getColValueByName(next(self.spreadsheetParser.dataRows()), "col2"))

    def test_GetNonExistentColumn(self):
        with self.assertRaises(ColumnDoesNotExist):
            self.assertEqual("", self.spreadsheetParser.getColValueByName(next(self.spreadsheetParser.dataRows()), "colz"))

    def test_GetColumnThatHasSpace(self):
        self.assertEqual("d", self.spreadsheetParser.getColValueByName(list(self.spreadsheetParser.dataRows())[1], "col3"))

    def test_GetIntCellValue(self):
        self.assertEqual("5", self.spreadsheetParser.getColValueByName(list(self.spreadsheetParser.dataRows())[1], "col 4"))