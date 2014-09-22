from openpyxl import load_workbook

__author__ = 'doreper'

class ColumnDoesNotExist(Exception):
    def __init__(self, colName, availableColumns):
        self.colName = colName
        self.availableColumns = availableColumns

    def __str__(self):
        return "Column '{}' does not exist. Available columns: {}".format(self.colName, self.availableColumns)

class SpreadsheetParser:
    def __init__(self, workbookName, dataSheetName):
        """

        :type workbookName: str
        :type dataSheetName: str
        """
        wb = load_workbook(filename=workbookName)
        self.sheet_ranges = wb[dataSheetName]
        self.spreadsheetColumns = ColumnHeaders(self.sheet_ranges.rows[0])

    def getColValueByName(self, aRow, colName):
        """

        :type aRow: tuple
        :type colName: str
        :return:
        """
        colEntry = self.spreadsheetColumns.getColumn(colName)
        if colEntry is None or aRow[colEntry].value is None:
            return ""

        return str(aRow[colEntry].value).strip()

    def columnExists(self, colName):
        return self.spreadsheetColumns.columnExists(colName)

    def dataRows(self):
        for aRow in self.sheet_ranges.rows[1:]:
            yield aRow

    def filteredDataRows(self, colValFilterDict):
        """ Get all the rows that match the criteria

        :param colValFilterDict: {Manager: Dave Oreper, Type: Consultant}
        :type colValFilterDict: {str: str}
        """
        for aRow in self.dataRows():
            isMatch = False
            for aKey, aVal in colValFilterDict.iteritems():
                if not self.getColValueByName(aRow, aKey) == aVal:
                    isMatch = False
                    break
                isMatch = True
            if isMatch:
                yield aRow

class ColumnHeaders:
    """
    This class takes a row and creates a dictionary of {colName: colIndex}. This mapping is used to
    extract cell data by column name instead of index
    """
    def __init__(self, topRow):
        """

        :param topRow:
        """
        self.colDict = {}

        colIndex = 0
        for aCol in topRow:
            colVal = aCol.value
            if colVal == "" or colVal == None:
                colIndex += 1
                continue
            self.colDict[colVal.strip()] = colIndex
            colIndex += 1

    def getColumn(self, colName):
        """

        :param colName:
        :return:
        """
        if not colName in self.colDict:
            raise ColumnDoesNotExist(colName, self.colDict.keys())
        return self.colDict.get(colName, None)

    def columnExists(self, colName):
        return colName in self.colDict