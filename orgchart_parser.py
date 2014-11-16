#!/usr/bin/python

from spreadsheet_parser import SpreadsheetParser


__author__ = 'doreper'

NOT_SET = "!!NOT SET!!"

class PeopleDataKeys:
    def __init__(self):
        pass

    MANAGER = "Manager"
    NAME = "Name"
    NICK_NAME = NAME
    LEVEL = "Level"
    FUNCTION = "Function"
    PROJECT = "Project"
    TYPE = "Type"
    REQ = "Requisition Number"
    CONSULTANT = "Consultant"
    EXPAT_TYPE = "Expat"
    INTERN_TYPE = "Intern"

    CROSS_FUNCTIONS = ["admin", "inf", "infrastructure"]
    CROSS_FUNCT_TEAM = "Cross"
    CONSULTANT_DECORATOR = "**"
    MANAGER_DECORATOR = "=="
    FLOORS = {}


class PeopleDataKeysBellevue(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)

    NAME = "HR Name"
    NICK_NAME = NAME


class PeopleDataKeysWaltham(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    NAME = "HR Name"
    NICK_NAME = "Name"
    CROSS_FUNCTIONS = ["Technology", "DevOps", "Admin"]
    FLOORS = { "Second Floor": ["Anderson, Vic", "Burnham, John", "Kostadinov, Alex", "Lin, Wayzen",
                                "Pfahl, Matt"],
               "Third Floor":  ["Chestna, Wayne", "Isherwood, Ben", "Kohli, Nishant", "Liang, Candy",
                                "Lin, Wayzen", "Pannese, Donald", "Pinkney, Dave"]
    }


class PersonRowWrapper:
    def __init__(self, spreadsheetParser, peopleDataKeys, aRow):
        self.spreadsheetParser = spreadsheetParser
        self.peopleDataKeys = peopleDataKeys
        self.aRow = aRow
        self.manager = False

    def isConsultant(self):
        """


        :return:
        """
        typeStr = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.TYPE) or ""
        return typeStr.lower() == self.peopleDataKeys.CONSULTANT.lower()

    def setManager(self):
        """


        """
        self.manager = True

    def isManager(self):
        """


        :return:
        """
        return self.manager

    def getReqNumber(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.REQ).split(".")[0]

    def getFirstName(self):
        fullName = self.getRawNickName()
        if "," in fullName:
            return " ".join(fullName.split(",")[1:])
        return fullName.split(" ")[0]

    def getLastName(self):
        fullName = self.getRawNickName()
        if "," in fullName:
            return fullName.split(",")[0]
        return " ".join(fullName.split(" ")[1:])

    def getFullName(self):
        return "{} {}".format(self.getFirstName(), self.getLastName())

    def getRawName(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.NAME)

    def getRawNickName(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.NICK_NAME)

    def isExpat(self):
        typeStr = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.TYPE) or ""
        return typeStr.lower() == self.peopleDataKeys.EXPAT_TYPE.lower()

    def isIntern(self):
        typeStr = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.TYPE) or ""
        return typeStr.lower() == self.peopleDataKeys.INTERN_TYPE.lower()

    def getTitle(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.LEVEL)

    def getFunction(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.FUNCTION)

    def getProduct(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.PROJECT)

    def __lt__(self, other):
        if self.isIntern() and not other.isIntern():
            return False;
        elif not self.isIntern() and other.isIntern():
            return True;

        if self.getFullName().startswith("TBH"):
            if other.getFullName().startswith("TBH"):
                return self.getFullName() < other.getFullName()
            return False

        if self.getFullName().startswith("TBD"):
            if other.getFullName().startswith("TBD"):
                return self.getFullName() < other.getFullName()
            return False

        return self.getFullName() < other.getFullName()

    def __gt__(self, other):
        return not self.__lt__(other)

    def __eq__(self, other):
        return self.getFullName() == other.getFullName()

    def __ne__(self, other):
        return not self.__eq__(other)


class OrgParser:
    def __init__(self, workbookName, dataSheetName):
        """

        :type workbookName: str
        :type dataSheetName: str
        """
        self.peopleDataKeys = PeopleDataKeys
        if "waltham" in workbookName.lower():
            self.peopleDataKeys = PeopleDataKeysWaltham

        self.spreadsheetParser = SpreadsheetParser(workbookName, dataSheetName)
        self.managerList = self.getManagerSet()

    def getManagerSet(self):
        """
        :return:
        """
        managerSet = set()
        for aRow in self.spreadsheetParser.dataRows():
            managerName = self.spreadsheetParser.getColValueByName(aRow, self.peopleDataKeys.MANAGER)
            managerSet.add(managerName)
        return managerSet

    def getPerson(self, aRow):
        aPerson = PersonRowWrapper(self.spreadsheetParser, self.peopleDataKeys, aRow)
        if aPerson.getRawName() in self.managerList or aPerson.getRawNickName() in self.managerList:
            aPerson.setManager()
        return aPerson

    def getDirectReports(self, managerName, productName=""):
        directReportList = []
        filterDict = {self.peopleDataKeys.MANAGER: managerName}
        if productName:
            filterDict[self.peopleDataKeys.PROJECT] = productName
        directReportRows = self.spreadsheetParser.filteredDataRows(filterDict)
        for aDirectReport in directReportRows:
            directReportList.append(self.getPerson(aDirectReport))
        return directReportList

    def getProductSet(self):
        """

        :return:
        """
        productList = set()
        for aRow in self.spreadsheetParser.dataRows():
            productName = self.spreadsheetParser.getColValueByName(aRow, self.peopleDataKeys.PROJECT)
            productList.add(productName)
        return productList

    def getFunctionSet(self, productName=None):
        functionSet = set()
        for aRow in self.spreadsheetParser.dataRows():
            functionName = self.spreadsheetParser.getColValueByName(aRow, self.peopleDataKeys.FUNCTION)
            if productName:
                if self.spreadsheetParser.getColValueByName(aRow, self.peopleDataKeys.PROJECT) != productName:
                    continue

            functionSet.add(functionName)
        return functionSet

    def getFilteredPeople(self, productName=None, functionName=None, isExpat=None):
        """ Get all the people that match the criteria

        :type productName: str
        :type functionName: str
        """
        matchingPeople = []
        filterDict = {}
        if productName is not None:
            filterDict[self.peopleDataKeys.PROJECT] = productName
        if functionName is not None:
            filterDict[self.peopleDataKeys.FUNCTION] = functionName
        if isExpat is not None:
            filterDict[self.peopleDataKeys.TYPE] = self.peopleDataKeys.EXPAT_TYPE

        for aRow in self.spreadsheetParser.filteredDataRows(filterDict):
            matchingPeople.append(self.getPerson(aRow))

        return matchingPeople
