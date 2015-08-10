#!/usr/bin/python
import os
from people_filter_criteria import ProductCriteria, FunctionalGroupCriteria, IsInternCriteria, IsExpatCriteria, \
    FeatureTeamCriteria, IsCrossFuncCriteria, ManagerCriteria, IsTBHCriteria, LocationCriteria, IsManagerCriteria

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
    FEATURE_TEAM = "Feature Team"
    TYPE = "Type"
    REQ = "Requisition Number"
    CONSULTANT = "Consultant"
    CONTRACTOR = "Contractor"
    EXPAT_TYPE = "Expat"
    INTERN_TYPE = "Intern"
    LOCATION = "Location"

    CROSS_FUNCTIONS = ["admin", "inf", "infrastructure"]
    CROSS_FUNCT_TEAM = "Cross"
    FLOORS = {}


class PeopleDataKeysBellevue(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)

    CROSS_FUNCTIONS = ["technology", "admin", "inf", "infrastructure", "cross functional"]


class PeopleDataKeysWaltham(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    FUNCTION = "Function - Actual"
    NAME = "HR Name"
    NICK_NAME = "Name"
    CROSS_FUNCTIONS = ["Technology", "DevOps", "Admin"]
    FLOORS = {"Second Floor": ["Anderson, Vic", "Burnham, John", "Kostadinov, Alex", "Lin, Wayzen",
                               "Pfahl, Matt"],
              "Third Floor": ["Chestna, Wayne", "Isherwood, Ben", "Kohli, Nishant", "Liang, Candy",
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
        return (typeStr.lower() == self.peopleDataKeys.CONSULTANT.lower()) or (self.peopleDataKeys.CONTRACTOR.lower()
                                                                               in typeStr.lower())

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
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.REQ).split(".")[0].strip()

    def getFirstName(self, aName=None):
        if not aName:
            aName = self.getRawNickName()

        if "," in aName:
            return " ".join(aName.split(",")[1:]).strip()
        return aName.split(" ")[0].strip()

    def getLastName(self, aName=None):
        if not aName:
            aName = self.getRawNickName()

        if "," in aName:
            return aName.split(",")[0].strip()
        return " ".join(aName.split(" ")[1:]).strip()

    def getFullName(self, fullName=None):
        return "{} {}".format(self.getFirstName(fullName), self.getLastName(fullName)).strip()

    def getRawName(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.NAME).strip()

    def getNormalizedRawName(self):
        return "{} {}".format(self.getFirstName(self.getRawName()), self.getLastName(self.getRawName()))

    def getRawNickName(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.NICK_NAME).strip()

    def isExpat(self):
        typeStr = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.TYPE) or ""
        return typeStr.lower() == self.peopleDataKeys.EXPAT_TYPE.lower()

    def isIntern(self):
        typeStr = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.TYPE) or ""
        return typeStr.lower() == self.peopleDataKeys.INTERN_TYPE.lower()

    def isTBH(self):
        if (self.getFullName().lower().startswith("tbh")
            or self.getFullName().lower().startswith("tbd")):
            return True
        return False

    def isUnfunded(self):
        if (self.getFullName().lower().startswith("unfunded")
            or self.getFullName().lower().startswith("unfunded")):
            return True
        return False

    def isCrossFunc(self):
        return ((self.getFunction().lower() in self.peopleDataKeys.CROSS_FUNCTIONS)
                or (self.getProduct().lower() == self.peopleDataKeys.CROSS_FUNCT_TEAM.lower()))

    def getTitle(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.LEVEL).strip()

    def getFunction(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.FUNCTION).strip()

    def getFeatureTeam(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.FEATURE_TEAM).strip()

    def getManagerRawName(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.MANAGER).strip()

    def getManagerFullName(self):
        """ Return the manager name in the form [first] [last], even if it's listed as [last],[first]
        in source data
        """
        managerRawName = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.MANAGER)
        if not managerRawName:
            return ""
        return "{} {}".format(self.getFirstName(managerRawName), self.getLastName(managerRawName)).strip()

    def getProduct(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.PROJECT).strip()

    def getFloor(self):
        for aFloor, managerNames in self.peopleDataKeys.FLOORS.iteritems():
            for aManagerName in managerNames:
                if (self.getFullName() == self.getFullName(aManagerName)
                or (self.getRawName() == self.getFullName(aManagerName))
                or (self.getNormalizedRawName() == self.getFullName(aManagerName))):
                    return aFloor
        return ""

    def getLocation(self):
        if not self.spreadsheetParser.columnExists(self.peopleDataKeys.LOCATION):
            return ""
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.LOCATION).strip() or ""

    def __str__(self):
        return "Person: {}".format(self.getFullName())

    def __repr__(self):
        return self.__str__()

    def __lt__(self, other):

        if self.isUnfunded() and not other.isUnfunded():
            return False
        elif not self.isUnfunded() and other.isUnfunded():
            return True

        if self.isIntern() and not other.isIntern():
            return False
        elif not self.isIntern() and other.isIntern():
            return True

        if self.isTBH() and not other.isTBH():
            return False
        elif not self.isTBH() and other.isTBH():
            return True




        return self.getFullName() < other.getFullName()

    def __gt__(self, other):
        return not self.__lt__(other)

    def __eq__(self, other):
        return self.getFullName() == other.getFullName()

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash(self.getFullName())

class OrgParser:
    def __init__(self, workbookName, dataSheetName):
        """

        :type workbookName: str
        :type dataSheetName: str
        """
        self.peopleDataKeys = PeopleDataKeys
        self.orgName = os.path.basename(workbookName.split("Staff")[0].strip())

        if "waltham" in workbookName.lower():
            self.peopleDataKeys = PeopleDataKeysWaltham

        if "bellevue" in workbookName.lower():
            self.peopleDataKeys = PeopleDataKeysBellevue

        self.spreadsheetParser = SpreadsheetParser(workbookName, dataSheetName)
        self.managerList = self.getManagerSet()

    def getManagerSet(self):
        """
        :return:
        """
        managerSet = set()
        for aRow in self.spreadsheetParser.dataRows():
            managerName = self.spreadsheetParser.getColValueByName(aRow, self.peopleDataKeys.MANAGER)
            if managerName:
                managerSet.add(managerName)
        return managerSet

    def getPerson(self, aRow):
        aPerson = PersonRowWrapper(self.spreadsheetParser, self.peopleDataKeys, aRow)
        if (aPerson.getRawName() in self.managerList
            or aPerson.getRawNickName() in self.managerList
            or aPerson.getFullName() in self.managerList):
            aPerson.setManager()
        return aPerson

    def getProductSet(self):
        """

        :return:
        """
        productSet = set()
        for aPerson in self.getPeople():
            productSet.add(aPerson.getProduct())
        return productSet

    def getFeatureTeamSet(self, productName):
        featureSet = set()
        for aPerson in self.getFilteredPeople(PeopleFilter().addProductFilter(productName)):
            featureSet.add(aPerson.getFeatureTeam())
        return featureSet

    def getFunctionSet(self, productName=None):
        functionSet = set()
        for aPerson in self.getFilteredPeople(PeopleFilter().addProductFilter(productName)):
            functionSet.add(aPerson.getFunction())
        return functionSet

    def getLocationSet(self, productName=""):
        locationSet = set()
        filter = PeopleFilter()
        if productName:
            filter.addProductFilter(productName)

        for aPerson in self.getFilteredPeople():
            locationSet.add(aPerson.getLocation())
        return locationSet

    def getPeople(self):
        for aRow in self.spreadsheetParser.dataRows():
            aPerson = self.getPerson(aRow)
            yield aPerson

    def getFilteredPeople(self, peopleFilter=None):
        """ Get all the people that match the filter
        """
        if not peopleFilter:
            peopleFilter = PeopleFilter()

        matchingPeople = []

        for aPerson in self.getPeople():
            if (aPerson.getRawNickName() or aPerson.getRawName()) and peopleFilter.isMatch(aPerson):
                matchingPeople.append(aPerson)
        matchingPeople.sort()
        return matchingPeople


class PeopleFilter:
    def __init__(self):
        self.filterList = []


    def addManagerFilter(self, manager):
        self.filterList.append(ManagerCriteria(manager))
        return self

    def addProductFilter(self, productName):
        self.filterList.append(ProductCriteria(productName))
        return self

    def addFunctionFilter(self, functionName):
        self.filterList.append(FunctionalGroupCriteria(functionName))
        return self

    def addFeatureTeamFilter(self, featureTeam):
        self.filterList.append(FeatureTeamCriteria(featureTeam))
        return self

    def addIsManagerFilter(self, isManager=True):
        self.filterList.append(IsManagerCriteria(isManager))
        return self

    def addIsInternFilter(self, isIntern=True):
        self.filterList.append(IsInternCriteria(isIntern))
        return self

    def addIsCrossFuncFilter(self, isCrossFunc=True):
        self.filterList.append(IsCrossFuncCriteria(isCrossFunc))
        return self

    def addIsExpatFilter(self, isExpat=True):
        self.filterList.append(IsExpatCriteria(isExpat))
        return self

    def addIsTBHFilter(self, isTBH=True):
        self.filterList.append(IsTBHCriteria(isTBH))
        return self

    def addLocationFilter(self, location):
        self.filterList.append(LocationCriteria(location))
        return self

    def isMatch(self, aPerson):
        for criterion in self.filterList:
            if not criterion.matches(aPerson):
                return False
        return True

