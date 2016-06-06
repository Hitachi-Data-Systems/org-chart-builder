#!/usr/bin/python
import os
import sys
import datetime
import dateutil.parser

from people_filter_criteria import ProductCriteria, FunctionalGroupCriteria, IsInternCriteria, IsExpatCriteria, \
    FeatureTeamCriteria, IsCrossFuncCriteria, ManagerCriteria, IsTBHCriteria, LocationCriteria, IsManagerCriteria, \
    IsProductManagerCriteria

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
    VENDOR_TYPE = "Vendor"
    INTERN_TYPE = "Intern"
    LOCATION = "Location"
    START_DATE = "Start Date"

    CROSS_FUNCTIONS = ["admin", "inf", "infrastructure"]
    CROSS_FUNCT_TEAM = "Cross"
    FLOORS = {}
    TEAM_MODEL = {}
    PRODUCT_SORT_ORDER = []

class PeopleDataKeysBellevue(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)

    CROSS_FUNCTIONS = ["technology", "admin", "inf", "infrastructure", "cross functional"]

class PeopleDataKeysSantaClara(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)

    LEVEL = "Title"
    TEAM_MODEL = {
    "UCP" : "1 Tracks @ (1 PO, 1 TA,  4 Dev, 1 QA, 2 Char, 2 Auto)",
    "HID" : "2 Tracks @ (1 PO, 5 Dev, 2 QA, 2 Auto, 1 UX)",
    "HVS" : "Q1:20; Q2:25; Q3:27; Q4:32 -- 1 Tracks @ (1 PO, 5 Dev, 1 QA, 1 Auto)",
    "Evidence Management" : "1 Tracks @ (1 PO, 4 Dev, 1 QA, 1 Auto)",
    "HCmD" : "1 Tracks @ (1 Head Coach, 2 PO, 2 Dev, 1 QA, 1 UX)",

    }


class PeopleDataKeysWaltham(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    FUNCTION = "Function"
    NAME = "HR Name"
    NICK_NAME = "Name"
    CROSS_FUNCTIONS = ["Technology", "DevOps", "Admin", "Seal" ]
    FLOORS = {
        "- Mobility": [
            "Anderson, Vic", 
            "Kostadinov, Alex",
            "Lin, Wayzen",
            "Pfahl, Matt",
            "Van Thong, Adrien",
        ],

        "- Content Part 1": [
            "Bronner, Mark",
            "Burnham, John",
            "Chestna, Wayne",
            "Lin, Wayzen",
        ],

        "- Content Part 2": [
            "Hartford, Joe",
            "Pannese, Donald",
            "Pinkney, Dave",
        ],

        "- Ensemble": [
            "Isherwood, Ben",
            "Liang, Candy",
        ]
    }

    TEAM_MODEL = {
        "Aspen" : "1 Track @ (1 PO, 4 Dev, 2 QA, 1 Char, 2 Auto)",
        "Ensemble" : "2 Tracks @ (1 PO, 4 Dev, 1 QA, 1 Char, 1 Auto)",
        "HCP" : "4 Tracks @ (1 PO, 4 Dev, 2 QA, 1 Char, 2 Auto)",
        "HCP (Rhino)" : "1 Tracks @ (1 PO, 4 Dev, 2 QA, 1 Char, 2 Auto)",
        "HCP-AW" : "3 Tracks @ (1 PO, 4 Dev, 2 QA, 1 Char, 2 Auto)",
        }

class PeopleDataKeysSIBU(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    NAME = "HR Name"
    NICK_NAME = NAME
    REQ = "Requisition"
    TEAM_MODEL = {
            "HVS" : "[Forecast: Q1:20; Q2:25; Q3:28; Q4:32] -- 1 Tracks @ (1 PO, 5 Dev, 1 QA, 1 Auto)",
            "HVS EM" : "2 Tracks @ (1 PO, 4 Dev, 1 QA, 1 Char, 1 Auto)",
            "Lumada" : "[Forecast: Q1:7; Q2:10; Q3:43; Q4:110]",
            "City Data Exchange" : "[Forecast: Q1:5; Q2:20; Q3:25; Q4:31]",
            "Predictive Maintenance" : "[Forecast: Q1:5; Q2:17; Q3:22; Q4:27]",
            "Optimized Factory" : "[Forecast: Q1:1; Q2:11; Q3:13; Q4:15]",
        }

    PRODUCT_SORT_ORDER = ["hvs", "hvs em", "lumada", "city data exchange", "cde", "optimized factory",
                          "opf", "predictive maintenance", "pdm"]

class PeopleDataKeysHPP(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    FUNCTION = "Function"
    NAME = "HR Name"
    NICK_NAME = "Name"
    #CROSS_FUNCTIONS = ["Technology", "DevOps", "Admin", "Seal" ]

    def __init__(self, useActualFunction):
        PeopleDataKeys.__init__(self, useActualFunction)

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
            #if Nickname is blank but raw name is populated - use that
            aName = self.getRawNickName() or self.getRawName()

        if "," in aName:
            return " ".join(aName.split(",")[1:]).strip()
        return aName.split(" ")[0].strip()

    def getLastName(self, aName=None):
        if not aName:
            #if Nickname is blank but raw name is populated - use that
            aName = self.getRawNickName() or self.getRawName()

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

    def isVendor(self):
        typeStr = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.TYPE) or ""
        return typeStr.lower() == self.peopleDataKeys.VENDOR_TYPE.lower()

    def isIntern(self):
        typeStr = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.TYPE) or ""
        return typeStr.lower() == self.peopleDataKeys.INTERN_TYPE.lower()

    def isLead(self):
        return self.spreadsheetParser.getColByName(self.aRow, self.peopleDataKeys.NAME).style.font.bold

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

    def isProductManager(self):
        return self.getFunction().lower() in ["pm", "product manager", "product management"]

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

    def getStartDate(self):
        """


        :return: DateTime Object. Return empty datetime object if date is not set
        """
        if not self.spreadsheetParser.columnExists(self.peopleDataKeys.START_DATE):
            return ""

        startDateStr = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.START_DATE).strip()

        if startDateStr:
            try:
                return dateutil.parser.parse(startDateStr)
            except ValueError:
                print "Warning: can not parse start date for {}: '{}'".format(self.getFullName(), startDateStr)

        return datetime.datetime.min

    def __str__(self):
        personStr = "Person: {}, Product: {}, Location: {}".format(self.getFullName(), self.getProduct(), self.getLocation())
        if self.isTBH():
            personStr = "{} Req:{}".format(personStr, self.getReqNumber())
        personStr = "{} Row: {}".format(personStr, self.aRow[00].row)
        return personStr

    def __repr__(self):
        return self.__str__()

    def __lt__(self, other):

        if self.isUnfunded() and not other.isUnfunded():
            return False
        elif not self.isUnfunded() and other.isUnfunded():
            return True

        # # Uncomment if we want to sort interns to the bottom of each list...currently, we put interns on own slide
        # if self.isIntern() and not other.isIntern():
        #     return False
        # elif not self.isIntern() and other.isIntern():
        #     return True

        if self.isTBH() and not other.isTBH():
            return False
        elif not self.isTBH() and other.isTBH():
            return True

        return self.getFullName() < other.getFullName()

    def __gt__(self, other):
        return not self.__lt__(other)

    def __eq__(self, other):
        #return self.getFullName() == other.getFullName()
        return self.aRow[00].row == other.aRow[00].aRow

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash(self.aRow[00].row)

class OrgParser:
    def __init__(self, workbookName, dataSheetName, ):
        """

        :type workbookName: str
        :type dataSheetName: str
        """
        self.peopleDataKeys = PeopleDataKeys()
        self.orgName = os.path.basename(workbookName.split("Staff")[0].strip())

        if "waltham" in workbookName.lower():
            self.peopleDataKeys = PeopleDataKeysWaltham()

        if "hpp" in workbookName.lower():
            self.peopleDataKeys = PeopleDataKeysHPP(useActualFunction)

        if "bellevue" in workbookName.lower():
            self.peopleDataKeys = PeopleDataKeysBellevue()

        if "clara" in workbookName.lower():
            self.peopleDataKeys = PeopleDataKeysSantaClara()

        if "sibu" in workbookName.lower():
            self.peopleDataKeys = PeopleDataKeysSIBU()

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
        people = self.getFilteredPeople(PeopleFilter().addProductFilter(productName))
        for aPerson in people:
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

    def addIsProductManagerFilter(self, isPM=True):
        self.filterList.append(IsProductManagerCriteria(isPM))
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

