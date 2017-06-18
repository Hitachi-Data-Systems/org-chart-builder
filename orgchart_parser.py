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
    FLOOR_SORT_ORDER = []

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
    CROSS_FUNCTIONS = ["Technology", "DevOps", "Admin", "Sustaining" ]
    FLOORS = {
        "- Mobility": [
            "Anderson, Vic",
            "Kostadinov, Alex",
            "Lin, Wayzen",
            "Manjanatha, Sowmya",
            "Maruca, Fran",
            "Pfahl, Matt",
            "Van Thong, Adrien",
        ],

        "- Content Part 1": [
            "Boba, Andrew",
            "Bronner, Mark",
            "Burnham, John",
            "Chestna, Wayne",
        ],

        "- Content Part 2": [
            "Lee, Jonathan",
            "Shea, Kevin",
        ],

        "- Aspen": [
            "Hartford, Joe",
            "Liang, Candy",
        ],
        "- HPP": [
            "Wesley, Joe",
            "Moore, Jim",
        ],
	"- HDID": [
	    "Agashe, Sujata",
            "Caswell, Paul",
            "Chappell, Simon",
            "Gothoskar, Chandrashekhar",
            "Helliker, Fabrice",
            "Mason, Bill",
            "Melville, Andrew",
            "Pendlebury, Ian",
            "Pfaff, Florian",
            "Sinkar, Milind",
	],
    }

    TEAM_MODEL = {
        "Aspen" : "4 Tracks @ (1 PO, 3 Dev, 1 QA, 1 Char, 1 Auto)",
        "HCP-Rhino" : "4 Tracks @ (1 PO, 4 Dev, 2 QA, 1 Char, 2 Auto)",
        "HCP-India" : "1 Track @ (1 PO, 3 Dev, 1 QA, 1 Auto)",
#        "HCP (Rhino)" : "1 Track @ (1 PO, 4 Dev, 2 QA, 2 Char, 2 Auto)",
        "HCP-AW" : "4 Tracks @ (1 PO, 4 Dev, 2 QA, 1 Char, 2 Auto)",
        }

    # names should be lower case here
    PRODUCT_SORT_ORDER = ["aspen", "ensemble", "hcp-rhino", "hcp-india", "hcp-aw", "aw-japan","hpp", "hpp-india", "hdid-uk", "hdid-waltham", "hdid-germany", "hdid-pune", "future funding"]
    FLOOR_SORT_ORDER = ["- ensemble", "- content part 1", "- content part 2", "- mobility", "- hpp" ]


class PeopleDataKeysSIBU(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    NAME = "HR Name"
    NICK_NAME = "Nickname"
    REQ = "Requisition"
    LEVEL = "Title"
    TEAM_MODEL = {
            "HVS" : "[Forecast: Q1:20; Q2:26; Q3:29; Q4:34] -- 1 Tracks @ (1 PO, 5 Dev, 1 QA, 1 Auto)",
            "HVS EM" : "2 Tracks @ (1 PO, 4 Dev, 1 QA, 1 Char, 1 Auto)",
            "Lumada - System" : "[Forecast: Q1:7; Q2:6; Q3:43; Q4:110]",
            "Lumada - Studio" : "[Forecast: Q1:7; Q2:10; Q3:43; Q4:110]",
            "City Data Exchange" : "[Forecast: Q1:6; Q2:19; Q3:6; Q4:6]",
            "Predictive Maintenance" : "[Forecast: Q1:5; Q2:17; Q3:22; Q4:27]",
            "Optimized Factory" : "[Forecast: Q1:1; Q2:6; Q3:13; Q4:15]",
        }

    PRODUCT_SORT_ORDER = ["hvs", "hvs em", "vmp", "hvp", "smart city technology", "technology", "tactical integration",
                          "tactical integrations",  "lumada - system", "sc iiot", "bel iiot", "lumada platform", "pdm", "predictive maintenance",
                          "lumada - studio", "lumada - microservices", "optimized factory", "opf", "city data exchange",
                          "cde", "denver", "lumada - ai", "lumada - analytics", "lumada - di", "lumada - hci", "hci", "lumada - machine intelligence", "lumada", "cross", "lumada cross", "global"]

class PeopleDataKeysHPP(PeopleDataKeys):
    def __init__(self):
        PeopleDataKeys.__init__(self)
    FUNCTION = "Function"
    NAME = "HR Name"
    NICK_NAME = "Name"
    #CROSS_FUNCTIONS = ["Technology", "DevOps", "Admin", "Seal" ]

    def __init__(self):
        PeopleDataKeys.__init__(self)

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
            #if HRName  is blank but nick name is populated - use that
            aName = self.getRawName() or self.getRawNickName()

        if "," in aName:
            return " ".join(aName.split(",")[1:]).strip()
        return aName.split(" ")[0].strip()

    def getLastName(self, aName=None):
        if not aName:
            #if HR Name is blank but nick name is populated - use that
            aName = self.getRawName() or self.getRawNickName()

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

    def getFloors(self):
        floors = list()
        for aFloor, managerNames in self.peopleDataKeys.FLOORS.iteritems():
            for aManagerName in managerNames:
                aFullName = self.getFullName(aManagerName)
                if (self.getFullName() == aFullName
                or (self.getRawName() == aFullName)
                or (self.getNormalizedRawName() == aFullName)):
                    floors.append(aFloor)
        if not floors:
            floors.append("")
        return floors

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
        """
        Compare entries in the spreadsheet based on their fullname. If the entry is 'TBH', assume it's unique.

        :param other:
        :return:
        """
        if not isinstance(other, PersonRowWrapper):
            return False

        # All TBHs have the same name so we assume each one is unique or they would all
        # get merged into 1
        if other.isTBH():
            return False

        if self.getFullName() == other.getFullName():
            return True

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return hash(self.getFullName())


class OrgParser:
    def getManagerSet(self):
        raise NotImplementedError

    def getPerson(self, aRow):
        raise NotImplementedError

    def getProductSet(self):
        raise NotImplementedError

    def getFeatureTeamSet(self, productName):
        raise NotImplementedError

    def getFunctionSet(self, productName=None):
        raise NotImplementedError

    def getLocationSet(self, productName=""):
        raise NotImplementedError

    def getPeople(self):
        raise NotImplementedError

    def getFilteredPeople(self, peopleFilter=None):
        raise NotImplementedError

    def getOrgName(self):
        raise NotImplementedError

    def getCrossFuncPeople(self):
        raise NotImplementedError

    def getTeamModel(self):
        raise NotImplementedError

    def getCrossFunctions(self):
        raise NotImplementedError

    def getFloorSortOrder(self):
        raise NotImplementedError

    def getProductSortOrder(self):
        raise NotImplementedError

class MultiOrgParser(OrgParser):
    def __init__(self, workbookNames, dataSheetName):
        self.orgSheets = []
        for aWorkbook in workbookNames:
            self.orgSheets.append(SingleOrgParser(aWorkbook, dataSheetName))

    def getOrgName(self):
        # If there is only 1 org sheet, use that name
        if 1 == len(self.orgSheets):
            return self.orgSheets[0].getOrgName()
        # If there are multiple, leave it empty since it's multiple orgs
        return ""

    def getManagerSet(self):
        managerSet = set()
        for orgSheet in self.orgSheets:
            managerSet.update(orgSheet.getManagerSet())
        return managerSet

    def getProductSet(self):
        productSet = set()
        for orgSheet in self.orgSheets:
            productSet.update(orgSheet.getProductSet())
        return productSet

    def getFeatureTeamSet(self, productName):
        featureTeamSet = set()
        for orgSheet in self.orgSheets:
            featureTeamSet.update(orgSheet.getFeatureTeamSet(productName))
        return featureTeamSet

    def getFunctionSet(self, productName=None):
        functionSet = []
        for orgSheet in self.orgSheets:
            functionSet.extend(orgSheet.getFunctionSet(productName))
        return functionSet

    def getLocationSet(self, productName=""):
        locationSet = set()
        for orgSheet in self.orgSheets:
            locationSet.update(orgSheet.getLocationSet(productName))
        return locationSet

    def getFilteredPeople(self, peopleFilter=None):
        filteredPeople = set()
        for orgSheet in self.orgSheets:
            filteredPeople.update(orgSheet.getFilteredPeople(peopleFilter))

        filteredPeople = list(filteredPeople)
        filteredPeople.sort()
        return filteredPeople

    def getCrossFuncPeople(self):
        crossFuncPeople = set()
        for orgSheet in self.orgSheets:
            crossFuncPeople.update(orgSheet.getCrossFuncPeople())
        return crossFuncPeople

    def getTeamModel(self):
        teamModel = {}
        for orgSheet in self.orgSheets:
            teamModel.update(orgSheet.getTeamModel())
        return teamModel

    def getCrossFunctions(self):
        crossFunctions = set()
        for orgSheet in self.orgSheets:
            crossFunctions.update(orgSheet.getCrossFunctions())
        return crossFunctions

    def getProductSortOrder(self):
        productSortOrder = []
        for orgSheet in self.orgSheets:
            productSortOrder.extend(orgSheet.getProductSortOrder())
        return productSortOrder

    def getFloorSortOrder(self):
        floorSortOrder = []
        for orgSheet in self.orgSheets:
            floorSortOrder.extend(orgSheet.getFloorSortOrder())
        return floorSortOrder

    def getCrossFuncTeams(self):
        crossFuncTeams = set()
        for orgSheet in self.orgSheets:
            crossFuncTeams.update(orgSheet.getCrossFuncTeam())
        return crossFuncTeams



class SingleOrgParser(OrgParser):
    def __init__(self, workbookName, dataSheetName, ):
        """

        :type workbookName: str
        :type dataSheetName: str
        """

        self.peopleDataKeys = PeopleDataKeys()
        filename = os.path.basename(workbookName).lower()
        self.orgName = filename.split("Staff")[0].strip()



        if "waltham" in filename or "content" in filename:
            self.peopleDataKeys = PeopleDataKeysWaltham()

        if "hpp" in filename:
            self.peopleDataKeys = PeopleDataKeysHPP()

        if "bellevue" in filename:
            self.peopleDataKeys = PeopleDataKeysBellevue()

        if "clara" in filename:
            self.peopleDataKeys = PeopleDataKeysSantaClara()

        if "sibu" in filename:
            self.peopleDataKeys = PeopleDataKeysSIBU()

        self.spreadsheetParser = SpreadsheetParser(workbookName, dataSheetName)

        # Manager set needs to be created and cached at the start so when an individual person is created, we can check
        # whether the person is a manager
        self.managerSet = set()
        self.managerSet = self.getManagerSet()

    def getOrgName(self):
        return self.orgName

    def getManagerSet(self):
        """
        :return:
        """
        if self.managerSet:
            return self.managerSet
        managerSet = set()
        for aRow in self.spreadsheetParser.dataRows():
            managerName = self.spreadsheetParser.getColValueByName(aRow, self.peopleDataKeys.MANAGER)
            if managerName:
                managerSet.add(managerName)
        return managerSet

    def getProductSet(self):
        """

        :return:
        """
        productSet = set()
        for aPerson in self._getPeople():
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

    def getFilteredPeople(self, peopleFilter=None):
        """ Get all the people that match the filter
        """
        if not peopleFilter:
            peopleFilter = PeopleFilter()

        matchingPeople = []

        for aPerson in self._getPeople():
            if (aPerson.getRawNickName() or aPerson.getRawName()) and peopleFilter.isMatch(aPerson):
                matchingPeople.append(aPerson)
        return matchingPeople

    def getCrossFuncPeople(self):
        crossFuncPeople = []
        # All the global 'cross func' people
        for aFunc in self.peopleDataKeys.CROSS_FUNCTIONS:
            crossFuncPeople.extend(self.getFilteredPeople(PeopleFilter().addFunctionFilter(aFunc)))

        # Add folks directly on the cross func team
        crossFuncTeam = self.getFilteredPeople(PeopleFilter().addProductFilter(self.peopleDataKeys.CROSS_FUNCT_TEAM))
        crossFuncPeople.extend(crossFuncTeam)

        return crossFuncPeople

    def getTeamModel(self):
        return self.peopleDataKeys.TEAM_MODEL

    def getCrossFunctions(self):
        return self.peopleDataKeys.CROSS_FUNCTIONS

    def getFloorSortOrder(self):
        return self.peopleDataKeys.FLOOR_SORT_ORDER

    def getProductSortOrder(self):
        return self.peopleDataKeys.PRODUCT_SORT_ORDER

    def getCrossFuncTeam(self):
        return self.peopleDataKeys.CROSS_FUNCT_TEAM

    def _getPerson(self, aRow):
        aPerson = PersonRowWrapper(self.spreadsheetParser, self.peopleDataKeys, aRow)
        managerSet = self.getManagerSet()
        if (aPerson.getRawName() in managerSet
            or aPerson.getRawNickName() in managerSet
            or aPerson.getFullName() in managerSet):
            aPerson.setManager()
        return aPerson

    def _getPeople(self):

        for aRow in self.spreadsheetParser.dataRows():
            aPerson = self._getPerson(aRow)
            yield aPerson

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

