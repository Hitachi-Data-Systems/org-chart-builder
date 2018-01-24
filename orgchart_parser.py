#!/usr/bin/python
import os
import orgchart_keys
from orgchart_person import EngineeringPersonRowWrapper

from people_filter_criteria import ProductCriteria, FunctionalGroupCriteria, IsInternCriteria, IsExpatCriteria, \
    FeatureTeamCriteria, IsCrossFuncCriteria, ManagerCriteria, IsTBHCriteria, LocationCriteria, IsManagerCriteria, \
    IsProductManagerCriteria, ManagerEmptyCriteria

from spreadsheet_parser import SpreadsheetParser


__author__ = 'doreper'

NOT_SET = "!!NOT SET!!"


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

    ManagerSet = set()

    def __init__(self, workbookNames, dataSheetName):
        self.orgSheets = []
        for aWorkbook in workbookNames:
            self.orgSheets.append(SingleOrgParser(aWorkbook, dataSheetName))
        MultiOrgParser.ManagerSet = self.getManagerSet()

    def getOrgName(self):
        # If there is only 1 org sheet, use that name
        if 1 == len(self.orgSheets):
            return self.orgSheets[0].getOrgName()
        # If there are multiple, leave it empty since it's multiple orgs
        return ""

    def getManagerSet(self):
        if not MultiOrgParser.ManagerSet:
            for orgSheet in self.orgSheets:
                MultiOrgParser.ManagerSet.update(orgSheet.getManagerSet())
        return MultiOrgParser.ManagerSet

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
        functionSet = set()
        for orgSheet in self.orgSheets:
            functionSet.update(orgSheet.getFunctionSet(productName))
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
            crossFuncTeams.add(orgSheet.getCrossFuncTeam())
        return crossFuncTeams



class SingleOrgParser(OrgParser):
    def __init__(self, workbookName, dataSheetName, ):
        """

        :type workbookName: str
        :type dataSheetName: str
        """

        self.peopleDataKeys = orgchart_keys.PeopleDataKeys()
        filename = os.path.basename(workbookName).lower()
        self.orgName = filename.split("staff")[0].strip()

        if "waltham" in filename or "content" in filename:
            self.peopleDataKeys = orgchart_keys.PeopleDataKeysWaltham()

        if "hpp" in filename:
            self.peopleDataKeys = orgchart_keys.PeopleDataKeysHPP()

        if "bellevue" in filename or "converged" in filename:
            self.peopleDataKeys = orgchart_keys.PeopleDataKeysBellevue()

        if "clara" in filename:
            self.peopleDataKeys = orgchart_keys.PeopleDataKeysSantaClara()

        if "sibu" in filename or "insight" in filename:
            self.peopleDataKeys = orgchart_keys.PeopleDataKeysSIBU()

        self.spreadsheetParser = SpreadsheetParser(workbookName, dataSheetName)

        # Manager set needs to be created and cached at the start so when an individual person is created, we can check
        # whether the person is a manager
        self.managerSet = set()
        self.managerSet = self.getManagerSet()
        self.peopleCache = []

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
        aPerson = EngineeringPersonRowWrapper(self.spreadsheetParser, self.peopleDataKeys, aRow)
        managerSet = MultiOrgParser.ManagerSet

        if (aPerson.getRawName() in managerSet
            or aPerson.getRawNickName() in managerSet
            or aPerson.getFullName() in managerSet
            or aPerson.getPreferredName() in managerSet):
            aPerson.setManager()
        return aPerson

    def _getPeople(self):
        if not self.peopleCache:
            self._populatePeopleCache()

        if self.peopleCache:
            for aPerson in self.peopleCache:
                yield aPerson

    def _populatePeopleCache(self):
        for aRow in self.spreadsheetParser.dataRows():
            aPerson = self._getPerson(aRow)
            self.peopleCache.append(aPerson)


class PeopleFilter:
    def __init__(self):
        self.filterList = []

    def addManagerFilter(self, manager):
        self.filterList.append(ManagerCriteria(manager))
        return self

    def addManagerEmptyFilter(self):
        self.filterList.append(ManagerEmptyCriteria())
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

