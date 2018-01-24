import datetime
import re
import dateutil.parser

__author__ = 'David Oreper'

class SkeletonPerson:
    def __init__(self, fullName):
        self.fullName = fullName


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
        return self.fullName

    def getNormalizedRawName(self):
        return "{} {}".format(self.getFirstName(self.getRawName()), self.getLastName(self.getRawName()))

    def getPreferredName(self):
        if self.getRawNickName():
            return "{} {}".format(self.getRawNickName(), self.getLastName())
        return self.getFullName()

    def getRawNickName(self):
        return self.getFirstName()

    def __str__(self):
        return self.getPreferredName()

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
        if not isinstance(other, SkeletonPerson):
            return False


        # All TBHs have the same name so we assume each one is unique or they would all
        # get merged into 1
        if other.isTBH():
            return False

        if self.getFullName().lower() == other.getFullName().lower():
            return True

        if self.getFullName() and (other.getFullName().lower() == self.getFullName().lower()):
            return True

        if self.getRawName() and (other.getRawName().lower() == self.getRawName().lower()):
            return True

        if self.getNormalizedRawName() and (other.getNormalizedRawName().lower() == self.getNormalizedRawName().lower()):
            return True

        if self.getRawNickName().lower() or other.getRawNickName().lower():
            if other.getPreferredName().lower() == self.getPreferredName().lower():
                return True

        return False

    def getFloors(self):
        #TODO - Either get rid of floors concept or find a way to map skeleton person to a floor
        return [""]


    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        # Use the last name for hash so we can further compare collisions with the __eq__ algorithm and identify
        # people as the same if their fullname is different from their preferred name
        #print "TODO: HASH: {} - {}".format(self.getLastName(), hash(self.getLastName()))
        return hash(self.getLastName().lower())

    def isTBH(self):
        if (self.getRawName().lower().startswith("tbh")
            or self.getRawName().lower().startswith("tbd")):
            return True
        return False

    def isUnfunded(self):
        if ("unfunded" in self.getFullName().lower()):
            return True
        return False



class EngineeringPersonRowWrapper(SkeletonPerson):
    def __init__(self, spreadsheetParser, peopleDataKeys, aRow):

        self.spreadsheetParser = spreadsheetParser
        self.peopleDataKeys = peopleDataKeys
        self.aRow = aRow
        self.manager = False
        SkeletonPerson.__init__(self, spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.NAME))

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
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.REQ)

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
        return u"{} {}".format(self.getFirstName(fullName), self.getLastName(fullName)).strip()

    def getRawName(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.NAME).strip()

    def getNormalizedRawName(self):
        return u"{} {}".format(self.getFirstName(self.getRawName()), self.getLastName(self.getRawName()))

    def getPreferredName(self):
        if self.getRawNickName():
            return u"{} {}".format(self.getRawNickName(), self.getLastName())
        return self.getFullName()

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

    def isProductManager(self):
        return self.getFunction().lower() in ["pm", "product manager", "product management"]

    def isCrossFunc(self):
        return ((self.getFunction().lower() in self.peopleDataKeys.CROSS_FUNCTIONS)
                or (self.getProduct().lower() == self.peopleDataKeys.CROSS_FUNCT_TEAM.lower()))

    def getTitle(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.TITLE).strip()

    def getLevel(self):
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
                if SkeletonPerson(aManagerName) == self:
                    floors.append(aFloor)
        if not floors:
            floors.append("")
        return floors

    def getLocation(self):
        if not self.spreadsheetParser.columnExists(self.peopleDataKeys.LOCATION):
            return ""
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.LOCATION).strip() or ""

    def getCostCenter(self):
        if not self.spreadsheetParser.columnExists(self.peopleDataKeys.COST_CENTER):
            return ""

        rawCostCenter = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.COST_CENTER).strip()
        rawCostCenter = rawCostCenter.replace(".", "-")
        rawCostCenter.strip()

        costCenterMatch = re.search('((\d+-?)+)', rawCostCenter)
        costCenter = ""
        if costCenterMatch:
            costCenter = costCenterMatch.group(0)
        return costCenter


    def getStartDate(self):
        """


        :return: DateTime Object. Return empty datetime object if date is not set
        """
        if not self.spreadsheetParser.columnExists(self.peopleDataKeys.START_DATE):
            return ""

        startDateStr = self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.START_DATE).strip()

        if not startDateStr:
            return dateutil.parser.parse("1/1/1900")

        try:
            return dateutil.parser.parse(startDateStr)
        except ValueError:
            print "Warning: can not parse start date for {}: '{}'".format(self.getFullName(), startDateStr)

    def __str__(self):
        personStr = "Person: {}".format(self.getFullName())
        return personStr

    def __repr__(self):
        return self.__str__()

class FinancePersonRowWrapper(EngineeringPersonRowWrapper):
    def getFirstName(self, aName=None):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.FIRST_NAME).strip()

    def getLastName(self, aName=None):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.LAST_NAME).strip()

    def getRawName(self):
        return u"{} {}".format(self.getFirstName(), self.getLastName())

    def getPreferredName(self):
        if self.getRawNickName():
            return u"{} {}".format(self.getRawNickName(), self.getLastName())
        return self.getFullName()

    def getRawNickName(self):
        return self.spreadsheetParser.getColValueByName(self.aRow, self.peopleDataKeys.NICK_NAME).strip()

    def isConsultant(self):
        """
        :return:
        """
        title = self.getTitle()
        return title.lower() == "contractor"