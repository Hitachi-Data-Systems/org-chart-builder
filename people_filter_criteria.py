__author__ = 'David Oreper'


class FilterCriteria:
    def __init__(self,):
        pass

    def matchesCriteria(self, aPerson):
        return True

class KeyMatchesCriteria(FilterCriteria):
    def __init__(self, expectedValue):
        FilterCriteria.__init__(self)
        self.expectedValue = expectedValue

    def _matches(self, actualValue):
        return actualValue == self.expectedValue


class ProductCriteria(KeyMatchesCriteria):
    def matches(self, aPerson):
        return self._matches(aPerson.getProduct())

class FunctionalGroupCriteria(KeyMatchesCriteria):
    def matches(self, aPerson):
        return self._matches(aPerson.getFunction())

class FeatureTeamCriteria(KeyMatchesCriteria):
    def matches(self, aPerson):
        return self._matches(aPerson.getFeatureTeam())

class ManagerCriteria(KeyMatchesCriteria):
    def getNormalizedFullName(self, aName):
        fullName = aName
        if "," in aName:
            firstName = aName.split(",")[1]
            lastName = aName.split(",")[0]
            fullName = "{} {}".format(firstName, lastName)
        return fullName.strip()

    def matches(self, aPerson):
        self.expectedValue = self.getNormalizedFullName(self.expectedValue)
        return self._matches(aPerson.getManagerFullName())

class IsInternCriteria(FilterCriteria):
    def __init__(self, isIntern):
        FilterCriteria.__init__(self)
        self.isIntern = isIntern

    def matches(self, aPerson):
        return aPerson.isIntern() == self.isIntern

class IsExpatCriteria(FilterCriteria):
    def __init__(self, isExpat):
        FilterCriteria.__init__(self)
        self.isExpat = isExpat

    def matches(self, aPerson):
        return aPerson.isExpat() == self.isExpat

class IsTBHCriteria(FilterCriteria):
    def __init__(self, isTBH):
        FilterCriteria.__init__(self)
        self.isTBH = isTBH

    def matches(self, aPerson):
        return aPerson.isTBH() == self.isTBH


class IsCrossFuncCriteria(FilterCriteria):
    def __init__(self, isCrossFunc):
        FilterCriteria.__init__(self)
        self.isCrossFunc = isCrossFunc

    def matches(self, aPerson):
        return aPerson.isCrossFunc() == self.isCrossFunc


