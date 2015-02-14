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
        return actualValue.lower() == self.expectedValue.lower()


class ProductCriteria(KeyMatchesCriteria):
    def matches(self, aPerson):
        return self._matches(aPerson.getProduct())

class FunctionalGroupCriteria(KeyMatchesCriteria):
    def matches(self, aPerson):
        return self._matches(aPerson.getFunction())

class FeatureTeamCriteria(KeyMatchesCriteria):
    def matches(self, aPerson):
        return self._matches(aPerson.getFeatureTeam())

class LocationCriteria(KeyMatchesCriteria):
    def matches(self, aPerson):
        return self._matches(aPerson.getLocation())

class ManagerCriteria(FilterCriteria):
    def __init__(self, manager):
        FilterCriteria.__init__(self)
        self.manager = manager

    def matches(self, aPerson):
        personManager = aPerson.getManagerFullName()

        return (personManager == self.manager.getFullName()
                    or personManager == self.manager.getRawName()
                    or personManager == self.manager.getRawNickName()
                    or personManager == self.manager.getNormalizedRawName())

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

class IsManagerCriteria(FilterCriteria):
    def __init__(self, isManager):
        FilterCriteria.__init__(self)
        self.isManager = isManager

    def matches(self, aPerson):
        return aPerson.isManager() == self.isManager


