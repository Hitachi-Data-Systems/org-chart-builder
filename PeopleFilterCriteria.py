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

class IsInternCriteria(FilterCriteria):
    def matches(self, aPerson):
        return aPerson.isIntern()

class IsExpatCriteria(FilterCriteria):
    def matches(self, aPerson):
        return aPerson.isExpat()

