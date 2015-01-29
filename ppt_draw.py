#!/usr/bin/env python

import argparse
import glob
import os

from orgchart_parser import OrgParser, PeopleFilter
import sys
from pptx import Presentation
import orgchart_parser
from ppt_slide import DrawChartSlide

__author__ = 'David Oreper'


class OrgDraw:
    def __init__(self, workbookPath, sheetName, draftMode=False):
        """

        :type workbookPath: str
        :type sheetName: str
        """
        self.presentation = Presentation('HDSPPTTemplate.pptx')
        self.presentation.slide_height = DrawChartSlide.MAX_HEIGHT_INCHES
        self.presentation.slide_width = DrawChartSlide.MAX_WIDTH_INCHES
        self.slideLayout = self.presentation.slide_layouts[4]
        self.orgParser = OrgParser(workbookPath, sheetName)
        self.draftMode = draftMode

    def save(self, filename):
        self.presentation.save(filename)


    def getFirstLastName(self, aName):
        """

        :type aName: str
        :return:
        """
        if "," in aName:
            nameParts = aName.split(",")
            aName = "{} {}".format(nameParts[-1], " ".join(nameParts[0:-1]))
        return aName

    def _sortManagers(self, a, b):
        a = self.getFirstLastName(a)
        b = self.getFirstLastName(b)
        if a < b:
            return -1
        if b > a:
            return 1
        return 0

    def _getDirects(self, aManagerName):
        """

        :type aManagerName: str
        :return:
        """
        peopleFilter = PeopleFilter()
        peopleFilter.addManagerFilter(aManagerName)
        peopleFilter.addIsTBHFilter(False)
        directReports = self.orgParser.getFilteredPeople(peopleFilter)

        return directReports

    def drawAdmin(self):
        managersLeft = self.orgParser.getManagerSet()
        completedManagers = set()

        for aFloor in self.orgParser.peopleDataKeys.FLOORS.keys():
            chartDrawer = DrawChartSlide(self.presentation, "{} Admin {}".format(self.orgParser.orgName, aFloor), self.slideLayout)
            managerList = self.orgParser.peopleDataKeys.FLOORS[aFloor]
            managerList.sort(cmp=self._sortManagers)
            for aManagerName in managerList:
                directReports = []
                directReports.extend(self._getDirects(aManagerName))
                completedManagers.add(aManagerName)

                aManagerName = self.getFirstLastName(aManagerName)
                if not aManagerName:
                    if not self.draftMode:
                        continue
                    aManagerName = orgchart_parser.NOT_SET
                self.buildGroup(aManagerName, directReports, chartDrawer)
            chartDrawer.drawSlide()

        managersLeft = list(set(managersLeft) - completedManagers)
        managersLeft.sort(cmp=self._sortManagers)

        if len(managersLeft):
            chartDrawer = DrawChartSlide(self.presentation, "{} Admin".format(self.orgParser.orgName), self.slideLayout)
            for aManagerName in managersLeft:
                directReports = self._getDirects(aManagerName)
                if not len(directReports):
                    continue
                aManagerName = self.getFirstLastName(aManagerName)
                if not aManagerName:
                    if not self.draftMode:
                        continue
                    aManagerName = orgchart_parser.NOT_SET
                self.buildGroup(aManagerName, directReports, chartDrawer)
            chartDrawer.drawSlide()

    def drawExpat(self):
        expats = self.orgParser.getFilteredPeople(PeopleFilter().addIsExpatFilter())
        self._drawMiscGroups("ExPat", expats)

    def drawIntern(self):
        interns = self.orgParser.getFilteredPeople(PeopleFilter().addIsInternFilter())
        self._drawMiscGroups("Intern", interns)

    def _drawMiscGroups(self, slideName, peopleList):
        if not len(peopleList):
            return
        chartDrawer = DrawChartSlide(self.presentation, slideName, self.slideLayout)

        peopleFunctions = set([aPerson.getFunction() for aPerson in peopleList])

        for aFunction in peopleFunctions:
            self.buildGroup(aFunction, [aPerson for aPerson in peopleList if aPerson.getFunction() == aFunction], chartDrawer)
        chartDrawer.drawSlide()

    def drawCrossFunc(self):
        crossFuncPeople = []

        for aFunc in self.orgParser.peopleDataKeys.CROSS_FUNCTIONS:
            crossFuncPeople.extend(self.orgParser.getFilteredPeople(
                PeopleFilter().addFunctionFilter(aFunc)))

        if not len(crossFuncPeople):
            return

        chartDrawer = DrawChartSlide(self.presentation, "Cross Functional", self.slideLayout)

        functions = set([aPerson.getFunction() for aPerson in crossFuncPeople])

        for aFunction in functions:
            peopleFilter = PeopleFilter()
            peopleFilter.addFunctionFilter(aFunction)
            peopleFilter.addIsCrossFuncFilter()
            peopleFilter.addIsExpatFilter(False)
            peopleFilter.addIsInternFilter(False)

            funcPeople = self.orgParser.getFilteredPeople(peopleFilter)
            self.buildGroup(aFunction, funcPeople, chartDrawer)
        chartDrawer.drawSlide()

    def drawAllProducts(self, drawFeatureTeams):
        productList = list(self.orgParser.getProductSet())
        if self.orgParser.peopleDataKeys.CROSS_FUNCT_TEAM in productList:
            productList.remove(self.orgParser.peopleDataKeys.CROSS_FUNCT_TEAM)
        productList.sort()

        for aProductName in productList:
            self.drawProduct(aProductName, drawFeatureTeams)

    def sortByFunc(self, a, b):
        funcOrder = ["lead", "head coach", "po", "product owner", "product owner/qa", "technology", "ta", "tech", "sw architecture", "dev",
                     "development", "qa", "quality assurance", "stress",
                     "characterization", "auto", "aut", "automation", "sustaining", "solutions and sustaining",
                     "ui", "ux", "ui/ux", "inf", "infrastructure", "devops", "cross functional", "cross", "doc",
                     "documentation"]

        if a.lower() in funcOrder:
            if b.lower() in funcOrder:
                if funcOrder.index(a.lower()) > funcOrder.index(b.lower()):
                    return 1
            return -1

        if b in funcOrder:
            return -1

        return 0

    def drawProduct(self, productName, drawFeatureTeams=False):
        """

        :type productName: str
        :type chartDrawer: ppt_draw.DrawChartSlide
        """
        teamName = ""
        if not productName:
            if not self.draftMode:
                return

        featureTeamList = [""]
        if drawFeatureTeams:
            featureTeamList = list(self.orgParser.getFeatureTeamSet(productName))
        functionList = list(self.orgParser.getFunctionSet(productName))
        functionList.sort(cmp=self.sortByFunc)

        for aFeatureTeam in featureTeamList:
            if not productName:
                slideTitle = orgchart_parser.NOT_SET
            elif drawFeatureTeams:
                teamName = "- {} ".format(aFeatureTeam)
                if not aFeatureTeam:
                    if len(featureTeamList) > 1:
                        teamName = "- Cross "
                    else:
                        teamName = ""
                slideTitle = "{} {}Feature Team".format(productName, teamName)
            else:
                slideTitle = "{} Functional Team".format(productName)

            chartDrawer = DrawChartSlide(self.presentation, slideTitle, self.slideLayout)
            for aFunction in functionList:
                if aFunction.lower() in self.orgParser.peopleDataKeys.CROSS_FUNCTIONS:
                    continue

                peopleFilter = PeopleFilter()
                peopleFilter.addProductFilter(productName)
                peopleFilter.addFunctionFilter(aFunction)

                if drawFeatureTeams:
                    peopleFilter.addFeatureTeamFilter(aFeatureTeam)
                else:
                    peopleFilter.addIsExpatFilter(False)
                    peopleFilter.addIsInternFilter(False)

                functionPeople = self.orgParser.getFilteredPeople(peopleFilter)

                if not functionPeople:
                    print "WARNING: No members added to '{}' for product: '{}{}'".format(aFunction, productName, teamName)
                    continue
                self.buildGroup(aFunction, functionPeople, chartDrawer)

            chartDrawer.drawSlide()

    def buildGroup(self, functionName, functionPeople, chartDrawer):
        """

        :type functionName: str
        :type functionPeople: list
        :type chartDrawer: ppt_draw.DrawChartSlide
        :return:
        """
        functionPeople = [person for person in functionPeople if person.getRawName().strip() != ""]
        if len(functionPeople) == 0:
            return

        if not functionName:
            if not self.draftMode:
                return
            functionName = orgchart_parser.NOT_SET

        chartDrawer.addGroup(functionName, functionPeople)

def main(argv):
    userDir = os.environ.get("USERPROFILE") or os.environ.get("HOME")
    defaultSheetName = "PeopleData"
    defaultDir = os.path.join(userDir, "Documents/HCP Anywhere/Org Charts and Hiring History")
    defaultOutputFile = "OrgChart.pptx"

    examples = """
    Examples:
    # Print functional layout for Waltham Staff. Uses unique identifier for a file in default director: {}
        %prog Waltham Staff -f
    # Print Admin layout for Bellevue. Uses fully qualified path to the spreadsheet
        %prog C:\Users\doreper\Documents\HCP Anywhere\Org Charts and Hiring History\Bellevue Staff.xlsm -a
    # Print Admin layout for Waltham. Uses fully qualified path to the spreadsheet. Uses unique identifier
    # for file in the specified directory
        %prog ham staff -d {}\Documents\HCP Anywhere\Org Charts and Hiring History\ -a
    """.format(defaultDir, userDir)

    parser = argparse.ArgumentParser(description="""This tool is used to parse staff spreadsheet and display
    information in a format that can easily be pasted into an excel smartArt chart builder""",
                                     epilog=examples, formatter_class=argparse.RawDescriptionHelpFormatter)

    parser.add_argument("path", nargs="+", type=str,
                        help="unique file token for file in directory specified by '-d [default={}]' ".format(
                            defaultDir))
    parser.add_argument("-d", "--directory", type=str, help="directory for the spreadsheet",
                        default=defaultDir)

    parser.add_argument("-s", "--sheetName", type=str, default=defaultSheetName, help="Sheet Name")

    parser.add_argument("-o", "--outputFile", type=str, default=None, help="output file")
    parser.add_argument("-f", "--featureTeam", action="store_true", default=False, help="Show products by feature team")


    parser.add_argument("--draftMode", type=bool, default=False,
                        help="Show {} for people that don't have manager, product, function set. Otherwise, "
                             "people with missing fields are not represented on the chart".format(
                            orgchart_parser.NOT_SET))

    options = parser.parse_args(argv)

    specifiedPath = " ".join(options.path)
    if os.path.exists(specifiedPath):
        workbookPath = specifiedPath
    else:
        fileMatch = glob.glob(os.path.join(options.directory, "*{}*".format(specifiedPath)))
        fileMatch = [aFile for aFile in fileMatch if
                     (not ((os.path.basename(aFile).startswith("~")) or "conflict" in os.path.basename(aFile).lower()))]
        if not fileMatch:
            raise OSError("Could not find any files in directory: '{}' that contain string: '{}'"
                          .format(options.directory, specifiedPath))
        if len(fileMatch) > 1:
            raise OSError(
                "Too many files found in dir: '{}' that contain string '{}' : \n\t\t{}".format(options.directory,
                                                                                               specifiedPath,
                                                                                               "\n\t\t".join(
                                                                                                   fileMatch)))
        workbookPath = fileMatch[0]

    orgDraw = OrgDraw(workbookPath, options.sheetName, options.draftMode)

    orgDraw.drawAllProducts(options.featureTeam)
    orgDraw.drawCrossFunc()
    if not options.featureTeam:
        orgDraw.drawExpat()
        orgDraw.drawIntern()
    orgDraw.drawAdmin()

    outputFileName = options.outputFile
    if not outputFileName:
        outputFileName = "{}{}".format(orgDraw.orgParser.orgName, defaultOutputFile)
    orgDraw.save(outputFileName)


if __name__ == "__main__":
    # sys.argv = ["", 'Z:\Documents\HCP Anywhere\Org Charts and Hiring History\SantaClara Staff.xlsm', "-f"]
    # sys.argv = ["", 'Z:\Documents\HCP Anywhere\Org Charts and Hiring History\WalthamStaff.xlsm']
    # sys.argv = ["", 'Z:\Documents\HCP Anywhere\Org Charts and Hiring History\Bellevue Staff.xlsm']
    #
    # for davep:
    #sDir = '/Users/dpinkney/Documents/HCP Anywhere/SharedWithMe/Org Charts and Hiring History'
    #sys.argv = ['', '-d', sDir, '-o%s/WalthamChartGen.pptx' % sDir, 'WalthamStaff.xlsm']
    #sys.argv = ['', '-d', sDir, '-oSantaClara Staff Gen-dbp.pptx', 'SantaClara Staff.xlsm']
    main(sys.argv[1:])
