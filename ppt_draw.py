#!/usr/bin/env python

import argparse
import glob
import os
import pprint
from unittest import TestCase
import datetime

from orgchart_parser import OrgParser, PeopleFilter
import sys
from pptx import Presentation
import orgchart_parser
from ppt_slide import DrawChartSlide, DrawChartSlideAdmin, DrawChartSlideTBH, DrawChartSlideExpatIntern

__author__ = 'David Oreper'


class OrgDraw:
    def __init__(self, workbookPath, sheetName, draftMode):
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

    def _getDirects(self, aManager, location=None):
        """

        :type aManager: str
        :return:
        """
        peopleFilter = PeopleFilter()
        directReports = []
        peopleFilter.addManagerFilter(aManager)
        peopleFilter.addIsTBHFilter(False)
        peopleFilter.addLocationFilter(location)
        directReports.extend(self.orgParser.getFilteredPeople(peopleFilter))

        return directReports

    def drawAdmin(self):
        managerList = self.orgParser.getFilteredPeople(PeopleFilter().addIsManagerFilter())
        allManagerNames = list(self.orgParser.managerList)

        # Make sure we're adding all managers.
        # Could be a problem if a person has a manger that isn't entered as a row
        for aManagerName in self.orgParser.managerList:
            for aManager in managerList:
                if aManager.getFullName() == aManager.getFullName(aManagerName):
                    allManagerNames.remove(aManagerName)

        if allManagerNames:
            print "WARNING: Managers not drawn because they are not entered as row: {}".format(pprint.pformat(allManagerNames))


        managersByFloor = {}
        for aManager in managerList:
            if not aManager.getFloor() in managersByFloor:
                managersByFloor[aManager.getFloor()] = set()
            managersByFloor[aManager.getFloor()].add(aManager)

        # Location: There can be people across locations reporting to the same manager:
        # Example: People report to Arno in Santa Clara and in Milan.
        # There will be a single slide for each unique location. Only direct reports in the specified location will be
        # drawn
        for aLocation in self.orgParser.getLocationSet():
            locationName = aLocation or self.orgParser.orgName

            # Floor: The floor (or other grouping) that separates manager.
            # Example: There are a lot of managers in Waltham so we break them up across floors
            # NOTE: All of the direct reports in the location will be drawn
            # Example: If Dave is on floor1 and floor2 in Waltham, then the same direct reports will be drawn both times

            sortedFloors = list(managersByFloor.keys())
            sortedFloors.sort(cmp=self._sortByFloor)

            for aFloor in sortedFloors:
                managerList = managersByFloor[aFloor]
                chartDrawer = DrawChartSlideAdmin(self.presentation, "{} Admin {}".format(locationName, aFloor), self.slideLayout)
                managerList = list(managerList)
                managerList.sort()
                for aManager in managerList:
                    directReports = []
                    directReports.extend(self._getDirects(aManager, aLocation))
                    self.buildGroup(aManager.getFullName(), directReports, chartDrawer)
                chartDrawer.drawSlide()

        peopleMissingManager = self._getDirects("")
        if peopleMissingManager:
            print "People missing manager: {}".format(pprint.pformat(peopleMissingManager))


    def drawExpat(self):
        expats = self.orgParser.getFilteredPeople(PeopleFilter().addIsExpatFilter())
                    #.addIsProductManagerFilter(False))
        self._drawMiscGroups("ExPat", expats)


    def drawProductManger(self):
        productManagers = self.orgParser.getFilteredPeople(PeopleFilter().addIsProductManagerFilter())
        if not len(productManagers):
            return
        chartDrawer = DrawChartSlide(self.presentation, "Product Management", self.slideLayout)

        peopleProducts = list(set([aPerson.getProduct() for aPerson in productManagers]))
        for aProduct in peopleProducts:
            self.buildGroup(aProduct, [aPerson for aPerson in productManagers if aPerson.getProduct() == aProduct], chartDrawer)
        chartDrawer.drawSlide()

    def drawIntern(self):
        interns = self.orgParser.getFilteredPeople(PeopleFilter().addIsInternFilter())
        self._drawMiscGroups("Intern", interns)

    def _drawMiscGroups(self, slideName, peopleList):
        if not len(peopleList):
            return
        chartDrawer = DrawChartSlideExpatIntern(self.presentation, slideName, self.slideLayout)
        peopleFunctions = list(set([aPerson.getFunction() for aPerson in peopleList]))
        peopleFunctions.sort(cmp=self._sortByFunc)
        for aFunction in peopleFunctions:
            self.buildGroup(aFunction, [aPerson for aPerson in peopleList if aPerson.getFunction() == aFunction], chartDrawer)
        chartDrawer.drawSlide()

    def drawCrossFunc(self):
        crossFuncPeople = []

        for aFunc in self.orgParser.peopleDataKeys.CROSS_FUNCTIONS:
            crossFuncPeople.extend(self.orgParser.getFilteredPeople(
                PeopleFilter().addFunctionFilter(aFunc)))

        crossFuncTeam = self.orgParser.getFilteredPeople(PeopleFilter().addProductFilter(self.orgParser.peopleDataKeys.CROSS_FUNCT_TEAM))
        crossFuncPeople.extend(crossFuncTeam)

        if not len(crossFuncPeople):
            return

        chartDrawer = DrawChartSlide(self.presentation, "Cross Functional", self.slideLayout)

        functions = list(set([aPerson.getFunction() for aPerson in crossFuncPeople]))
        functions.sort(cmp=self._sortByFunc)

        for aFunction in functions:
            peopleFilter = PeopleFilter()
            peopleFilter.addFunctionFilter(aFunction)
            peopleFilter.addIsCrossFuncFilter()
            peopleFilter.addIsExpatFilter(False)
            peopleFilter.addIsInternFilter(False)

            funcPeople = self.orgParser.getFilteredPeople(peopleFilter)
            self.buildGroup(aFunction, funcPeople, chartDrawer)
        chartDrawer.drawSlide()

    def drawAllProducts(self, drawFeatureTeams, drawLocations, drawExpatsInTeam):
        #Get all the products except the ones where a PM is the only member
        people = self.orgParser.getFilteredPeople(PeopleFilter().addIsProductManagerFilter(False))

        productList = list(set([aPerson.getProduct() for aPerson in people]))

        if self.orgParser.peopleDataKeys.CROSS_FUNCT_TEAM in productList:
            productList.remove(self.orgParser.peopleDataKeys.CROSS_FUNCT_TEAM)
        productList.sort(cmp=self._sortByProduct)

        for aProductName in productList:
            self.drawProduct(aProductName, drawFeatureTeams, drawLocations, drawExpatsInTeam)

    def drawProduct(self, productName, drawFeatureTeams=False, drawLocations=False, drawExpatsInTeam=True):
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
        functionList.sort(cmp=self._sortByFunc)

        teamModelText = None

        locations = [""]
        if drawLocations:
            locations = self.orgParser.getLocationSet(productName)

        for aLocation in locations:
            locationName = aLocation.strip() or self.orgParser.orgName

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
                    slideTitle = "{}".format(productName)
                    modelDict = self.orgParser.peopleDataKeys.TEAM_MODEL
                    if productName in modelDict:
                        teamModelText = modelDict[productName]

                chartDrawer = DrawChartSlide(self.presentation, slideTitle, self.slideLayout, teamModelText)
                if len(locations) > 1 and aLocation:
                   chartDrawer.setLocation(locationName)

                for aFunction in functionList:
                    if aFunction.lower() in self.orgParser.peopleDataKeys.CROSS_FUNCTIONS:
                        continue

                    peopleFilter = PeopleFilter()
                    peopleFilter.addProductFilter(productName)
                    peopleFilter.addFunctionFilter(aFunction)
                    if drawLocations:
                        peopleFilter.addLocationFilter(aLocation)

                    if drawFeatureTeams:
                        peopleFilter.addFeatureTeamFilter(aFeatureTeam)
                    else:
                        if not drawExpatsInTeam:
                            peopleFilter.addIsExpatFilter(False)
                        peopleFilter.addIsInternFilter(False)
                        # peopleFilter.addIsProductManagerFilter(False)

                    functionPeople = self.orgParser.getFilteredPeople(peopleFilter)
                    self.buildGroup(aFunction, functionPeople, chartDrawer)

                chartDrawer.drawSlide()

    def drawTBH(self):

        totalTBHSet = set(self.orgParser.getFilteredPeople(PeopleFilter().addIsTBHFilter()))
        tbhLocations = set([aTBH.getLocation() for aTBH in totalTBHSet]) or [""]
        tbhProducts = list(set([aTBH.getProduct() for aTBH in totalTBHSet])) or [""]
        tbhProducts.sort(cmp=self._sortByProduct)

        for aLocation in tbhLocations:
            title = "Hiring"

            # Location might not be set.
            if aLocation:
                title = "{} - {}".format(title, aLocation)

            chartDrawer = DrawChartSlideTBH(self.presentation, title, self.slideLayout)
            for aProduct in tbhProducts:
                productTBHList = self.orgParser.getFilteredPeople(PeopleFilter().addIsTBHFilter().addProductFilter(aProduct).addLocationFilter(aLocation))
                productTBHList = sorted(productTBHList, self._sortByFunc, lambda tbh: tbh.getProduct())
                self.buildGroup(aProduct, productTBHList, chartDrawer)

            chartDrawer.drawSlide()

    def _sortByFunc(self, a, b):
        funcOrder = ["lead", "leadership", "head coach", "product management", "pm", "po", "product owner", "product owner/qa", "technology", "ta", "technology architect", "tech", "sw architecture", "dev",
                     "development", "qa", "quality assurance", "stress",
                     "characterization", "auto", "aut", "automation", "sustaining", "solutions and sustaining",
                     "ui", "ux", "ui/ux", "inf", "infrastructure", "devops", "cross functional", "cross", "doc",
                     "documentation"]

        if a.lower() in funcOrder:
            if b.lower() in funcOrder:
                if funcOrder.index(a.lower()) > funcOrder.index(b.lower()):
                    return 1
            return -1

        if b.lower() in funcOrder:
            return 1

        return 0

    def _sortByFloor(self, a, b):
        floorOrder = self.orgParser.peopleDataKeys.FLOOR_SORT_ORDER

        if a.lower() in floorOrder:
            if b.lower() in floorOrder:
                if floorOrder.index(a.lower()) > floorOrder.index(b.lower()):
                    return 1
            return -1

        if b.lower() in floorOrder:
            return 1

        return 0

    def _sortByProduct(self, a, b):
        productOrder = self.orgParser.peopleDataKeys.PRODUCT_SORT_ORDER

        if a.lower() in productOrder:
            if b.lower() in productOrder:
                if productOrder.index(a.lower()) > productOrder.index(b.lower()):
                    return 1
            return -1

        if b.lower() in productOrder:
            return 1

        return 0

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
    parser.add_argument("-l", "--location", action="store_true", default=False, help="Show products by location")
    parser.add_argument("-t", "--tbh", action="store_true", default=False, help="Add a TBH slide")
    parser.add_argument("-e", "--expatsInTeam", action="store_true", default=False, help="Include expats in Product team slide")
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

    orgDraw.drawAllProducts(options.featureTeam, options.location, options.expatsInTeam)
    orgDraw.drawCrossFunc()
    if not options.featureTeam:
        orgDraw.drawExpat()
        orgDraw.drawIntern()
        orgDraw.drawProductManger()

    if options.tbh:
        orgDraw.drawTBH()

    orgDraw.drawAdmin()


    outputFileName = options.outputFile
    if not outputFileName:
        outputFileName = "{}{}".format(orgDraw.orgParser.orgName, defaultOutputFile)
    orgDraw.save(outputFileName.strip())


if __name__ == "__main__":
    # for davep:
    #sDir = '/Users/dpinkney/Documents/HCPAnywhere/SharedWithMe/Waltham Engineering Org Charts/'
    #sys.argv += ['-t', '-d', sDir, '-o%s/WalthamChartGen.pptx' % sDir, 'WalthamStaff.xlsm']
    #sys.argv += ['-d', sDir, '-o%s/HPP_Charts.pptx' % sDir, 'HPP_Staff.xlsm']
    #sDir = '/Users/dpinkney/Documents/HCPAnywhere/SharedWithMe/Waltham Engineering Org Charts/tmp/'
    #sys.argv += ['-d', sDir, '-o%s/WalthamChartGen-moves.pptx' % sDir, 'WalthamStaff-moves.xlsm']
    main(sys.argv[1:])


class GenChartCommandline(TestCase):

    def testSantaClara(self):
        todayDate = datetime.date.today().strftime("%Y-%m-%d")
        outputFileName = "{cwd}{slash}{dateStamp}_SantaClaraOrgChart.pptx".format(cwd=os.getcwd(), slash=os.sep, dateStamp=todayDate)
        #main(['C:\SantaClara Staff.xlsm', "-o {}".format(outputFileName)])
        #main(['C:\SantaClara StaffRainier_Model.xlsm', '-f'])
        main(['Z:\Documents\HCP Anywhere\Org Charts and Hiring History\Santa Clara\SantaClara Staff.xlsm', "-t", "-o {}".format(outputFileName)])
        # main(['C:\SantaClara Staff - Remodel.xlsm'])

        startCmd = 'start {}'.format(outputFileName)
        os.system(startCmd)

    def testSantaClaraFeatures(self):
        outputFileName = "{}{}SantaClaraOrgChart.feature.pptx".format(os.getcwd(),os.sep)
        main(['Z:\Documents\HCP Anywhere\Org Charts and Hiring History\Santa Clara\SantaClara Staff.xlsm', "-f", "-o {}".format(outputFileName)])
        os.system("start " + outputFileName)

    def testWaltham(self):
        main(['Z:\Documents\HCP Anywhere\Org Charts and Hiring History\Waltham\WalthamStaff.xlsm', "-t"])

    def testBellevue(self):
        #main(['Z:\Documents\HCP Anywhere\Org Charts and Hiring History\Bellevue\Bellevue Staff.xlsm'])
        main(['Z:\Documents\HCP Anywhere\Org Charts and Hiring History\Bellevue\Bellevue Staff.xlsm'])

    def testSIBU(self):
        todayDate = datetime.date.today().strftime("%Y-%m-%d")
        outputFileName = "{cwd}{slash}{dateStamp}_SIBUOrgChart.pptx".format(cwd=os.getcwd(), slash=os.sep, dateStamp=todayDate)
        main(['Z:\doreper On My Mac\Documents\HCP Anywhere\SIBU Org Charts and Hiring History\SIBUEngStaff.xlsm', "-t", "-o {}".format(outputFileName)])
        #main(['C:\SIBUEngStaff.xlsm', "-o {}".format(outputFileName)])
        os.system("start " + outputFileName)

    def testSIBU10M(self):
        todayDate = datetime.date.today().strftime("%Y-%m-%d")
        outputFileName = "{cwd}{slash}{dateStamp}_SIBUOrgChart10M.pptx".format(cwd=os.getcwd(), slash=os.sep, dateStamp=todayDate)
        main(['Z:\doreper On My Mac\Documents\HCP Anywhere\SIBU Org Charts and Hiring History\SIBUEngStaff10M.xlsm', "-t", "-o {}".format(outputFileName)])
        #main(['C:\SIBUEngStaff.xlsm', "-o {}".format(outputFileName)])
        os.system("start " + outputFileName)

    def testSIBU40M(self):
        todayDate = datetime.date.today().strftime("%Y-%m-%d")
        outputFileName = "{cwd}{slash}{dateStamp}_SIBUOrgChart40M.pptx".format(cwd=os.getcwd(), slash=os.sep, dateStamp=todayDate)
        main(['Z:\doreper On My Mac\Documents\HCP Anywhere\SIBU Org Charts and Hiring History\SIBUEngStaff40M.xlsm', "-t", "-o {}".format(outputFileName)])
        #main(['C:\SIBUEngStaff.xlsm', "-o {}".format(outputFileName)])
        os.system("start " + outputFileName)

