#!/usr/bin/env python

import argparse
import glob
import os
import pprint
from unittest import TestCase
import datetime

from orgchart_parser import MultiOrgParser, PeopleFilter
import sys
from pptx import Presentation
import orgchart_parser
from ppt_slide import DrawChartSlide, DrawChartSlideAdmin, DrawChartSlideTBH, DrawChartSlideExpatIntern, DrawChartSlidePM

__author__ = 'David Oreper'


class OrgDraw:
    def __init__(self, workbookPaths, sheetName, draftMode):
        """

        :type workbookPath: str
        :type sheetName: str
        """
        self.presentation = Presentation('HDSPPTTemplate.pptx')
        self.presentation.slide_height = DrawChartSlide.MAX_HEIGHT_INCHES
        self.presentation.slide_width = DrawChartSlide.MAX_WIDTH_INCHES
        self.slideLayout = self.presentation.slide_layouts[4]
        self.multiOrgParser = MultiOrgParser(workbookPaths, sheetName)
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
        directReports.extend(self.multiOrgParser.getFilteredPeople(peopleFilter))

        return directReports

    def drawAdmin(self):
        managerList = self.multiOrgParser.getFilteredPeople(PeopleFilter().addIsManagerFilter())
        allManagerNames = list(self.multiOrgParser.getManagerSet())

        # Make sure we're adding all managers.
        # Could be a problem if a person has a manger that isn't entered as a row
        for aManagerNameStr in self.multiOrgParser.getManagerSet():
            for aManager in managerList:
                aManagerFullName = aManager.getFullName(aManagerNameStr)
                if aManager.getFullName() == aManagerFullName or aManager.getNormalizedRawName() == aManagerFullName:
                    allManagerNames.remove(aManagerNameStr)

        if allManagerNames:
            print "WARNING: Managers not drawn because they are not entered as row: {}".format(pprint.pformat(allManagerNames))

        # A manager can have reports on more than one floor
        managersByFloor = {}
        for aManager in managerList:
            floors = aManager.getFloors()
            for floor in floors:
                if not floor in managersByFloor:
                    managersByFloor[floor] = set()
                managersByFloor[floor].add(aManager)

        # Location: There can be people across locations reporting to the same manager:
        # Example: People report to Arno in Santa Clara and in Milan.
        # There will be a single slide for each unique location. Only direct reports in the specified location will be
        # drawn
        for aLocation in self.multiOrgParser.getLocationSet():
            locationName = aLocation or self.multiOrgParser.getOrgName()

            # Floor: The floor (or other grouping) that separates manager.
            # Example: There are a lot of managers in Waltham so we break them up across floors
            # NOTE: All of the direct reports in the location will be drawn
            # Example: If Dave is on floor1 and floor2 in Waltham, then the same direct reports will be drawn both times

            sortedFloors = list(managersByFloor.keys())
            sortedFloors.sort(cmp=self._sortByFloor)

            maxManagersPerSlide = 7
            managersOnSlide = 0
            slideNameAddendum = "pt2"

            for aFloor in sortedFloors:
                managerList = managersByFloor[aFloor]
                chartDrawer = DrawChartSlideAdmin(self.presentation, "{} Admin {}".format(locationName, aFloor), self.slideLayout)
                managerList = list(managerList)
                managerList.sort()
                for aManager in managerList:
                    directReports = []
                    directReports.extend(self._getDirects(aManager, aLocation))
                    if not directReports:
                        continue
                    self.buildGroup(aManager.getPreferredName(), directReports, chartDrawer)
                    managersOnSlide += 1

                    # Split the slide into multiple parts if it's getting too crowded
                    if managersOnSlide >= maxManagersPerSlide:
                        managersOnSlide = 0
                        chartDrawer.drawSlide()
                        chartDrawer = DrawChartSlideAdmin(self.presentation, "{} Admin {} {}".format(locationName, aFloor, slideNameAddendum), self.slideLayout)
                        slideNameAddendum = "pt3"

                # Keep track of whether this floor has any people so that we avoid spamming "WARNING" messages because
                # a slide is being drawn that's empty
                if managersOnSlide > 0:
                    chartDrawer.drawSlide()
                managersOnSlide = 0

            emptyManagerPeople = (PeopleFilter()
                                  .addManagerEmptyFilter()
                                  .addIsTBHFilter(False)
                                  .addLocationFilter(locationName)
                                    # Someone's name might be entered in the spreadsheet so that their direct reports are drawn but the person
                                    # could be assigned to a different org so their other information is blank. In this case, they aren't
                                    # really missing a manager
                                  .addIsManagerFilter(False))

            peopleMissingManager = (self.multiOrgParser.getFilteredPeople(emptyManagerPeople))

            if peopleMissingManager:
                #Draw people who are missing a manager on their own slide
                chartDrawer = DrawChartSlideAdmin(self.presentation, "{} Missing Admin Manager".format(locationName), self.slideLayout)

                self.buildGroup("Chuck Norris", peopleMissingManager, chartDrawer)
                chartDrawer.drawSlide()


    def drawExpat(self):
        expats = self.multiOrgParser.getFilteredPeople(PeopleFilter().addIsExpatFilter())
                    #.addIsProductManagerFilter(False))
        self._drawMiscGroups("ExPat", expats)

    def drawProductManger(self):
        productManagers = self.multiOrgParser.getFilteredPeople(PeopleFilter().addIsProductManagerFilter())
        if not len(productManagers):
            return
        chartDrawer = DrawChartSlidePM(self.presentation, "Product Management", self.slideLayout)

        productBuckets = list(set([aPerson.getFeatureTeam() for aPerson in productManagers]))

        for aBucket in productBuckets:
            peopleList = [aPerson for aPerson in productManagers if aPerson.getFeatureTeam() == aBucket]
            # peopleList = sorted(peopleList, key=lambda x: x.getProduct(), cmp=self._sortByProduct)
            peopleList.sort()

            # Set default name for PM who don't have 'feature team' set
            # If default name isn't set, these people are accidentally filtered out in the
            # buildGroup function
            if not aBucket:
                aBucket = orgchart_parser.NOT_SET
            self.buildGroup(aBucket, peopleList, chartDrawer)

        # peopleProducts = list(set([aPerson.getProduct() for aPerson in productManagers]))
        # for aProduct in peopleProducts:
        #     self.buildGroup(aProduct, [aPerson for aPerson in productManagers if aPerson.getProduct() == aProduct], chartDrawer)
        chartDrawer.drawSlide()

    def drawIntern(self):
        interns = self.multiOrgParser.getFilteredPeople(PeopleFilter().addIsInternFilter())
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
        crossFuncPeople = self.multiOrgParser.getCrossFuncPeople()

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

            funcPeople = self.multiOrgParser.getFilteredPeople(peopleFilter)
            self.buildGroup(aFunction, funcPeople, chartDrawer)
        chartDrawer.drawSlide()

    def drawAllProducts(self, drawFeatureTeams, drawLocations, drawExpatsInTeam):
        #Get all the products except the ones where a PM is the only member
        people = self.multiOrgParser.getFilteredPeople(PeopleFilter().addIsProductManagerFilter(False))

        productList = list(set([aPerson.getProduct() for aPerson in people]))

        for aCrossFuncTeam in self.multiOrgParser.getCrossFuncTeams():
            if aCrossFuncTeam in productList:
                productList.remove(aCrossFuncTeam)
            productList.sort(cmp=self._sortByProduct)

        for aProductName in productList:
            self.drawProduct(aProductName, drawFeatureTeams, drawLocations, drawExpatsInTeam)

    def drawProduct(self, productName, drawFeatureTeams=False, drawLocations=False, drawExpatsInTeam=True):
        """

        :type productName: str
        """
        if not productName:
            if not self.draftMode:
                return

        featureTeamList = [""]
        if drawFeatureTeams:
            featureTeamList = list(self.multiOrgParser.getFeatureTeamSet(productName))

        functionList = list(self.multiOrgParser.getFunctionSet(productName))
        functionList.sort(cmp=self._sortByFunc)

        teamModelText = None

        locations = [""]
        if drawLocations:
            locations = self.multiOrgParser.getLocationSet(productName)

        for aLocation in locations:
            locationName = aLocation.strip() or self.multiOrgParser.getOrgName()

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
                    modelDict = self.multiOrgParser.getTeamModel()
                    if productName in modelDict:
                        teamModelText = modelDict[productName]

                chartDrawer = DrawChartSlide(self.presentation, slideTitle, self.slideLayout, teamModelText)
                if len(locations) > 1 and aLocation:
                   chartDrawer.setLocation(locationName)

                for aFunction in functionList:
                    if aFunction.lower() in self.multiOrgParser.getCrossFunctions():
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

                    functionPeople = self.multiOrgParser.getFilteredPeople(peopleFilter)
                    self.buildGroup(aFunction, functionPeople, chartDrawer)

                chartDrawer.drawSlide()

    def drawTBH(self):

        totalTBHSet = set(self.multiOrgParser.getFilteredPeople(PeopleFilter().addIsTBHFilter()))
        tbhLocations = set([aTBH.getLocation() for aTBH in totalTBHSet]) or [""]
        tbhProducts = list(set([aTBH.getProduct() for aTBH in totalTBHSet])) or [""]
        tbhProducts.sort(cmp=self._sortByProduct)
        tbhFunctions = list(set([aTBH.getFunction() for aTBH in totalTBHSet]))
        tbhFunctions.sort(cmp=self._sortByFunc)


        for aLocation in tbhLocations:
            title = "Hiring"

            # Location might not be set.
            if aLocation:
                title = "{} - {}".format(title, aLocation)

            chartDrawer = DrawChartSlideTBH(self.presentation, title, self.slideLayout)

            for aFunction in tbhFunctions:
                productTBHList = self.multiOrgParser.getFilteredPeople(PeopleFilter().addIsTBHFilter().addFunctionFilter(aFunction).addLocationFilter(aLocation))
                productTBHList = sorted(productTBHList, self._sortByFunc, lambda tbh: tbh.getProduct())
                self.buildGroup(aFunction, productTBHList, chartDrawer)

            chartDrawer.drawSlide()

            # for aProduct in tbhProducts:
            #     productTBHList = self.orgParser.getFilteredPeople(PeopleFilter().addIsTBHFilter().addProductFilter(aProduct).addLocationFilter(aLocation))
            #     productTBHList = sorted(productTBHList, self._sortByFunc, lambda tbh: tbh.getProduct())
            #     self.buildGroup(aProduct, productTBHList, chartDrawer)
            #
            # chartDrawer.drawSlide()

    def _sortByFunc(self, a, b):
        funcOrder = ["lead", "leadership", "head coach", "product management", "pm", "po", "product owner", "product owner/qa", "technology", "ta", "technology architect", "tech", "sw architecture", "dev",
                     "development", "development (connectors)", "qa", "quality assurance", "stress",
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
        floorOrder = self.multiOrgParser.getFloorSortOrder()

        if a.lower() in floorOrder:
            if b.lower() in floorOrder:
                if floorOrder.index(a.lower()) > floorOrder.index(b.lower()):
                    return 1
            return -1

        if b.lower() in floorOrder:
            return 1

        return 0

    def _sortByProduct(self, a, b):
        productOrder = self.multiOrgParser.getProductSortOrder()

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

    parser.add_argument("path", nargs="+",
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



    workbooks = []
    for aPath in options.path:
        if os.path.exists(aPath):
            workbooks.append(aPath)
        else:
            fileMatch = glob.glob(os.path.join(options.directory, "*{}*".format(aPath)))
            fileMatch = [aFile for aFile in fileMatch if
                         (not ((os.path.basename(aFile).startswith("~")) or "conflict" in os.path.basename(aFile).lower()))]
            if not fileMatch:
                raise OSError("Could not find any files in directory: '{}' that contain string: '{}'"
                              .format(options.directory, aPath))
            if len(fileMatch) > 1:
                raise OSError(
                    "Too many files found in dir: '{}' that contain string '{}' : \n\t\t{}".format(options.directory,
                                                                                                   aPath,
                                                                                                   "\n\t\t".join(
                                                                                                       fileMatch)))
            workbooks.append(fileMatch[0])

    orgDraw = OrgDraw(workbooks, options.sheetName, options.draftMode)

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
        outputFileName = "{}{}".format(orgDraw.multiOrgParser.getOrgName(), defaultOutputFile)
    outputFileName = outputFileName.strip()
    modifiedName = outputFileName

    for i in xrange(0, 10):
        try:
            orgDraw.save(modifiedName)
            print "{} SAVED".format(modifiedName)
            break
        except Exception, e:
            print "{} Save FAILED. Is it open already? Retrying".format(modifiedName)
            outputFileNameParts = outputFileName.split(".")
            outputFileNameParts.insert(-1, str(i))
            modifiedName = ".".join(outputFileNameParts)
    return modifiedName



if __name__ == "__main__":
    # for davep:
    #sDir = '/Users/dpinkney/Documents/HCPAnywhere/SharedWithMe/Waltham Engineering Org Charts/'
    #sys.argv += ['-t', '-d', sDir, '-o%s/WalthamChartGen.pptx' % sDir, 'WalthamStaff.xlsm']
    #sys.argv += ['-d', sDir, '-o%s/HPP_Charts.pptx' % sDir, 'HPP_Staff.xlsm']
    #sDir = '/Users/dpinkney/Documents/HCPAnywhere/SharedWithMe/Waltham Engineering Org Charts/tmp/'
    #sys.argv += ['-d', sDir, '-o%s/WalthamChartGen-moves.pptx' % sDir, 'WalthamStaff-moves.xlsm']
    # test against SC
    #sDir = '/Users/dpinkney/Documents/HCPAnywhere/SharedWithMe/Waltham Engineering Org Charts/Old/Santa Clara/'
    #sys.argv += ['-t', '-d', sDir, '-o%s/SantaClaraChartGen.pptx' % sDir, 'SantaClara Staff_04_28_2016.xlsm']
    main(sys.argv[1:])


class GenChartCommandline(TestCase):


    def testSantaClaraFeatures(self):
        outputFileName = "{}{}SantaClaraOrgChart.feature.pptx".format(os.getcwd(),os.sep)
        main(['Z:\Documents\HCP Anywhere\Org Charts\Insight Group\SibuEngStaff.xlsm', "-f", "-o {}".format(outputFileName)])
        os.system("start " + outputFileName)

    def testConverged(self):
        todayDate = datetime.date.today().strftime("%Y-%m-%d")
        outputFileName = "{cwd}{slash}{dateStamp}_Converged_OrgChart.pptx".format(cwd=os.getcwd(), slash=os.sep, dateStamp=todayDate)
        main(['Z:\doreper On My Mac\Documents\HCP Anywhere\Org Charts\Converged\ConvergedEngStaff.xlsm', "-t","-e", "-o {}".format(outputFileName)])
        startCmd = 'start {}'.format(outputFileName)
        os.system(startCmd)

    def testContent(self):
        todayDate = datetime.date.today().strftime("%Y-%m-%d")
        outputFileName = "{cwd}{slash}{dateStamp}_ContentOrgChart.pptx".format(cwd=os.getcwd(), slash=os.sep, dateStamp=todayDate)
        outputFileName = main(['Z:\Documents\HCP Anywhere\Org Charts\Content\ContentStaff.xlsm',
                               "-e", "-t", "-o {}".format(outputFileName)])
        os.system("start " + outputFileName)

    def testInsightContent(self):
        todayDate = datetime.date.today().strftime("%Y-%m-%d")
        outputFileName = "{cwd}{slash}{dateStamp}_Insight_ContentOrgChart.pptx".format(cwd=os.getcwd(), slash=os.sep, dateStamp=todayDate)
        outputFileName = main(['Z:\Documents\HCP Anywhere\Org Charts\Insight Group\SibuEngStaff.xlsm',
                               'Z:\Documents\HCP Anywhere\Org Charts\Content\ContentStaff.xlsm',
                               "-e", "-t", "-o {}".format(outputFileName)])
        os.system("start " + outputFileName)


    def testInsight(self):
        todayDate = datetime.date.today().strftime("%Y-%m-%d")
        outputFileName = "{cwd}{slash}{dateStamp}_InsightOrgChart.pptx".format(cwd=os.getcwd(), slash=os.sep, dateStamp=todayDate)
        outputFileName = main(['Z:\doreper On My Mac\Documents\HCP Anywhere\Org Charts\Insight Group\SIBUEngStaff.xlsm',
                               "-e", "-t", "-o {}".format(outputFileName)])
        os.system("start " + outputFileName)

    def testALLOrg(self):
        todayDate = datetime.date.today().strftime("%Y-%m-%d")
        outputFileName = "{cwd}{slash}{dateStamp}ALLOrgChart.pptx".format(cwd=os.getcwd(), slash=os.sep, dateStamp=todayDate)
        outputFileName = main(['Z:\doreper On My Mac\Documents\HCP Anywhere\Org Charts\Insight Group\SIBUEngStaff.xlsm',
                               'Z:\doreper On My Mac\Documents\HCP Anywhere\Org Charts\Converged\ConvergedEngStaff.xlsm',
                               'Z:\Documents\HCP Anywhere\Org Charts\Content\ContentStaff.xlsm',
                               "-e", "-t", "-o {}".format(outputFileName)])
        os.system("start " + outputFileName)

    def testInsightConverged(self):
        todayDate = datetime.date.today().strftime("%Y-%m-%d")
        outputFileName = "{cwd}{slash}{dateStamp}_Insight_Converged_OrgChart.pptx".format(cwd=os.getcwd(), slash=os.sep, dateStamp=todayDate)
        outputFileName = main(['Z:\doreper On My Mac\Documents\HCP Anywhere\Org Charts\Insight Group\SIBUEngStaff.xlsm',
                               'Z:\doreper On My Mac\Documents\HCP Anywhere\Org Charts\Converged\ConvergedEngStaff.xlsm',
                               "-e", "-t", "-o {}".format(outputFileName)])
        os.system("start " + outputFileName)
