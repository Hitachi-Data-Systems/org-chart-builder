#!/usr/bin/env python

import argparse
import glob
import os
import math
from pptx.dml.color import RGBColor
import re
from orgchart_parser import OrgParser
import sys
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_ANCHOR

__author__ = 'David Oreper'

HARD_WRAP_NUM = 7


class ColorPicker:
    def __init__(self):
        self.index = 0
        self.colors = [(RGBColor(245, 227, 226), RGBColor(166, 75, 71)),
                       (RGBColor(209, 225, 243), RGBColor(42, 94, 142)),
                       (RGBColor(239, 243, 229), RGBColor(138, 160, 81)),
                       (RGBColor(227, 235, 244), RGBColor(72, 117, 161)),
                       (RGBColor(235, 231, 240), RGBColor(116, 96, 140)),
                       (RGBColor(226, 241, 246), RGBColor(65, 151, 171)),
                       (RGBColor(253, 238, 226), RGBColor(235, 128, 35))]

    def getBackgroundColor(self):
        colorIndex = self.index % len(self.colors)
        return self.colors[colorIndex][0]

    def getForegroundColor(self):
        colorIndex = self.index % len(self.colors)
        return self.colors[colorIndex][1]

    def nextColor(self):
        self.index += 1


class ShapeBuffer:
    def __init__(self):
        pass

    FOREGROUND_WIDTH = Inches(.2)
    BACKGROUND_WIDTH = Inches(.05)
    HEIGHT = Inches(.05)


class BackgroundShape:
    def __init__(self):
        pass

    TOP = Inches(1.2)
    WIDTH = Inches(1.17)
    HEIGHT = Inches(4.3)


class ForegroundShape:
    def __init__(self):
        pass

    TOP = BackgroundShape.TOP + ShapeBuffer.HEIGHT + Inches(.4)
    WIDTH = BackgroundShape.WIDTH - (ShapeBuffer.FOREGROUND_WIDTH / 2)
    HEIGHT = Inches(.48)


class RectBuilder(object):
    """
    this class does stuff
    """

    def __init__(self, slide, left, top, width, height, rgbFillColor):
        """

        :param left:
        :param top:
        :param width:
        :param height:
        :param rgbFillColor:
        """
        self.slide = slide
        self.width = width
        self.height = height
        self.top = top
        self.left = left
        self.rgbFillColor = rgbFillColor
        self.rgbTextColor = RGBColor(255, 255, 255)
        self.rgbFirstNameColor = self.rgbTextColor
        self.brightness = 0
        self.firstName = None
        self.lastName = None
        self.title = None
        self.consultant = None
        self.expat = None
        self.heading = None

    def setRGBTextColor(self, rgbColor):
        """

        :type rgbColor: pptx.dml.color.RGBColor
        """
        self.rgbTextColor = rgbColor

    def setRGBFirstNameColor(self, rgbColor):
        self.rgbFirstNameColor = rgbColor

    def setFirstName(self, firstName):
        """

        :type firstName: str
        """
        self.firstName = firstName

    def setLastName(self, lastName):
        """

        :type lastName: list or str
        """
        self.lastName = lastName

    def setTitle(self, title):
        """

        :type title: str
        """
        self.title = title

    def setHeading(self, heading):
        """

        :type heading: str
        """
        self.heading = heading

    def setBrightness(self, brightness):
        """

        :type brightness: float or int
        """
        self.brightness = brightness

    def _buildShape(self):
        shape = self.slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.left, self.top, self.width, self.height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.rgbFillColor
        shape.fill.fore_color.brightness = self.brightness
        shape.line.color.rgb = self.rgbFillColor
        return shape

    def addRun(self, textFrame, text, size, isBold, isItalic, rgbTextColor=None):
        """

        :param textFrame:
        :type text: str
        :type size: int
        :type isBold: bool
        :type isItalic: bool
        :type rgbTextColor: __builtin__.NoneType
        """
        paragraph = textFrame.paragraphs[0]
        aRun = paragraph.add_run()

        if not rgbTextColor:
            rgbTextColor = self.rgbTextColor

        text = "{}\r".format(text.strip())
        aRun.text = text

        aRun.font.size = Pt(size)
        aRun.font.bold = isBold
        aRun.font.italic = isItalic
        aRun.font.color.rgb = rgbTextColor
        aRun.alignment = PP_ALIGN.CENTER
        aRun.font.name = "Arial (Body)"

    def _buildFirstName(self, textFrame):
        self.addRun(textFrame, self.firstName, 9, True, False, self.rgbFirstNameColor)

    def _buildLastName(self, textFrame):
        self.addRun(textFrame, self.lastName, 7, False, False)

    def _buildTitle(self, textFrame):
        title = self.title
        self.addRun(textFrame, title, 5, False, True)

    def _buildHeading(self, textFrame):
        self.addRun(textFrame, self.heading, 8, True, False)

    def build(self):
        shape = self._buildShape()
        shape.textframe.vertical_anchor = MSO_ANCHOR.TOP

        if self.heading:
            self._buildHeading(shape.textframe)

        if self.firstName:
            self._buildFirstName(shape.textframe)

        if self.lastName:
            self._buildLastName(shape.textframe)

        if self.title:
            self._buildTitle(shape.textframe)


class DrawChartSlide:
    def __init__(self, aPresentation, slideTitle, titleSlide):
        """

        :type aPresentation: pptx.api.Presentation
        :type slideTitle: str
        :type titleSlide: pptx.parts.slidelayout.SlideLayout
        """
        self.slide = aPresentation.slides.add_slide(titleSlide)
        shapes = self.slide.shapes
        shapes.title.text = slideTitle
        self.colorPicker = ColorPicker()

        self.backgroundShapeLeft = Inches(.1)
        self.foregroundShapeTop = ForegroundShape.TOP
        self.foregroundShapeLeft = self.backgroundShapeLeft + ShapeBuffer.FOREGROUND_WIDTH
        self.foregroundColor = self.colorPicker.getForegroundColor()
        self.backgroundColor = self.colorPicker.getBackgroundColor()


    def addBackgroundShape(self, shapeName, width=1):
        self.foregroundColor = self.colorPicker.getForegroundColor()
        self.backgroundColor = self.colorPicker.getBackgroundColor()
        self.foregroundShapeLeft = self.backgroundShapeLeft + ShapeBuffer.BACKGROUND_WIDTH
        newRect = RectBuilder(self.slide, self.backgroundShapeLeft, BackgroundShape.TOP,
                              BackgroundShape.WIDTH * width, BackgroundShape.HEIGHT, self.backgroundColor)

        newRect.setHeading(shapeName)
        newRect.setRGBTextColor(RGBColor(0, 0, 0))
        newRect.setBrightness(.2)
        newRect.build()
        self.backgroundShapeLeft += (BackgroundShape.WIDTH * width) + ShapeBuffer.BACKGROUND_WIDTH
        self.foregroundShapeTop = ForegroundShape.TOP
        self.colorPicker.nextColor()

    def nextColumn(self):
        self.foregroundShapeTop = ForegroundShape.TOP
        self.foregroundShapeLeft = self.foregroundShapeLeft + ForegroundShape.WIDTH + ShapeBuffer.BACKGROUND_WIDTH

    def addForegroundShape(self, aPerson):
        aPersonRect = RectBuilder(self.slide, self.foregroundShapeLeft, self.foregroundShapeTop,
                                  ForegroundShape.WIDTH, ForegroundShape.HEIGHT, self.foregroundColor)

        aPersonRect.setFirstName(aPerson.getFirstName())
        aPersonRect.setLastName(aPerson.getLastName())
        aPersonRect.setTitle(aPerson.getTitle())
        aPersonRect.setBrightness(0)

        if aPerson.getRawName().startswith("TBH") or aPerson.getRawName().startswith("TBH"):
            if not re.search('\d', aPerson.getRawName()):
                aPersonRect.setLastName(aPerson.getReqNumber())

        if aPerson.isExpat():
            aPersonRect.setTitle(aPerson.getProduct())

        if aPerson.isConsultant():
            aPersonRect.setTitle(aPerson.getTitle() + " (c)")

        if aPerson.isManager():
            aPersonRect.setRGBFirstNameColor(RGBColor(255, 238, 0))

        aPersonRect.build()
        self.foregroundShapeTop += ForegroundShape.HEIGHT + ShapeBuffer.HEIGHT

MAX_WIDTH_INCHES = Inches(10)
MAX_HEIGHT_INCHES = Inches(5.63)

class OrgDraw:
    def __init__(self, workbookPath, sheetName):
        """

        :type workbookPath: str
        :type sheetName: str
        """
        self.presentation = Presentation()
        self.presentation.slide_height = MAX_HEIGHT_INCHES
        self.presentation.slide_width = MAX_WIDTH_INCHES
        self.slideLayout = self.presentation.slide_layouts[5]
        self.orgParser = OrgParser(workbookPath, sheetName)

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
        directReports = self.orgParser.getDirectReports(aManagerName)
        directReports.sort()
        directReports = [person for person in directReports if not person.getRawName().strip().startswith("TBH")]
        directReports = [person for person in directReports if not person.getRawName().strip().startswith("TBD")]
        return directReports

    def drawAdmin(self):
        managersLeft = self.orgParser.getManagerSet()
        completedManagers = set()

        for aFloor in self.orgParser.peopleDataKeys.FLOORS.keys():
            chartDrawer = DrawChartSlide(self.presentation, "Admin {}".format(aFloor), self.slideLayout)
            managerList = self.orgParser.peopleDataKeys.FLOORS[aFloor]
            managerList.sort(cmp=self._sortManagers)
            for aManagerName in managerList:
                directReports = []
                directReports.extend(self._getDirects(aManagerName))
                completedManagers.add(aManagerName)

                aManagerName = self.getFirstLastName(aManagerName)
                if not aManagerName:
                    aManagerName = "!!NOT SET!!"
                self.drawFunction(aManagerName, directReports, chartDrawer)

        managersLeft = list(set(managersLeft) - completedManagers)
        managersLeft.sort(cmp=self._sortManagers)

        if len(managersLeft):
            chartDrawer = DrawChartSlide(self.presentation, "Admin", self.slideLayout)
            for aManagerName in managersLeft:
                directReports = self._getDirects(aManagerName)
                if not len(directReports):
                    continue
                aManagerName = self.getFirstLastName(aManagerName)
                if not aManagerName:
                    aManagerName = "!!NOT SET!!"
                self.drawFunction(aManagerName, directReports, chartDrawer)

    def drawExpat(self):
        chartDrawer = DrawChartSlide(self.presentation, "EXPAT", self.slideLayout)
        expats = self.orgParser.getFilteredPeople(isExpat=True)
        expatFunctions = set([expat.getFunction() for expat in expats])

        for aFunction in expatFunctions:
            self.drawFunction(aFunction, [expat for expat in expats if expat.getFunction() == aFunction], chartDrawer)

    def drawAllProducts(self):
        productList = list(self.orgParser.getProductSet())
        productList.sort()

        for aProductName in productList:
            slideTitle = aProductName
            if not slideTitle:
                slideTitle = "!!Product not set!!"
            chartDrawer = DrawChartSlide(self.presentation, slideTitle, self.slideLayout)
            self.drawProduct(aProductName, chartDrawer)

    def getSortedFuncList(self):
        return ["Lead", "Head Coach", "PO", "Product Owner", "Technology", "TA", "Tech", "SW Architecture", "Dev",
                "Development", "QA", "Quality Assurance", "Stress",
                "Characterization", "Auto", "Aut", "Automation", "Sustaining", "Solutions and Sustaining",
                "UI", "UX", "UI/UX", "Inf", "Infrastructure", "DevOps", "Cross Functional", "Cross", "Doc",
                "Documentation"]

    def drawProduct(self, productName, chartDrawer):
        """

        :type productName: str
        :type chartDrawer: ppt_draw.DrawChartSlide
        """
        functionList = self.orgParser.getFunctionSet(productName)
        for aFunction in self.getSortedFuncList():
            if aFunction in functionList:
                functionPeople = self.orgParser.getFilteredPeople(productName, aFunction)
                functionPeople.sort()
                self.drawFunction(aFunction, functionPeople, chartDrawer)
                functionList.remove(aFunction)

        for aFunction in functionList:
            functionPeople = self.orgParser.getFilteredPeople(productName, aFunction)
            functionPeople = [person for person in functionPeople if not person.isExpat()]
            functionPeople.sort()
            self.drawFunction(aFunction, functionPeople, chartDrawer)

    def drawFunction(self, functionName, functionPeople, chartDrawer):
        """

        :type functionName: str
        :type functionPeople: list
        :type chartDrawer: ppt_draw.DrawChartSlide
        :return:
        """
        functionPeople = [person for person in functionPeople if person.getRawName().strip() != ""]
        if len(functionPeople) == 0:
            return
        backgroundShapeWidth = max(math.ceil(len(functionPeople) / float(HARD_WRAP_NUM)), 1)
        if not functionName:
            functionName = "!!NOT SET!!"
        chartDrawer.addBackgroundShape(functionName, backgroundShapeWidth)

        count = 1
        wrapCount = math.ceil(len(functionPeople) / float(backgroundShapeWidth))
        for aPerson in functionPeople:
            chartDrawer.addForegroundShape(aPerson)
            count += 1
            if count > wrapCount:
                chartDrawer.nextColumn()
                count = 1


def main(argv):
    userDir = os.environ.get("USERPROFILE") or os.environ.get("HOME")
    defaultSheetName = "PeopleData"
    defaultDir = os.path.join(userDir, "Documents/HCP Anywhere/Org Charts and Hiring History")
    defaultOutputFile = "orgChart.pptx"

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

    parser.add_argument("-o", "--outputFile", type=str, default=defaultOutputFile, help="output file")

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

    orgDraw = OrgDraw(workbookPath, options.sheetName)

    orgDraw.drawAllProducts()
    orgDraw.drawExpat()
    orgDraw.drawAdmin()
    orgDraw.save(options.outputFile)

if __name__ == "__main__":
    # sys.argv = ["", 'C:\Code\OrgChartBuilder\\test\Bellevue Staff.xlsm']
    # sys.argv = ["", 'Z:\Documents\HCP Anywhere\Org Charts and Hiring History\SantaClara Staff.xlsm']
    # sys.argv = ["", 'Z:\Documents\HCP Anywhere\Org Charts and Hiring History\Waltham Staff.xlsm']
    #  sys.argv = ["", 'Z:\Documents\HCP Anywhere\Org Charts and Hiring History\Bellevue Staff.xlsm']
    main(sys.argv[1:])