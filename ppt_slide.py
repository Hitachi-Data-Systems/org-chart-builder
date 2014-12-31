import math
import datetime
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
import re
from shape import RectangleBuilder, ColorPicker

__author__ = 'David Oreper'


class DrawChartSlide:
    MAX_WIDTH_INCHES = Inches(10)
    MAX_HEIGHT_INCHES = Inches(5.63)
    PAGE_BUFFER = Inches(.2)
    RIGHT_EDGE = MAX_WIDTH_INCHES - PAGE_BUFFER

    def __init__(self, aPresentation, slideTitle, slideLayout):
        """

        :type aPresentation: pptx.api.Presentation
        :type slideTitle: str
        :type slideLayout: pptx.parts.slidelayout.SlideLayout
        """
        self.presentation = aPresentation
        self.slideTitle = slideTitle
        self.slideLayout = slideLayout
        self.colorPicker = ColorPicker()
        self.groupList = []

        self.groupLeft = Inches(.2)
        self.groupTop = GroupShapeDimensions.TOP

    def addGroup(self, title, groupMembers):
        peopleGroup = PeopleGroup(title, self.groupLeft, GroupShapeDimensions.TOP, GroupShapeDimensions.HEIGHT)
        peopleGroup.setGroupColor(self.colorPicker.getBackgroundColor())
        peopleGroup.setMemberColor(self.colorPicker.getForegroundColor())
        self.colorPicker.nextColor()

        peopleGroup.addMembers(groupMembers)

        self.groupLeft = peopleGroup.getRightEdge() + GroupShapeDimensions.BUFFER_WIDTH
        self.groupList.append(peopleGroup)

    def _adjustGroupWidth(self, reductionRatio):
        print("Reducing width by {}% for: {}".format(100 - int(reductionRatio*100), self.slideTitle))
        rightEdge = Inches(.2)
        for aGroup in self.groupList:
            aGroup.adjustWidth(reductionRatio)
            aGroup.setLeft(rightEdge)
            rightEdge = aGroup.getRightEdge() + GroupShapeDimensions.BUFFER_WIDTH * reductionRatio

    def _center(self):
        rightEdge = self.groupList[-1].getRightEdge()
        leftAdjust = (DrawChartSlide.RIGHT_EDGE - rightEdge) / 2
        for aGroup in self.groupList:
            aGroup.setLeft(aGroup.getLeft() + leftAdjust)

    def addTitle(self, aSlide):

        aSlide.shapes.title.text = self.slideTitle

        effectiveDateRun = aSlide.shapes.title.textframe.paragraphs[0].add_run()

        # A little ugly but does the trick - add suffix to the number: 2 -> 2nd
        # http://stackoverflow.com/a/20007730
        suffixFormatter = lambda n: "%d%s" % (n,"tsnrhtdd"[(n/10%10!=1)*(n%10<4)*n%10::4])
        ordinalDay = suffixFormatter(int(datetime.datetime.now().strftime("%d").lstrip("0")))

        month = datetime.datetime.now().strftime("%B")
        effectiveDateRun.text = "\rEffective {} {}".format(month, ordinalDay)
        effectiveDateRun.font.size = Pt(11)
        effectiveDateRun.font.italic = True
        effectiveDateRun.font.bold = False


    def drawSlide(self):
        if not self.groupList:
            print "WARNING: NO Groups added for product: {}".format(self.slideTitle)
            return

        slide = self.presentation.slides.add_slide(self.slideLayout)
        self.addTitle(slide)

        rightEdge = self.groupList[-1].getRightEdge()
        if rightEdge > DrawChartSlide.RIGHT_EDGE:
            reductionRatio = 1 / (float(rightEdge) / DrawChartSlide.RIGHT_EDGE)
            self._adjustGroupWidth(reductionRatio)

        self._center()

        for aGroup in self.groupList:
            aGroup.build(slide)


class GroupShapeDimensions:
    def __init__(self):
        pass

    TOP = Inches(1.1)
    HEIGHT = Inches(4.4)
    WIDTH = Inches(1.17)
    BUFFER_WIDTH = Inches(.05)
    BUFFER_HEIGHT = Inches(.4)


class MemberShapeDimensions:
    def __init__(self):
        pass

    HEIGHT = Inches(.45)
    BUFFER_WIDTH = Inches(.03)
    BUFFER_HEIGHT = Inches(.03)
    WIDTH = GroupShapeDimensions.WIDTH - (BUFFER_WIDTH * 2)
    HARD_WRAP_NUM = (GroupShapeDimensions.HEIGHT - GroupShapeDimensions.BUFFER_HEIGHT) / (HEIGHT + BUFFER_HEIGHT)

class PeopleGroup(object):
    def __init__(self, title, left, top, height):
        """

        :type title: str
        :type left: Inches or float
        :type top: pptx.util.Inches
        :type height: pptx.util.Inches
        """
        self.name = title
        self.title = title
        self.groupWidthUnits = 1
        self.groupUnitWidth = GroupShapeDimensions.WIDTH
        self.groupTop = top
        self.groupLeft = left
        self.groupHeight = height

        self.memberLeft = self.groupLeft + MemberShapeDimensions.BUFFER_WIDTH
        self.memberTop = self.groupTop + GroupShapeDimensions.BUFFER_HEIGHT
        self.memberShapeList = []
        self.fontReduce = 1

    def _nextColumn(self):
        self.memberTop = self.groupTop + GroupShapeDimensions.BUFFER_HEIGHT
        self.memberLeft = self.memberLeft + MemberShapeDimensions.WIDTH + MemberShapeDimensions.BUFFER_WIDTH * 2

    def setGroupColor(self, groupColorRGB):
        """

        :type groupColorRGB: pptx.dml.color.RGBColor
        """
        self.groupColorRGB = groupColorRGB

    def setMemberColor(self, memberColorRGB):
        """

        :type memberColorRGB: pptx.dml.color.RGBColor
        """
        self.memberColor = memberColorRGB

    def setHeight(self, heightInches):
        """

        :param heightInches:
        """
        self.groupHeight = heightInches

    def getLeft(self):
        return self.groupLeft

    def setLeft(self, leftCoord):
        """

        :type leftCoord: float or Inches
        """
        self.groupLeft = leftCoord

        leftStart = self.memberShapeList[0].getLeft()
        memberLeft = self.groupLeft + MemberShapeDimensions.BUFFER_WIDTH
        for aMember in self.memberShapeList:
            if (aMember.getLeft() - leftStart) > 0:
                leftStart = aMember.getLeft()
                memberLeft += aMember.getWidth() + MemberShapeDimensions.BUFFER_WIDTH

            aMember.setLeft(memberLeft)

    def adjustWidth(self, reductionRatio):
        """

        :type reductionRatio: float
        """
        self.groupUnitWidth = self.groupUnitWidth * reductionRatio
        self.fontReduce = reductionRatio + ((1 - reductionRatio) / 2)
        for aMember in self.memberShapeList:
            aMember.adjustWidth(reductionRatio)
            aMember.adjustFontSizes(self.fontReduce)

    def getWidth(self):
        return self.groupWidthUnits * self.groupUnitWidth

    def getRightEdge(self):
        return self.groupLeft + self.getWidth()

    def addMembers(self, peopleList):
        count = 1

        # Calculate the total width of the group so that we can distribute columns evenly
        self.groupWidthUnits = max(math.ceil(len(peopleList) / float(MemberShapeDimensions.HARD_WRAP_NUM)), 1)

        # Calculate how many people will be in each column
        wrapCount = math.ceil(len(peopleList) / float(self.groupWidthUnits))
        for aPerson in peopleList:
            self.addMember(aPerson)
            count += 1
            if count > wrapCount:
                self._nextColumn()
                count = 1

    def addMember(self, aPerson):
        aPersonRect = RectangleBuilder(self.memberLeft, self.memberTop, MemberShapeDimensions.WIDTH,
                                       MemberShapeDimensions.HEIGHT)

        aPersonRect.setFirstName(aPerson.getFirstName())
        aPersonRect.setLastName(aPerson.getLastName())
        aPersonRect.setBrightness(0)
        aPersonRect.setRGBFillColor(self.memberColor)
        aPersonRect.setRGBTextColor(RGBColor(255, 255, 255))
        aPersonRect.setRGBFirstNameColor(RGBColor(255, 255, 255))

        if aPerson.getRawName().startswith("TBH") or aPerson.getRawName().startswith("TBD"):
            if not re.search('\d', aPerson.getRawName()):
                aPersonRect.setTitle(aPerson.getReqNumber())
        else:
            if aPerson.isExpat() or aPerson.isIntern():
                aPersonRect.setTitle(aPerson.getProduct())
            else:
                aPersonRect.setTitle(aPerson.getTitle())

            if aPerson.isConsultant():
                aPersonRect.setTitle(aPerson.getTitle() + " (c)")

            if aPerson.isManager():
                aPersonRect.setRGBFirstNameColor(RGBColor(255, 238, 0))

        self.memberShapeList.append(aPersonRect)
        self.memberTop += MemberShapeDimensions.HEIGHT + MemberShapeDimensions.BUFFER_HEIGHT

    def build(self, aSlide):
        groupRect = RectangleBuilder(self.groupLeft, self.groupTop, self.getWidth(), self.groupHeight)
        groupRect.setRGBFillColor(self.groupColorRGB)
        groupRect.setRGBTextColor(RGBColor(0, 0, 0))
        groupRect.setHeading(self.title)
        groupRect.setBrightness(.2)
        groupRect.adjustFontSizes(self.fontReduce)
        groupRect.build(aSlide)
        for aMember in self.memberShapeList:
            aMember.build(aSlide)


