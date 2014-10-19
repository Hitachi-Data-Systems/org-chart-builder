import math
from pptx.dml.color import RGBColor
from pptx.util import Inches
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

    def drawSlide(self):
        if not self.groupList:
            print "WARNING: NO Groups added for product: {}".format(self.slideTitle)
            return

        self.slide = self.presentation.slides.add_slide(self.slideLayout)
        shapes = self.slide.shapes
        shapes.title.text = self.slideTitle
        rightEdge = self.groupList[-1].getRightEdge()
        if rightEdge > DrawChartSlide.RIGHT_EDGE:
            # TODO? Consolidate groups
            reductionRatio = 1 / (float(rightEdge) / DrawChartSlide.RIGHT_EDGE)
            self._adjustGroupWidth(reductionRatio)

        self._center()

        for aGroup in self.groupList:
            aGroup.build(self.slide)


class GroupShapeDimensions:
    def __init__(self):
        pass

    TOP = Inches(1.2)
    HEIGHT = Inches(4.3)
    WIDTH = Inches(1.17)
    BUFFER_WIDTH = Inches(.05)
    BUFFER_HEIGHT = Inches(.4)


class MemberShapeDimensions:
    def __init__(self):
        pass

    HEIGHT = Inches(.43)
    BUFFER_WIDTH = Inches(.03)
    BUFFER_HEIGHT = Inches(.05)
    WIDTH = GroupShapeDimensions.WIDTH - (BUFFER_WIDTH * 2)
    HARD_WRAP_NUM = GroupShapeDimensions.HEIGHT / (HEIGHT+ BUFFER_HEIGHT)


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

        for aMember in self.memberShapeList:
            aMember.adjustWidth(reductionRatio)
            aMember.adjustFontSizes(reductionRatio + ((1 - reductionRatio) / 2))

    def getWidth(self):
        return self.groupWidthUnits * self.groupUnitWidth

    def getRightEdge(self):
        return self.groupLeft + self.getWidth()

    def addMembers(self, peopleList):
        count = 1
        self.groupWidthUnits = max(math.ceil(len(peopleList) / float(MemberShapeDimensions.HARD_WRAP_NUM)), 1)
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

        if aPerson.getRawName().startswith("TBH") or aPerson.getRawName().startswith("TBD"):
            if not re.search('\d', aPerson.getRawName()):
                aPersonRect.setTitle(aPerson.getReqNumber())
        else:
            if aPerson.isExpat():
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

        groupRect.build(aSlide)
        for aMember in self.memberShapeList:
            aMember.build(aSlide)


