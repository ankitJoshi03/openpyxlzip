# Copyright (c) 2010-2020 openpyxlzip

from openpyxlzip.descriptors import (
    Bool,
    Integer,
    Typed,
    Sequence
)
from openpyxlzip.descriptors.excel import ExtensionList
from openpyxlzip.descriptors.serialisable import Serialisable


class ChartsheetView(Serialisable):
    tagname = "sheetView"

    tabSelected = Bool(allow_none=True)
    zoomScale = Integer(allow_none=True)
    workbookViewId = Integer()
    zoomToFit = Bool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('extLst',)

    def __init__(self,
                 tabSelected=None,
                 zoomScale=None,
                 workbookViewId=0,
                 zoomToFit=None,
                 extLst=None,
                 ):
        self.tabSelected = tabSelected
        self.zoomScale = zoomScale
        self.workbookViewId = workbookViewId
        self.zoomToFit = zoomToFit
        self.extLst = extLst


class ChartsheetViewList(Serialisable):
    tagname = "sheetViews"

    sheetView = Sequence(expected_type=ChartsheetView, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('sheetView', 'extLst',)

    def __init__(self,
                 sheetView=None,
                 extLst=None,
                 ):
        if sheetView is None:
            sheetView = [ChartsheetView()]
        self.sheetView = sheetView
        self.extLst = extLst
