# Copyright (c) 2010-2020 openpyxlzip

from openpyxlzip.xml.constants import CHART_NS

from openpyxlzip.descriptors.serialisable import Serialisable
from openpyxlzip.descriptors.excel import Relation


class ChartRelation(Serialisable):

    tagname = "chart"
    namespace = CHART_NS

    id = Relation()

    def __init__(self, id):
        self.id = id
