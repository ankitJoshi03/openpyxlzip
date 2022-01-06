# Copyright (c) 2010-2020 openpyxlzip

from openpyxlzip.descriptors.serialisable import Serialisable
from openpyxlzip.descriptors import (
    Sequence
)
from openpyxlzip.descriptors.excel import (
    Relation,
)

class ExternalReference(Serialisable):

    tagname = "externalReference"

    id = Relation()

    def __init__(self, id):
        self.id = id
