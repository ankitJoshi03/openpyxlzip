# Copyright (c) 2010-2020 openpyxlzip

from openpyxlzip.descriptors import Bool
from openpyxlzip.descriptors.serialisable import Serialisable


class Protection(Serialisable):
    """Protection options for use in styles."""

    tagname = "protection"

    locked = Bool()
    hidden = Bool()

    def __init__(self, locked=True, hidden=False):
        self.locked = locked
        self.hidden = hidden
