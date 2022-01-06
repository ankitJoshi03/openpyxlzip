#copyright openpyxlzip 2010-2018

"""
Excel office art descriptors
"""

from openpyxlzip.xml.constants import REL_NS, DRAWING_NS, DRAWING_16_NS, DRAWING_14_NS
from openpyxlzip.compat import safe_string
from openpyxlzip.xml.functions import Element

from openpyxlzip.descriptors import (
    Typed,)
from . import (
    MatchPattern,
    MinMax,
    Integer,
    String,
    Sequence,
)
from .serialisable import Serialisable
from openpyxlzip.drawing.effect import EffectList


class HiddenEffects(Serialisable):
    
    tagname = "hiddenEffects"
    namespace = DRAWING_14_NS
    effectLst = Typed(expected_type=EffectList, allow_none=True)

    def __init__(self,
                 effectLst=None,
                ):
        self.effectLst = effectLst


class CreationId(Serialisable):
    
    tagname = "creationId"
    namespace = DRAWING_16_NS
    id = String()

    def __init__(self,
                 id=None,
                ):
        self.id = id


class CompatExt(Serialisable):
    
    tagname = "compatExt"
    namespace = DRAWING_14_NS
    spid = String()

    def __init__(self,
                 spid=None,
                ):
        self.spid = spid


class OfficeArtExtension(Serialisable):

    tagname = "ext"
    namespace = DRAWING_NS
    uri = String()
    compatExt = Typed(expected_type=CompatExt, allow_none=True)
    creationId = Typed(expected_type=CreationId, allow_none=True)
    __elements__ = ("compatExt", "creationId", "hiddenEffects")

    def __init__(self,
                 uri=None,
                 compatExt=None,
                 creationId=None,
                 hiddenEffects=None
                ):
        self.uri = uri
        self.compatExt = compatExt
        self.creationId = creationId
        self.hiddenEffects = hiddenEffects


class OfficeArtExtensionList(Serialisable):

    tagname = "extLst"
    namespace = DRAWING_NS
    ext = Sequence(expected_type=OfficeArtExtension)

    def __init__(self,
                 ext=(),
                ):
        self.ext = ext
