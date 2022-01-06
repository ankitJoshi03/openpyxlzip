# Copyright (c) 2010-2020 openpyxlzip

from openpyxlzip.descriptors.serialisable import Serialisable
from openpyxlzip.descriptors import (
    Typed,
    Integer,
    String,
    Set,
    NoneSet,
    Bool,
    Sequence,
)
from openpyxlzip.descriptors.excel import Relation
from openpyxlzip.drawing.spreadsheet_drawing import AnchorMarker
from openpyxlzip.xml.constants import SHEET_DRAWING_NS, MC_NS, SHEET_MAIN_NS
from openpyxlzip.xml.functions import (
    Element,
    localname,
    get_namespace,
)

class ObjectAnchor(Serialisable):

    tagname = "anchor"

    _from = Typed(expected_type=AnchorMarker, namespace=SHEET_DRAWING_NS)
    to = Typed(expected_type=AnchorMarker, namespace=SHEET_DRAWING_NS)
    moveWithCells = Bool(allow_none=True)
    sizeWithCells = Bool(allow_none=True)
    z_order = Integer(allow_none=True, hyphenated=True)


    def __init__(self,
                 _from=None,
                 to=None,
                 moveWithCells=False,
                 sizeWithCells=False,
                 z_order=None,
                ):
        self._from = _from
        self.to = to
        self.moveWithCells = moveWithCells
        self.sizeWithCells = sizeWithCells
        self.z_order = z_order


class ObjectPr(Serialisable):

    tagname = "objectPr"

    anchor = Typed(expected_type=ObjectAnchor, )
    locked = Bool(allow_none=True)
    defaultSize = Bool(allow_none=True)
    _print = Bool(allow_none=True)
    disabled = Bool(allow_none=True)
    uiObject = Bool(allow_none=True)
    autoFill = Bool(allow_none=True)
    autoLine = Bool(allow_none=True)
    autoPict = Bool(allow_none=True)
    macro = String(allow_none=True)
    altText = String(allow_none=True)
    dde = Bool(allow_none=True)
    id = Relation()

    __elements__ = ('anchor',)
    __attrs__ = ("locked", "defaultSize", "_print", "disabled", "uiObject", "autoFill", "autoLine", "autoPict", "macro", "altText", "dde", "id")

    def __init__(self,
                 anchor=None,
                 locked=True,
                 defaultSize=True,
                 _print=True,
                 disabled=False,
                 uiObject=False,
                 autoFill=True,
                 autoLine=True,
                 autoPict=True,
                 macro=None,
                 altText=None,
                 dde=False,
                 id=None,
                 path=None,
                 imgData=None,
                ):
        self.anchor = anchor
        self.locked = locked
        self.defaultSize = defaultSize
        self._print = _print
        self.disabled = disabled
        self.uiObject = uiObject
        self.autoFill = autoFill
        self.autoLine = autoLine
        self.autoPict = autoPict
        self.macro = macro
        self.altText = altText
        self.dde = dde
        self.id = id
        self.path = path
        self.imgData = imgData


class OleObject(Serialisable):

    tagname = "oleObject"

    objectPr = Typed(expected_type=ObjectPr, allow_none=True)
    progId = String(allow_none=True)
    dvAspect = NoneSet(values=(['DVASPECT_CONTENT', 'DVASPECT_ICON']))
    link = String(allow_none=True)
    oleUpdate = NoneSet(values=(['OLEUPDATE_ALWAYS', 'OLEUPDATE_ONCALL']))
    autoLoad = Bool(allow_none=True)
    shapeId = Integer()
    id = Relation()
    mime_type = "application/vnd.openxmlformats-officedocument.oleObject"


    __elements__ = ('objectPr',)
    __attrs__ = ("progId", "dvAspect", "link", "oleUpdate", "autoLoad", "shapeId", "id")

    def __init__(self,
                 objectPr=None,
                 progId=None,
                 dvAspect='DVASPECT_CONTENT',
                 link=None,
                 oleUpdate=None,
                 autoLoad=False,
                 shapeId=None,
                 id=None,
                 path=None,
                 oleObj=None,
                ):
        self.objectPr = objectPr
        self.progId = progId
        self.dvAspect = dvAspect
        self.link = link
        self.oleUpdate = oleUpdate
        self.autoLoad = autoLoad
        self.shapeId = shapeId
        self.id = id
        self.path = path
        self.oleObj = oleObj


class OleObjects(Serialisable):

    tagname = "oleObjects"

    oleObject = Sequence(expected_type=OleObject)

    __elements__ = ('oleObject',)

    def __init__(self,
                 oleObject=(),
                ):
        self.oleObject = oleObject

    def to_tree(self, elem_type=None):
        tagname = "oleObjects"
        root_node = Element(tagname, {}, nsmap={})
        for obj in self.oleObject:
            tagname = "{%s}%s" % (MC_NS, "AlternateContent")
            alt_content = Element(tagname, {}, nsmap={"mc": MC_NS})

            tagname = "{%s}%s" % (MC_NS, "Choice")
            choice = Element(tagname, {"Requires": "x14"})
            alt_content.append(choice)

            sub_tree = obj.to_tree()
            for child in sub_tree.iterdescendants():
                if child.tag.endswith("}from"):
                    child.tag = "from"
                    # child.set("xmlns", SHEET_DRAWING_NS)
                elif child.tag.endswith("}to"):
                    child.tag = "to"
                    # child.set("xmlns", SHEET_DRAWING_NS)
            choice.append(sub_tree)


            tagname = "{%s}%s" % (MC_NS, "Fallback")
            fallback = Element(tagname, {})

            sub_tree_copy = obj.to_tree()
            child = sub_tree_copy.getchildren()[0]
            sub_tree_copy.remove(child)

            fallback.append(sub_tree_copy)
            
            alt_content.append(fallback)

            root_node.append(alt_content)
        # root_node.set("xmlns", SHEET_MAIN_NS)
        return root_node


VML_OLE_DOC_FORMAT = """<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75"
  o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
  <v:stroke joinstyle="miter"/>
  <v:formulas>
   <v:f eqn="if lineDrawn pixelLineWidth 0"/>
   <v:f eqn="sum @0 1 0"/>
   <v:f eqn="sum 0 0 @1"/>
   <v:f eqn="prod @2 1 2"/>
   <v:f eqn="prod @3 21600 pixelWidth"/>
   <v:f eqn="prod @3 21600 pixelHeight"/>
   <v:f eqn="sum @0 0 1"/>
   <v:f eqn="prod @6 1 2"/>
   <v:f eqn="prod @7 21600 pixelWidth"/>
   <v:f eqn="sum @8 21600 0"/>
   <v:f eqn="prod @7 21600 pixelHeight"/>
   <v:f eqn="sum @10 21600 0"/>
  </v:formulas>
  <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
  <o:lock v:ext="edit" aspectratio="t"/>
 </v:shapetype>{shapes}</xml>"""

VML_OLE_SHAPE_FORMAT = """<v:shape id="{shape_id}" type="#_x0000_t75" style='position:absolute;
  margin-left:{margin_left}pt;margin-top:{margin_top}pt;width:{width}pt;height:{height}pt;z-index:{z_index}'
  filled="t" fillcolor="window [65]" stroked="t" strokecolor="windowText [64]"
  o:insetmode="auto">
  <v:fill color2="window [65]"/>
  <v:imagedata o:relid="{relid}" o:title=""/>
  <x:ClientData ObjectType="Pict">
   <x:SizeWithCells/>
   <x:Anchor>
    {anchor}</x:Anchor>
   <x:CF>Pict</x:CF>
   <x:AutoPict/>
  </x:ClientData>
 </v:shape>"""