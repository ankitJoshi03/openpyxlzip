# Copyright (c) 2010-2020 openpyxlzip

import atexit
from collections import defaultdict
from io import BytesIO
import os
from tempfile import NamedTemporaryFile
from warnings import warn

from openpyxlzip import LXML
from openpyxlzip.xml.functions import xmlfile, tostring
from openpyxlzip.xml.constants import SHEET_MAIN_NS, REL_NS

from openpyxlzip.comments.comment_sheet import CommentRecord
from openpyxlzip.packaging.relationship import Relationship, RelationshipList
from openpyxlzip.styles.differential import DifferentialStyle
from openpyxlzip.drawing.spreadsheet_drawing import SpreadsheetDrawing

from .dimensions import SheetDimension
from .hyperlink import HyperlinkList
from .merge import MergeCell, MergeCells
from .related import Related
from .table import TablePartList

from openpyxlzip.cell._writer import write_cell

from openpyxlzip.xml.schema import ALL_DEFINITIONS, ROOT_ELEMS

PROPERTIES_TAG = '{%s}sheetPr' % SHEET_MAIN_NS
DIMENSION_TAG = '{%s}dimension' % SHEET_MAIN_NS
VIEWS_TAG = '{%s}sheetViews' % SHEET_MAIN_NS
FORMAT_TAG = '{%s}sheetFormatPr' % SHEET_MAIN_NS
COLS_TAG = '{%s}cols' % SHEET_MAIN_NS
DATA_TAG = '{%s}sheetData' % SHEET_MAIN_NS
SHEET_CALC_PR_TAG = '{%s}sheetCalcPr' % SHEET_MAIN_NS
PROT_TAG = '{%s}sheetProtection' % SHEET_MAIN_NS
PROT_RANGES_TAG = '{%s}protectedRanges' % SHEET_MAIN_NS
SCENARIOS_TAG = '{%s}scenarios' % SHEET_MAIN_NS
FILTER_TAG = '{%s}autoFilter' % SHEET_MAIN_NS
SORT_STATE_TAG = '{%s}sortState' % SHEET_MAIN_NS
CONSOLIDATE_TAG = '{%s}dataConsolidate' % SHEET_MAIN_NS
CUSTOM_VIEWS_TAG = '{%s}customSheetViews' % SHEET_MAIN_NS
MERGE_TAG = '{%s}mergeCells' % SHEET_MAIN_NS
PHONETIC_TAG = '{%s}phoneticPr' % SHEET_MAIN_NS
CF_TAG = '{%s}conditionalFormatting' % SHEET_MAIN_NS
VALIDATION_TAG = '{%s}dataValidations' % SHEET_MAIN_NS
HYPERLINK_TAG = "{%s}hyperlinks" % SHEET_MAIN_NS
PRINT_TAG = '{%s}printOptions' % SHEET_MAIN_NS
MARGINS_TAG = '{%s}pageMargins' % SHEET_MAIN_NS
PAGE_TAG = '{%s}pageSetup' % SHEET_MAIN_NS
HEADER_TAG = '{%s}headerFooter' % SHEET_MAIN_NS
ROW_BREAK_TAG = '{%s}rowBreaks' % SHEET_MAIN_NS
COL_BREAK_TAG = '{%s}colBreaks' % SHEET_MAIN_NS
CUSTOM_PR_TAG = '{%s}customProperties' % SHEET_MAIN_NS
CELL_WATCH_TAG = '{%s}cellWatches' % SHEET_MAIN_NS
IGNORED_ERRORS_TAG = '{%s}ignoredErrors' % SHEET_MAIN_NS
SMART_TAGS_TAG = '{%s}smartTags' % SHEET_MAIN_NS
DRAWING_TAG = '{%s}drawing' % SHEET_MAIN_NS
DRAWING_HF_TAG = '{%s}drawingHF' % SHEET_MAIN_NS
PICTURE_TAG = '{%s}picture' % SHEET_MAIN_NS
OLE_OBJECTS_TAG = '{%s}oleObjects' % SHEET_MAIN_NS
CONTROLS_TAG = '{%s}controls' % SHEET_MAIN_NS
WEB_PUBLISH_ITEMS_TAG = '{%s}webPublishItems' % SHEET_MAIN_NS
LEGACY_TAG = '{%s}legacyDrawing' % SHEET_MAIN_NS
TABLE_TAG = "{%s}tableParts" % SHEET_MAIN_NS
EXT_LIST_TAG = "{%s}extLst" % SHEET_MAIN_NS

ALL_TEMP_FILES = []

@atexit.register
def _openpyxl_shutdown():
    for path in ALL_TEMP_FILES:
        if os.path.exists(path):
            os.remove(path)


def create_temporary_file(suffix=''):
    fobj = NamedTemporaryFile(mode='w+', suffix=suffix,
                              prefix='openpyxlzip.', delete=False)
    filename = fobj.name
    fobj.close()
    ALL_TEMP_FILES.append(filename)
    return filename


class WorksheetWriter:


    def __init__(self, ws, out=None):
        self.ws = ws
        self.ws._hyperlinks = []
        self.ws._comments = []
        if out is None:
            out = create_temporary_file()
        self.out = out
        self._rels = ws._rels
        self.xf = self.get_stream()
        next(self.xf) # start generator


    def write_properties(self):
        props = self.ws.sheet_properties
        self.xf.send(props.to_tree())


    def write_dimensions(self):
        """
        Write worksheet size if known
        """
        ref = getattr(self.ws, 'calculate_dimension', None)
        if ref:
            dim = SheetDimension(ref())
            self.xf.send(dim.to_tree())


    def write_format(self):
        self.ws.sheet_format.outlineLevelCol = self.ws.column_dimensions.max_outline
        fmt = self.ws.sheet_format
        self.xf.send(fmt.to_tree())


    def write_views(self):
        views = self.ws.views
        self.xf.send(views.to_tree())


    def write_cols(self):
        cols = self.ws.column_dimensions
        self.xf.send(cols.to_tree())


    def write_top(self):
        """
        Write all elements up to rows:
        properties
        dimensions
        views
        format
        cols
        """
        self.write_properties(None)
        self.write_dimensions(None)
        self.write_views(None)
        self.write_format(None)
        self.write_cols(None)


    def rows(self):
        """Return all rows, and any cells that they contain"""
        # order cells by row
        rows = defaultdict(list)
        for (row, col), cell in sorted(self.ws._cells.items()):
            rows[row].append(cell)

        # add empty rows if styling has been applied
        for row in self.ws.row_dimensions.keys() - rows.keys():
            rows[row] = []

        return sorted(rows.items())


    def write_rows(self):
        xf = self.xf.send(True)

        with xf.element("sheetData"):
            for row_idx, row in self.rows():
                self.write_row(xf, row, row_idx)

        self.xf.send(None) # return control to generator


    def write_row(self, xf, row, row_idx):
        attrs = {'r': f"{row_idx}"}
        dims = self.ws.row_dimensions
        attrs.update(dims.get(row_idx, {}))

        with xf.element("row", attrs):

            for cell in row:
                if cell._comment is not None:
                    comment = CommentRecord.from_cell(cell)
                    self.ws._comments.append(comment)
                if (
                    cell._value is None
                    and not cell.has_style
                    and not cell._comment
                    ):
                    continue
                write_cell(xf, self.ws, cell, cell.has_style)


    def write_protection(self):
        prot = self.ws.protection
        if prot:
            self.xf.send(prot.to_tree())


    def write_extra(self, tag=None):
        extra_elem = self.ws.extra_elem
        if extra_elem:
            if tag is None:
                for key in extra_elem:
                    self.xf.send(extra_elem[key])
            else:
                self.xf.send(extra_elem[tag])


    def write_scenarios(self):
        scenarios = self.ws.scenarios
        if scenarios:
            self.xf.send(scenarios.to_tree())


    def write_filter(self):
        flt = self.ws.auto_filter
        if flt:
            self.xf.send(flt.to_tree())


    def write_sort(self):
        """
        As per discusion with the OOXML Working Group global sort state is not required.
        openpyxlzip never reads it from existing files
        """
        pass


    def write_merged_cells(self):
        merged = self.ws.merged_cells
        if merged:
            cells = [MergeCell(str(ref)) for ref in self.ws.merged_cells]
            self.xf.send(MergeCells(mergeCell=cells).to_tree())


    def write_formatting(self):
        df = DifferentialStyle()
        wb = self.ws.parent
        for cf in self.ws.conditional_formatting:
            for rule in cf.rules:
                if rule.dxf and rule.dxf != df:
                    rule.dxfId = wb._differential_styles.add(rule.dxf)
            self.xf.send(cf.to_tree())


    def write_validations(self):
        dv = self.ws.data_validations
        if dv:
            self.xf.send(dv.to_tree())


    def write_hyperlinks(self):
        links = HyperlinkList()

        for link in self.ws._hyperlinks:
            if link.target:
                found = False
                for rel in self._rels.Relationship:
                    if rel.Type == REL_NS + "/hyperlink" and rel.TargetMode == "External" and rel.Target == link.target:
                        found = True
                if not found:
                    max_rel_id = None
                    for rel in self._rels.Relationship:
                        if max_rel_id is None:
                            max_rel_id = int(rel.id.replace("rId", ""))
                        max_rel_id = max(max_rel_id, int(rel.id.replace("rId", "")))
                    if max_rel_id is None:
                        max_rel_id = 1
                    else:
                        max_rel_id += 1
                    rel = Relationship(type="hyperlink", TargetMode="External", Target=link.target, Id="rId{0}".format(max_rel_id))
                    self._rels.append(rel)
                    link.id = rel.id
            links.hyperlink.append(link)

        if links:
            self.xf.send(links.to_tree())


    def write_print(self):
        print_options = self.ws.print_options
        if print_options:
            self.xf.send(print_options.to_tree())


    def write_margins(self):
        margins = self.ws.page_margins
        if margins:
            self.xf.send(margins.to_tree())


    def write_page(self):
        setup = self.ws.page_setup
        if setup:
            self.xf.send(setup.to_tree())


    def write_header(self):
        hf = self.ws.HeaderFooter
        if hf:
            self.xf.send(hf.to_tree())


    def write_row_breaks(self):
        if self.ws.row_breaks:
            self.xf.send(self.ws.row_breaks.to_tree())

    def write_col_breaks(self):
        if self.ws.col_breaks:
            self.xf.send(self.ws.col_breaks.to_tree())


    def write_ole_objects(self):
        if self.ws.ole_objects is not None:
            for ole_obj in self.ws.ole_objects.oleObject:
                found = False
                final_rel = None
                for rel in self._rels.Relationship:
                    if rel.Type == REL_NS + "/oleObject" and rel.Target == ole_obj.path:
                        final_rel = rel
                        found = True
                if found:
                    rel = final_rel
                    rel.Target = ole_obj.path.replace("xl/", "../")
                    rel.BackupTarget = ole_obj.path.replace("xl/", "../")
                else:
                    max_rel_id = None
                    for rel in self._rels.Relationship:
                        if max_rel_id is None:
                            max_rel_id = int(rel.id.replace("rId", ""))
                        max_rel_id = max(max_rel_id, int(rel.id.replace("rId", "")))
                    if max_rel_id is None:
                        max_rel_id = 1
                    else:
                        max_rel_id += 1
                    rel = Relationship(type="oleObject", Target=ole_obj.path.replace("xl/", "../"), Id="rId{0}".format(max_rel_id))
                    self._rels.append(rel)
                ole_obj.id = rel.id


                found = False
                final_rel = None
                for rel in self._rels.Relationship:
                    if rel.Type == REL_NS + "/image" and rel.Target == ole_obj.objectPr.path:
                        final_rel = rel
                        found = True
                if found:
                    rel = final_rel
                    rel.Target = ole_obj.objectPr.path.replace("xl/", "../")
                    rel.BackupTarget = ole_obj.objectPr.path.replace("xl/", "../")
                else:
                    max_rel_id = None
                    for rel in self._rels.Relationship:
                        if max_rel_id is None:
                            max_rel_id = int(rel.id.replace("rId", ""))
                        max_rel_id = max(max_rel_id, int(rel.id.replace("rId", "")))
                    if max_rel_id is None:
                        max_rel_id = 1
                    else:
                        max_rel_id += 1
                    rel = Relationship(type="image", Target=ole_obj.objectPr.path.replace("xl/", "../"), Id="rId{0}".format(max_rel_id))
                    self._rels.append(rel)
                ole_obj.objectPr.id = rel.id

            self.xf.send(self.ws.ole_objects.to_tree())


    def write_drawings(self):
        # print("Write drawing")
        if self.ws._charts or self.ws._images or self.ws._drawing is not None:
            found = False
            final_rel = None
            #MattJ this seems to only be reachable during testing
            if self.ws._drawing is None:
                self.ws._drawing = SpreadsheetDrawing()
                self.ws._drawing.charts = self.ws._charts
                self.ws._drawing.images = self.ws._images
            for rel in self._rels.Relationship:
                if rel.Type == REL_NS + "/drawing":
                    final_rel = rel
                    found = True
            if found:
                rel = final_rel
                rel.Target = self.ws._drawing.path.replace("/xl/", "../")
                rel.BackupTarget = self.ws._drawing.path.replace("/xl/", "../")
            else:
                max_rel_id = None
                for rel in self._rels.Relationship:
                    if max_rel_id is None:
                        max_rel_id = int(rel.id.replace("rId", ""))
                    max_rel_id = max(max_rel_id, int(rel.id.replace("rId", "")))
                if max_rel_id is None:
                    max_rel_id = 1
                else:
                    max_rel_id += 1
                rel = Relationship(type="drawing", Target=self.ws._drawing.path.replace("/xl/", "../"), Id="rId{0}".format(max_rel_id))
                self._rels.append(rel)
            drawing = Related()
            drawing.id = rel.id
            self.xf.send(drawing.to_tree("drawing"))


    def write_legacy(self):
        """
        Comments & VBA controls use VML and require an additional element
        that is no longer in the specification.
        """
        if (self.ws.legacy_drawing is not None or self.ws._comments):
            # legacy = Related(id="rId3")
            found = False
            for rel in self._rels.Relationship:
                if rel.Type == REL_NS + "/vmlDrawing":
                    final_rel = rel
                    found = True
            if found:
                rel = final_rel
                if self.ws.legacy_drawing is None:
                    print("**************************************************************")
                    print("**************************************************************")
                    print("**************************************************************")
                    print("**************************************************************")
                    print(rel)
                    print("**************************************************************")
                    print("**************************************************************")
                    print("**************************************************************")
                    print("**************************************************************")
                    print("**************************************************************")
                else:
                    rel.Target = self.ws.legacy_drawing.replace("/xl/", "../")
                    rel.BackupTarget = self.ws.legacy_drawing.replace("/xl/", "../")
            else:
                max_rel_id = None
                for rel in self._rels.Relationship:
                    if max_rel_id is None:
                        max_rel_id = int(rel.id.replace("rId", ""))
                    max_rel_id = max(max_rel_id, int(rel.id.replace("rId", "")))
                if max_rel_id is None:
                    max_rel_id = 1
                else:
                    max_rel_id += 1
                rel = Relationship(type="vmlDrawing", Target=self.ws.legacy_drawing.replace("/xl/", "../"), Id="rId{0}".format(max_rel_id))
                self._rels.append(rel)
            legacy = Related()
            legacy.id = rel.id
            self.xf.send(legacy.to_tree("legacyDrawing"))
        elif hasattr(self.ws, "extra_elem"):
            if LEGACY_TAG in self.ws.extra_elem:
                self.write_extra(tag=LEGACY_TAG)



    def write_tables(self):
        tables = TablePartList()

        for table in self.ws._tables.values():
            if not table.tableColumns:
                table._initialise_columns()
                if table.headerRowCount:
                    try:
                        row = self.ws[table.ref][0]
                        for cell, col in zip(row, table.tableColumns):
                            if cell.data_type != "s":
                                warn("File may not be readable: column headings must be strings.")
                            col.name = str(cell.value)
                    except TypeError:
                        warn("Column headings are missing, file may not be readable")
            added = False
            for rel in self._rels.Relationship:
                if rel.Type == table._rel_type:
                    #TODO this might be a bug
                    tables.append(Related(id=rel.Id))
                    added = True
            if not added:
                rel = Relationship(Type=table._rel_type, Target="")
                self._rels.append(rel)
                table._rel_id = rel.Id
                tables.append(Related(id=rel.Id))

        if tables:
            self.xf.send(tables.to_tree())


    def get_stream(self):
        with xmlfile(self.out,encoding="UTF-8") as xf:
            xf.write_declaration(standalone=True)
            temp_nsmap = {}
            if hasattr(self.ws, "nsmaps") and self.ws.nsmaps is not None:
                for key in self.ws.nsmaps:
                    if self.ws.nsmaps[key] != SHEET_MAIN_NS:
                        temp_nsmap[key] = self.ws.nsmaps[key]
            with xf.element("worksheet", self.ws.extra_attrib, xmlns=SHEET_MAIN_NS, nsmap=temp_nsmap) as root:
                try:
                    while True:
                        el = (yield)
                        if el is True:
                            yield xf
                        elif el is None: # et_xmlfile chokes
                            continue
                        else:
                            xf.write(el)
                except GeneratorExit:
                    pass


    def write_tail(self):
        """
        Write all elements after the rows
        calc properties
        protection
        protected ranges #
        scenarios
        filters
        sorts # always ignored
        data consolidation #
        custom views #
        merged cells
        phonetic properties #
        conditional formatting
        data validation
        hyperlinks
        print options
        page margins
        page setup
        header
        row breaks
        col breaks
        custom properties #
        cell watches #
        ignored errors #
        smart tags #
        drawing
        drawingHF #
        background #
        OLE objects #
        controls #
        web publishing #
        tables
        """
        self.write_protection(None)
        self.write_scenarios(None)
        self.write_filter(None)
        self.write_merged_cells(None)
        self.write_formatting(None)
        self.write_validations(None)
        self.write_hyperlinks(None)
        self.write_print(None)
        self.write_margins(None)
        self.write_page(None)
        self.write_header(None)
        self.write_row_breaks(None)
        self.write_col_breaks(None)
        self.write_drawings(None)
        self.write_legacy(None)
        self.write_tables(None)
        self.write_extra(None)


    def write(self):
        """
        High level
        """
        dispatcher = [
            (PROPERTIES_TAG, self.write_properties),
            (DIMENSION_TAG, self.write_dimensions),
            (VIEWS_TAG, self.write_views),
            (FORMAT_TAG, self.write_format),
            (COLS_TAG, self.write_cols),
            (DATA_TAG, self.write_rows),
            (SHEET_CALC_PR_TAG, None),
            (PROT_TAG, None),
            # (PROT_TAG, self.write_protection),
            (PROT_RANGES_TAG, None),
            (SCENARIOS_TAG, self.write_scenarios),
            (FILTER_TAG, self.write_filter),
            (SORT_STATE_TAG, None),
            (CONSOLIDATE_TAG, None),
            (CUSTOM_VIEWS_TAG, None),
            (MERGE_TAG, self.write_merged_cells),
            (PHONETIC_TAG, None),
            (CF_TAG, self.write_formatting),
            (VALIDATION_TAG, self.write_validations),
            (HYPERLINK_TAG, self.write_hyperlinks),
            (PRINT_TAG, self.write_print),
            (MARGINS_TAG, self.write_margins),
            (PAGE_TAG, self.write_page),
            (HEADER_TAG, self.write_header),
            (ROW_BREAK_TAG, self.write_row_breaks),
            (COL_BREAK_TAG, self.write_col_breaks),
            (CUSTOM_PR_TAG, None),
            (CELL_WATCH_TAG, None),
            (IGNORED_ERRORS_TAG, None),
            (SMART_TAGS_TAG, None),
            (DRAWING_TAG, self.write_drawings),
            (DRAWING_HF_TAG, None),
            (PICTURE_TAG, None),
            # (LEGACY_TAG, None),
            (LEGACY_TAG, self.write_legacy),
            (OLE_OBJECTS_TAG, self.write_ole_objects),
            (CONTROLS_TAG, None),
            (WEB_PUBLISH_ITEMS_TAG, None),
            (TABLE_TAG, self.write_tables),
            (EXT_LIST_TAG, None),
        ]
        # root_tagname = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}chartsheet"
        # root_typename = ROOT_ELEMS[root_tagname]
        # element_order = sorted([int(key) for key in ALL_DEFINITIONS[root_typename]["element_order"].keys()])
        # for key in element_order:
        #     elem_tag, elem_type = ALL_DEFINITIONS[root_typename]["element_order"][str(key)]
        #     for tag, handler in dispatcher:
        #         if tag == elem_tag:
        #             if handler is None:
        #                 if hasattr(self.ws, "extra_elem"):
        #                     if tag in self.ws.extra_elem:
        #                         self.write_extra(tag=tag, elem_type=elem_type)
        #             else:
        #                 handler(elem_type)
        # if hasattr(self.ws, "extra_elem"):
        #     print(self.ws.extra_elem.keys())
            
        for tag, handler in dispatcher:
            # print(tag)
            if handler is None:
                if hasattr(self.ws, "extra_elem"):
                    if tag in self.ws.extra_elem:
                        self.write_extra(tag=tag)
            else:
                handler()
        self.close()


    def close(self):
        """
        Close the context manager
        """
        if self.xf:
            self.xf.close()


    def read(self):
        """
        Close the context manager and return serialised XML
        """
        self.close()
        if isinstance(self.out, BytesIO):
            return self.out.getvalue()
        with open(self.out, "rb") as src:
            out = src.read()

        return out


    def cleanup(self):
        """
        Remove tempfile
        """
        os.remove(self.out)
        ALL_TEMP_FILES.remove(self.out)
