# Copyright (c) 2010-2020 openpyxlzip


"""Read an xlsx file into Python"""

# Python stdlib imports
from zipfile import ZipFile, ZIP_DEFLATED, BadZipfile
from sys import exc_info
from io import BytesIO
import os.path
import warnings
from parse import parse

from openpyxlzip.pivot.table import TableDefinition

# Allow blanket setting of KEEP_VBA for testing
try:
    from ..tests import KEEP_VBA
except ImportError:
    KEEP_VBA = False


# package imports
from openpyxlzip.utils.exceptions import InvalidFileException
from openpyxlzip.xml.constants import (
    ARC_SHARED_STRINGS,
    ARC_CORE,
    ARC_APP,
    ARC_CUSTOM,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
    ARC_THEME,
    COMMENTS_NS,
    SHARED_STRINGS,
    EXTERNAL_LINK,
    PACKAGE_PRINTER_SETTINGS,
    PACKAGE_CUSTOM_XML,
    PACKAGE_IMAGES,
    PACKAGE_DRAWINGS,
    XLTM,
    XLTX,
    XLSM,
    XLSX,
)
from openpyxlzip.cell import MergedCell
from openpyxlzip.comments.comment_sheet import CommentSheet

from .strings import read_string_table
from .workbook import WorkbookParser
from openpyxlzip.styles.stylesheet import apply_stylesheet

from openpyxlzip.packaging.core import DocumentProperties
from openpyxlzip.packaging.manifest import Manifest, Override

from openpyxlzip.packaging.relationship import (
    RelationshipList,
    get_dependents,
    get_rels_path,
)

from openpyxlzip.worksheet._read_only import ReadOnlyWorksheet
from openpyxlzip.worksheet._reader import WorksheetReader
from openpyxlzip.chartsheet import Chartsheet
from openpyxlzip.worksheet.table import Table
from openpyxlzip.drawing.spreadsheet_drawing import SpreadsheetDrawing

from openpyxlzip.xml.functions import fromstring

from .drawings import find_images


SUPPORTED_FORMATS = ('.xlsx', '.xlsm', '.xltx', '.xltm')

def _validate_archive(filename):
    """
    Does a first check whether filename is a string or a file-like
    object. If it is a string representing a filename, a check is done
    for supported formats by checking the given file-extension. If the
    file-extension is not in SUPPORTED_FORMATS an InvalidFileException
    will raised. Otherwise the filename (resp. file-like object) will
    forwarded to zipfile.ZipFile returning a ZipFile-Instance.
    """
    is_file_like = hasattr(filename, 'read')
    if not is_file_like:
        file_format = os.path.splitext(filename)[-1].lower()
        if file_format not in SUPPORTED_FORMATS:
            if file_format == '.xls':
                msg = ('openpyxlzip does not support the old .xls file format, '
                       'please use xlrd to read this file, or convert it to '
                       'the more recent .xlsx file format.')
            elif file_format == '.xlsb':
                msg = ('openpyxlzip does not support binary format .xlsb, '
                       'please convert this file to .xlsx format if you want '
                       'to open it with openpyxlzip')
            else:
                msg = ('openpyxlzip does not support %s file format, '
                       'please check you can open '
                       'it with Excel first. '
                       'Supported formats are: %s') % (file_format,
                                                       ','.join(SUPPORTED_FORMATS))
            raise InvalidFileException(msg)

    archive = ZipFile(filename, 'r')
    return archive


def _find_workbook_part(package):
    workbook_types = [XLTM, XLTX, XLSM, XLSX]
    for ct in workbook_types:
        part = package.find(ct)
        if part:
            return part

    # some applications reassign the default for application/xml
    defaults = {p.ContentType for p in package.Default}
    workbook_type = defaults & set(workbook_types)
    if workbook_type:
        return Override("/" + ARC_WORKBOOK, workbook_type.pop())

    raise IOError("File contains no valid workbook part")


class ExcelReader:

    """
    Read an Excel package and dispatch the contents to the relevant modules
    """

    def __init__(self,  fn, read_only=False, keep_vba=KEEP_VBA,
                  data_only=False, keep_links=True):
        self.archive = _validate_archive(fn)
        self.valid_files = self.archive.namelist()
        self.read_only = read_only
        self.keep_vba = keep_vba
        self.data_only = data_only
        self.keep_links = keep_links
        self.shared_strings = []


    def read_manifest(self):
        src = self.archive.read(ARC_CONTENT_TYPES)
        root = fromstring(src)
        self.package = Manifest.from_tree(root)


    def read_strings(self):
        ct = self.package.find(SHARED_STRINGS)
        if ct is not None:
            strings_path = ct.PartName[1:]
            with self.archive.open(strings_path,) as src:
                self.shared_strings = read_string_table(src)


    def read_workbook(self):
        wb_part = _find_workbook_part(self.package)
        self.parser = WorkbookParser(self.archive, wb_part.PartName[1:], keep_links=self.keep_links)
        self.parser.parse()
        wb = self.parser.wb
        wb._sheets = []
        wb._data_only = self.data_only
        wb._read_only = self.read_only
        wb.template = wb_part.ContentType in (XLTX, XLTM)

        # If are going to preserve the vba then attach a copy of the archive to the
        # workbook so that is available for the save.
        if self.keep_vba:
            wb.vba_archive = ZipFile(BytesIO(), 'a', ZIP_DEFLATED)
            for name in self.valid_files:
                wb.vba_archive.writestr(name, self.archive.read(name))

        if ARC_APP in self.valid_files:
            wb.app_archive = ZipFile(BytesIO(), 'a', ZIP_DEFLATED)
            wb.app_archive.writestr(ARC_APP, self.archive.read(ARC_APP))

        if self.read_only:
            wb._archive = self.archive

        self.wb = wb


    def read_properties(self):
        if ARC_CORE in self.valid_files:
            src = fromstring(self.archive.read(ARC_CORE))
            self.wb.properties = DocumentProperties.from_tree(src)
        if ARC_CUSTOM in  self.valid_files:
            self.wb.arc_custom = ZipFile(BytesIO(), 'a', ZIP_DEFLATED)
            self.wb.arc_custom.writestr(ARC_CUSTOM, self.archive.read(ARC_CUSTOM))


    def read_theme(self):
        if ARC_THEME in self.valid_files:
            self.wb.loaded_theme = self.archive.read(ARC_THEME)


    def read_chartsheet(self, sheet, rel):
        sheet_path = rel.target
        rels_path = get_rels_path(sheet_path)
        rels = []
        if rels_path in self.valid_files:
            rels = get_dependents(self.archive, rels_path)

        with self.archive.open(sheet_path, "r") as src:
            xml = src.read()
        node = fromstring(xml)
        cs = Chartsheet.from_tree(node)
        cs._parent = self.wb
        cs.title = sheet.name
        self.wb._add_sheet(cs)

        drawings = rels.find(SpreadsheetDrawing._rel_type)
        for rel in drawings:
            drawing, charts, images = find_images(self.archive, rel.target)
            if not hasattr(sheet, "drawings"):
                sheet.drawings = {}
            sheet.drawings[rel.target] = drawing
            if self.wb.drawings is None:
                self.wb.drawings = {}
            self.wb.drawings[rel.target] = drawing
            for c in charts:
                cs.add_chart(c)


    def read_worksheets(self):
        comment_warning = """Cell '{0}':{1} is part of a merged range but has a comment which will be removed because merged cells cannot contain any data."""
        for sheet, rel in self.parser.find_sheets():
            if rel.target not in self.valid_files:
                continue

            if "chartsheet" in rel.Type:
                self.read_chartsheet(sheet, rel)
                continue

            rels_path = get_rels_path(rel.target)
            rels = RelationshipList()
            if rels_path in self.valid_files:
                rels = get_dependents(self.archive, rels_path)

            if self.read_only:
                ws = ReadOnlyWorksheet(self.wb, sheet.name, rel.target, self.shared_strings)
                ws.sheet_state = sheet.state
                ws.sheetId = sheet.sheetId
                self.wb._sheets.append(ws)
                continue
            else:
                fh = self.archive.open(rel.target)
                ws = self.wb.create_sheet(sheet.name)
                ws._rels = rels
                ws_parser = WorksheetReader(ws, fh, self.shared_strings, self.data_only)
                ws_parser.bind_all()
                ws.sheetId = sheet.sheetId

            # assign any comments to cells
            for r in rels.find(COMMENTS_NS):
                if r.target in self.valid_files:
                    src = self.archive.read(r.target)
                else:
                    continue
                comment_sheet = CommentSheet.from_tree(fromstring(src))
                for ref, comment in comment_sheet.comments:
                    try:
                        ws[ref].comment = comment
                    except AttributeError:
                        c = ws[ref]
                        if isinstance(c, MergedCell):
                            warnings.warn(comment_warning.format(ws.title, c.coordinate))
                            continue

            # preserve link to VML file if VBA
            if self.wb.vba_archive and ws.legacy_drawing:
                ws.legacy_drawing = rels[ws.legacy_drawing].target

            for t in ws_parser.tables:
                src = self.archive.read(t)
                xml = fromstring(src)
                table = Table.from_tree(xml)
                ws.add_table(table)

            drawings = rels.find(SpreadsheetDrawing._rel_type)
            for rel in drawings:
                drawing, charts, images = find_images(self.archive, rel.target)
                if self.wb.drawings is None:
                    self.wb.drawings = {}
                ws.drawings[rel.target] = drawing
                self.wb.drawings[rel.target] = drawing
                for c in charts:
                    ws.add_chart(c, c.anchor)
                for im in images:
                    ws.add_image(im, im.anchor)
                drawing_id = parse(PACKAGE_DRAWINGS + "/drawing{:d}.xml", rel.target)[0]
                if self.wb._all_drawings is None:
                    self.wb._all_drawings = {}
                self.wb._all_drawings[rel.target] = ZipFile(BytesIO(), 'a', ZIP_DEFLATED)
                self.wb._all_drawings[rel.target].writestr(rel.target, self.archive.read(rel.target))
                sub_rels_path = get_rels_path(rel.target)
                if sub_rels_path not in self.archive.namelist():
                    continue
                sub_rels = get_dependents(self.archive, sub_rels_path)
                if self.wb._all_drawings_rels is None:
                    self.wb._all_drawings_rels = {}
                self.wb._all_drawings_rels[sub_rels_path] = sub_rels
                for sub_rel in sub_rels.Relationship:
                    if sub_rel.target in self.archive.namelist():
                        if self.wb._all_drawing_dependencies is None:
                            self.wb._all_drawing_dependencies = {}
                        if sub_rel.target in self.wb._all_drawing_dependencies:
                            continue
                        self.wb._all_drawing_dependencies[sub_rel.target] = ZipFile(BytesIO(), 'a', ZIP_DEFLATED)
                        self.wb._all_drawing_dependencies[sub_rel.target].writestr(sub_rel.target, self.archive.read(sub_rel.target))


            pivot_rel = rels.find(TableDefinition.rel_type)
            for r in pivot_rel:
                pivot_path = r.Target
                src = self.archive.read(pivot_path)
                tree = fromstring(src)
                pivot = TableDefinition.from_tree(tree)
                pivot.cache = self.parser.pivot_caches[pivot.cacheId]
                ws.add_pivot(pivot)

            ws.sheet_state = sheet.state

    #MattJ added to preserve printer settings
    def read_printer_settings(self):
        for full_filename in self.valid_files:
            if full_filename.startswith(PACKAGE_PRINTER_SETTINGS):
                if self.wb._printer_settings is None:
                    self.wb._printer_settings = {}
                printer_setting_id = parse(PACKAGE_PRINTER_SETTINGS + "/printerSettings{:d}.bin", full_filename)[0]
                self.wb._printer_settings[printer_setting_id] = ZipFile(BytesIO(), 'a', ZIP_DEFLATED)
                self.wb._printer_settings[printer_setting_id].writestr(full_filename, self.archive.read(full_filename))


    #MattJ added to handle custom XML when present
    def read_custom_xml(self):
        #TODO add the relationships
        for full_filename in self.valid_files:
            if full_filename.startswith(PACKAGE_CUSTOM_XML):
                if self.wb._custom_xml is None:
                    self.wb._custom_xml = {}
                self.wb._custom_xml[full_filename] = ZipFile(BytesIO(), 'a', ZIP_DEFLATED)
                self.wb._custom_xml[full_filename].writestr(full_filename, self.archive.read(full_filename))

    def read(self):
        self.read_manifest()
        self.read_strings()
        self.read_workbook()
        self.read_printer_settings()
        self.read_custom_xml()
        self.read_properties()
        self.read_theme()
        apply_stylesheet(self.archive, self.wb)
        self.read_worksheets()
        self.parser.assign_names()
        if not self.read_only:
            self.archive.close()


def load_workbook(filename, read_only=False, keep_vba=KEEP_VBA,
                  data_only=False, keep_links=True):
    """Open the given filename and return the workbook

    :param filename: the path to open or a file-like object
    :type filename: string or a file-like object open in binary mode c.f., :class:`zipfile.ZipFile`

    :param read_only: optimised for reading, content cannot be edited
    :type read_only: bool

    :param keep_vba: preseve vba content (this does NOT mean you can use it)
    :type keep_vba: bool

    :param data_only: controls whether cells with formulae have either the formula (default) or the value stored the last time Excel read the sheet
    :type data_only: bool

    :param keep_links: whether links to external workbooks should be preserved. The default is True
    :type keep_links: bool

    :rtype: :class:`openpyxl.workbook.Workbook`

    .. note::

        When using lazy load, all worksheets will be :class:`openpyxlzip.worksheet.iter_worksheet.IterableWorksheet`
        and the returned workbook will be read-only.

    """
    reader = ExcelReader(filename, read_only, keep_vba,
                        data_only, keep_links)
    reader.read()
    return reader.wb
