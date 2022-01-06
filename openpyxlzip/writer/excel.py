# Copyright (c) 2010-2020 openpyxlzip

"""Write a .xlsx file."""

# Python stdlib imports
import re
from tempfile import TemporaryFile
from zipfile import ZipFile, ZIP_DEFLATED

# package imports
from openpyxlzip.xml import LXML
from openpyxlzip.compat import deprecated
from openpyxlzip.utils.exceptions import InvalidFileException
from openpyxlzip.xml.constants import (
    ARC_SHARED_STRINGS,
    ARC_CONTENT_TYPES,
    ARC_ROOT_RELS,
    ARC_WORKBOOK_RELS,
    ARC_APP, ARC_CORE, ARC_CUSTOM,
    ARC_THEME,
    ARC_STYLE,
    ARC_WORKBOOK,
    PACKAGE_WORKSHEETS,
    PACKAGE_PRINTER_SETTINGS,
    PACKAGE_CUSTOM_XML,
    PACKAGE_CHARTSHEETS,
    PACKAGE_DRAWINGS,
    PACKAGE_CHARTS,
    PACKAGE_IMAGES,
    PACKAGE_XL,
    VBA,
    SHEET_MAIN_NS,
    )
from openpyxlzip.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxlzip.xml.functions import tostring, fromstring, Element
from openpyxlzip.packaging.manifest import Manifest
from openpyxlzip.packaging.relationship import (
    get_rels_path,
    RelationshipList,
    Relationship,
)
from openpyxlzip.comments.comment_sheet import CommentSheet
from openpyxlzip.packaging.extended import ExtendedProperties
from openpyxlzip.styles.stylesheet import write_stylesheet
from openpyxlzip.worksheet._writer import WorksheetWriter
from openpyxlzip.workbook._writer import WorkbookWriter
from .theme import theme_xml


class ExcelWriter(object):
    """Write a workbook object to an Excel file."""

    def __init__(self, workbook, archive):
        self._archive = archive
        self.workbook = workbook
        self.manifest = Manifest()
        self.vba_modified = set()
        self._tables = []
        self._charts = []
        self._images = []
        self._drawings = []
        self._comments = []
        self._pivots = []
        self.drawing_id = 1


    def write_data(self):
        """Write the various xml files into the zip archive."""
        # cleanup all worksheets
        archive = self._archive

        if self.workbook.app_archive is None:
            props = ExtendedProperties()
            if LXML:
                archive.writestr(ARC_APP, tostring(props.to_tree(), pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
            else:
                archive.writestr(ARC_APP, tostring(props.to_tree()))

        else:
            archive.writestr(ARC_APP, self.workbook.app_archive.read(ARC_APP))

        if LXML:
            archive.writestr(ARC_CORE, tostring(self.workbook.properties.to_tree(), pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
        else:
            archive.writestr(ARC_CORE, tostring(self.workbook.properties.to_tree()))
        if self.workbook.loaded_theme:
            archive.writestr(ARC_THEME, self.workbook.loaded_theme)
        else:
            archive.writestr(ARC_THEME, theme_xml)

        self._write_worksheets()
        self._write_chartsheets()
        self._write_printer_settings()
        self._write_custom_xml()
        self._write_images()
        self._write_charts()

        #MattJ this ensures that the original versions are all preserved
        self._write_drawings_and_dependencies()

        #self._archive.writestr(ARC_SHARED_STRINGS,
                              #write_string_table(self.workbook.shared_strings))
        self._write_external_links()

        stylesheet = write_stylesheet(self.workbook)
        if LXML:
            archive.writestr(ARC_STYLE, tostring(stylesheet, pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
        else:
            archive.writestr(ARC_STYLE, tostring(stylesheet))

        writer = WorkbookWriter(self.workbook)
        archive.writestr(ARC_ROOT_RELS, writer.write_root_rels())
        archive.writestr(ARC_WORKBOOK, writer.write())
        archive.writestr(ARC_WORKBOOK_RELS, writer.write_rels())

        self._merge_vba()

        self.manifest._write(archive, self.workbook)

    def _merge_vba(self):
        """
        If workbook contains macros then extract associated files from cache
        of old file and add to archive
        """
        ARC_VBA = re.compile("|".join(
            ('xl/vba', r'xl/drawings/.*vmlDrawing\d\.vml',
             'xl/ctrlProps', 'customUI', 'xl/activeX', r'xl/media/.*\.emf')
        )
                             )

        if self.workbook.vba_archive:
            for name in set(self.workbook.vba_archive.namelist()) - self.vba_modified:
                if ARC_VBA.match(name):
                    # print("VBA**************", name)
                    if name not in self._archive.namelist():
                        if hasattr(self.workbook, "checkbox_values") and name.startswith("xl/activeX/activeX") and name.endswith(".bin"):
                            idx = int(name.replace("xl/activeX/activeX", "").replace(".bin", ""))
                            byte_array = bytearray(self.workbook.vba_archive.read(name))
                            if idx in self.workbook.checkbox_values:
                                byte_idx = 56
                                if idx == 1:
                                    byte_idx = 60
                                # print(idx, type(byte_array), byte_array[byte_idx], self.workbook.checkbox_values)
                                if self.workbook.checkbox_values[idx] == True:
                                    byte_array[byte_idx] = byte_array[byte_idx] | 1
                                    # print("Switching true")
                                else:
                                    byte_array[byte_idx] = byte_array[byte_idx] & 254
                                    # print("Switching false")
                                # print(idx, type(byte_array), byte_array[byte_idx])
                            self._archive.writestr(name, bytes(byte_array))
                        else:
                            self._archive.writestr(name, self.workbook.vba_archive.read(name))
                    if name == "xl/vbaProject.bin":
                        self.manifest.append_manual("/" + name, VBA)


    def _write_printer_settings(self):
        if self.workbook._printer_settings is not None:
            for key in self.workbook._printer_settings:
                printer_setting = self.workbook._printer_settings[key]
                full_filename = "{}/printerSettings{}.bin".format(PACKAGE_PRINTER_SETTINGS, key)
                self._archive.writestr(full_filename, printer_setting.read(full_filename))


    def _write_custom_xml(self):
        if self.workbook._custom_xml is not None:
            for full_filename in self.workbook._custom_xml:
                custom_xml = self.workbook._custom_xml[full_filename]
                self._archive.writestr(full_filename, custom_xml.read(full_filename))
                if full_filename.startswith(PACKAGE_CUSTOM_XML + "/" + "itemProps"):
                    self.manifest.append_manual("/" + full_filename, "application/vnd.openxmlformats-officedocument.customXmlProperties+xml")
        if self.workbook.arc_custom is not None:
            self._archive.writestr(ARC_CUSTOM, self.workbook.arc_custom.read(ARC_CUSTOM))
            self.manifest.append_manual("/" + ARC_CUSTOM, "application/vnd.openxmlformats-officedocument.custom-properties+xml")


    def _write_drawings_and_dependencies(self):
        if self.workbook._all_drawings is not None:
            for pathname in self.workbook._all_drawings:
                if pathname not in self._archive.namelist():
                    drawing = self.workbook._all_drawings[pathname]
                    self._archive.writestr(pathname, drawing.read(pathname))
        if self.workbook._all_drawings_rels is not None:
            for pathname in self.workbook._all_drawings_rels:
                if pathname not in self._archive.namelist():
                    rels = self.workbook._all_drawings_rels[pathname]
                    tree = rels.to_tree()
                    if LXML:
                        self._archive.writestr(pathname, tostring(tree, pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
                    else:
                        self._archive.writestr(pathname, tostring(tree))
        if self.workbook._all_drawing_dependencies is not None:
            for pathname in self.workbook._all_drawing_dependencies:
                if pathname not in self._archive.namelist():
                    drawing = self.workbook._all_drawing_dependencies[pathname]
                    self._archive.writestr(pathname, drawing.read(pathname))


    def _write_images(self):
        # delegate to object
        print("_write_images")
        for img in self._images:
            print("photo")
            self._archive.writestr(img.path[1:], img._data())


    def _write_charts(self):
        # delegate to object
        if len(self._charts) != len(set(self._charts)):
            raise InvalidFileException("The same chart cannot be used in more than one worksheet")
        for chart in self._charts:
            if LXML:
                self._archive.writestr(chart.path[1:], tostring(chart._write(), pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
            else:
                self._archive.writestr(chart.path[1:], tostring(chart._write()))
            self.manifest.append(chart)


    def _write_drawing(self, drawing):
        """
        Write a drawing
        """
        print("Writing drawing")
        self._drawings.append(drawing)
        if drawing._id is None:
            drawing._id = len(self._drawings)
        for chart in drawing.charts:
            self._charts.append(chart)
            chart._id = len(self._charts)
        for img in drawing.images:
            self._images.append(img)
            img._id = len(self._images)
        rels_path = get_rels_path(drawing.path)[1:]
        if LXML:
            self._archive.writestr(drawing.path[1:], tostring(drawing._write(), pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
        else:
            self._archive.writestr(drawing.path[1:], tostring(drawing._write()))
        if LXML:
            self._archive.writestr(rels_path, tostring(drawing._write_rels(), pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
        else:
            self._archive.writestr(rels_path, tostring(drawing._write_rels()))
        self.manifest.append(drawing)


    def _write_chartsheets(self):
        for idx, sheet in enumerate(self.workbook.chartsheets, 1):

            sheet._id = idx
            if LXML:
                xml = tostring(sheet.to_tree(), pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True)
            else:
                xml = tostring(sheet.to_tree())

            self._archive.writestr(sheet.path[1:], xml)
            self.manifest.append(sheet)

            if sheet._drawing:
                self._write_drawing(sheet._drawing)

                rel = Relationship(type="drawing", Target=sheet._drawing.path)
                rels = RelationshipList()
                rels.append(rel)
                tree = rels.to_tree()

                rels_path = get_rels_path(sheet.path[1:])
                if LXML:
                    self._archive.writestr(rels_path, tostring(tree, pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
                else:
                    self._archive.writestr(rels_path, tostring(tree))


    def _write_comment(self, ws):

        cs = CommentSheet.from_comments(ws._comments)
        self._comments.append(cs)
        total_vml_already_existing = 0
        if self.workbook.vba_archive is not None:
            for name in self.workbook.vba_archive.namelist():
                if "xl/drawings/vmlDrawing" in name:
                    total_vml_already_existing += 1

        cs._id = len(self._comments) + total_vml_already_existing

        if LXML:
            self._archive.writestr(cs.path[1:], tostring(cs.to_tree(), pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
        else:
            self._archive.writestr(cs.path[1:], tostring(cs.to_tree()))
        self.manifest.append(cs)

        if ws.legacy_drawing is None or self.workbook.vba_archive is None:
            ws.legacy_drawing = 'xl/drawings/vmlDrawing{0}.vml'.format(cs._id)
            vml = None
        else:
            vml = fromstring(self.workbook.vba_archive.read(ws.legacy_drawing))

        vml = cs.write_shapes(vml)

        self._archive.writestr(ws.legacy_drawing, vml)
        self.vba_modified.add(ws.legacy_drawing)

        comment_rel = Relationship(Id="comments", type=cs._rel_type, Target=cs.path)
        ws._rels.append(comment_rel)


    def write_worksheet(self, ws):
        if len(ws.drawings) == 0:
            ws._drawing = SpreadsheetDrawing()
        elif len(ws.drawings) == 1:
            for key in ws.drawings:
                ws._drawing = ws.drawings[key]
        else:
            raise Exception("Multiple drawings for a single worksheet")
        if len(ws._charts) > 0 or len(ws._images) > 0:
            ws._drawing.charts = ws._charts
            ws._drawing.images = ws._images
            ws._drawing._id = self.drawing_id
            self.drawing_id += 1
        if self.workbook.write_only:
            if not ws.closed:
                ws.close()
            writer = ws._writer
        else:
            writer = WorksheetWriter(ws)
            writer.write()

        ws._rels = writer._rels
        self._archive.write(writer.out, ws.path[1:])
        self.manifest.append(ws)

        if ws.ole_objects is not None:
            for ole_obj in ws.ole_objects.oleObject:
                self.manifest.append_manual('/' + ole_obj.path, ole_obj.mime_type)

        writer.cleanup()


    def _write_worksheets(self):

        pivot_caches = set()

        for idx, ws in enumerate(self.workbook.worksheets, 1):

            ws._id = idx
            self.write_worksheet(ws)

            if ws.ole_objects is not None:
                for ole_obj in ws.ole_objects.oleObject:

                    self._archive.write(ole_obj.oleObj, arcname=ole_obj.path)
                    # self._archive.write(ole_obj.objectPr.imgData, arcname=ole_obj.objectPr.path)

                    found = False
                    for r in ws._rels.Relationship:
                        if "oleObject" in r.Type and r.Target == ole_obj.path.replace("xl/", "../"):
                            r.Target = ole_obj.path.replace("xl/", "../")
                            r.BackupTarget = ole_obj.path.replace("xl/", "../")
                            found = True
                    if not found:
                        ole_obj_rel = Relationship(type="oleObject", Id="rId{}".format(ole_obj.id),
                                                Target=ole_obj.path.replace("xl/", "../"))
                        ws._rels.append(ole_obj_rel)



                    found = False
                    for r in ws._rels.Relationship:
                        if "image" in r.Type and r.Target == ole_obj.objectPr.path.replace("xl/", "../"):
                            r.Target = ole_obj.objectPr.path.replace("xl/", "../")
                            r.BackupTarget = ole_obj.objectPr.path.replace("xl/", "../")
                            found = True
                    if not found:
                        ole_obj_objectPr_rel = Relationship(type="image", Id="rId{}".format(ole_obj.objectPr.id),
                                                Target=ole_obj.objectPr.path.replace("xl/", "../"))
                        ws._rels.append(ole_obj_objectPr_rel)


            if ws._drawing is not None:
                # print(list(ws._drawing.twoCellAnchor))
                self._write_drawing(ws._drawing)

                found = False
                for r in ws._rels.Relationship:
                    if "drawing" in r.Type:
                        r.Target = ws._drawing.path.replace("/xl/", "../")
                        r.BackupTarget = ws._drawing.path.replace("/xl/", "../")
                        found = True
                if not found:
                    drawing_rel = Relationship(type="drawing", Id="rId{}".format(ws._drawing._id),
                                            Target=ws._drawing.path.replace("/xl/", "../"))
                    ws._rels.append(drawing_rel)

            if ws._comments:
                self._write_comment(ws)

            if ws.legacy_drawing is not None:
                target = "/" + ws.legacy_drawing
                target = target.replace("/xl/", "../")
                found = False
                for r in ws._rels.Relationship:
                    if "vmlDrawing" in r.Type:
                        r.Target = target
                        r.BackupTarget = target
                        shape_rel = r
                        found = True
                if not found:
                    shape_rel = Relationship(type="vmlDrawing", Id="rId{}".format(len(ws._rels) +  1),
                                            Target=target)
                    ws._rels.append(shape_rel)

                # shape_rel = Relationship(type="vmlDrawing", Id="rId3",
                #                          Target=target)
                # ws._rels.append(shape_rel)
                vml_out_path = ws.legacy_drawing
                print("Printing vml", vml_out_path)
                if hasattr(ws, "vml") and ws.vml is not None:
                    self._archive.writestr(vml_out_path, ws.vml)
                if hasattr(ws, "vml_rels") and ws.vml_rels is not None:
                    tree = ws.vml_rels.to_tree()
                    vml_rels_out_path = vml_out_path.replace("/drawings/", "/drawings/_rels/") + ".rels"
                    if LXML:
                        self._archive.writestr(vml_rels_out_path, tostring(tree, pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
                    else:
                        self._archive.writestr(vml_rels_out_path, tostring(tree))
                if hasattr(ws, "_legacy_images") and ws._legacy_images is not None:
                    for filename in ws._legacy_images:
                        print(filename, type(ws._legacy_images[filename]))
                        self._archive.write(ws._legacy_images[filename], arcname=filename)

            for t in ws._tables.values():
                self._tables.append(t)
                t.id = len(self._tables)
                t._write(self._archive)
                self.manifest.append(t)
                #TODO probably a bug
                if t._rel_id in ws._rels:
                    ws._rels[t._rel_id].Target = t.path
                    ws._rels[t._rel_id].BackupTarget = t.path

            for p in ws._pivots:
                if p.cache not in pivot_caches:
                    pivot_caches.add(p.cache)
                    p.cache._id = len(pivot_caches)

                self._pivots.append(p)
                p._id = len(self._pivots)
                p._write(self._archive, self.manifest)
                self.workbook._pivots.append(p)
                r = Relationship(Type=p.rel_type, Target=p.path)
                ws._rels.append(r)

            if ws._rels:
                rels_path = get_rels_path(ws.path)[1:]
                tree = ws._rels.to_tree()
                if LXML:
                    self._archive.writestr(rels_path, tostring(tree, pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
                else:
                    self._archive.writestr(rels_path, tostring(tree))


    def _write_external_links(self):
        # delegate to object
        """Write links to external workbooks"""
        wb = self.workbook
        for idx, link in enumerate(wb._external_links, 1):
            link._id = idx
            rels_path = get_rels_path(link.path[1:])

            xml = link.to_tree()
            if LXML:
                self._archive.writestr(link.path[1:], tostring(xml, pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
            else:
                self._archive.writestr(link.path[1:], tostring(xml))
            rels = RelationshipList()
            rels.append(link.file_link)
            if LXML:
                self._archive.writestr(rels_path, tostring(rels.to_tree(), pretty_print = True, xml_declaration = True, encoding='UTF-8', standalone=True))
            else:
                self._archive.writestr(rels_path, tostring(rels.to_tree()))
            self.manifest.append(link)


    def save(self):
        """Write data into the archive."""
        self.write_data()
        self._archive.close()


def save_workbook(workbook, filename):
    """Save the given workbook on the filesystem under the name filename.

    :param workbook: the workbook to save
    :type workbook: :class:`openpyxl.workbook.Workbook`

    :param filename: the path to which save the workbook
    :type filename: string

    :rtype: bool

    """
    archive = ZipFile(filename, 'w', ZIP_DEFLATED, allowZip64=True)
    writer = ExcelWriter(workbook, archive)
    writer.save()
    return True


@deprecated("Use a NamedTemporaryFile")
def save_virtual_workbook(workbook):
    """Return an in-memory workbook, suitable for a Django response."""
    tmp = TemporaryFile()
    archive = ZipFile(tmp, 'w', ZIP_DEFLATED, allowZip64=True)

    writer = ExcelWriter(workbook, archive)
    writer.save()

    tmp.seek(0)
    virtual_workbook = tmp.read()
    tmp.close()

    return virtual_workbook
