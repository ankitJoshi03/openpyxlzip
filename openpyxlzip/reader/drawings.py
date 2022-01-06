from __future__ import absolute_import
# Copyright (c) 2010-2020 openpyxlzip


from io import BytesIO
from warnings import warn

from openpyxlzip.xml.functions import fromstring
from openpyxlzip.xml.constants import IMAGE_NS
from openpyxlzip.packaging.relationship import get_rel, get_rels_path, get_dependents
from openpyxlzip.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxlzip.drawing.image import Image, PILImage
from openpyxlzip.chart.chartspace import ChartSpace
from openpyxlzip.chart.reader import read_chart


def find_images(archive, path):
    """
    Given the path to a drawing file extract charts and images

    Ingore errors due to unsupported parts of DrawingML
    """

    src = archive.read(path)
    tree = fromstring(src)
    try:
        drawing = SpreadsheetDrawing.from_tree(tree)
    except TypeError:
        warn("DrawingML support is incomplete and limited to charts and images only. Shapes and drawings will be lost.")
        return None, [], []

    rels_path = get_rels_path(path)
    deps = []
    if rels_path in archive.namelist():
        deps = get_dependents(archive, rels_path)

    charts = []
    for rel in drawing._chart_rels:
        cs = get_rel(archive, deps, rel.id, ChartSpace)
        chart = read_chart(cs)
        chart.anchor = rel.anchor
        charts.append(chart)

    images = []
    if not PILImage: # Pillow not installed, drop images
        return drawing, charts, images

    for rel in drawing._blip_rels:
        dep = deps[rel.embed]
        if dep.Type == IMAGE_NS:
            try:
                image = Image(BytesIO(archive.read(dep.target)))
                image.filename = dep.target
            except OSError:
                msg = "The image {0} will be removed because it cannot be read".format(dep.target)
                warn(msg)
                continue
            #MattJ I think this should not be the case, although this does seem to add files
            # if image.format.upper() == "WMF": # cannot save
            #     msg = "{0} image format is not supported so the image is being dropped".format(image.format)
            #     warn(msg)
            #     continue
            image.anchor = rel.anchor
            images.append(image)
    return drawing, charts, images
