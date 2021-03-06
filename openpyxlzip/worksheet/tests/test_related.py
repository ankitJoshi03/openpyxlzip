# Copyright (c) 2010-2020 openpyxlzip

from openpyxlzip.tests.helper import compare_xml

from openpyxlzip.xml.functions import tostring


def test_related():
    from ..related import Related
    rel = Related(id="rId1")
    expected = """
    <drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
    """
    xml = tostring(rel.to_tree("drawing"))
    diff = compare_xml(xml, expected)
    assert diff is None, diff
