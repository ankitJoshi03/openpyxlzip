# Copyright (c) 2010-2020 openpyxlzip

"""
XML compatability functions
"""

# Python stdlib imports
import re
from functools import partial

from openpyxlzip import DEFUSEDXML, LXML

if LXML is True:
    from lxml.etree import (
    Element,
    SubElement,
    register_namespace,
    QName,
    xmlfile,
    XMLParser,
    iterparse
    )
    from lxml.etree import fromstring, tostring
    # do not resolve entities
    safe_parser = XMLParser(resolve_entities=False)
    fromstring = partial(fromstring, parser=safe_parser)

else:
    from xml.etree.ElementTree import (
    Element,
    SubElement,
    fromstring,
    tostring,
    QName,
    register_namespace,
    iterparse
    )
    from et_xmlfile import xmlfile
    if DEFUSEDXML is True:
        from defusedxml.ElementTree import fromstring
        from defusedxml.ElementTree import iterparse


from openpyxlzip.xml.constants import (
    CHART_NS,
    DRAWING_NS,
    SHEET_DRAWING_NS,
    CHART_DRAWING_NS,
    SHEET_MAIN_NS,
    REL_NS,
    VTYPES_NS,
    COREPROPS_NS,
    DCTERMS_NS,
    DCTERMS_PREFIX,
    XML_NS,
    XR_NS,
    XR2_NS,
    XR3_NS,
    XR6_NS,
    XR10_NS,
    X14_NS,
    X14AC_NS,
    X15_NS,
    X15AC_NS,
    MC_NS,
    XCALCF_NS,
    XMLNS_NS,
    DRAWING_14_NS,
    DRAWING_16_NS,
)

register_namespace(DCTERMS_PREFIX, DCTERMS_NS)
register_namespace('dcmitype', 'http://purl.org/dc/dcmitype/')
register_namespace('cp', COREPROPS_NS)
register_namespace('c', CHART_NS)
register_namespace('a', DRAWING_NS)
register_namespace('a14', DRAWING_14_NS)
register_namespace('a16', DRAWING_16_NS)
register_namespace('s', SHEET_MAIN_NS)
register_namespace('r', REL_NS)
register_namespace('vt', VTYPES_NS)
register_namespace('xdr', SHEET_DRAWING_NS)
register_namespace('cdr', CHART_DRAWING_NS)
register_namespace('xml', XML_NS)
#New namespaces to support later versions
register_namespace('xr', XR_NS)
register_namespace('xr2', XR2_NS)
register_namespace('xr3', XR3_NS)
register_namespace('xr6', XR6_NS)
register_namespace('xr10', XR10_NS)
register_namespace('x14', X14_NS)
register_namespace('x14ac', X14AC_NS)
register_namespace('x15', X15_NS)
register_namespace('x15ac', X15AC_NS)
register_namespace('mc', MC_NS)
register_namespace('xcalcf', XCALCF_NS)


tostring = partial(tostring, encoding="UTF-8")

NS_REGEX = re.compile("({(?P<namespace>.*)})?(?P<localname>.*)")

def localname(node):
    if type(node) is str:
        tag = node
    else:
        tag = node.tag
    if callable(tag):
        return "comment"
    m = NS_REGEX.match(tag)
    return m.group('localname')


def get_namespace(node):
    if callable(node.tag):
        return "comment"
    m = NS_REGEX.match(node.tag)
    return m.group('namespace')


def whitespace(node):
    if node.text != node.text.strip():
        node.set("{%s}space" % XML_NS, "preserve")
