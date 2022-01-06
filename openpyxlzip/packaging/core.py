# Copyright (c) 2010-2020 openpyxlzip

import datetime

from openpyxlzip.compat import safe_string
from openpyxlzip.utils.datetime import (
    CALENDAR_WINDOWS_1900,
    to_ISO8601,
    from_ISO8601,
)
from openpyxlzip.descriptors import (
    String,
    DateTime,
    Alias,
    )
from openpyxlzip.descriptors.serialisable import Serialisable
from openpyxlzip.descriptors.nested import NestedText
from openpyxlzip.xml.functions import (Element, QName, tostring)
from openpyxlzip.xml.constants import (
    COREPROPS_NS,
    DCORE_NS,
    XSI_NS,
    DCTERMS_NS,
    DCTERMS_PREFIX
)

class NestedDateTime(DateTime, NestedText):

    expected_type = datetime.datetime

    def to_tree(self, tagname=None, value=None, namespace=None, elem_type=None):
        namespace = getattr(self, "namespace", namespace)
        if namespace is not None:
            tagname = "{%s}%s" % (namespace, tagname)
        el = Element(tagname)
        if value is not None:
            el.text = to_ISO8601(value)
            return el


class QualifiedDateTime(NestedDateTime):

    """In certain situations Excel will complain if the additional type
    attribute isn't set"""

    def to_tree(self, tagname=None, value=None, namespace=None, elem_type=None):
        el = super(QualifiedDateTime, self).to_tree(tagname, value, namespace)
        el.set("{%s}type" % XSI_NS, QName(DCTERMS_NS, "W3CDTF"))
        return el


class DocumentProperties(Serialisable):
    """High-level properties of the document.
    Defined in ECMA-376 Par2 Annex D
    """

    tagname = "coreProperties"
    namespace = COREPROPS_NS

    category = NestedText(expected_type=str, allow_none=True)
    contentStatus = NestedText(expected_type=str, allow_none=True)
    keywords = NestedText(expected_type=str, allow_none=True)
    lastModifiedBy = NestedText(expected_type=str, allow_none=True)
    lastPrinted = NestedDateTime(allow_none=True)
    revision = NestedText(expected_type=str, allow_none=True)
    version = NestedText(expected_type=str, allow_none=True)
    last_modified_by = Alias("lastModifiedBy")

    # Dublin Core Properties
    subject = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    title = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    creator = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    description = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    identifier = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    language = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    # Dublin Core Terms
    created = QualifiedDateTime(allow_none=True, namespace=DCTERMS_NS)
    modified = QualifiedDateTime(allow_none=True, namespace=DCTERMS_NS)

    __elements__ = ("creator", "title", "description", "subject","identifier",
                  "language", "created", "modified", "lastModifiedBy", "category",
                  "contentStatus", "version", "revision", "keywords", "lastPrinted",
                  )


    def __init__(self,
                 category=None,
                 contentStatus=None,
                 keywords=None,
                 lastModifiedBy=None,
                 lastPrinted=None,
                 revision=None,
                 version=None,
                 created=datetime.datetime.now(),
                 creator="openpyxlzip",
                 description=None,
                 identifier=None,
                 language=None,
                 modified=datetime.datetime.now(),
                 subject=None,
                 title=None,
                 ):
        self.contentStatus = contentStatus
        self.lastPrinted = lastPrinted
        self.revision = revision
        self.version = version
        self.creator = creator
        self.lastModifiedBy = lastModifiedBy
        self.modified = modified
        self.created = created
        self.title = title
        self.subject = subject
        self.description = description
        self.identifier = identifier
        self.language = language
        self.keywords = keywords
        self.category = category