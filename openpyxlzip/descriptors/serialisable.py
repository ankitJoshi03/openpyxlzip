# copyright openpyxlzip 2010-2015

from copy import copy
from keyword import kwlist
KEYWORDS = frozenset(kwlist)

from . import Descriptor
from . import _Serialiasable
from .sequence import (
    Sequence,
    NestedSequence,
    MultiSequencePart,
)
from .namespace import namespaced

from openpyxlzip.compat import safe_string
from openpyxlzip.xml.functions import (
    Element,
    localname,
    get_namespace,
)
from openpyxlzip.xml.constants import (
    SHEET_MAIN_NS,
    X14_NS,
    DRAWING_NS,
    DRAWING_14_NS,
    DRAWING_16_NS,
    X15AC_NS,
    MC_NS,
)
from openpyxlzip.xml.schema import ALL_DEFINITIONS, ROOT_ELEMS, ROOT_NAMESPACES, ALL_NAMESPACES_PREFIXES

seq_types = (list, tuple)
TOP_LEVEL_NSMAP_TYPES = set(["workbook", "worksheet", "table", "styleSheet", "datastoreItem", "externalLink", "wsDr", "coreProperties"])
SUB_LEVEL_NSMAP_TYPES = set([   (SHEET_MAIN_NS, "sortState"),
                            (SHEET_MAIN_NS, "ext"),
                            (X14_NS,"dataValidations"),
                            (DRAWING_NS, "theme"),
                            (SHEET_MAIN_NS, "externalBook"),
                            (X15AC_NS, "absPath"),
                            (MC_NS, "AlternateContent")])

class Serialisable(_Serialiasable):
    """
    Objects can serialise to XML their attributes and child objects.
    The following class attributes are created by the metaclass at runtime:
    __attrs__ = attributes
    __nested__ = single-valued child treated as an attribute
    __elements__ = child elements
    """

    __attrs__ = None
    __nested__ = None
    __elements__ = None
    __namespaced__ = None

    idx_base = 0

    @property
    def tagname(self):
        raise(NotImplementedError)

    @property
    # Getter method
    def extra_attr(self):
        return self.__extra_attr

    # Setter method
    @extra_attr.setter
    def extra_attr(self, val):
        self.__extra_attr = val

    # Deleter method
    @extra_attr.deleter
    def extra_attr(self):
       del self.__extra_attr

    @property
    # Getter method
    def extra_elem(self):
        return self.__extra_elem

    # Setter method
    @extra_elem.setter
    def extra_elem(self, val):
        self.__extra_elem = val

    # Deleter method
    @extra_elem.deleter
    def extra_elem(self):
       del self.__extra_elem

    @property
    # Getter method
    def nsmaps(self):
        return self.__nsmaps

    # Setter method
    @nsmaps.setter
    def nsmaps(self, val):
        self.__nsmaps = val

    # Deleter method
    @nsmaps.deleter
    def nsmaps(self):
       del self.__nsmaps

    @property
    # Getter method
    def elem_order(self):
        return self.__elem_order

    # Setter method
    @elem_order.setter
    def elem_order(self, val):
        self.__elem_order = val

    # Deleter method
    @elem_order.deleter
    def elem_order(self):
       del self.__elem_order

    namespace = None

    @classmethod
    def from_tree(cls, node):
        """
        Create object from XML
        """
        # strip known namespaces from attributes
        attrib = dict(node.attrib)
        extra_attr = {}
        extra_elem = {}
        elem_order = {}
        for key, ns in cls.__namespaced__:
            if ns in attrib:
                attrib[key] = attrib[ns]
                del attrib[ns]

        # strip attributes with unknown namespaces
        for key in list(attrib):
            if key.startswith('{'):
                extra_attr[key] = attrib[key]
                del attrib[key]
            elif key in KEYWORDS:
                attrib["_" + key] = attrib[key]
                del attrib[key]
            elif "-" in key:
                n = key.replace("-", "_")
                attrib[n] = attrib[key]
                del attrib[key]
            else:
                extra_attr[key] = attrib[key]

        if node.text and "attr_text" in cls.__attrs__:
            attrib["attr_text"] = node.text

        tags_seen = set()
        for el in node:
            tag = localname(el)
            if tag in KEYWORDS:
                tag = "_" + tag
            desc = getattr(cls, tag, None)
            if desc is None or isinstance(desc, property):
                if el.tag not in tags_seen:
                    elem_order[len(elem_order)] = el.tag
                    tags_seen.add(el.tag)
                if el.tag in extra_elem:
                    extra_elem[el.tag].append(el)
                else:
                    extra_elem[el.tag] = [el]
                continue

            if el.tag not in tags_seen:
                elem_order[len(elem_order)] = el.tag
                tags_seen.add(el.tag)

            if hasattr(desc, 'from_tree'):
                #descriptor manages conversion
                obj = desc.from_tree(el)
            else:
                if hasattr(desc.expected_type, "from_tree"):
                    #complex type
                    obj = desc.expected_type.from_tree(el)
                else:
                    #primitive
                    obj = el.text

            if isinstance(desc, NestedSequence):
                attrib[tag] = obj
            elif isinstance(desc, Sequence):
                attrib.setdefault(tag, [])
                attrib[tag].append(obj)
            elif isinstance(desc, MultiSequencePart):
                attrib.setdefault(desc.store, [])
                attrib[desc.store].append(obj)
            else:
                attrib[tag] = obj

        new_obj = cls(**attrib)
        new_obj.extra_attr = extra_attr
        new_obj.extra_elem = extra_elem
        new_obj.elem_order = elem_order
        if hasattr(node, "nsmap"):
            if localname(node) in TOP_LEVEL_NSMAP_TYPES or (localname(node), get_namespace(node)) in SUB_LEVEL_NSMAP_TYPES:
                new_obj.nsmaps = node.nsmap
        return new_obj


    def to_tree(self, tagname=None, idx=None, namespace=None, verbose=False, elem_type=None):
        if tagname is None:
            tagname = self.tagname

        # keywords have to be masked
        if tagname.startswith("_"):
            tagname = tagname[1:]

        #Get the proper namespace and tagname
        tagname = namespaced(self, tagname, namespace)
        namespace = getattr(self, "namespace", namespace)

        #Get the attributes
        attrs = dict(self)
        for key, ns in self.__namespaced__:
            if key in attrs:
                #Change to namespaced version
                attrs[ns] = attrs[key]
                del attrs[key]

        #Add the attrs that weren't documented in the class
        if (hasattr(self, "extra_attr")) and len(self.extra_attr) > 0:
            for key in self.extra_attr:
                if key not in attrs: #Don't overwrite
                    attrs[key] = self.extra_attr[key]

        #Make sure we only use the namespaced version
        to_delete = []
        for key in attrs:
            if key != localname(key) and localname(key) in attrs:
                to_delete.append(localname(key))
        for key in to_delete:
            del attrs[key]

        #Make the element, if possible include the namespace
        if hasattr(self, "nsmaps"):
            temp_nsmap = {}
            for key in self.nsmaps:
                if self.nsmaps[key] != SHEET_MAIN_NS:
                    temp_nsmap[key] = self.nsmaps[key]
            el = Element(tagname, attrs, nsmap=temp_nsmap)
        else:
            el = Element(tagname, attrs)
        #Add the text
        if "attr_text" in self.__attrs__:
            el.text = safe_string(getattr(self, "attr_text"))

        #Check if we can get the order
        # if tagname in ROOT_ELEMS:
        #     # print("Found tagname as root", tagname)
        #     pass
        # else:
        #     # print("Could not find tagname as root", tagname)
        #     pass
        # if elem_type is not None:
        #     if elem_type in ALL_DEFINITIONS:
        #         print("Found elem", elem_type)
        #         pass
        #     else:
        #         pass
        #         # print("Could not find", elem_type)
        added_tags = set()
        if hasattr(self, "elem_order"):
            # if verbose or "wsDr" in tagname:
            #     print("Using elem order", self.elem_order)
            found_in_elem_order = set()
            for i in range(len(self.elem_order)):
                child_tag = self.elem_order[i]
                full_tag = self.elem_order[i]
                found_in_elem_order.add(full_tag)
                if full_tag in added_tags:
                    raise Exception("Trying to add", full_tag, added_tags)
                added_tags.add(full_tag)
                local_child_tag = localname(self.elem_order[i])
                if local_child_tag in self.__elements__:
                    child_tag = local_child_tag
                if "_" + local_child_tag in self.__elements__:
                    child_tag = "_" + local_child_tag
                if child_tag in self.__elements__:
                    desc = getattr(self.__class__, child_tag, None)
                    obj = getattr(self, child_tag)
                    if hasattr(desc, "namespace") and hasattr(obj, 'namespace'):
                        obj.namespace = desc.namespace

                    if isinstance(obj, seq_types):
                        if isinstance(desc, NestedSequence):
                            # wrap sequence in container
                            if not obj:
                                continue
                            nodes = [desc.to_tree(child_tag, obj, namespace)]
                        elif isinstance(desc, Sequence):
                            # sequence
                            desc.idx_base = self.idx_base
                            nodes = (desc.to_tree(child_tag, obj, namespace))
                        else: # property
                            nodes = (v.to_tree(child_tag, namespace) for v in obj)
                        for node in nodes:
                            node.tag = full_tag.replace("{" + SHEET_MAIN_NS + "}", "")
                            for sub_child in node.iterchildren():
                                if SHEET_MAIN_NS in sub_child.tag:
                                    old_tag = sub_child.tag
                                    sub_child.tag = sub_child.tag.replace("{" + SHEET_MAIN_NS + "}", "")
                            el.append(node)
                    else:
                        if child_tag in self.__nested__:
                            node = desc.to_tree(child_tag, obj, namespace)
                        elif obj is None:
                            continue
                        else:
                            node = obj.to_tree(child_tag)
                        if node is not None:
                            node.tag = full_tag.replace("{" + SHEET_MAIN_NS + "}", "")
                            for sub_child in node.iterchildren():
                                if SHEET_MAIN_NS in sub_child.tag:
                                    old_tag = sub_child.tag
                                    sub_child.tag = sub_child.tag.replace("{" + SHEET_MAIN_NS + "}", "")
                            el.append(node)
                elif hasattr(self, "extra_elem") and child_tag in self.extra_elem:
                    for original_child_node in self.extra_elem[child_tag]:
                        child_node = copy(original_child_node)
                        child_node.tag = child_node.tag.replace("{" + SHEET_MAIN_NS + "}", "")
                        for sub_child in child_node.iterchildren():
                            if SHEET_MAIN_NS in sub_child.tag:
                                old_tag = sub_child.tag
                                sub_child.tag = sub_child.tag.replace("{" + SHEET_MAIN_NS + "}", "")
                        if hasattr(child_node, "nsmap") and (localname(child_node), get_namespace(child_node)) in SUB_LEVEL_NSMAP_TYPES:
                            #This dele
                            for key in child_node.nsmap:
                                if child_node.nsmap[key] == SHEET_MAIN_NS:
                                    del child_node.nsmap[key]
                        elif hasattr(child_node, "nsmap"):
                            for key in child_node.nsmap:
                                del child_node.nsmap[key]
                        el.append(child_node)
                else:
                    # raise Exception("Unknown element", child_tag, local_child_tag)
                    pass
            if hasattr(self, "extra_elem"):
                for child_tag in self.extra_elem:
                    if child_tag not in found_in_elem_order:
                        for original_child_node in self.extra_elem[child_tag]:
                            child_node = copy(original_child_node)
                            child_node.tag = child_node.tag.replace("{" + SHEET_MAIN_NS + "}", "")
                            for sub_child in child_node.iterchildren():
                                if SHEET_MAIN_NS in sub_child.tag:
                                    old_tag = sub_child.tag
                                    sub_child.tag = sub_child.tag.replace("{" + SHEET_MAIN_NS + "}", "")
                            if hasattr(child_node, "nsmap") and (localname(child_node), get_namespace(child_node)) in SUB_LEVEL_NSMAP_TYPES:
                                #This dele
                                for key in child_node.nsmap:
                                    if child_node.nsmap[key] == SHEET_MAIN_NS:
                                        del child_node.nsmap[key]
                            elif hasattr(child_node, "nsmap"):
                                for key in child_node.nsmap:
                                    del child_node.nsmap[key]
                            el.append(child_node)

        #This is an element that we made some other way.
        else:
            # if verbose or "wsDr" in tagname:
            #     print("Not using elem order")
            if (hasattr(self, "extra_elem")) and len(self.extra_elem) > 0:
                for child_tag in self.extra_elem:
                    # if verbose or "wsDr" in tagname:
                    #     print("Extra elem", child_tag)
                    for original_child_node in self.extra_elem[child_tag]:
                        child_node = copy(original_child_node)
                        if hasattr(child_node, "nsmap") and (localname(child_node), get_namespace(child_node)) in SUB_LEVEL_NSMAP_TYPES:
                            for key in child_node.nsmap:
                                if child_node.nsmap[key] == SHEET_MAIN_NS:
                                    del child_node.nsmap[key]
                        elif hasattr(child_node, "nsmap"):
                            for key in child_node.nsmap:
                                del child_node.nsmap[key]
                        el.append(child_node)

            for child_tag in self.__elements__:
                # if verbose or "wsDr" in tagname:
                #     print("Elements", child_tag)
                desc = getattr(self.__class__, child_tag, None)
                obj = getattr(self, child_tag)
                if hasattr(desc, "namespace") and hasattr(obj, 'namespace'):
                    obj.namespace = desc.namespace

                if isinstance(obj, seq_types):
                    if isinstance(desc, NestedSequence):
                        # wrap sequence in container
                        if not obj:
                            continue
                        nodes = [desc.to_tree(child_tag, obj, namespace)]
                    elif isinstance(desc, Sequence):
                        # sequence
                        desc.idx_base = self.idx_base
                        nodes = (desc.to_tree(child_tag, obj, namespace))
                    else: # property
                        nodes = (v.to_tree(child_tag, namespace) for v in obj)
                    for node in nodes:
                        el.append(node)
                else:
                    if child_tag in self.__nested__:
                        node = desc.to_tree(child_tag, obj, namespace)
                    elif obj is None:
                        continue
                    else:
                        if not hasattr(obj, "to_tree"):
                            # print(obj, child_tag)
                            continue
                        node = obj.to_tree(child_tag)
                    if node is not None:
                        el.append(node)
        return el


    def __iter__(self):
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if attr.startswith("_"):
                attr = attr[1:]
            elif attr != "attr_text" and "_" in attr:
                desc = getattr(self.__class__, attr)
                if getattr(desc, "hyphenated", False):
                    attr = attr.replace("_", "-")
            if attr != "attr_text" and value is not None:
                yield attr, safe_string(value)


    def __eq__(self, other):
        if not self.__class__ == other.__class__:
            return False
        elif not dict(self) == dict(other):
            return False
        for el in self.__elements__:
            if getattr(self, el) != getattr(other, el):
                return False
        return True


    def __ne__(self, other):
        return not self == other


    def __repr__(self):
        s = u"<{0}.{1} object>\nParameters:".format(
            self.__module__,
            self.__class__.__name__
        )
        args = []
        for k in self.__attrs__ + self.__elements__:
            v = getattr(self, k)
            if isinstance(v, Descriptor):
                v = None
            args.append(u"{0}={1}".format(k, repr(v)))
        args = u", ".join(args)

        return u"\n".join([s, args])


    def __hash__(self):
        fields = []
        for attr in self.__attrs__ + self.__elements__:
            val = getattr(self, attr)
            if isinstance(val, list):
                val = tuple(val)
            fields.append(val)

        return hash(tuple(fields))


    def __add__(self, other):
        if type(self) != type(other):
            raise TypeError("Cannot combine instances of different types")
        vals = {}
        for attr in self.__attrs__:
            vals[attr] = getattr(self, attr) or getattr(other, attr)
        for el in self.__elements__:
            a = getattr(self, el)
            b = getattr(other, el)
            if a and b:
                vals[el] = a + b
            else:
                vals[el] = a or b
        return self.__class__(**vals)


    def __copy__(self):
        # serialise to xml and back to avoid shallow copies
        xml = self.to_tree(tagname="dummy")
        cp = self.__class__.from_tree(xml)
        # copy any non-persisted attributed
        for k in self.__dict__:
            if k not in self.__attrs__ + self.__elements__:
                v = copy(getattr(self, k))
                setattr(cp, k, v)
        return cp
