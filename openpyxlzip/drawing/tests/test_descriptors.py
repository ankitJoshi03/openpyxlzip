# Copyright (c) 2010-2020 openpyxlzip


def test_color_descriptor():
    from ..colors import ColorChoiceDescriptor
    from ..colors import SRGBClr

    class DummyStyle(object):

        value = ColorChoiceDescriptor('value')

    style = DummyStyle()
    style.value = "efefef"
    style.value.RGB.val == "efefef"
