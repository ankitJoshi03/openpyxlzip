# Copyright (c) 2010-2020 openpyxlzip


try:
    from abc import ABC
except ImportError:
    from abc import ABCMeta
    ABC = ABCMeta('ABC', (object, ), {})
