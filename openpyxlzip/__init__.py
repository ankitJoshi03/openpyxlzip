# Copyright (c) 2010-2020 openpyxlzip


from openpyxlzip.compat.numbers import NUMPY, PANDAS
from openpyxlzip.xml import DEFUSEDXML, LXML
from openpyxlzip.workbook import Workbook
from openpyxlzip.reader.excel import load_workbook as open
from openpyxlzip.reader.excel import load_workbook
import openpyxlzip._constants as constants

# Expose constants especially the version number

__author__ = constants.__author__
__author_email__ = constants.__author_email__
__license__ = constants.__license__
__maintainer_email__ = constants.__maintainer_email__
__url__ = constants.__url__
__version__ = constants.__version__
