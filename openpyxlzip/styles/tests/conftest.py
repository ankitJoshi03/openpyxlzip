# Copyright (c) 2010-2020 openpyxlzip
import pytest


@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    import os
    here = os.path.split(__file__)[0]
    DATADIR = os.path.join(here, "data")
    from py._path.local import LocalPath
    return LocalPath(DATADIR)


# objects under test


@pytest.fixture
def FormatRule():
    """Formatting rule class"""
    from openpyxlzip.formatting.rules import FormatRule
    return FormatRule
