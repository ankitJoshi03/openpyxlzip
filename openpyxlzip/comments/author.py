# Copyright (c) 2010-2020 openpyxlzip


from openpyxlzip.descriptors.serialisable import Serialisable
from openpyxlzip.descriptors import (
    Sequence,
    Alias
)


class AuthorList(Serialisable):

    tagname = "authors"

    author = Sequence(expected_type=str)
    authors = Alias("author")

    def __init__(self,
                 author=(),
                ):
        self.author = author
