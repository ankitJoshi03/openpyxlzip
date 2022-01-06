# Copyright (c) 2010-2020 openpyxlzip

from .. protection import hash_password


def test_password():
    enc = hash_password('secret')
    assert enc == 'DAA7'
