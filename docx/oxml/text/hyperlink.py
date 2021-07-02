# encoding: utf-8

"""Custom element classes for hyperlinks"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..simpletypes import ST_String
from ..xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne


class CT_Hyperlink(BaseOxmlElement):
    """
    ``<w:hyperlink>`` element.
    """

    rid = OptionalAttribute('r:id', ST_String)
    r = ZeroOrOne('w:r')
