# encoding: utf-8

"""Custom element classes for fields (and properties)"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..simpletypes import ST_String
from ..xmlchemy import BaseOxmlElement, RequiredAttribute


class CT_FldSimple(BaseOxmlElement):
    """
    ``<w:fldSimple>`` element, used for document properties etc.
    """
    instr = RequiredAttribute('w:instr', ST_String)


class CT_FldChar(BaseOxmlElement):
    """
    ``<w:fldChar>`` element, used to bracket field references etc.
    """
    fldCharType = RequiredAttribute('w:fldCharType', ST_String)


class CT_InstrText(BaseOxmlElement):
    """
    ``<w:instrText>`` element, used for cross-references etc.
    """
