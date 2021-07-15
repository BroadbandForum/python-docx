# encoding: utf-8

"""
Custom element classes related to endnotes and footnotes parts (note
definitions) and note referen
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .simpletypes import ST_DecimalNumber
from ..oxml.xmlchemy import BaseOxmlElement, OneAndOnlyOne, \
    RequiredAttribute, ZeroOrMore


class CT_Endnotes(BaseOxmlElement):
    """
    ``<w:endnotes>`` element, the root element of the endnotes part, i.e.
    endnotes.xml.
    """
    endnote = ZeroOrMore('w:endnote')


class CT_Endnote(BaseOxmlElement):
    """
    ``<w:endnote>`` element, an element of the endnotes part, i.e. an
    endnote definition.
    """
    id = RequiredAttribute('w:id', ST_DecimalNumber)
    p = OneAndOnlyOne('w:p')


class CT_EndnoteReference(BaseOxmlElement):
    """
    ``<w:endnoteReference>`` element.
    """
    id = RequiredAttribute('w:id', ST_DecimalNumber)


class CT_Footnotes(BaseOxmlElement):
    """
    ``<w:footnotes>`` element, the root element of the footnotes part, i.e.
    footnotes.xml.
    """
    footnote = ZeroOrMore('w:footnote')


class CT_Footnote(BaseOxmlElement):
    """
    ``<w:footnote>`` element, an element of the footnotes part, i.e. a
    footnote definition.
    """
    id = RequiredAttribute('w:id', ST_DecimalNumber)
    p = OneAndOnlyOne('w:p')


class CT_FootnoteReference(BaseOxmlElement):
    """
    ``<w:footnoteReference>`` element.
    """
    id = RequiredAttribute('w:id', ST_DecimalNumber)
