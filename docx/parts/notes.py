# encoding: utf-8

"""Endnotes and footnotes part objects"""

from __future__ import absolute_import, division, print_function, unicode_literals

import os

from docx.opc.constants import CONTENT_TYPE as CT
from docx.oxml import parse_xml
from docx.parts.story import BaseStoryPart


class EndnotesPart(BaseStoryPart):
    """Endnotes definitions."""

    @classmethod
    def new(cls, package):
        """Return newly created endnotes part."""
        partname = package.next_partname("/word/endnotes.xml")
        content_type = CT.WML_ENDNOTES
        element = parse_xml(cls._default_endnotes_xml())
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_endnotes_xml(cls):
        """Return bytes containing XML for a default endnotes part."""
        path = os.path.join(
                os.path.split(__file__)[0], '..', 'templates',
                'default-endnotes.xml'
        )
        with open(path, 'rb') as f:
            xml_bytes = f.read()
        return xml_bytes


class FootnotesPart(BaseStoryPart):
    """Footnotes definitions."""

    @classmethod
    def new(cls, package):
        """Return newly created footnotes part."""
        partname = package.next_partname("/word/footnotes.xml")
        content_type = CT.WML_FOOTNOTES
        element = parse_xml(cls._default_footnotes_xml())
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_footnotes_xml(cls):
        """Return bytes containing XML for a default footnotes part."""
        path = os.path.join(
                os.path.split(__file__)[0], '..', 'templates',
                'default-footnotes.xml'
        )
        with open(path, 'rb') as f:
            xml_bytes = f.read()
        return xml_bytes
