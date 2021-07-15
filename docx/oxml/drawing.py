# encoding: utf-8

"""Custom element classes for drawings"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, ZeroOrOne
)


class CT_Drawing(BaseOxmlElement):
    """
    ``<w:drawing>`` element.
    """
    inline = ZeroOrOne('wp:inline')


class CT_EmbeddedObject(BaseOxmlElement):
    """
    ``<w:object>`` element.
    """
    ole_object = OneAndOnlyOne('o:OLEObject')


class CT_Picture(BaseOxmlElement):
    """
    ``<w:pict>`` element.
    """
