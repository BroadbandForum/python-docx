# encoding: utf-8

"""
The |Drawing| object and related proxy classes.
"""

from __future__ import absolute_import, print_function, unicode_literals

import re

from .shared import Parented


class Drawing(Parented):
    """
    Proxy class for a WordprocessingML ``<w:drawing>`` element.
    """
    def __init__(self, drawing, parent):
        super(Drawing, self).__init__(parent)
        self._element = self._drawing = drawing

    @property
    def markdown(self):
        name, descr = 2 * ('unknown',)
        inline = self._drawing.inline
        if inline is not None:
            docPr = inline.docPr
            name = docPr.name
            descr = docPr.descr
            if descr is not None:
                descr = re.sub(r'\n.*', r'', docPr.descr)
        return '{{drawing|name=%s|descr=%s}}' % (name, descr)


class EmbeddedObject(Parented):
    """
    Proxy class for a WordprocessingML ``<w:object>`` element.
    """
    def __init__(self, obj, parent):
        super(EmbeddedObject, self).__init__(parent)
        self._element = self._obj = obj

    @property
    def markdown(self):
        # XXX this hasn't been wrapped yet
        ole_object = self._element.ole_object
        prog_id = ole_object.attrib['ProgID']
        return '{{embeddedObject|progId=%s}}' % prog_id


class Picture(Parented):
    """
    Proxy class for a WordprocessingML ``<w:pict>`` element.
    """
    def __init__(self, pict, parent):
        super(Picture, self).__init__(parent)
        self._element = self._pict = pict

    # XXX I don't think that this is used yet
    @property
    def markdown(self):
        return '{{picture|XXX}}'
