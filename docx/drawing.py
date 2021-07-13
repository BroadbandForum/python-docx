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
