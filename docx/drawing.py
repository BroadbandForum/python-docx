# encoding: utf-8

"""
The |Drawing| object and related proxy classes.
"""

from __future__ import absolute_import, print_function, unicode_literals

from .blkcntnr import BlockItemContainer
from .shared import Parented


class Drawing(Parented):
    """
    Proxy class for a WordprocessingML ``<w:drawing>`` element.
    """
    def __init__(self, hyperlink, parent):
        super(Drawing, self).__init__(parent)
        self._element = self._hyperlink = hyperlink

    @property
    def text(self):
        return '{{drawing|XXX}}'
