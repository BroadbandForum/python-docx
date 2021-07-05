# encoding: utf-8

"""
The |Hyperlink| object and related proxy classes.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..blkcntnr import BlockItemContainer
from ..shared import Parented


class Hyperlink(Parented):
    """
    Proxy class for a WordprocessingML ``<w:hyperlink>`` element.
    """
    def __init__(self, hyperlink, parent):
        super(Hyperlink, self).__init__(parent)
        self._element = self._hyperlink = hyperlink

    @property
    def markdown(self):
        return '{{hyperlink|%s|%s}}' % (self._element.rid,
                                        self._element.r.text)
