# encoding: utf-8

"""
The |FldSimple| object and related proxy classes.
"""

from __future__ import absolute_import, print_function, unicode_literals

import re

from ..shared import Parented


class SimpleField(Parented):
    """
    Proxy class for a WordprocessingML ``<w:fldSimple>`` element.
    """
    def __init__(self, field, parent):
        super(SimpleField, self).__init__(parent)
        self._element = field

    @property
    def markdown(self):
        return '{{simpleField|%s|%s}}' % (self._element.instr.strip(),
                                          ''.join(i.markdown for i in
                                                  self.items))


class FieldChar(Parented):
    """
    Proxy class for a WordprocessingML ``<w:fldChar>`` element.
    """
    def __init__(self, field, parent):
        super(FieldChar, self).__init__(parent)
        self._element = field

    @property
    def markdown(self):
        return '{{fieldChar|%s}}' % self._element.fldCharType


class FieldCode(Parented):
    """
    Proxy class for a WordprocessingML ``<w:instrText>`` element.
    """
    def __init__(self, field, parent):
        super(FieldCode, self).__init__(parent)
        self._element = field

    @property
    def markdown(self):
        return '{{fieldCode|%s}}' % re.sub(r' +', ' ',
                                           self._element.text.strip())
