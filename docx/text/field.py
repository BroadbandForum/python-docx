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
    def items(self):
        """
        Sequence of all child items.
        """
        items = []
        for elem in self._element:
            # XXX want a factory method
            from ..oxml.ns import qn
            from .run import Run
            cls_map = {
                qn('w:r'): Run}
            cls = cls_map.get(elem.tag)
            if cls is None:
                # XXX need logging
                if elem.tag not in cls_map:
                    import sys
                    sys.stderr.write("%s: couldn't find class for element "
                                     "%r\n" % (self.__class__.__name__,
                                               elem.tag))
            else:
                items += [cls(elem, self)]
        return items

    @property
    def text(self):
        return '{{simpleField|%s|%s}}' % (self._element.instr.strip(),
                                          ''.join(i.text for i in self.items))


class FieldChar(Parented):
    """
    Proxy class for a WordprocessingML ``<w:fldChar>`` element.
    """
    def __init__(self, field, parent):
        super(FieldChar, self).__init__(parent)
        self._element = field

    @property
    def text(self):
        return '{{fieldChar|%s}}' % self._element.fldCharType


class FieldCode(Parented):
    """
    Proxy class for a WordprocessingML ``<w:instrText>`` element.
    """
    def __init__(self, field, parent):
        super(FieldCode, self).__init__(parent)
        self._element = field

    @property
    def text(self):
        return '{{fieldCode|%s}}' % re.sub(r' +', ' ',
                                           self._element.text.strip())
