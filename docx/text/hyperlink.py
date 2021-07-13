# encoding: utf-8

"""
The |Hyperlink| object and related proxy classes.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..compat import is_string
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
        # XXX this is similar to the Paragraph.markdown logic
        text = ''
        markdown_lst = [item.markdown for item in self.items]
        for value in markdown_lst:
            if not isinstance(value, list):
                value = [value]
            for val in value:
                if is_string(val):
                    text += val

        anchor = self._element.anchor
        rid = self._element.rid
        assert (anchor or rid) and not (anchor and rid)
        if anchor:
            # XXX use of [] rather than () is deliberate here; bookmark
            #     resolution will change them to () if necessary
            return '[%s][%s]' % (text, '{{ref|%s}}' % anchor)
        else:
            assert rid in self.part.rels
            target_ref = self.part.rels[rid].target_ref
            if text == target_ref:
                return '<%s>' % target_ref
            else:
                return '[%s](%s)' % (text, target_ref)
