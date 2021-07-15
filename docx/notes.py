# encoding: utf-8

"""
The |EndnoteReference| and |FootnoteReference| objects, and related proxy
classes.
"""

from __future__ import absolute_import, print_function, unicode_literals

from docx.shared import Parented


class Endnote(Parented):
    """
    Proxy class for a WordprocessingML ``<w:endnote>`` element.
    """
    def __init__(self, note, parent):
        super(Endnote, self).__init__(parent)
        self._element = note


class EndnoteReference(Parented):
    """
    Proxy class for a WordprocessingML ``<w:endnoteReference>`` element.
    """
    def __init__(self, ref, parent):
        super(EndnoteReference, self).__init__(parent)
        self._element = self._ref = ref

    # XXX need to update this in line with footnotes (should use a base class?)
    @property
    def markdown(self):
        return '{{endnote=%d}}' % self._ref.id


class Footnote(Parented):
    """
    Proxy class for a WordprocessingML ``<w:footnote>`` element.
    """
    def __init__(self, note, parent):
        super(Footnote, self).__init__(parent)
        self._element = note


class FootnoteReference(Parented):
    """
    Proxy class for a WordprocessingML ``<w:footnoteReference>`` element.
    """
    def __init__(self, ref, parent):
        super(FootnoteReference, self).__init__(parent)
        self._element = self._ref = ref

    # XXX this currently only supports inline footnotes
    @property
    def markdown(self):
        # XXX it would be better if document.footnotes was a dict keyed by
        #     footnote id (a lazy property?)
        footnotes = self.part.document.footnotes
        matching = [f for f in footnotes if f._element.id == self._ref.id]
        assert len(matching) == 1
        footnote = matching[0]
        # return '{{footnote=%d,text=%r}}' % (self._ref.id,
        #                                     footnote.markdown.strip())
        return '^[%s]' % footnote.markdown.strip()
