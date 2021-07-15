# encoding: utf-8

"""
The |FldSimple| object and related proxy classes.
"""

from __future__ import absolute_import, print_function, unicode_literals

import re

from ..shared import Parented
from ..text.parfmt import BookmarkStart


class SimpleField(Parented):
    """
    Proxy class for a WordprocessingML ``<w:fldSimple>`` element.
    """
    def __init__(self, field, parent):
        super(SimpleField, self).__init__(parent)
        self._element = field

    @property
    def markdown(self):
        # 'instr' can have leading and trailing space, so ignore this and
        # then split into words (the first word indicates the field type)
        words = self._element.instr.strip().split()

        # custom property: second word is the (quoted) property name
        # example: DOCPROPERTY "BBF-name" \* MERGEFORMAT
        if words[0] == 'DOCPROPERTY':
            assert len(words) > 1
            name = re.sub(r'^"|"$', '', words[1])
            text = '%%%s%%' % self.mapname(name)

        # sequence number: second word is the counter
        # example: SEQ Table \* ARABIC
        # XXX will need some special logic for table and figure captions, for
        #     which the leading 'Table ' and 'Figure ' are plain text that
        #     precede this field
        # XXX for now use a fake variable; assume that this will be removed
        #     later
        # XXX what about header numbers?
        elif words[0] == 'SEQ':
            assert len(words) > 1
            name = words[1]
            text = '%%%s%%' % self.mapname(name, prefix='docx-seq-')

        # otherwise assume core property
        # example: AUTHOR  \* MERGEFORMAT
        else:
            name = words[0].lower()
            text = '%%%s%%' % self.mapname(name, prefix='docx-')

        # this is a metadata variable reference (not the field value)
        return text

    @staticmethod
    def mapname(name, prefix=''):
        """
        Map 'BBF-name' etc. to 'bbfName'.
        """
        return re.sub(r'([-_ ].)', lambda m: m.group(1)[1].upper(),
                      (prefix + name).lower())

    def __str__(self):
        return '%s' % self._element.instr.strip()

    def __repr__(self):
        return '<%s %s>' % (type(self).__name__, str(self))


class FieldChar(Parented):
    """
    Proxy class for a WordprocessingML ``<w:fldChar>`` element.
    """
    def __init__(self, field, parent):
        super(FieldChar, self).__init__(parent)
        self._element = field

    @property
    def markdown(self):
        return self

    def __str__(self):
        return '%s' % self._element.fldCharType

    def __repr__(self):
        return '<%s %s>' % (type(self).__name__, str(self))


class FieldCode(Parented):
    """
    Proxy class for a WordprocessingML ``<w:instrText>`` element.
    """
    def __init__(self, field, parent):
        super(FieldCode, self).__init__(parent)
        self._element = field

    @property
    def markdown(self):
        # text can have leading and trailing space, so ignore this and
        # then split into words (the first word indicates the field type)
        words = self._element.text.strip().split()

        # empty field codes can happen, so can plain 'DOCPROPERTY',
        # 'REF' etc., so just ignore them
        # XXX return self for debugging
        if len(words) < 2:
            text = self

        # custom property: second word is the (quoted) property name
        # example: DOCPROPERTY "BBF-name" \* MERGEFORMAT
        # XXX this logic is copied from SimpleField; should re-factor
        elif words[0] == 'DOCPROPERTY':
            assert len(words) > 1
            name = re.sub(r'^"|"$', '', words[1])
            text = '%%%s%%' % SimpleField.mapname(name)

        # sequence number: second word is the counter
        # example: SEQ Table \* ARABIC
        # XXX will need some special logic for table and figure captions, for
        #     which the leading 'Table ' and 'Figure ' are plain text that
        #     precede this field
        # XXX for now use a fake variable; assume that this will be removed
        #     later
        # XXX what about header numbers?
        elif words[0] == 'SEQ':
            assert len(words) > 1
            name = words[1]
            text = '%%%s%%' % SimpleField.mapname(name, prefix='docx-seq-')

        # cross-reference: second word is the bookmark name
        # example: REF _Ref76048121 \h
        elif words[0] == 'REF':
            # XXX forward references can only work if there's already been
            #     a document pass to find and store the bookmarks
            bookmarks = BookmarkStart.bookmarks
            name = words[1]
            # this is a BookmarkStart object (or None); Paragraph.markdown
            # will process it
            # XXX it's better simply always to return self, and process all
            #     bookmarks later
            # text = (self, bookmarks.get(name, None))
            text = (self, None)

        # other: ignore
        # XXX should definitely support PAGEREF
        # XXX return self for debugging
        else:
            text = self

        return text

    def __str__(self):
        return '%s' % self._element.text.strip().split()

    def __repr__(self):
        return '<%s %s>' % (type(self).__name__, str(self))
