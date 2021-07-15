# encoding: utf-8

"""
Paragraph-related proxy types.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import re

from ..compat import is_string
from ..enum.style import WD_STYLE_TYPE
from .field import FieldChar, FieldCode, SimpleField
from .parfmt import BookmarkEnd, BookmarkStart, ParagraphFormat
from .run import Break, Run
from ..shared import Parented


class Paragraph(Parented):
    """
    Proxy object wrapping ``<w:p>`` element.
    """
    def __init__(self, p, parent):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

    def add_run(self, text=None, style=None):
        """
        Append a run to this paragraph containing *text* and having character
        style identified by style ID *style*. *text* can contain tab
        (``\\t``) characters, which are converted to the appropriate XML form
        for a tab. *text* can also include newline (``\\n``) or carriage
        return (``\\r``) characters, each of which is converted to a line
        break.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    @property
    def alignment(self):
        """
        A member of the :ref:`WdParagraphAlignment` enumeration specifying
        the justification setting for this paragraph. A value of |None|
        indicates the paragraph has no directly-applied alignment value and
        will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value):
        self._p.alignment = value

    def clear(self):
        """
        Return this same paragraph after removing all its content.
        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    def insert_paragraph_before(self, text=None, style=None):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph. If *text* is supplied, the new paragraph contains that
        text in a single run. If *style* is provided, that style is assigned
        to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    @property
    def paragraph_format(self):
        """
        The |ParagraphFormat| object providing access to the formatting
        properties for this paragraph, such as line spacing and indentation.
        """
        return ParagraphFormat(self._element)

    @property
    def runs(self):
        """
        Sequence of |Run| instances corresponding to the <w:r> elements in
        this paragraph.
        """
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def style(self):
        """
        Read/Write. |_ParagraphStyle| object representing the style assigned
        to this paragraph. If no explicit style is assigned to this
        paragraph, its value is the default paragraph style for the document.
        A paragraph style name can be assigned in lieu of a paragraph style
        object. Assigning |None| removes any applied style, making its
        effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        return self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)

    @style.setter
    def style(self, style_or_name):
        style_id = self.part.get_style_id(
            style_or_name, WD_STYLE_TYPE.PARAGRAPH
        )
        self._p.style = style_id

    @property
    def text(self):
        """
        String formed by concatenating the text of each run in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n``
        characters respectively.

        Assigning text to this property causes all existing paragraph content
        to be replaced with a single run containing the assigned text.
        A ``\\t`` character in the text is mapped to a ``<w:tab/>`` element
        and each ``\\n`` or ``\\r`` character is mapped to a line break.
        Paragraph-level formatting, such as style, is preserved. All
        run-level formatting, such as bold or italic, is removed.
        """
        text = ''
        for run in self.runs:
            text += run.text
        return text

    @text.setter
    def text(self, text):
        self.clear()
        self.add_run(text)

    def _insert_paragraph_before(self):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph.
        """
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)

    # previous paragraph's style name (lower case)
    previous_style_name = None

    @property
    def markdown(self):
        cls = self.__class__

        # note that this is lower-case
        style_name = self.style.name.lower()

        # use a fenced div for unrecognised styles
        # XXX for now this is opt-in rather than opt-out
        # XXX have removed 'issue description' because it can contain a list,
        #     and using a div per paragraph forces the numbering to be
        #     restarted (also, it's used within table cells)
        text = ''
        use_div = False
        if style_name in {'subheading'}:
            text += '\n:::::: %s ::::::\n' % style_name.replace(' ', '-')
            use_div = True

        markdown_lst = [item.markdown for item in self.items]

        state = 'outer'
        bookmarks = []
        words = []
        anchor = None
        added_begin_text = False
        # XXX should merge anchor and added_open_bracket logic?
        added_open_bracket = False
        need_page_break = False
        # XXX this handles the case where a reference is split across
        #     multiple field codes
        for value in markdown_lst:
            if not isinstance(value, list):
                value = [value]
            for val in value:
                if isinstance(val, FieldChar):
                    # XXX should define a property for this
                    state = val._element.fldCharType
                    if state == 'begin':
                        words = []
                        anchor = None
                        added_open_bracket = False
                        added_begin_text = False
                    elif state == 'end':
                        if added_open_bracket:
                            # it's possible that no text has been added
                            # XXX standard logic should be able to handle this
                            if text[-1] == '[':
                                text = text[:-1]
                            else:
                                text += ']'
                        # XXX could this add an anchor without an earlier
                        #     bracketed section? would that be a problem?
                        if anchor:
                            text += '[%s]' % anchor
                elif isinstance(val, FieldCode):
                    # XXX this is duplicate code from the tuple (bookmark)
                    #     case
                    field_code = val
                    # XXX temporary; need an improved interface
                    words += field_code._element.text.strip().split()
                    if len(words) <= 1:
                        pass
                    elif words[0] == 'DOCPROPERTY':
                        # 'begin' text is preferred to 'separate' text
                        # XXX this is copied from string handling below
                        # XXX this isn't supporting MERGEFORMAT or text
                        #     formatting (which is down at the run level)
                        if state in {'outer', 'begin', 'separate', 'end'}:
                            if state != 'separate' or not added_begin_text:
                                name = re.sub(r'^"|"$', '', words[1])
                                text += '%%%s%%' % SimpleField.mapname(name)
                            if state == 'begin':
                                added_begin_text = True
                    elif words[0] == 'REF':
                        if anchor is None:
                            if not added_open_bracket:
                                text += '['
                                added_open_bracket = True
                            anchor = '{{ref|%s}}' % words[1]
                    # XXX and the rest?
                elif isinstance(val, BookmarkStart):
                    bookmarks += [val]
                    # XXX can't do this unconditionally; there are too many
                    #     cases, plus bookmarks can be nested
                    # text += '['
                elif isinstance(val, BookmarkEnd):
                    # text += ']{}'
                    pass
                elif isinstance(val, tuple):
                    assert len(val) == 2
                    field_code, bookmark_start = val
                    if not added_open_bracket:
                        text += '['
                        added_open_bracket = True
                    if not bookmark_start:
                        # XXX temporary; need an improved interface
                        # XXX we ALWAYS take this path; see field.py
                        words += field_code._element.text.strip().split()
                        if len(words) > 1 and anchor is None:
                            anchor = '{{ref|%s}}' % words[1]
                    else:
                        # XXX need a proper utility function here
                        # XXX we NEVER take this path; see field.py
                        assert False, 'The impossible happened'
                        anchor = bookmark_start._parent.markdown.strip()
                        anchor = re.sub(r'^#*\s*', '', anchor)
                elif isinstance(val, Break):
                    # XXX can assume it's a page break, because line breaks
                    #     were already handled
                    need_page_break = True
                elif is_string(val):
                    # 'begin' text is preferred to 'separate' text
                    if state in {'outer', 'begin', 'separate', 'end'}:
                        if state != 'separate' or not added_begin_text:
                            text += val
                        if state == 'begin' and val.strip():
                            added_begin_text = True
                else:
                    # XXX unsupported type
                    assert False, 'Unsupported paragraph markdown value: %r' \
                                  % val

        # determine the first line indent from the paragraph properties,
        # falling back on the paragraph style's paragraph properties
        # XXX should always use self.paragraph_format to access pPr where
        #     possible (I hadn't really realized what it was!)
        # XXX it would be better instead to use pandoc '| ' line blocks, but
        #     this requires some CSS fixes to make it look nice
        # XXX it would also be good to be able to detect cases where 'Code'
        #     or 'Plain Text' are being simulated by other styles and font
        #     overrides
        left_indent = self.paragraph_format.left_indent or \
                      self.style.paragraph_format.left_indent or 0
        first_line_indent = self.paragraph_format.first_line_indent or \
                            self.style.paragraph_format.first_line_indent or 0
        total_indent = left_indent + first_line_indent
        # XXX the 50 here means about 20 spaces per inch; this is heuristic;
        #     the 4.00/6.35 is further heuristic to cause tabs and indent to
        #     be about the same, at least in one case
        spaces = int((total_indent/1440/50) * (4.00/6.35) + 0.5)
        indent = '{{indent=%r}}' % spaces if spaces else ''

        # naively indent 'Code' etc. paragraphs
        # XXX this makes assumptions about style names
        # XXX there can be problems with transitions from lists to code
        if style_name in {'code', 'plain text'}:
            text = '    ' + indent + text
        # naively replace leading TAB characters and spaces with escaped spaces
        # XXX don't do this for list paragraphs (they already have their '* '
        #     etc.); this logic should be earlier
        # XXX TAB replacement has been suppressed
        elif text.startswith('{{tab}}') or indent != '':
            if style_name != 'list paragraph':
                text = indent + text

        # XXX this is a hack to add CSS classes etc. (assumes first item is
        #     paragraph properties)
        elif self.items:
            text = text.rstrip() + self.items[0].markdown_suffix

        # (fairly) naively collapse and substitute character attributes
        # XXX should escape '`' and '*' before doing this
        # XXX this could misbehave in complex cases? should scan the input
        for ch, md in {'c': '`', 'i': '*', 'b': '**', 'B': '***'}.items():
            # {{endb}}{{b}} -> nothing
            text = text.replace('{{end%s}}{{%s}}' % (ch, ch), '')
            # {{b}}__XYZ__{{endb}} -> __{{b}}XYZ{{endb}}__ (_ = whitespace)
            text = re.sub(r'({{%s}})(\s*)(.*?)(\s*)({{end%s}})' % (ch, ch),
                          r'\2\1\3\5\4', text)
            # {{b}}{{endb}} -> nothing
            text = text.replace('{{%s}}{{end%s}}' % (ch, ch), '')
            # replace with markup
            text = text.replace('{{%s}}' % ch, md)
            text = text.replace('{{end%s}}' % ch, md)

        # transform figure and table captions to what pandoc expects
        # XXX these will mostly be hyperlinks rather than use @sec:xxx
        #     references; should transform them (or some of them)
        # XXX allowing optional number before/after the variable ref is
        #     defensive: this can happen if there are additional field codes,
        #     e.g. STYLEREF
        # XXX may also want to transform 'Section [n]' to '[Section n]' etc.?
        if style_name in {'caption'}:
            text = re.sub(r'Figure \d*(?:%\w+%)?\d*\W+(.*)',
                          r'![\1](missing.png)', text)
            text = re.sub(r'Table \d*%\w+%\d*\W+', r':', text)

        # if a page break was seen, insert it at the beginning
        # XXX or just insert a []{.new-page} span at the break location?
        if need_page_break:
            newline_maybe = '' if text.strip() == '' else '\n'
            text = '::: new-page :::\n:::%s' % newline_maybe + text

        # if using a div, terminate it
        if use_div:
            text += '\n::::::'

        # unless suppressed, insert a newline (the idea is to prevent
        # unwanted newlines in lists and code blocks)
        else:
            no_newline_styles = {'code', 'list paragraph', 'plain text'}
            if cls.previous_style_name not in no_newline_styles or style_name \
                    not in no_newline_styles:
                text = '\n' + text

        # remember the style name
        cls.previous_style_name = style_name

        # ignore ToC and list of figures (which is also used for tables)
        # XXX this is at the bottom so the previous style name is saved
        if re.match(r'toc \d+|table of figures', style_name):
            text = None

        # ignore empty paragraphs (people abuse them as an alternative to
        # "after paragraph space")
        # XXX it's a bit messy to have to deal with the fact that {{xxx=n}} or
        #     {{yyy}} might have been inserted
        elif re.sub(r'{{\w+(=.+?)?}}\s*', r'', text).strip() == '':
            text = None

        return text

    def __str__(self):
        return self.text

    __repr__ = __str__
