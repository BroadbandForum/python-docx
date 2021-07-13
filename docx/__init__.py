# encoding: utf-8

from docx.api import Document  # noqa

__version__ = "0.8.11"


# register custom Part classes with opc package reader

from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docx.opc.part import PartFactory
from docx.opc.parts.coreprops import CorePropertiesPart
from docx.opc.parts.customprops import CustomPropertiesPart

from docx.parts.document import DocumentPart
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.parts.image import ImagePart
from docx.parts.numbering import NumberingPart
from docx.parts.settings import SettingsPart
from docx.parts.styles import StylesPart


def part_class_selector(content_type, reltype):
    if reltype == RT.IMAGE:
        return ImagePart
    return None


PartFactory.part_class_selector = part_class_selector
PartFactory.part_type_for[CT.OPC_CORE_PROPERTIES] = CorePropertiesPart
PartFactory.part_type_for[CT.OPC_CUSTOM_PROPERTIES] = CustomPropertiesPart
PartFactory.part_type_for[CT.WML_DOCUMENT_MAIN] = DocumentPart
PartFactory.part_type_for[CT.WML_FOOTER] = FooterPart
PartFactory.part_type_for[CT.WML_HEADER] = HeaderPart
PartFactory.part_type_for[CT.WML_NUMBERING] = NumberingPart
PartFactory.part_type_for[CT.WML_SETTINGS] = SettingsPart
PartFactory.part_type_for[CT.WML_STYLES] = StylesPart

del (
    CT,
    CorePropertiesPart,
    CustomPropertiesPart,
    DocumentPart,
    FooterPart,
    HeaderPart,
    NumberingPart,
    PartFactory,
    SettingsPart,
    StylesPart,
    part_class_selector,
)

# register Parented classes with the corresponding element tag (this allows
# Parented.items to instantiate them); if the second argument (stop) is True,
# the "items" property won't include child elements
# XXX this isn't a nice interface; should use something more like the oxml
#     registration function

# XXX ignored elements; some of these shouldn't be ignored
from docx.shared import Ignored  # noqa
Ignored.register('mc:AlternateContent')
Ignored.register('w:commentRangeEnd')
Ignored.register('w:commentRangeStart')
Ignored.register('w:commentReference')
Ignored.register('w:footnoteReference')
Ignored.register('w:lastRenderedPageBreak')
Ignored.register('w:noBreakHyphen')
Ignored.register('w:object')
Ignored.register('w:pict')
Ignored.register('w:proofErr')
Ignored.register('w:rPr')
Ignored.register('w:sectPr')
Ignored.register('w:softHyphen')
Ignored.register('w:tblGrid')
Ignored.register('w:tcPr')
Ignored.register('w:tblPr')
Ignored.register('w:trPr')

from docx.drawing import (  # noqa
    Drawing
)
Drawing.register('w:drawing', True)

from docx.table import (  # noqa
    _Cell,
    _Row,
    Table,
)
_Cell.register('w:tc')
_Row.register('w:tr')
Table.register('w:tbl')

from docx.text.field import (  # noqa
    FieldChar,
    FieldCode,
    SimpleField,
)
FieldChar.register('w:fldChar')
FieldCode.register('w:instrText')
SimpleField.register('w:fldSimple')

from docx.text.hyperlink import (  # noqa
    Hyperlink,
)
Hyperlink.register('w:hyperlink')

from docx.text.paragraph import (  # noqa
    Paragraph,
)
Paragraph.register('w:p')

from docx.text.parfmt import (  # noqa
    BookmarkEnd,
    BookmarkStart,
    ParagraphProperties,
    SymbolChar,
    TabChar,
)
BookmarkEnd.register('w:bookmarkEnd')
BookmarkStart.register('w:bookmarkStart')
ParagraphProperties.register('w:pPr', True)
SymbolChar.register('w:sym')
TabChar.register('w:tab')

from docx.text.run import (  # noqa
    Break,
    DeletedRun,
    DeletedText,
    InsertedRun,
    Run,
    _Text,
)
Break.register('w:br')
DeletedRun.register('w:del')
DeletedText.register('w:delText')
InsertedRun.register('w:ins')
Run.register('w:r')
_Text.register('w:t')

# XXX should remove (del) all imported Parented classes?
del (
    _Cell,
    _Row,
    _Text
)
