# encoding: utf-8

"""Custom element classes for custom properties-related XML elements"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore


class CT_CustomProperties(BaseOxmlElement):
    """
    ``<cusp:Properties>`` element, the root element of the Custom Properties
    part stored as ``/docProps/custom.xml``.
    """

    # XXX this isn't working for me, possibly because the 'property' elements
    #     have no namespace prefixes?; I find these XML elements extremely
    #     difficult to debug! I've worked around it for now (see below)
    property_ = ZeroOrMore('cusp:property', successors=())

    _customProperties_tmpl = (
        '<cusp:Properties %s/>\n' % nsdecls('cusp', 'vt')
    )

    @classmethod
    def new(cls):
        """
        Return a new ``<cusp:Properties>`` element
        """
        xml = cls._customProperties_tmpl
        customProperties = parse_xml(xml)
        return customProperties

    # XXX this is a workaround; see the earlier comment
    @property
    def properties(self):
        """
        Dictionary mapping property names to values.
        """
        import lxml.etree as etree
        dct = {}
        for elem in self:
            tag = etree.QName(elem).localname
            assert tag == 'property' and 'name' in elem.attrib
            name = elem.attrib['name']
            assert len(elem) == 1
            typ = etree.QName(elem[0]).localname
            assert typ in {'lpwstr', 'i4'}
            value = int(elem[0].text) if typ == 'i4' else elem[0].text
            dct[name] = value
        return dct
