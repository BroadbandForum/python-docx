# encoding: utf-8

"""
Custom properties.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)


class CustomProperties(object):
    """
    Corresponds to part named ``/docProps/custom.xml``, containing the custom
    document properties for this document package.
    """
    def __init__(self, element):
        self._element = element

    @property
    def properties(self):
        return self._element.properties
