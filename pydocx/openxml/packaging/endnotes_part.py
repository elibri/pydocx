# coding: utf-8
from __future__ import (
    absolute_import,
    print_function,
    unicode_literals,
)

from pydocx.openxml.packaging.open_xml_part import OpenXmlPart
from pydocx.openxml.wordprocessing import Endnotes


class EndnotesPart(OpenXmlPart):
    '''
    Represents a Endnotes part within a Word document container.

    See also: http://msdn.microsoft.com/en-us/library/documentformat.openxml.packaging.endnotespart%28v=office.14%29.aspx  # noqa
    '''

    relationship_type = '/'.join([
        'http://schemas.openxmlformats.org',
        'officeDocument',
        '2006',
        'relationships',
        'endnotes',
    ])

    def __init__(self, *args, **kwargs):
        super(EndnotesPart, self).__init__(*args, **kwargs)
        self._endnotes = None

    @property
    def endnotes(self):
        if not self._endnotes:
            self._endnotes = self.load_endnotes()
        return self._endnotes

    def load_endnotes(self):
        self._endnotes = Endnotes.load(self.root_element, container=self)

        return self._endnotes
