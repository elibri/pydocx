# coding: utf-8
from __future__ import (
    absolute_import,
    print_function,
    unicode_literals,
)

from pydocx.models import XmlModel, XmlAttribute


class EndnoteReference(XmlModel):
    XML_TAG = 'endnoteReference'

    endnote_id = XmlAttribute(name='id')

    @property
    def endnote(self):
        if not self.endnote_id:
            return
        part = self.container.endnotes_part
        if not part:
            return
        endnotes = part.endnotes
        endnote = endnotes.get_endnote_by_id(endnote_id=self.endnote_id)
        return endnote
