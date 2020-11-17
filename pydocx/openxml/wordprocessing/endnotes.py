# coding: utf-8
from __future__ import (
    absolute_import,
    print_function,
    unicode_literals,
)

from pydocx.models import XmlModel, XmlCollection
from pydocx.openxml.wordprocessing.endnote import Endnote


class Endnotes(XmlModel):
    XML_TAG = 'endnotes'

    children = XmlCollection(Endnote)

    def __init__(self, *args, **kwargs):
        super(Endnotes, self).__init__(*args, **kwargs)

        endnote_by_id = {}
        for endnote in self.children:
            if endnote.endnote_id:
                endnote_by_id[endnote.endnote_id] = endnote
        self._endnote_by_id = endnote_by_id

    def get_endnote_by_id(self, endnote_id):
        return self._endnote_by_id.get(endnote_id)
