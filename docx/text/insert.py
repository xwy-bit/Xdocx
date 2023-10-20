# encoding: utf-8

"""
Run-related proxy objects for python-docx, Run in particular.
"""

from __future__ import absolute_import, print_function, unicode_literals
from datetime import datetime

from docx.oxml.ns import qn
from docx.opc.part import *

from ..enum.style import WD_STYLE_TYPE
from ..enum.text import WD_BREAK
from .font import Font
from ..shape import InlineShape
from ..shared import Parented

from .comment import Comment

class Insert(Parented):

    def __init__(self, ins, parent):
        super(Insert, self).__init__(parent)
        self._ins = self._element = self.element = ins
    def add_r(self):
        """
        Return a newly appended ``<w:r>`` child element.
        """
        return self._element._add_r()