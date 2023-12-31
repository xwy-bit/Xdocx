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


class Delete(Parented):
    def __init__(self, dele , parent):
        super().__init__(parent)
        self.dele = self._element = self.element = dele

    def add_r(self):
        return self._element._add_r()