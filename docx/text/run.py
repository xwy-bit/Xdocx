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


class Run(Parented):
    """
    Proxy object wrapping ``<w:r>`` element. Several of the properties on Run
    take a tri-state value, |True|, |False|, or |None|. |True| and |False|
    correspond to on and off respectively. |None| indicates the property is
    not specified directly on the run and its effective value is taken from
    the style hierarchy.
    """

    def __init__(self, r, parent):
        super(Run, self).__init__(parent)
        self._r = self._element = self.element = r

    def add_break(self, break_type=WD_BREAK.LINE):
        """
        Add a break element of *break_type* to this run. *break_type* can
        take the values `WD_BREAK.LINE`, `WD_BREAK.PAGE`, and
        `WD_BREAK.COLUMN` where `WD_BREAK` is imported from `docx.enum.text`.
        *break_type* defaults to `WD_BREAK.LINE`.
        """
        type_, clear = {
            WD_BREAK.LINE:             (None,           None),
            WD_BREAK.PAGE:             ('page',         None),
            WD_BREAK.COLUMN:           ('column',       None),
            WD_BREAK.LINE_CLEAR_LEFT:  ('textWrapping', 'left'),
            WD_BREAK.LINE_CLEAR_RIGHT: ('textWrapping', 'right'),
            WD_BREAK.LINE_CLEAR_ALL:   ('textWrapping', 'all'),
        }[break_type]
        br = self._r.add_br()
        if type_ is not None:
            br.type = type_
        if clear is not None:
            br.clear = clear

    def add_picture(self, image_path_or_stream, width=None, height=None):
        """
        Return an |InlineShape| instance containing the image identified by
        *image_path_or_stream*, added to the end of this run.
        *image_path_or_stream* can be a path (a string) or a file-like object
        containing a binary image. If neither width nor height is specified,
        the picture appears at its native size. If only one is specified, it
        is used to compute a scaling factor that is then applied to the
        unspecified dimension, preserving the aspect ratio of the image. The
        native size of the picture is calculated using the dots-per-inch
        (dpi) value specified in the image file, defaulting to 72 dpi if no
        value is specified, as is often the case.
        """
        inline = self.part.new_pic_inline(image_path_or_stream, width, height)
        self._r.add_drawing(inline)
        return InlineShape(inline)

    def add_tab(self):
        """
        Add a ``<w:tab/>`` element at the end of the run, which Word
        interprets as a tab character.
        """
        self._r._add_tab()

    def add_text(self, text):
        """
        Returns a newly appended |_Text| object (corresponding to a new
        ``<w:t>`` child element) to the run, containing *text*. Compare with
        the possibly more friendly approach of assigning text to the
        :attr:`Run.text` property.
        """
        t = self._r.add_t(text)
        return _Text(t)

    def add_comment(self, text, author='python-docx', initials='pd', dtime=None):
        comment_part = self.part._comments_part.element
        if dtime is None:
            dtime = str(datetime.now()).replace(' ', 'T')
        comment = self._r.add_comm(author, comment_part, initials, dtime, text)

        return comment

    @property
    def bold(self):
        """
        Read/write. Causes the text of the run to appear in bold.
        """
        return self.font.bold

    @bold.setter
    def bold(self, value):
        self.font.bold = value

    def clear(self):
        """
        Return reference to this run after removing all its content. All run
        formatting is preserved.
        """
        self._r.clear_content()
        return self

    @property
    def font(self):
        """
        The |Font| object providing access to the character formatting
        properties for this run, such as font name and size.
        """
        return Font(self._element)

    @property
    def italic(self):
        """
        Read/write tri-state value. When |True|, causes the text of the run
        to appear in italics.
        """
        return self.font.italic

    @italic.setter
    def italic(self, value):
        self.font.italic = value

    @property
    def style(self):
        """
        Read/write. A |_CharacterStyle| object representing the character
        style applied to this run. The default character style for the
        document (often `Default Character Font`) is returned if the run has
        no directly-applied character style. Setting this property to |None|
        removes any directly-applied character style.
        """
        style_id = self._r.style
        return self.part.get_style(style_id, WD_STYLE_TYPE.CHARACTER)

    @style.setter
    def style(self, style_or_name):
        style_id = self.part.get_style_id(
            style_or_name, WD_STYLE_TYPE.CHARACTER
        )
        self._r.style = style_id

    @property
    def text(self):
        """
        String formed by concatenating the text equivalent of each run
        content child element into a Python string. Each ``<w:t>`` element
        adds the text characters it contains. A ``<w:tab/>`` element adds
        a ``\\t`` character. A ``<w:cr/>`` or ``<w:br>`` element each add
        a ``\\n`` character. Note that a ``<w:br>`` element can indicate
        a page break or column break as well as a line break. All ``<w:br>``
        elements translate to a single ``\\n`` character regardless of their
        type. All other content child elements, such as ``<w:drawing>``, are
        ignored.

        Assigning text to this property has the reverse effect, translating
        each ``\\t`` character to a ``<w:tab/>`` element and each ``\\n`` or
        ``\\r`` character to a ``<w:cr/>`` element. Any existing run content
        is replaced. Run formatting is preserved.
        """
        return self._r.text

    @text.setter
    def text(self, text):
        self._r.text = text

    @property
    def underline(self):
        """
        The underline style for this |Run|, one of |None|, |True|, |False|,
        or a value from :ref:`WdUnderline`. A value of |None| indicates the
        run has no directly-applied underline value and so will inherit the
        underline value of its containing paragraph. Assigning |None| to this
        property removes any directly-applied underline value. A value of
        |False| indicates a directly-applied setting of no underline,
        overriding any inherited value. A value of |True| indicates single
        underline. The values from :ref:`WdUnderline` are used to specify
        other outline styles such as double, wavy, and dotted.
        """
        return self.font.underline

    @underline.setter
    def underline(self, value):
        self.font.underline = value

    @property
    def footnote(self):
        _id = self._r.footnote_id

        if _id is not None:
            footnotes_part = self._parent._parent.part._footnotes_part.element
            footnote = footnotes_part.get_footnote_by_id(_id)
            return footnote.paragraph.text
        else:
            return None

    @property
    def is_hyperlink(self):
        '''
        checks if the run is nested inside a hyperlink element
        '''
        return self.element.getparent().tag.split('}')[1] == 'hyperlink'

    def get_hyperLink(self):
        """
        returns the text of the hyperlink of the run in case of the run has a hyperlink
        """
        document = self._parent._parent.document
        parent = self.element.getparent()
        linkText = ''
        if self.is_hyperlink:
            if parent.attrib.__contains__(qn('r:id')):
                rId = parent.get(qn('r:id'))
                linkText = document._part._rels[rId].target_ref
                return linkText, True
            elif parent.attrib.__contains__(qn('w:anchor')):
                linkText = parent.get(qn('w:anchor'))
                return linkText, False
            else:
                print('No Link in Hyperlink!')
                print(self.text)
                return '', False
        else:
            return 'None'

    @property
    def comments(self):
        comment_part = self._parent._parent.part._comments_part.element
        comment_refs = self._element.findall(qn('w:commentReference'))
        ids = [int(ref.get(qn('w:id'))) for ref in comment_refs]
        coms = [com for com in comment_part if com._id in ids]
        return [Comment(com, comment_part) for com in coms]

    def add_ole_object_to_run(self, ole_object_path):
        """
        Add saved OLE Object in the disk to an run and retun the newly created relationship ID
        Note: OLE Objects must be stored in the disc as `.bin` file
        """
        reltype: str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
        pack_path: str = "/word/embeddings/" + ole_object_path.split("\\")[-1]
        partname = PackURI(pack_path)
        content_type: str = "application/vnd.openxmlformats-officedocument.oleObject"

        with open(ole_object_path, "rb") as f:
            blob = f.read()
        target_part = Part(partname=partname, content_type=content_type, blob=blob)
        rel_id: str = self.part.rels._next_rId
        self.part.rels.add_relationship(reltype=reltype, target=target_part, rId=rel_id)
        return rel_id

    def add_fldChar(self, fldCharType, fldLock: bool = False, dirty: bool = False):

        fldChar = self._r.add_fldChar(fldCharType, fldLock, dirty)
        return fldChar

    @property
    def instr_text(self):
        return self._r.instr_text

    @instr_text.setter
    def instr_text(self, instr_text_val):
        self._r.instr_text = instr_text_val

    def remove_instr_text(self):
        if self.instr_text is None:
            return None
        else:
            self._r._remove_instr_text()


class _Text(object):
    """
    Proxy object wrapping ``<w:t>`` element.
    """

    def __init__(self, t_elm):
        super(_Text, self).__init__()
        self._t = t_elm
