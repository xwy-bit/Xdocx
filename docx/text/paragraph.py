# encoding: utf-8

"""
Paragraph-related proxy types.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..enum.style import WD_STYLE_TYPE
from .parfmt import ParagraphFormat
from .run import Run
from ..shared import Parented


from docx.text.insert import Insert
from docx.text.delete import Delete

from datetime import datetime
import re

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
    
    def delete(self):
        """
        delete the content of the paragraph
        """
        self._p.getparent().remove(self._p)
        self._p = self._element = None
    
    def add_comment(self, text, author='python-docx', initials='pd', dtime=None ,rangeStart=0, rangeEnd=0, comment_part=None):
        if comment_part is None:
            comment_part = self.part._comments_part.element
        if dtime is None:
            dtime = str( datetime.now() ).replace(' ', 'T')
        comment =  self._p.add_comm(author, comment_part, initials, dtime, text, rangeStart, rangeEnd)

        return comment
    
    def get_matched_run_info(self , query_text):

        run_text_ranges = []
        text_counter = 0
        for run in self.runs:
            run_text_ranges.append([text_counter,len(run.text)+ text_counter])
            text_counter += len(run.text)

        query_text_range = [self.text.index(query_text)
                            ,self.text.index(query_text)+len(query_text)] 
        print(query_text_range)
        print(run_text_ranges)
        begin_run_idx = -1
        begin_run_offset = 0
        end_run_idx = -1
        end_run_offset = 0
        for idx, run_text_range in enumerate(run_text_ranges):
            if run_text_range[0] <= query_text_range[0] and run_text_range[1] > query_text_range[0]:
                begin_run_idx = idx
                begin_run_offset = query_text_range[0] - run_text_range[0]
            if run_text_range[0] < query_text_range[1] and run_text_range[1] >= query_text_range[1]:
                end_run_idx = idx
                end_run_offset = query_text_range[1] - run_text_range[0]
                continue
        if begin_run_idx == -1 or end_run_idx == -1:
            raise ValueError('query text not found in paragraph')
        return begin_run_idx, begin_run_offset, end_run_idx, end_run_offset

    def add_comment_by_text(self, text, author='Wayen', initials='W', dtime=None,query_text="WenShuTech Comment TEST"):
        begin_run_idx, begin_run_offset, end_run_idx, end_run_offset \
                = self.get_matched_run_info(query_text)
        new_run = self.reconstruct_paragraph(query_text, begin_run_idx, begin_run_offset, end_run_idx, end_run_offset)
        comment = new_run.add_comment(text, author, initials, dtime)

        return comment

    def add_comment_by_range(self, text, author='Wayen', initials='W', dtime=None ,rangeStart=0, rangeEnd=0):
        paragraph_text = self.text
        if dtime is None:
            dtime = str( datetime.now() ).replace(' ', 'T')

        if rangeStart == rangeEnd:
            raise ValueError('rangeStart and rangeEnd can not be equal')

        qury_text = paragraph_text[rangeStart:rangeEnd]
        begin_run_idx, begin_run_offset, end_run_idx, end_run_offset \
                = self.get_matched_run_info(qury_text)
        new_run = self.reconstruct_paragraph(qury_text, begin_run_idx, begin_run_offset, end_run_idx, end_run_offset)
        comment = new_run.add_comment(text, author, initials, dtime)

        return comment
       
    def add_insert_by_range(self, text, author='Wayen', initials='W', dtime=None ,ins_index = 0):
        paragraph_text = self.text

        # get run text ranges, e.g. [[0,3],[3,6],[6,9]]
        run_text_ranges = []
        text_counter = 0
        for run in self.runs:
            run_text_ranges.append([text_counter,len(run.text)+ text_counter])
            text_counter += len(run.text)
        
        if dtime is None:
            dtime = str( datetime.now() ).replace(' ', 'T')

        if ins_index < 0 or ins_index > len(paragraph_text):
            raise ValueError('ins_index should be in range of paragraph')

        print('run_text_ranges',run_text_ranges)
        # if ins_index not in the middle of a run, then insert a new run
        if ins_index == 0:
            nins = self._p._new_ins()
            nins._id = 1
            nins.author = author
            nins.date = dtime
            nins.add_r(text)
            New_Insert = Insert(nins,self._p)
            self._p.insert(0,nins)
        elif ins_index in [run_text_range[1] for run_text_range in run_text_ranges]:
            run_idx = [run_text_range[1] 
                       for run_text_range in run_text_ranges].index(ins_index)
            run = self.runs[run_idx]
            nins = self._p._new_ins()
            nins._id = 1
            nins.author = author
            nins.date = dtime
            nins.add_r(text)
            New_Insert = Insert(nins,self._p)
            run._r.addnext(nins)
        # if ins_index in the middle of a run, then split the run
        else:
            for idx, run_text_range in enumerate(run_text_ranges):
                if run_text_range[0] <= ins_index and run_text_range[1] > ins_index:
                    run_idx = idx
                    run_offset = ins_index - run_text_range[0]
                    break
            # split run
            anchor_Run = self.runs[run_idx]
            nrun = self._p._new_r()
            nrun.text = anchor_Run.text[run_offset:]
            anchor_Run.text = anchor_Run.text[:run_offset]
            New_Run = Run(nrun,self._p)
            nins = self._p._new_ins()
            nins._id = 1
            nins.author = author
            nins.date = dtime
            nins.add_r(text)
            New_Insert = Insert(nins,self._p)
            anchor_Run._r.addnext(nins)
            nins.addnext(nrun)

            # alian font
            New_Run.bold = anchor_Run.bold
            New_Run.italic = anchor_Run.italic
            New_Run.underline = anchor_Run.underline
            New_Run.font.name = anchor_Run.font.name
            New_Run.font.size = anchor_Run.font.size
            New_Run.font.color.rgb = anchor_Run.font.color.rgb

            New_Insert.bold = anchor_Run.bold
            New_Insert.italic = anchor_Run.italic
            New_Insert.underline = anchor_Run.underline
            New_Insert.font.name = anchor_Run.font.name
            New_Insert.font.size = anchor_Run.font.size
            New_Insert.font.color.rgb = anchor_Run.font.color.rgb










        
        
    def reconstruct_paragraph(self ,query_text, begin_run_idx, begin_run_offset, end_run_idx, end_run_offset):
        if begin_run_idx == end_run_idx:
            runs = self.runs
            original_text = runs[begin_run_idx].text
            anchor_run = runs[begin_run_idx]
            nrun_middle= self._element._new_r()
            nrun_end = self._element._new_r()
            new_run_middle = Run(nrun_middle, anchor_run._parent)
            new_run_end = Run(nrun_end, anchor_run._parent)
            anchor_run._element.addnext(nrun_middle)
            new_run_middle._element.addnext(nrun_end)
            if end_run_idx + 1 < len(runs) :
                print('add next')
                new_run_end._element.addnext(runs[end_run_idx + 1]._element)

            anchor_run.text = anchor_run.text[:begin_run_offset]
            new_run_middle.text =  query_text
            new_run_end.text = original_text[end_run_offset:]

            # alian font
            new_run_middle.bold = anchor_run.bold
            new_run_middle.italic = anchor_run.italic
            new_run_middle.underline = anchor_run.underline
            new_run_middle.font.name = anchor_run.font.name
            new_run_middle.font.size = anchor_run.font.size
            new_run_middle.font.color.rgb = anchor_run.font.color.rgb

            new_run_end.bold = anchor_run.bold
            new_run_end.italic = anchor_run.italic
            new_run_end.underline = anchor_run.underline
            new_run_end.font.name = anchor_run.font.name
            new_run_end.font.size = anchor_run.font.size
            new_run_end.font.color.rgb = anchor_run.font.color.rgb

            return new_run_middle

        elif begin_run_idx < end_run_idx:

            runs = self.runs
            anchor_run = runs[begin_run_idx]
            nrun= self._element._new_r()
            new_run = Run(nrun, anchor_run._parent)
            runs[begin_run_idx]._element.addnext(nrun)
            new_run._element.addnext(runs[end_run_idx]._element)
            new_run.text =  query_text
            runs[begin_run_idx].text = runs[begin_run_idx].text[:begin_run_offset]
            runs[end_run_idx].text = runs[end_run_idx].text[end_run_offset:]
            
            # alian font
            new_run.bold = anchor_run.bold
            new_run.italic = anchor_run.italic
            new_run.underline = anchor_run.underline
            new_run.font.name = anchor_run.font.name
            new_run.font.size = anchor_run.font.size
            new_run.font.color.rgb = anchor_run.font.color.rgb

            return new_run  
    
    def add_footnote(self, text):
        footnotes_part = self.part._footnotes_part.element
        footnote = self._p.add_fn(text, footnotes_part)

        return footnote

    def merge_paragraph(self, otherParagraph):
        r_lst = otherParagraph.runs
        self.append_runs(r_lst)
    
    def append_runs(self, runs):
        self.add_run(' ')
        for run in runs:
            self._p.append(run._r)
            
    
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
    def all_runs(self):
        return [Run(r, self) for r in self._p.xpath('.//w:r[not(ancestor::w:r)]')]
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

    @property
    def header_level(self):
        '''
        input Paragraph Object
        output Paragraph level in case of header or returns None
        '''
        headerPattern = re.compile(".*Heading (\d+)$")
        level = 0
        if headerPattern.match(self.style.name):
            level = int(self.style.name.lower().split('heading')[-1].strip())
        return level
    
    @property
    def NumId(self):
        '''
        returns NumId val in case of paragraph has numbering
        else: return None
        '''
        try:
            return self._p.pPr.numPr.numId.val
        except:
            return None
    
    @property
    def list_lvl(self):
        '''
        returns ilvl val in case of paragraph has a numbering level
        else: return None
        '''
        try:
            return self._p.pPr.numPr.ilvl.val
        except :
            return None
    
    @property
    def list_info(self):
        '''
        returns tuple (has numbering info, numId value, ilvl value)
        '''
        if self.NumId and self.list_lvl:
            return True, self.NumId, self.list_lvl
        else:
            return False, 0, 0
    
    @property
    def is_heading(self):
        return True if self.header_level else False
    
    @property
    def full_text(self):
        return u"".join([r.text for r in self.all_runs])
    
    @property
    def footnotes(self):
        if self._p.footnote_ids is not None :
            return True
        else :
            return False

    @property
    def comments(self):
        runs_comments = [run.comments for run in self.runs]
        return [comment for comments in runs_comments for comment in comments]

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
