from . import OxmlElement
from .simpletypes import ST_DecimalNumber, ST_String
from ..opc.constants import NAMESPACE
from ..text.paragraph import Paragraph
from ..text.run import Run
from .xmlchemy import (
	BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrMore, ZeroOrOne
)

class CT_Ins(BaseOxmlElement):
    '''
    A ``<w:ins>`` element, a container for Insert properties
    '''
    _id = RequiredAttribute('w:id', ST_DecimalNumber)
    date = RequiredAttribute('w:date', ST_String)
    author = RequiredAttribute('w:author', ST_String)
	
    @classmethod
    def new(cls, initials, comm_id, date, author):
        """
        Return a new ``<w:comment>`` element having _id of *comm_id* and having
        the passed params as meta data 
        """
        ins = OxmlElement('w:ins')
        ins.date = date
        ins._id = comm_id
        ins.author = author
        return ins
    
    def add_r(self, text):
        _r = OxmlElement('w:r')
        _t = _r.add_t(text)
        self.insert(0,_r)
        return _r