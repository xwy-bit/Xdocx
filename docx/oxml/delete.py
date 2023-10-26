from . import OxmlElement
from .simpletypes import ST_DecimalNumber, ST_String
from ..opc.constants import NAMESPACE
from ..text.paragraph import Paragraph
from ..text.run import Run
from .xmlchemy import (
	BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrMore, ZeroOrOne
)

class CT_Dele(BaseOxmlElement):
    '''
    A ``<w:dele>`` element, a container for Delete properties
    '''
    _id = RequiredAttribute('w:id', ST_DecimalNumber)
    date = RequiredAttribute('w:date', ST_String)
    author = RequiredAttribute('w:author', ST_String)

    @classmethod
    def new(cls, initials, comm_id, date, author):
        dele = OxmlElement('w:del')
        dele.date = date
        dele._id = comm_id
        dele.author = author
        return dele
    def add_r(self,text):
        _r = OxmlElement('w:r')
        _t = _r.add_dele_t(text)
        self.insert(0,_r)
        return _r