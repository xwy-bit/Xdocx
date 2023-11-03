from ..shared import Parented


class Comment(Parented):
    """[summary]

    :param Parented: [description]
    :type Parented: [type]
    """
    def __init__(self, com, parent):
        super(Comment, self).__init__(parent)
        self._com = self._element = self.element = com
    
    def link_comment(self, _id, rangeStart=0, rangeEnd=0):
        pass
    def add_comm(self, author, comment_part, initials, dtime, comment_text, rangeStart, rangeEnd):
       pass