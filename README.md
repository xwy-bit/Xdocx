Xdocx
==========

Python library forked from [bayoo-docx](https://github.com/BayooG/bayoo-docx).
(original repository [python-docx](https://github.com/python-openxml/python-docx))

Description
-----------
add some features to the original library.(**To Be Continued ...**)

High Level API
 - add comment in any position of the paragraph
 - add insert in any position of the paragraph

Low Level API
 - Insert class & openxml element
 - Delete class & openxml element

Installation
------------

```bash
git clone https://github.com/xwy-bit/Xdocx.git
cd Xdocx
pip install .
```

Usage
-----

**Insert Demo**

[Method] 

paragraph.add_insert_by_range 

```
INPUT:
    @ text : the text you want to insert
    @ ins_index : the index of the insert position. The index is the text position in the paragraph. The index is start from 0.

OUTPUT:
    @ None
```

[Demo]

```python
from docx.text.paragraph import Paragraph
import docx

doc = docx.Document('path/to/original.docx')
for idx , para in enumerate(doc.paragraphs):
    # you can use other method to select the paragraph , in this demo , we just use the index
    if idx == 6:
        print('=' * 40,'paragraph',idx,'=' * 40)
        '''
        example:
        ABCD  EFG 
            ⬆️ (TEST TEXT)
        '''
        para.add_insert_by_range(text = 'TEST TEXT' , ins_index = 4)
        print(para._element.xml)

doc.save('path/to/Insert_Demo.docx')
```

Comment Demo

> this is an navie demo ,not the final version

```python
import docx

doc = docx.Document('path/to/original.docx')
for idx , para in enumerate(doc.paragraphs):
    if idx ==6 : # find the paragraph you want to start comment 
        comment , comment_id = para._p.add_cross_paragraph_comment_start('Wayen Xu', para.part._comments_part._element, 'WX', '2023-11-03T00:00:00Z', 'THIS IS A TEST COMMENT', 0)
    if idx == 8: # find the paragraph you want to end comment
        para._p.pull_overflow_comment(comment_id,rangeEnd=2)

doc.save('path/to/Comment_Demo.docx')
```