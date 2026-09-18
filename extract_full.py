import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

SRC = '/root/.hermes/cache/documents/doc_bb12df3d56ae_Lesson 2 - Inversion (Student\'s).docx'
d = docx.Document(SRC)

def iter_block_items(parent):
    for child in parent.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)

def p_style(p):
    pPr = p._p.find(qn('w:pPr'))
    flags = []
    if pPr is not None:
        if pPr.find(qn('w:pBdr')) is not None: flags.append('BORDER')
        shd = pPr.find(qn('w:shd'))
        if shd is not None: flags.append('SHADED:' + shd.get(qn('w:fill')))
        if pPr.find(qn('w:numPr')) is not None: flags.append('NUM')
    return ','.join(flags)

for b in iter_block_items(d):
    if isinstance(b, Paragraph):
        print('P[%s][%s] %r' % (b.style.name, p_style(b), b.text))
    else:
        for i, row in enumerate(b.rows):
            cells = []
            for c in row.cells:
                cells.append(repr(' / '.join(p.text for p in c.paragraphs)))
            print('T[%dx%d] r%d: %s' % (len(b.rows), len(b.columns), i, ' || '.join(cells)))