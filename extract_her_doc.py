import docx
import re
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

def fmt_run(r):
    f = r.font
    out = ''
    if r.bold: out += 'B'
    if r.italic: out += 'I'
    if f.name: out += '|' + str(f.name)
    if f.size: out += '|' + str(f.size.pt)
    return out

ti = 0
for b in iter_block_items(d):
    if isinstance(b, Paragraph):
        runs = [fmt_run(r) for r in b.runs[:4]]
        print('P[%s] %r runs=%s' % (b.style.name, b.text[:180], runs))
    else:
        ti += 1
        fills = set(re.findall(r'w:fill="([0-9A-F]{6})"', b._tbl.xml))
        print('== TABLE %d (%dx%d) fills=%s' % (ti, len(b.rows), len(b.columns), sorted(fills)))
        for i, row in enumerate(b.rows):
            cells = []
            for c in row.cells:
                txt = ' / '.join(p.text for p in c.paragraphs)
                cells.append(txt[:110])
            print('  r%d: %s' % (i, ' || '.join(cells)))