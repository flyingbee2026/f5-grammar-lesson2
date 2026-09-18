# Patch Helen's revised STUDENT docx -> Student v3 with the harder questions + labels.
import docx
from docx.shared import Pt
from docx.oxml.ns import qn

SRC = '/root/.hermes/cache/documents/doc_bb12df3d56ae_Lesson 2 - Inversion (Student\'s).docx'
OUT = '/opt/f5-grammar-lesson2/Grammar_Lesson_2_Inversions_Student_v3.docx'

d = docx.Document(SRC)
doc = d

def set_text(p, text, size=11):
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    r = p.add_run(text)
    r.font.name = 'Times New Roman'; r.font.size = Pt(size)

# new content map: heading text -> (simple sentence, prompt)
NEW = {
    'Example Set 2 (Not only ... but also ...)': (
        'The workshop improved students’ digital skills, and it boosted their confidence in online presentations.',
        'Use ‘Not only ... but also ...’ to combine the workshop’s two benefits.'),
    'Example Set 4 (Never / Rarely / Seldom)': (
        'The Internet has never made it so easy to spread false information.',
        'Begin with ‘Never’.'),
    'Example Set 6 (Only by ...)': (
        'Governments can fight online scams only by working closely with technology companies.',
        'Begin with ‘Only by’.'),
}
SUFFIX = ' – More challenging'

body = doc.element.body
last_heading = None
patched_tables = 0
for child in list(body):
    if child.tag == qn('w:p'):
        from docx.text.paragraph import Paragraph
        p = Paragraph(child, d)
        if p.style.name == 'Heading 3' and p.text.strip():
            last_heading = p.text.strip()
    elif child.tag == qn('w:tbl'):
        from docx.table import Table
        t = Table(child, d)
        # set tables: 3 rows, r0 col0 starts with 'Simple sentence:'
        if len(t.rows) == 3 and len(t.columns) == 2 and t.rows[0].cells[0].text.strip().startswith('Simple sentence'):
            if last_heading in NEW:
                simple, prompt = NEW[last_heading]
                set_text(t.rows[0].cells[1].paragraphs[0], simple)
                set_text(t.rows[1].cells[1].paragraphs[0], prompt)
                patched_tables += 1
            elif last_heading and last_heading.startswith('Example Set'):
                # heading gets the "More challenging" label for 2/4/6
                pass
        # pattern-1 example clarity (structure table, 3x2, r2 col1 = example)
        if len(t.rows) == 3 and len(t.columns) == 2 and 'builds confidence' in t.rows[2].cells[1].text:
            set_text(t.rows[2].cells[1].paragraphs[0],
                     'Not only does the scheme save money, but it also builds students’ confidence.')
            print('pattern-1 example patched in student')

# patch headings + intro (paragraph-level edits)
for p in list(d.paragraphs):
    if p.style.name == 'Heading 3':
        txt = p.text.strip()
        if txt in NEW:
            set_text(p, txt + SUFFIX)
    if p.style.name == 'Normal' and p.text.strip().startswith('Rewrite each simple sentence'):
        set_text(p, p.text.rstrip() +
                 ' The even-numbered sets (2, 4 and 6) are more challenging.')

d.save(OUT)
print('set tables patched:', patched_tables, '| saved', OUT)