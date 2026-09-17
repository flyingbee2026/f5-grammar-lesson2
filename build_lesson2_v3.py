# Build Grammar Lesson 2 v3 (teacher's version):
#  - Times New Roman 11 for main content, line spacing 1.15
#  - 8 practice sets (Not only x2, Rarely/Never x2, Only by x2, So...that x2)
#  - No fill on example paragraph boxes; no italics on targets (students highlight)
#  - Quiz item 6 = common 'not only' inversion mistake
#  - Model uses Rarely / Not only / Only by, flow-first
import docx
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC = "/root/.hermes/cache/documents/doc_70fd9299bc97_Revised_Lesson_1_Relative_Clauses_Teacher's_CORRECTED.docx"
OUT = "/opt/f5-grammar-lesson2/Grammar_Lesson_2_Inversions_Teacher_v3.docx"
import os
os.makedirs(os.path.dirname(OUT), exist_ok=True)

# palette (kept outside the example paragraph boxes)
QUIZ_HDR  = 'D9D9D9'   # grey header
CORRECT   = 'E2EFDA'   # light green answers/corrections
STRUCTURE = 'DEEBF7'   # light blue structure boxes
NOTEBOX   = 'F2F2F2'   # grey teacher's notes / optional patterns
INSTBOX   = 'E4DFEC'   # light lavender instructions
GREY = RGBColor(0x59, 0x59, 0x59)

doc = docx.Document(SRC)
# base style: Times New Roman 11, line spacing 1.15
st = doc.styles['Normal']
st.font.name = 'Times New Roman'
st.font.size = Pt(11)
st.paragraph_format.line_spacing = 1.15

body = doc.element.body
for child in list(body):
    if child.tag in (qn('w:p'), qn('w:tbl')):
        body.remove(child)

def style_run(r, font='Times New Roman', size=11, color=None):
    r.font.name = font
    r.font.size = Pt(size)
    if color is not None:
        r.font.color.rgb = color

def add_para(text="", style=None, bold=False, italic=False, size=11, font='Times New Roman',
             color=None, page_break_before=False):
    p = doc.add_paragraph(style=style)
    p.paragraph_format.line_spacing = 1.15
    if page_break_before:
        p.paragraph_format.page_break_before = True
    if text:
        r = p.add_run(text)
        r.bold, r.italic = bold, italic
        style_run(r, font, size, color)
    return p

def add_para_runs(runs, style=None, page_break_before=False):
    """runs: list of (text, bold, italic) — TNR 11, spacing 1.15."""
    p = doc.add_paragraph(style=style)
    p.paragraph_format.line_spacing = 1.15
    if page_break_before:
        p.paragraph_format.page_break_before = True
    for spec in runs:
        r = p.add_run(spec[0])
        r.bold, r.italic = spec[1], spec[2]
        style_run(r)
    return p

def shade_cell(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)

def cell_runs(cell, spec, color=None):
    cell.paragraphs[0].paragraph_format.line_spacing = 1.15
    for run in spec:
        r = cell.paragraphs[0].add_run(run[0])
        r.bold, r.italic = run[1], run[2]
        style_run(r, color=color)

def make_table(rows, fills=None):
    """rows: list of rows; each row = list of cell specs (list of run tuples)."""
    fills = fills or {}
    ncols = len(rows[0])
    t = doc.add_table(rows=len(rows), cols=ncols)
    t.style = doc.styles['Table Grid']
    for i, row in enumerate(rows):
        for j, spec in enumerate(row):
            c = t.rows[i].cells[j]
            cell_runs(c, spec)
            if (i, j) in fills:
                shade_cell(c, fills[(i, j)])
    return t

def box(runs, color=None):
    """Bordered paragraph box, NO fill (students highlight targets themselves)."""
    t = doc.add_table(rows=1, cols=1)
    t.style = doc.styles['Table Grid']
    cell_runs(t.rows[0].cells[0], runs, color=color)
    return t

R  = lambda t: [(t, False, False)]
RB = lambda t: [(t, True, False)]

# ============ HEADER (same as Lesson 1) ============
p = doc.add_paragraph(style="Normal"); p.paragraph_format.line_spacing = 1.15
r = p.add_run("St. Joseph’s Anglo-Chinese School"); r.bold = True; r.italic = True; style_run(r, font="Century Gothic")
p = doc.add_paragraph(style="Normal"); p.paragraph_format.line_spacing = 1.15
r = p.add_run("F.5 English Language"); r.bold = True; r.italic = True; style_run(r, font="Trebuchet MS")
p = doc.add_paragraph(style="Normal"); p.paragraph_format.line_spacing = 1.15
r = p.add_run("Grammar – Lesson 2: Inversions"); r.bold = True; r.italic = True; style_run(r, font="Trebuchet MS")
p = doc.add_paragraph(style="Normal"); p.paragraph_format.line_spacing = 1.15
r = p.add_run("Inversions"); r.bold = True; style_run(r, font="Times New Roman")

# ============ PART 1: PROOFREADING QUIZ ============
add_para("Part 1: Proofreading quiz", style="Heading 2")
add_para("Each sentence contains ONE error. Underline the error and write the full corrected sentence. "
         "Items 1–5 review relative clauses from Lesson 1; Item 6 previews today’s topic — inversions.")
quiz_rows = [
    [R("Sentence"), R("Correction (Teacher’s version)")],
    [R("Social media connects people instantly which makes the world feel smaller."),
     R("Social media connects people instantly, which makes the world feel smaller. "
       "(Add a comma before the connective clause.)")],
    [R("Online courses are ideal for busy adults, especially those which have little time to attend classes."),
     R("Online courses are ideal for busy adults, especially those who have little time to attend classes. "
       "(People take \u2018who\u2019, not \u2018which\u2019.)")],
    [R("There are dozens of messaging apps, all of them are free to use."),
     R("There are dozens of messaging apps, all of which are free to use. "
       "(A pronoun cannot join two clauses; use \u2018which\u2019.)")],
    [R("The forum which students share their homework is very popular."),
     [("Teacher’s Answer A: The forum on which students share their homework is very popular. ", True, False),
      ("Teacher’s Answer B: The forum where students share their homework is very popular. "
       "(Add the missing preposition \u2018on\u2019.)", False, False)]],
    [R("Many teenagers spend hours gaming, that worries their parents."),
     R("Many teenagers spend hours gaming, which worries their parents. "
       "(\u2018That\u2019 cannot begin a connective clause.)")],
    [R("Not only the internet connects people around the world, but it also spreads information quickly."),
     R("Not only does the internet connect people around the world, but it also spreads information quickly. "
       "(Move the auxiliary before the subject after \u2018Not only\u2019 \u2014 more on this today.)")],
]
fills_quiz = {(0, 0): QUIZ_HDR, (0, 1): QUIZ_HDR}
for i in range(1, 7):
    fills_quiz[(i, 1)] = CORRECT
make_table(quiz_rows, fills=fills_quiz)

# ============ PART 2: TWO CORE PATTERNS ============
add_para("Part 2: Two patterns for formal writing", style="Heading 2", page_break_before=True)
add_para("Both patterns follow the same rule: the auxiliary verb (or the verb \u2018be\u2019) moves before the "
         "subject. We focus on two patterns that are high-value and safe for formal writing; two more advanced "
         "patterns are kept in the teacher’s note at the end of this part.")

add_para("Pattern 1: Not only ... but also ...", style="Heading 3")
make_table([
    [R("Structure:"), R("Not only + auxiliary + subject + verb, but (subject) also + verb")],
    [R("Use in formal writing:"), R("Adds a second advantage in proposals, reports and articles.")],
    [R("Example:"), R("Not only does the scheme save money, but it also builds confidence.")],
], fills={(0, 0): STRUCTURE, (0, 1): STRUCTURE, (1, 0): STRUCTURE, (1, 1): STRUCTURE,
          (2, 0): STRUCTURE, (2, 1): STRUCTURE})
box([("Introducing a school e-learning platform would bring clear benefits. ", False, False),
     ("Not only does the platform give students round-the-clock access to learning materials, but it also allows teachers to track progress more closely", False, False),
     (". With such a system, revision becomes a habit rather than a scramble before exams.", False, False)])

add_para("Pattern 2: Never / Rarely / Seldom", style="Heading 3")
make_table([
    [R("Structure:"), R("Never / Rarely / Seldom + auxiliary + subject + verb")],
    [R("Use in formal writing:"), R("Turns a criticism or a surprising claim into a formal, striking opening.")],
    [R("Example:"), R("Rarely do we question our daily screen habits.")],
], fills={(0, 0): STRUCTURE, (0, 1): STRUCTURE, (1, 0): STRUCTURE, (1, 1): STRUCTURE,
          (2, 0): STRUCTURE, (2, 1): STRUCTURE})
box([("Many parents blame video games for every poor result. ", False, False),
     ("Rarely do they stop to consider how limited their children’s other entertainment options are", False, False),
     (". The real problem may be the lack of balance, not the games themselves.", False, False)])

# ---- Additional patterns: teacher's note (grey) ----
add_para("Additional patterns (teacher’s note — optional)", style="Heading 3", color=GREY)
add_para("Introduce these two only after the class is secure with Patterns 1 and 2. The mechanism is the same: "
         "front the expression, then place the auxiliary before the subject.", color=GREY)
add_para("Only + adverbial (only when / only after / only by)", style="Heading 4", color=GREY)
make_table([
    [R("Structure:"), R("Only + when / after / by + clause or phrase + auxiliary + subject + verb")],
    [R("Example:"), R("Only by issuing clear guidelines can schools protect students’ privacy.")],
], fills={(0, 0): NOTEBOX, (0, 1): NOTEBOX, (1, 0): NOTEBOX, (1, 1): NOTEBOX})
add_para("Trap: \u2018Only + noun\u2019 does NOT invert. \u2018Only students can join\u2019 is already correct; "
         "inversion happens only when \u2018only\u2019 is followed by an adverbial.", color=GREY)
add_para("So ... that / Such ... that", style="Heading 4", color=GREY)
make_table([
    [R("Structure:"), R("So + adjective + auxiliary + subject + verb + that ...  OR  Such + be + noun phrase + that ...")],
    [R("Example:"), R("So popular is online shopping that it has changed the way we buy things.")],
], fills={(0, 0): NOTEBOX, (0, 1): NOTEBOX, (1, 0): NOTEBOX, (1, 1): NOTEBOX})
add_para("Use it to open with a strong statement when emphasising the scale or importance of an issue.", color=GREY)

# ============ PART 3: PRACTICE SETS ============
add_para("Part 3: Practice sets", style="Heading 2", page_break_before=True)
add_para("Rewrite each simple sentence using the pattern given in the prompt. The sentences are all about "
         "technology and communication, so keep the register formal.")

sets = [
    ("Example Set 1 (Not only ... but also ...)",
     "The online platform saves paper, and it gives students instant feedback.",
     "Begin with \u2018Not only\u2019 to add a second advantage.",
     "Not only does the online platform save paper, but it also gives students instant feedback."),
    ("Example Set 2 (Not only ... but also ...)",
     "E-books are cheaper than printed textbooks, and they are far lighter to carry.",
     "Use \u2018Not only ... but also ...\u2019 to combine the two advantages of e-books.",
     "Not only are e-books cheaper than printed textbooks, but they are also far lighter to carry."),
    ("Example Set 3 (Never / Rarely / Seldom)",
     "People rarely question the amount of time they spend scrolling on their phones.",
     "Begin with \u2018Rarely\u2019.",
     "Rarely do people question the amount of time they spend scrolling on their phones."),
    ("Example Set 4 (Never / Rarely / Seldom)",
     "The workplace has never relied so heavily on technology.",
     "Begin with \u2018Never\u2019.",
     "Never has the workplace relied so heavily on technology."),
    ("Example Set 5 (Only by ...)",
     "Schools can protect students’ privacy by issuing clear guidelines on social media use.",
     "Begin with \u2018Only by\u2019.",
     "Only by issuing clear guidelines on social media use can schools protect students’ privacy."),
    ("Example Set 6 (Only by ...)",
     "Bookshops can survive in the digital age by offering experiences that apps cannot match.",
     "Begin with \u2018Only by\u2019.",
     "Only by offering experiences that apps cannot match can bookshops survive in the digital age."),
    ("Example Set 7 (So ... that)",
     "The problem of cyberbullying is so serious that every school must act.",
     "Begin with \u2018So serious\u2019.",
     "So serious is the problem of cyberbullying that every school must act."),
    ("Example Set 8 (So ... that)",
     "Online shopping is so popular that it has changed how we buy things.",
     "Begin with \u2018So popular\u2019.",
     "So popular is online shopping that it has changed how we buy things."),
]
for title, simple, prompt, answer in sets:
    add_para(title, style="Heading 3")
    make_table([
        [R("Simple sentence:"), R(simple)],
        [R("Prompt:"), [(prompt, False, True)]],
        [RB("Teacher’s Answer:"), R(answer)],
    ], fills={(2, 1): CORRECT})

# ============ PART 4: WRITING TASK ============
add_para("Part 4: Writing Task", style="Heading 1", page_break_before=True)
ti = doc.add_table(rows=1, cols=1); ti.style = doc.styles['Table Grid']
ci = ti.rows[0].cells[0]
shade_cell(ci, INSTBOX)
ci.paragraphs[0].paragraph_format.line_spacing = 1.15
r = ci.paragraphs[0].add_run("Instructions:"); r.bold = True; style_run(r)
r2 = ci.paragraphs[0].add_run(); r2.add_break()
r3 = ci.paragraphs[0].add_run("Social media plays a big part in students’ lives. Write ONE paragraph (about 80–100 "
                              "words) giving your view on whether students should limit their use of social media. "
                              "Use at least two inversions and one relative clause. "); style_run(r3)
r4 = ci.paragraphs[0].add_run("Remember: the flow of ideas comes first — never force in an inversion just for the "
                              "sake of using one."); r4.bold = True; style_run(r4)

box([("Many people accuse social media of destroying students’ concentration and turning them into passive "
      "consumers of short videos. ", False, False),
     ("Rarely do they mention the benefits it brings to learning, such as instant access to discussion groups and "
      "subject resources, which older generations could never enjoy. ", False, False),
     ("Not only does social media connect students with their classmates, but it also links them to news and ideas "
      "from around the world. ", False, False),
     ("Only by teaching young people to manage their digital habits, therefore, can schools turn these platforms "
      "from a distraction into a real learning tool.", False, False)])
box([("Teacher’s note — flow before form: ", True, False),
     ("every inversion in the model earns its place. \u2018Rarely\u2019 contrasts with the accusation in sentence 1; "
      "\u2018Not only ... but also ...\u2019 adds the second advantage; \u2018Only by ...\u2019 delivers the concluding "
      "recommendation (condition \u2192 result). The task asks for two inversions; the model uses three, but only "
      "because the flow stays natural. Ask students to read their paragraphs aloud: if an inversion makes the flow "
      "awkward, they should re-order the ideas instead of forcing the structure.", False, False)], color=GREY)

doc.save(OUT)
print("SAVED:", OUT)