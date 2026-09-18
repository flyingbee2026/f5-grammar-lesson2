# Produce build_teacher_v7.py from build_teacher_v6.py: harder 2nd questions + labels + clarity tweaks
src = open("/opt/f5-grammar-lesson2/build_teacher_v6.py", encoding="utf-8").read()

reps = [
    ('OUT = "/opt/f5-grammar-lesson2/Grammar_Lesson_2_Inversions_Teacher_v6.docx"',
     'OUT = "/opt/f5-grammar-lesson2/Grammar_Lesson_2_Inversions_Teacher_v7.docx"'),
    ('add_para("Rewrite each simple sentence using the pattern given in the prompt. The sentences are all about "\n         "technology and communication, so keep the register formal.")',
     'add_para("Rewrite each simple sentence using the pattern given in the prompt. The sentences are all about "\n         "technology and communication, so keep the register formal. The even-numbered sets (2, 4 and 6) are "\n         "more challenging.")'),
    # Set 2 - harder
    ('    ("Example Set 2 (Not only ... but also ...)",\n'
     '     "E-books are cheaper than printed textbooks, and they are far lighter to carry.",\n'
     '     "Use ‘Not only ... but also ...’ to combine the two advantages of e-books.",\n'
     '     "Not only are e-books cheaper than printed textbooks, but they are also far lighter to carry."),',
     '    ("Example Set 2 (Not only ... but also ...) – More challenging",\n'
     '     "The workshop improved students’ digital skills, and it boosted their confidence in online presentations.",\n'
     '     "Use ‘Not only ... but also ...’ to combine the workshop’s two benefits.",\n'
     '     "Not only did the workshop improve students’ digital skills, but it also boosted their confidence in online presentations."),'),
    # Set 4 - harder
    ('    ("Example Set 4 (Never / Rarely / Seldom)",\n'
     '     "The workplace has never relied so heavily on technology.",\n'
     '     "Begin with ‘Never’.",\n'
     '     "Never has the workplace relied so heavily on technology."),',
     '    ("Example Set 4 (Never / Rarely / Seldom) – More challenging",\n'
     '     "The Internet has never made it so easy to spread false information.",\n'
     '     "Begin with ‘Never’.",\n'
     '     "Never has the Internet made it so easy to spread false information."),'),
    # Set 6 - harder
    ('    ("Example Set 6 (Only by ...)",\n'
     '     "Bookshops can survive in the digital age by offering experiences that apps cannot match.",\n'
     '     "Begin with ‘Only by’.",\n'
     '     "Only by offering experiences that apps cannot match can bookshops survive in the digital age."),',
     '    ("Example Set 6 (Only by ...) – More challenging",\n'
     '     "Governments can fight online scams only by working closely with technology companies.",\n'
     '     "Begin with ‘Only by’.",\n'
     '     "Only by working closely with technology companies can governments fight online scams."),'),
    # clarity: pattern-1 example sentence
    ('"Not only does the scheme save money, but it also builds confidence."',
     '"Not only does the scheme save money, but it also builds students’ confidence."'),
]
for old, new in reps:
    assert old in src, "NOT FOUND: " + old[:60]
    src = src.replace(old, new)
open("/opt/f5-grammar-lesson2/build_teacher_v7.py", "w", encoding="utf-8").write(src)
print("build_teacher_v7.py written")