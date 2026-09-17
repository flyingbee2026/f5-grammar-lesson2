# Lesson 2 v4 content data (imported by build_lesson2_v4.py)
# All examples deduped across pattern boxes, paragraphs, sets, quiz and model.

QUIZ_ROWS = [
    [("Sentence",), ("Correction (Teacher’s version)",)],
    [("Social media connects people instantly which makes the world feel smaller.",),
     ("Social media connects people instantly, which makes the world feel smaller. "
      "(Add a comma before the connective clause.)",)],
    [("Online courses are ideal for busy adults, especially those which have little time to attend classes.",),
     ("Online courses are ideal for busy adults, especially those who have little time to attend classes. "
      "(People take ‘who’, not ‘which’.)",)],
    [("There are dozens of messaging apps, all of them are free to use.",),
     ("There are dozens of messaging apps, all of which are free to use. "
      "(A pronoun cannot join two clauses; use ‘which’.)",)],
    [("The forum which students share their homework is very popular.",),
     [("Teacher’s Answer A: The forum on which students share their homework is very popular. ", True, False),
      ("Teacher’s Answer B: The forum where students share their homework is very popular. "
       "(Add the missing preposition ‘on’.)", False, False)]],
    [("Many teenagers spend hours gaming, that worries their parents.",),
     ("Many teenagers spend hours gaming, which worries their parents. "
      "(‘That’ cannot begin a connective clause.)",)],
    [("Not only the internet connects people around the world, but it also spreads information quickly.",),
     ("Not only does the internet connect people around the world, but it also spreads information quickly. "
      "(Move the auxiliary before the subject after ‘Not only’ — more on this today.)",)],
]

# (title, structure, use, example, paragraph_run_tuples) — plain runs, students highlight targets
PATTERNS = [
    ("Pattern 1: Not only ... but also ...",
     "Not only + auxiliary + subject + verb, but (subject) also + verb",
     "Adds a second advantage in proposals, reports and articles.",
     "Not only does the scheme save money, but it also builds confidence.",
     [("Introducing a school e-learning platform would bring clear benefits. ", False, False),
      ("Not only does the platform give students round-the-clock access to learning materials, but it also allows teachers to track progress more closely", False, False),
      (". With such a system, revision becomes a habit rather than a scramble before exams.", False, False)]),
    ("Pattern 2: Never / Rarely / Seldom",
     "Never / Rarely / Seldom + auxiliary + subject + verb",
     "Turns a criticism or a surprising claim into a formal, striking opening.",
     "Rarely do we question our daily screen habits.",
     [("Many parents blame video games for every poor result. ", False, False),
      ("Rarely do they stop to consider how limited their children’s other entertainment options are", False, False),
      (". The real problem may be the lack of balance, not the games themselves.", False, False)]),
    ("Pattern 3: Only by ...",
     "Only by + -ing / noun phrase + auxiliary + subject + verb",
     "Sets a condition that leads to a result — ideal for conclusions and recommendations.",
     "Only by upgrading the campus network can the school offer reliable online lessons.",
     [("Working from home has become common in many industries. ", False, False),
      ("Only by investing in secure networks and clear communication tools can companies keep their teams productive", False, False),
      (". Technology can help, but good management matters more.", False, False)]),
]

# optional pattern in the teacher's note (grey)
NOTE_TITLE = "Additional pattern (teacher’s note — optional)"
NOTE_INTRO = ("Introduce this only after the class is secure with the three core patterns. The mechanism is the "
              "same: front the expression, then place the auxiliary before the subject.")
NOTE_NAME = "So ... that / Such ... that"
NOTE_STRUCTURE = ("So + adjective + auxiliary + subject + verb + that ...  OR  Such + be + noun phrase + that ...")
NOTE_EXAMPLE = "Such is the influence of social media that it shapes public opinion overnight."
NOTE_USE = ("Use it to open with a strong statement when emphasising the scale or importance of an issue.")
NOTE_EXTRA = ("You can also extend Pattern 3 with ‘Only when ...’ or ‘Only after ...’ for stronger classes.")

# (title, simple sentence, prompt, answer)
SETS = [
    ("Example Set 1 (Not only ... but also ...)",
     "The online platform saves paper, and it gives students instant feedback.",
     "Begin with ‘Not only’ to add a second advantage.",
     "Not only does the online platform save paper, but it also gives students instant feedback."),
    ("Example Set 2 (Not only ... but also ...)",
     "E-books are cheaper than printed textbooks, and they are far lighter to carry.",
     "Use ‘Not only ... but also ...’ to combine the two advantages of e-books.",
     "Not only are e-books cheaper than printed textbooks, but they are also far lighter to carry."),
    ("Example Set 3 (Never / Rarely / Seldom)",
     "Parents rarely check what their children are doing online.",
     "Begin with ‘Rarely’.",
     "Rarely do parents check what their children are doing online."),
    ("Example Set 4 (Never / Rarely / Seldom)",
     "The workplace has never relied so heavily on technology.",
     "Begin with ‘Never’.",
     "Never has the workplace relied so heavily on technology."),
    ("Example Set 5 (Only by ...)",
     "Schools can protect students’ privacy by issuing clear guidelines on social media use.",
     "Begin with ‘Only by’.",
     "Only by issuing clear guidelines on social media use can schools protect students’ privacy."),
    ("Example Set 6 (Only by ...)",
     "Bookshops can survive in the digital age by offering experiences that apps cannot match.",
     "Begin with ‘Only by’.",
     "Only by offering experiences that apps cannot match can bookshops survive in the digital age."),
]

INSTR_MAIN = ("Social media plays a big part in students’ lives. Write ONE paragraph (about 80–100 words) "
              "giving your view on whether students should limit their use of social media. "
              "Use ONE inversion and ONE relative clause. ")
INSTR_BOLD = ("Remember: the flow of ideas comes first — never force in an inversion for the sake of using one.")

MODEL = [
    ("Social media is often blamed for distracting students, and every few months someone suggests banning it from school. ", False, False),
    ("Banning, however, treats the symptom, not the cause. ", False, False),
    ("Rarely do students have a more convenient way to exchange ideas and share learning resources, which textbooks alone can never provide. ", False, False),
    ("Through discussion groups, they can prepare for exams together and correct each other’s mistakes in real time. ", False, False),
    ("The real task, therefore, is to teach young people to use social media responsibly, not to lock it away.", False, False),
]

MODEL_NOTE = [
    ("Teacher’s note — flow before form: ", True, False),
    ("the model uses exactly ONE inversion because the task asks for one. ‘Rarely’ opens the evidence section and "
     "contrasts with the accusation in sentence 1; the relative clause ‘which textbooks alone can never provide’ "
     "adds detail without disturbing the flow. Ask students to read their paragraphs aloud: if an inversion makes "
     "the flow awkward, they should re-order the ideas instead of forcing the structure.", False, False),
]