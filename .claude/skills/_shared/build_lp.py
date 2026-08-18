import os, json
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import brand_doc as B
import config as C
import lab_parser as LP
import paths as P

SP = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = P.OUT
man = json.load(open(P.MANIFEST))
LABS = {L['num']: L for L in LP.load_all(P.LABS)}

TOTAL = man['total_slides']
lab_slide = {int(k): v['slide'] for k, v in man['labs'].items()}
anchors = sorted(list(lab_slide.values()) + [d['section_slide'] for d in man['domains']])

first_lab_slide = {d['topic']: min(lab_slide[l] for l in d['labs']) for d in man['domains']}
section_slide = {d['topic']: d['section_slide'] for d in man['domains']}

_PRACTICE = man.get('practice_section')
def lab_range(ln):
    s = lab_slide[ln]
    nxt = [a for a in anchors if a > s]
    if nxt:
        e = min(nxt) - 1
    elif _PRACTICE:
        # last lab is followed by its topic recap, then the practice-exam section
        e = _PRACTICE - 2
    else:
        e = TOTAL - 8
    return s, e

def theory_range(dom):
    """Slides between a topic's Overview (section+1) and its first lab = the theory block."""
    s = dom['section_slide'] + 2  # section divider + 2-slide overview
    e = first_lab_slide[dom['topic']] - 1
    return s, e

def rng(a, b): return str(a) if a == b else f"{a}–{b}"
def short(ln):
    t = LABS[ln]['title']; return t.split('—',1)[-1].strip()
def span(labnums):
    s, _ = lab_range(labnums[0]); _, e = lab_range(labnums[-1]); return rng(s, e)

doc = B.new_doc()

def center_img(path, w):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(4)
    p.add_run().add_picture(path, width=Inches(w))
def center_text(t, s, b=False, col=None, sp=4):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after = Pt(sp)
    r = p.add_run(t); r.font.size = Pt(s); r.font.bold = b; r.font.name='Arial'
    if col: r.font.color.rgb = col

# ---- Cover ----
doc.add_paragraph().paragraph_format.space_after = Pt(2)
center_img(os.path.join(B.ASSETS, "tertiary_logo.png"), 1.5)
center_text("Tertiary Infotech Academy Pte Ltd", 13, True, B.INK, 1)
center_text("UEN: 201200696W", 9.5, False, B.GREY, 10)
center_text("LESSON PLAN", 26, True, B.BLUE, 2)
center_text("For", 11, False, B.GREY, 10)
center_img(os.path.join(B.ASSETS, "cysa_logo.png"), 1.3)
doc.add_paragraph().paragraph_format.space_after = Pt(2)
center_text(C.COURSE_TITLE, 20, True, B.INK, 4)
center_text(f"TGS Ref No: {C.COURSE_CODE}", 11, False, B.INK, 10)
center_text(f"Version {C.VERSION_DOC}", 12, True, B.BLUE, 4)
doc.add_page_break()

# ---- Version control ----
B.heading(doc, "Document Version Control Record", 1)
B.blue_table(doc, ["Version", "Effective Date", "Summary of Changes", "Author"], [
    ["1.0", "6 July 2026",
     "First edition — 5-day lesson plan aligned to the 30 hands-on CySA+ labs, with slide-number references to the Learner Guide Slides deck.",
     "Tertiary Infotech Academy Pte Ltd"],
    ["2.0", "18 Aug 2026",
     "Second edition — courseware updated to v7 Slides: full theory content from the legacy deck reordered to the four CS0-003 exam domains, exam domains shown up front, practice exam and CompTIA CySA+ exam registration added at the end.",
     "Tertiary Infotech Academy Pte Ltd"],
    ["2.1", "18 Aug 2026",
     "House-format fixes — daily schedule corrected to 9:30am-6:30pm with a 1-hour lunch (8 instructional hours/day), theory-delivery rows added per day, slide-reference summary corrected, Table of Contents and 11pt body applied.",
     "Tertiary Infotech Academy Pte Ltd"],
], widths=[Inches(0.8), Inches(1.2), Inches(3.6), Inches(1.4)])
doc.add_page_break()

# ---- Table of Contents ----
B.add_toc(doc, entries=[
    (1, "Course Details"), (1, "Course Overview"), (1, "Learning Outcomes"),
    (1, "Day 1"), (1, "Day 2"), (1, "Day 3"), (1, "Day 4"), (1, "Day 5"),
    (1, "Slide Reference Summary"),
])

# ---- Course details ----
B.heading(doc, "Course Details", 1)
B.blue_table(doc, ["Field", "Detail"], [
    ["Course Title", C.COURSE_TITLE],
    ["Course Code", C.COURSE_CODE],
    ["WSQ TSC", f"{C.TSC_TITLE} ({C.TSC_CODE})"],
    ["Duration", "5 Days (40 hours)"],
    ["Daily Schedule", "9:30 AM – 6:30 PM (8 instructional hours/day)"],
    ["Lunch Break", "1:00 PM – 2:00 PM (1 hour)"],
    ["Breaks", "Short tea breaks are scheduled within each day's training hours"],
    ["Delivery Mode", "Instructor-led presentation, live demonstration and hands-on labs (Killercoda / free tools)"],
    ["Assessment", "Written Assessment (SAQ) 1 hr + Case Study (CS) 1 hr — Open Book (Day 5)"],
    ["Prerequisites", "Basic networking and Linux command-line familiarity; interest in cybersecurity operations"],
    ["Slide Deck", f"{man['ppt_file']} ({TOTAL} slides)"],
], widths=[Inches(1.8), Inches(4.8)], label_col0=True)
B.para(doc, "Note: the 'Slides' column in each daily schedule references the exact slide numbers in the Learner Guide Slides deck for that activity.",
       italic=True, color=B.GREY, size=9.5)

# ---- Overview + LO ----
B.heading(doc, "Course Overview", 1)
B.para(doc, "This 5-day hands-on course develops practical cybersecurity analyst skills aligned to the CompTIA CySA+ "
            "CS0-003 exam. Across four domains — Security Operations, Vulnerability Management, Incident Response, and "
            "Reporting & Communication — learners complete 30 step-by-step labs using the free Killercoda playground and "
            "free industry tools, culminating in a written and case-study assessment.")
B.heading(doc, "Learning Outcomes", 1)
B.para(doc, "By the end of this course, participants will be able to:", bold=True, space=2)
B.bullets(doc, C.LEARNING_OUTCOMES)
doc.add_page_break()

# ---- Daily schedule tables ----
HDR = ["Time", "Topic / Activity", "Duration", "Slides"]
WID = [Inches(1.15), Inches(3.5), Inches(0.85), Inches(0.9)]

def sched_table(rows):
    t = doc.add_table(rows=1, cols=4); t.style = 'Table Grid'
    for c, h in zip(t.rows[0].cells, HDR):
        B.shade(c._tc.get_or_add_tcPr(), '1F6FEB'); c.paragraphs[0].text = ''
        r = c.paragraphs[0].add_run(h); r.font.bold = True; r.font.color.rgb = B.WHITE; r.font.size = Pt(9.5); r.font.name='Arial'
    for row in rows:
        kind = row[0]; vals = row[1:]
        cells = t.add_row().cells
        fill = {'topic': 'EAF1FD', 'break': 'FDF3E7'}.get(kind)
        for i, (cell, v) in enumerate(zip(cells, vals)):
            cell.paragraphs[0].text = ''
            r = cell.paragraphs[0].add_run(str(v)); r.font.size = Pt(9); r.font.name='Arial'
            if kind == 'topic': r.font.bold = True; r.font.color.rgb = B.BLUE
            elif kind == 'break': r.font.italic = True; r.font.color.rgb = B.GREY
            else: r.font.color.rgb = B.INK
            if fill: B.shade(cell._tc.get_or_add_tcPr(), fill)
    for row in t.rows:
        for i, wd in enumerate(WID): row.cells[i].width = wd
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def labs_rows(labnums, blocks):
    """distribute labnums across content 'blocks' [(time,dur)] -> rows list."""
    n = len(blocks); groups = [[] for _ in range(n)]
    for i, ln in enumerate(labnums): groups[i * n // len(labnums)].append(ln)
    rows = []
    for (time, dur), grp in zip(blocks, groups):
        if not grp: continue
        act = "; ".join(f"Lab {x}: {short(x)}" for x in grp)
        rows.append(('normal', time, act, dur, span(grp)))
    return rows

def day(dayno, subtitle, labnums, admin=False):
    B.heading(doc, f"Day {dayno}", 1)
    B.para(doc, subtitle, italic=True, color=B.GREY, space=6)
    dom = next(d for d in man['domains'] if d['labs'] == labnums) if labnums else None
    th = theory_range(dom) if dom else None
    # ---- Morning 9:30–1:00 (3.5h instruction incl. 15-min tea) ----
    B.para(doc, "Morning Session  ·  9:30 AM – 1:00 PM", bold=True, color=B.BLUE, space=3)
    rows = []
    if admin:
        rows += [('normal', "9:30 – 9:50", "Digital Attendance (AM); Trainer & Learner Introduction; Ground Rules", "20 min", "2–7"),
                 ('normal', "9:50 – 10:30", "Lesson Plan; Skills Framework; Learning Outcomes; CS0-003 Exam Domains; Course Outline", "40 min", "9–19"),
                 ('normal', "10:30 – 10:45", "Briefing for Assessment; Assessment; Criteria for Funding", "15 min", "20–23"),
                 ('break', "10:45 – 11:00", "Tea Break", "15 min", "—")]
        if dom:
            rows.append(('topic', "", f"{dom['topic']}: {dom['name']} — Topic Overview", "—", rng(dom['section_slide'], dom['section_slide']+1)))
            rows.append(('normal', "11:00 – 1:00", f"Theory: {dom['name']} concepts", "120 min", rng(*th)))
    else:
        rows += [('normal', "9:30 – 9:40", "Digital Attendance (AM); Recap of previous day", "10 min", "—")]
        if dom:
            rows.append(('topic', "", f"{dom['topic']}: {dom['name']} — Topic Overview", "—", rng(dom['section_slide'], dom['section_slide']+1)))
            rows.append(('normal', "9:40 – 10:45", f"Theory: {dom['name']} concepts (part 1)", "65 min", rng(th[0], (th[0]+th[1])//2)))
        rows.append(('break', "10:45 – 11:00", "Tea Break", "15 min", "—"))
        if dom:
            rows.append(('normal', "11:00 – 1:00", f"Theory: {dom['name']} concepts (part 2)", "120 min", rng((th[0]+th[1])//2 + 1, th[1])))
    rows.append(('break', "1:00 – 2:00", "Lunch Break", "60 min", "—"))
    sched_table(rows)
    # ---- Afternoon 2:00–6:30 (4.5h instruction incl. 15-min tea) ----
    B.para(doc, "Afternoon Session  ·  2:00 PM – 6:30 PM", bold=True, color=B.BLUE, space=3)
    rows2 = [('normal', "2:00 – 2:05", "Digital Attendance (PM)", "5 min", "—")]
    pm_blocks = [("2:05 – 4:00", "115 min"), ("4:15 – 6:25", "130 min")]
    pm_rows = labs_rows(labnums, pm_blocks)
    if pm_rows:
        rows2.append(('topic', "", f"Hands-On Labs {labnums[0]}–{labnums[-1]}", "—", span(labnums)))
        rows2.append(pm_rows[0]); rows2.append(('break', "4:00 – 4:15", "Tea Break", "15 min", "—"))
        rows2 += pm_rows[1:]
    rows2.append(('normal', "6:25 – 6:30", f"Day {dayno} Recap; Q&A", "5 min", rng(*lab_range(labnums[-1]))))
    rows2.append(('normal', "6:30", f"End of Day {dayno}", "—", "—"))
    sched_table(rows2)
    doc.add_page_break()

DOMS = [(1, "Topic 1: Security Operations", list(range(1,11)), True),
        (2, "Topic 2: Vulnerability Management", list(range(11,20)), False),
        (3, "Topic 3: Incident Response and Management", list(range(20,26)), False),
        (4, "Topic 4: Reporting and Communication", list(range(26,31)), False)]
for dn, sub, labs, adm in DOMS:
    day(dn, sub, labs, admin=adm)

# Day 5 — Revision, Practice Exam and Final Assessment
# Locate the practice-exam and exam-registration section dividers in the deck.
practice_sec = man.get('practice_section')
examreg_sec = man.get('examreg_section')
wrap_sec = man.get('wrap_section')
last_lab_end = lab_range(30)[1]
prac_start = practice_sec or (last_lab_end + 2)
examreg = examreg_sec or (prac_start + 17)
back = wrap_sec or (examreg + 3)

B.heading(doc, "Day 5", 1)
B.para(doc, "Revision, Practice Exam and Final Assessment", italic=True, color=B.GREY, space=6)
B.para(doc, "Morning Session  ·  9:30 AM – 1:00 PM", bold=True, color=B.BLUE, space=3)
sched_table([
    ('normal', "9:30 – 9:40", "Digital Attendance (AM)", "10 min", "—"),
    ('normal', "9:40 – 10:45", "Revision: Topic 1 Security Operations & Topic 2 Vulnerability Management", "65 min",
        f"{rng(section_slide['Topic 1'], lab_range(10)[1])}, {rng(section_slide['Topic 2'], lab_range(19)[1])}"),
    ('break', "10:45 – 11:00", "Tea Break", "15 min", "—"),
    ('normal', "11:00 – 1:00", "Revision: Topic 3 Incident Response & Topic 4 Reporting", "120 min",
        f"{rng(section_slide['Topic 3'], lab_range(25)[1])}, {rng(section_slide['Topic 4'], last_lab_end)}"),
    ('break', "1:00 – 2:00", "Lunch Break", "60 min", "—"),
])
B.para(doc, "Afternoon Session  ·  2:00 PM – 6:30 PM", bold=True, color=B.BLUE, space=3)
sched_table([
    ('normal', "2:00 – 2:05", "Digital Attendance (PM)", "5 min", "—"),
    ('normal', "2:05 – 3:15", "Practice Exam (warm-up quiz) and review", "70 min", rng(prac_start, examreg-1)),
    ('normal', "3:15 – 3:45", "CompTIA CySA+ exam registration walkthrough (Pearson VUE)", "30 min", rng(examreg, back-1)),
    ('break', "3:45 – 4:00", "Tea Break", "15 min", "—"),
    ('normal', "4:00 – 4:15", "Course Feedback & TRAQOM Survey; Assessment Briefing; Digital Attendance (Assessment)", "15 min", rng(back, TOTAL)),
    ('normal', "4:15 – 6:15", "Final Assessment: Written (SAQ) 1 hr + Case Study (CS) 1 hr", "120 min", "—"),
    ('normal', "6:15 – 6:30", "Assessment wrap-up; Recommended Courses; Support", "15 min", "—"),
    ('normal', "6:30", "End of Class", "—", "—"),
])

# ---- Slide reference summary ----
doc.add_page_break()
B.heading(doc, "Slide Reference Summary", 1)
rows = [["Front Admin", "Title, Attendance, Trainer, Ground Rules, Download Material, Lesson Plan, Skills Framework, Learning Outcomes, CS0-003 Exam Domains, Course Outline, Briefing, Assessment, Funding", "1–23"]]
for d in man['domains']:
    rows.append([f"{d['topic']}: {d['name']}",
                 f"Section + Overview + Theory + Labs {d['labs'][0]}–{d['labs'][-1]} + Recap",
                 rng(d['section_slide'], lab_range(d['labs'][-1])[1] + 1)])
for ln in range(1,31):
    s,e = lab_range(ln); rows.append([f"Lab {ln}", short(ln), rng(s,e)])
rows.append(["Practice Exam", "Practice section, warm-up quiz and answer key", rng(prac_start, examreg-1)])
rows.append(["Exam Registration", "CompTIA CySA+ exam facts and Pearson VUE registration steps", rng(examreg, back-1)])
rows.append(["Back Admin (Wrap-Up)", "Summary, TRAQOM, Attendance, Assessment, Assessment Flow, Recommended Courses, Support, Thank You", rng(back, TOTAL)])
B.blue_table(doc, ["Section", "Content", "Slides"], rows, widths=[Inches(1.9), Inches(3.9), Inches(0.9)])

B.add_footer(doc)
lp_name = f"LP-{C.COURSE_TITLE}.docx"
doc.save(os.path.join(OUT_DIR, lp_name))
print("Saved LP:", lp_name)
print("lab ranges:", {ln: lab_range(ln) for ln in [1,10,11,19,20,25,26,30]})
