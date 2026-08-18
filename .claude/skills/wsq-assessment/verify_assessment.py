#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Structural verifier for a WSQ assessment set — run it on the RENDERED PDFs.

Why this exists
---------------
The three-page rule and the "no For Official Use Only at the back" rule are
PAGINATION facts. They cannot be checked by reading the .docx: python-docx sees
paragraphs, not pages, so a paper whose questions spill onto page 2 looks
identical to a correct one. A CLSSBB paper shipped to Google Drive with the
questions on page 2 and the assessor sign-off at the back precisely because the
check was done on the source rather than the render.

So: render to PDF first, then run this.

    soffice --headless --convert-to pdf --outdir assessment assessment/*.docx
    python3 ~/.claude/skills/wsq-assessment/verify_assessment.py assessment/

Checks (question papers)
  1. Page 1 is the cover only — no Trainee Information, no questions.
  2. Page 2 carries Trainee Information AND Instructions AND Grading.
  3. No question/task appears before page 3.
  4. "For Official Use Only" appears NOWHERE.
  5. The assessor sign-off is NOT on the last page.
  6. The Instructions carry the LMS upload link.
Checks (answer keys)
  7. No Trainee Information / Instructions / Grading block.
  8. Every question/task carries "Suggestive answers".

Exit code 1 if anything fails, so it can gate a push.
"""
import glob
import os
import re
import subprocess
import sys


def page_text(pdf, n):
    r = subprocess.run(["pdftotext", "-f", str(n), "-l", str(n), pdf, "-"],
                       capture_output=True, text=True)
    return r.stdout


def n_pages(pdf):
    r = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True)
    m = re.search(r"Pages:\s+(\d+)", r.stdout)
    return int(m.group(1)) if m else 0


def full_text(pdf):
    return subprocess.run(["pdftotext", pdf, "-"], capture_output=True, text=True).stdout


QT = re.compile(r"^\s*(Question|Task)\s+\d+\s*:", re.M)


def check(pdf):
    name = os.path.basename(pdf)
    is_key = re.match(r"^\s*answers?\s+to\b", name, re.I) is not None
    total = n_pages(pdf)
    fails = []

    if total < 3 and not is_key:
        fails.append(f"only {total} pages — a question paper needs the cover, page 2 and questions from page 3")

    all_txt = full_text(pdf)
    p1 = page_text(pdf, 1)
    p2 = page_text(pdf, 2) if total >= 2 else ""
    last = page_text(pdf, total) if total else ""

    # 4 — applies to papers AND keys
    if re.search(r"for official use", all_txt, re.I):
        fails.append("'For Official Use Only' present — the assessor sign-off belongs in the page-2 Grading block")

    if not is_key:
        # 1
        if QT.search(p1) or re.search(r"Trainee Information", p1, re.I):
            fails.append("page 1 is not cover-only")
        # 2
        for need in ("Trainee Information", "Instructions to Candidate", "Grading"):
            if not re.search(need, p2, re.I):
                fails.append(f"page 2 is missing '{need}'")
        # 3
        if QT.search(p2):
            fails.append("a Question/Task appears on page 2 — content must start on page 3")
        # The rule is "assessment CONTENT starts on page 3" — for a Case Study the
        # scenario legitimately fills page 3 and the questions start on page 4 or
        # later. So require content (a question OR the scenario) on page 3, and
        # never require a question there specifically.
        p3 = page_text(pdf, 3) if total >= 3 else ""
        if total >= 3:
            has_q = QT.search(p3) is not None
            has_scenario = re.search(r"case study|scenario", p3, re.I) is not None
            if not (has_q or has_scenario):
                fails.append("page 3 carries neither the scenario nor a Question/Task — "
                             "assessment content must start on page 3")
        # 5
        if re.search(r"assessor name", last, re.I):
            fails.append("assessor sign-off on the LAST page — it belongs on page 2")
        # 6
        if "lms-tms.tertiaryinfotech.com" not in all_txt:
            fails.append("Instructions do not carry the LMS upload link")
    else:
        # 7
        for banned in ("Trainee Information", "Instructions to Candidate", "Grading"):
            if re.search(banned, all_txt, re.I):
                fails.append(f"answer key contains '{banned}' — that belongs to the candidate paper only")
        # 8
        n_items = len(QT.findall(all_txt))
        n_sugg = len(re.findall(r"Suggestive answers", all_txt, re.I))
        if n_items and n_sugg < n_items:
            fails.append(f"{n_items} questions/tasks but only {n_sugg} 'Suggestive answers' blocks")

    return name, total, fails


def main():
    target = sys.argv[1] if len(sys.argv) > 1 else "assessment"
    pdfs = sorted(glob.glob(os.path.join(target, "*.pdf")))
    if not pdfs:
        print(f"No PDFs in {target}/ — render the .docx first:")
        print(f"  soffice --headless --convert-to pdf --outdir {target} {target}/*.docx")
        return 1
    bad = 0
    for pdf in pdfs:
        name, total, fails = check(pdf)
        if fails:
            bad += 1
            print(f"FAIL  {name}  ({total} pages)")
            for f in fails:
                print(f"        - {f}")
        else:
            print(f"PASS  {name}  ({total} pages)")
    print()
    print("All assessment documents pass the structural checks."
          if not bad else f"{bad} document(s) FAILED — fix and re-render before pushing.")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
