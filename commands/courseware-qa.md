---
description: Audit the WSQ courseware (PPT, LP, LG) and the assessment set (WA + PP/CS) against the Tertiary Infotech house standards — renders every checked page to an image and reports pass/fail.
---

# /courseware-qa — WSQ courseware quality audit

Audit this course's artifacts against the house standards below. **Render the pages you check to images and LOOK at them** — a text-only check misses overlapping, clipped and missing visuals, which are the defects this command exists to catch.

Arguments (optional): `$ARGUMENTS` — a specific artifact to audit (e.g. `ppt`, `lp`, `lg`, `assessment`). With no argument, audit **everything**.

## How to run it

1. Locate the artifacts: the deck and documents in `courseware/` (ignore anything under `courseware/archive/`), the assessment DOCX in `assessment/`, and the original/reference papers in `reference/` if present.
2. Convert to PDF and render to PNG, then inspect the images:
   - `soffice --headless --convert-to pdf --outdir <scratchpad> <file>`
   - render pages with PyMuPDF; **always** look at the cover, the admin slides (front and end), the assessment pages, and every page you changed.
3. Work through the checklists. For each item report **PASS** or **FAIL — file + page + what is wrong**.
4. **Fix every failure, regenerate, and re-run this check** until it passes. Do not report completion on a failing audit.
5. Delegate to the **courseware-qa agent** for the page-by-page visual pass when the audit is large — it has the same standards.

## A. Slide deck (PPT) — hard rules

1. **Two trainer profiles** — a **General Trainer template** card *and* a named **Dr Alfred Ang** profile, as two separate visual profile pages (never a bullet list).
2. **Download Course Material** — a slide showing learners how to download the courseware from **lms-tms.tertiaryinfotech.com**, as a screenshot / step graphic, **not** a bare text link.
3. **Assessment Flow visual** — an Assessment Flow **diagram** (chevron/flow graphic) mapping the assessment path (WA → PP/CS → sign-off). Not a bulleted list.
4. **Practice Exam** — the matching practice exam from **exams.tertiaryinfotech.com** surfaced on a visual slide with the exam link.
5. **Version on cover** — the version number is printed on the cover page and **matches the `<<Course Title>>-vNN` filename**.
6. **Single version on cover** — exactly **one** version number appears on the cover; no duplicate or conflicting version labels anywhere on it.
7. **Admin slide order** — Briefing for Assessment **before** the Assessment slide; the closing block runs **Assessment → Assessment Flow → Digital Attendance (TRAQOM) → Thank You**; TRAQOM/digital attendance appears at the **front and the end**.
8. **Visual, not walls of text** — tile grids, flow diagrams, cards and profile cards; no bullet-wall slides.
9. **No overlapping, clipped or off-slide text**; copyright line + course title + UEN present.
10. **One version only** in `courseware/` — superseded versions moved to `courseware/archive/`.

## B. Lesson Plan (LP) and Learner Guide (LG)

1. WSQ **cover page** — course title, logos, UEN, TGS course code, version.
2. **Document Version Control Record** present, and the **version was bumped with a new row** for this change (LP/LG carry the record; assessments do not).
3. Auto **Table of Contents**; Arial 11pt body; copyright + page-number footer on every page.
4. LP: each training day totals the stated instructional hours; the daily schedule tables are colour-coded and consistent with the deck.
5. LG: content matches the labs and the deck — nothing assessed that is not taught.
6. No clipped/overflowing text or broken images.

## C. Assessment set (WA + PP **or** CS) — hard rules

1. **Assessment cover page** — **every** assessment paper (WA, PP, Case Study) and every answer key carries the WSQ cover page **before its questions**: course title, logos, UEN, TGS ref, version. Assessments carry the **cover page only — no Document Version Control Record**.
2. **The cover must name the instrument correctly** — "Written Assessment (SAQ)", "Practical Performance (PP)" or "Case Study (CS)", with answer keys marked as the Answer Key. A Case Study paper must never be labelled PP, and the reverse. Check the section headings too, not just the cover.
3. **Number of questions is UNCHANGED** — the count of WA questions and of PP/CS tasks must equal the **original / reference / previous** assessment (see `reference/`). If a count differs, that is a **FAIL** — flag it.
4. **K and A coverage (flag any gap as an issue):**
   - The **WA covers the Knowledge items** — every `K1 … Km` is assessed by at least one WA question.
   - The **PP / CS covers the Abilities** — every `A1 … An` is assessed by at least one task.
   - Every question/task **prints its own codes** on the paper (e.g. `(K1, K2)`, `(A3, A4)`), and the answer key repeats the **identical** codes.
   - Report the coverage as a table: code → question(s) → lab/slide. **A missing K or A is a FAIL, not a caveat.**
5. **Instrument type matches the original** — CS stays CS, PP stays PP.
6. **All questions OPEN-ENDED** — zero multiple choice anywhere.
7. **Traceability** — every WA question traces to a slide/module; every PP/CS task and its model answer traces to an in-class **lab** and cites it.
8. **Question paper structure — the page layout is FIXED.** Page 1 = the WSQ cover. **Page 2 = Trainee Information + Instructions to Candidate (with the clickable LMS link https://lms-tms.tertiaryinfotech.com/) + the Grading / For Official Use Only block (Grade C / NYC, assessor name, NRIC, date, signature) — and nothing else.** **Page 3 = where the scenario and the questions/tasks start.** A question or the scenario appearing on page 2, or the instructions/grading spilling onto page 3, is a **FAIL**. (Answer keys are trainer copies: cover, then model answers.) Every question/task has a boxed answer space.
9. **Timings match the original papers** (e.g. WA 60 minutes, CS 90 minutes).
10. **Model-answer tables render as real tables** — no wrapped ASCII columns, no row split across a page break, nothing past the right margin.

## Report format

Report concisely:

- A **PASS / FAIL** line per section (A, B, C).
- Every failure as: `file · page · defect · fix`.
- The **K/A coverage table**.
- End with the overall verdict. If anything failed: fix it, regenerate, and re-run this command.
