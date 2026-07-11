---
description: Update the course on lms-tms.tertiaryinfotech.com via its API — set Trainer Slides / Learner Slides / Learner Guide / Lesson Plan URLs from the course's Google Drive courseware folder
argument-hint: [course-code | drive-folder-link]   (both optional — resolved from the repo + the LMS)
---

# LMS Push

Write this course's Google Drive courseware links into its course record on
`lms-tms.tertiaryinfotech.com`.

**It links what is on Drive — it uploads nothing.** If the repo has a newer build than
Drive, the LMS will keep showing old material and the push will still report success. So
first confirm the Drive files ARE the current build; if not, run `/gdrive-push` first
(generating the deck/LG/LP PDFs it expects). The chain is build → `/gdrive-push` → `/lms-push`.

The course code is read from the courseware itself (deck cover / LG / LP) and must match
the course being written to — a mismatch is a hard abort.

**Optional argument:** `$ARGUMENTS` — a course code (`TGS-…`) and/or a Drive folder link.
Both are normally resolved automatically: the course code from the repo folder name, and
the Drive folder from the course's own **Courseware Link** field in the LMS. Supply one
only to override.

## Steps

1. Locate the `lms-push` skill (project `.claude/skills/lms-push/`, else
   `~/.claude/skills/lms-push/`) and follow its SKILL.md. Run from the course repo root:
   ```bash
   python3 <skill-dir>/lms_push.py --dry-run   # preview first — always
   python3 <skill-dir>/lms_push.py             # real push + verify
   ```
   Add `--course-code TGS-…` / `--drive-folder <link>` only if `$ARGUMENTS` overrides them.
2. Show the user the dry-run plan — the resolved LMS course (title + code), the Drive
   folder it read from the Courseware Link, and per field the **old → new** URL:

   | LMS-TMS field | Source file on Drive |
   |---|---|
   | Trainer Slides URL | the `.pptx` in `Master Trainer Slides` |
   | Learner Slides URL | the `.pdf` with "slide" in its name in `Learner Guide` |
   | Learner Guide URL | the `.pdf` without "slide" in its name in `Learner Guide` |
   | Lesson Plan URL | the `.pdf` in `Lesson Plan` |

   Facilitator Guide URL and Assessment Plan URL are left untouched. If the script reports
   it passed over other candidate files, surface that — it is how a stale deck gets caught.
3. **If a PDF is missing** the script aborts and names the field. Do not invent a link and
   do not silently fall back to the `.docx`. Tell the user, and offer: convert the newest
   DOCX in that Drive folder to PDF and upload it (recipe in SKILL.md — verify page count
   and cover page first), or re-run with `--allow-missing` to set only the found fields.
4. Confirm the course is the right one, then do the real push. It is a live, unauthenticated
   write to a production course page — a wrong course code silently overwrites a real course.
5. The script re-reads the course afterwards and prints ✓/✗ per URL plus a check that no
   other column was blanked. Report that verification; never claim success on a ✗.
