# wsqskills

Claude Code **skills, slash commands, an agent and hooks** for building WSQ courseware to the Tertiary Infotech Academy house standards — the slide deck (PPT), Lesson Plan (LP), Learner Guide (LG) and the assessment set (WA + PP/Case Study) — plus the publishing path to Google Drive and the LMS/TMS.

Everything here is course-agnostic: the generators derive the course title, code, version and file names from the course repo they are run in.

---

## Install

### With npx (recommended)

```bash
# into the current course repo (./.claude) — the usual case
npx github:tertiarycourses/wsqskills

# into your user config (~/.claude) — available in every project
npx github:tertiarycourses/wsqskills --user

# overwrite files that already exist (careful: may clobber a course-customised generator)
npx github:tertiarycourses/wsqskills --force
```

The installer copies `skills/`, `commands/`, `agents/` and `hooks/` into the target `.claude/` directory and registers the three hooks in its `settings.json`. **Existing files are never overwritten unless you pass `--force`** — a course repo often carries a course-customised generator, and clobbering it would destroy work.

### By hand

```bash
git clone https://github.com/tertiarycourses/wsqskills.git
cd wsqskills
node bin/install.js --user     # or: node bin/install.js  (project-level)
```

Or just copy the folders: `skills/`, `commands/`, `agents/`, `hooks/` → `~/.claude/` (user) or `<course-repo>/.claude/` (project), then add the hooks block below to `settings.json`.

### By prompt

Inside Claude Code, in the course repo:

> Install the WSQ courseware skills from https://github.com/tertiarycourses/wsqskills into this project — skills, commands, the courseware-qa agent and the three hooks, and register the hooks in .claude/settings.json. Don't overwrite any generator this repo has already customised.

Or, once installed, use the bundled command:

```
/wsq-setup          # import/update the WSQ skills from ~/.claude into this project
```

### Verify

```
/courseware-qa      # audits the course's PPT / LP / LG / assessments
```

---

## What's in the box

### Skills

| Skill | What it does |
|---|---|
| `courseware-build` | Single-source build pipeline — one content module drives the PPT, LP, LG and labs index so they can't drift apart. |
| `wsq-slides` | Slide-deck house standards (admin slide order, trainer profile cards, assessment flow, visual component system). |
| `wsq-lesson-plan` / `tertiary-lesson-plan` | Lesson Plan DOCX — WSQ cover, version-control record, TOC, daily schedule tables. |
| `wsq-learner-guide` / `tertiary-learner-guide` | Learner Guide DOCX + aligned Markdown mirror. |
| `wsq-assessment` | **WA (SAQ)** + **PP or Case Study**, each as a question paper and an answer key. Enforces: follow the original paper, do not change the question count, do not change the instrument type, and full K/A coverage. |
| `tertiary-course-slides` / `tertiary-ppt-design` | python-pptx deck generator and the visual design system it uses. |
| `create-tms-ato` | Client-branded Training Management System document for an ATO / SSG submission. |

### Commands

| Command | What it does |
|---|---|
| `/courseware-gen` | Generate the PPT, LP and LG **plus their PDFs**, archive superseded versions, then audit. |
| `/assessment-gen` | Generate the WA (SAQ) + PP/Case Study — question papers and answer keys — mirroring the original paper. |
| `/courseware-qa` | Audits the deck, LP, LG, labs **and the assessment set** against the published standards — renders pages to images and reports pass/fail. |
| `/gdrive-push` | Push the courseware to the Drive folder (archiving superseded versions) and emit the viewer links. |
| `/tms-push` | Set the courseware URLs on the LMS-TMS course record and attach the assessment — **question papers only; the answer keys never reach the LMS**. |
| `/wsq-setup`, `/importwsq` | Import/update the WSQ skills into the current project. |

The push logic lives in the **commands** (with their scripts in `scripts/`), not in skills.

**The pipeline:** `/courseware-gen` → `/assessment-gen` → `/courseware-qa` → `/gdrive-push` → `/tms-push`

### Agent

`courseware-qa` — the reviewer the command and the hooks delegate the page-by-page visual pass to.

### Hooks

| Hook | Event | What it enforces |
|---|---|---|
| `courseware-pre-hook.py` | PreToolUse (Bash) | Before a generator runs: reuse the reference deck components, read the assessment papers so content stays aligned, bump the version + version-control record. |
| `courseware-post-hook.py` | PostToolUse (Bash) | After a generator runs: **run `/courseware-qa`** on what was just produced — deck rules, and for assessments the cover/instrument name, unchanged question count, and full K/A coverage. Fix and re-run until it passes. |
| `courseware-tms-push-hook.py` | PreToolUse (Bash\|Skill) | Before any push to the TMS/LMS or its Drive folder: **the QA audit must have passed on the exact files being pushed.** A missing K or A is a blocking issue. |

They register in `settings.json` as:

```json
{
  "hooks": {
    "PreToolUse": [
      { "matcher": "Bash",
        "hooks": [{ "type": "command", "command": "python3 \"$CLAUDE_PROJECT_DIR/.claude/hooks/courseware-pre-hook.py\"" }] },
      { "matcher": "Bash|Skill",
        "hooks": [{ "type": "command", "command": "python3 \"$CLAUDE_PROJECT_DIR/.claude/hooks/courseware-tms-push-hook.py\"" }] }
    ],
    "PostToolUse": [
      { "matcher": "Bash",
        "hooks": [{ "type": "command", "command": "python3 \"$CLAUDE_PROJECT_DIR/.claude/hooks/courseware-post-hook.py\"" }] }
    ]
  }
}
```

For a user-level install, replace `$CLAUDE_PROJECT_DIR/.claude` with the absolute path to `~/.claude`.

---

## The quality gates, in short

1. **Assessments follow the original.** Same instrument (a Case Study stays a Case Study), **the same number of questions**, the same K/A codes, the same timings. Only the content is rewritten — from this course's slides and labs.
2. **Every K and every A is covered.** The WA covers the Knowledge items, the PP/CS covers the Abilities, each question prints its own codes. A missing K or A is flagged as an issue, not shipped.
3. **Every assessment paper carries the WSQ cover page**, naming the correct instrument — Written Assessment (SAQ), Practical Performance (PP) or Case Study (CS).
4. **Nothing is published until the audit passes.** The TMS-push hook blocks on it.

---

## Requirements

- Claude Code
- Python 3 with `python-docx`, `python-pptx`, `PyMuPDF`
- LibreOffice (`soffice`) for the DOCX/PPTX → PDF render used by the QA pass
- Node 18+ (for the `npx` installer only)

---

© Tertiary Infotech Academy Pte Ltd (UEN 201200696W) · [tertiarycourses.com.sg](https://www.tertiarycourses.com.sg/)
