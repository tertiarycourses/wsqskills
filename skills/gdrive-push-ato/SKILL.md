---
name: gdrive-push-ato
description: Upload one or more local documents (TMS, courseware and other ATO supporting docs) to a user-provided Google Drive folder, moving any superseded Drive file into an `archive` subfolder first (only when the file actually changed), and remove stray macOS .DS_Store / Office lock junk locally before upload. Then add the assessment to the LMS course record — the WA (SAQ) question paper and the PP or Case Study question paper ONLY; the answer keys are trainer-only and are NEVER uploaded to the LMS. Use when the user asks to "push to Drive", "gdrive push", "upload to Google Drive", "upload the TMS/document to Drive", "archive the old version", or "replace the material" in a Drive folder. ALWAYS requires the user to provide the destination Google Drive folder link.
---

# GDrive Push (ATO)

Uploads local files into the matching subfolders of a Google Drive folder the user names.
Before replacing any file that has **actually changed**, it moves the existing Drive copy
into an `archive` subfolder (one per target subfolder) so nothing is ever lost. Unchanged
files are skipped entirely — no re-upload, no archiving. **Upload-only: nothing on Drive is
ever deleted.** Strips `.DS_Store` and `~$*` junk from the local source first so it can't be
pushed.

## Transport — rclone, NOT the MCP connector

Use **rclone** (same as the `gdrive-push` skill), talking to Drive directly. This is
deliberate: the claude.ai Google Drive **MCP connector cannot move or delete files** (it only
exposes read/search/create/copy), so it *cannot* archive-and-replace — it would only pile up
duplicates. rclone supports server-side move (`moveto`), MD5 change-detection (`--checksum`),
and scoping to one folder (`--drive-root-folder-id`), which is exactly what this flow needs.
Every Drive operation is an upload (`copyto`) or a server-side move to `archive/` (`moveto`) —
**never a delete.** The MCP connector may still be used read-only to verify/enumerate.

- Remote name defaults to `gdrive` (override with env `GDRIVE_REMOTE`). Check with
  `rclone listremotes`.
- One-time setup if missing: `brew install rclone` then
  `rclone config create gdrive drive scope=drive` and complete the Google sign-in in the
  browser **using the account that owns (or has Editor on) the destination folder**.
- `--drive-root-folder-id <FOLDER_ID>` scopes every call to the user's folder, so the script
  can never touch anything outside it.

## HARD RULE — folder link required

**Never upload without the user-provided Google Drive folder link.** If it isn't given, ask
for it (AskUserQuestion) and wait. The link is
`https://drive.google.com/drive/folders/<FOLDER_ID>` (any `/u/N/` only selects a browser
profile and is NOT part of the ID). The folder ID is the last path segment. Confirm it before
the first real (non-dry-run) upload.

## HARD RULE — assessments on the LMS: QUESTION PAPERS ONLY, NEVER the answer keys

When this skill runs for a WSQ course, it also **adds the assessment to the LMS** course record
(lms-tms.tertiaryinfotech.com), after the Drive upload. Exactly **two** documents are attached —
both are the **question papers the learner sits**:

| LMS field | Upload | Local file |
|---|---|---|
| **Written Assessment** | the **WA (SAQ)** question paper | `assessment/WA (SAQ) - <Title> - <VER>.docx` |
| **Practical Performance** *or* **Case Study** | the **PP** *or* **CS** question paper — whichever instrument the course uses | `assessment/PP Assessment - <Title> - <VER>.docx` **or** `assessment/Case Study (CS) - <Title> - <VER>.docx` |

**NEVER upload an answer key / model answer / marking guide to the LMS.** Any file whose name
begins with `Answer to …` or `Answers to …` is a **trainer-only** document: it stays in the
Google Drive courseware folder (trainer access) and is **never** attached to the LMS, never
linked in a learner-visible field, and never pushed to GitHub. Before attaching anything, filter
the file list:

```bash
# the ONLY two files eligible for the LMS
ls assessment/*.docx | grep -viE '^.*/(Answer|Answers) to '
```

If that filter yields anything other than the WA paper plus exactly one PP/CS paper, **stop and
tell the user** — do not guess which file to attach.

Use the course's own **Courseware Link** field (read via the LMS API, as the `lms-push` skill
does) to locate the Drive folder, and set the assessment fields to the Drive links of those two
question papers. Report which two files were attached, and state explicitly that the answer keys
were withheld.

## Steps

1. **Confirm inputs**: the destination folder link, and which local folder/files to upload.
2. **Local junk sweep** — before uploading, delete junk so it can't be pushed:
   ```bash
   find "<local_folder>" -name ".DS_Store" -type f -delete
   find "<local_folder>" -name '~$*'      -type f -delete   # Word/Excel lock files
   ```
3. **Verify access**: `rclone lsd "gdrive:" --drive-root-folder-id <FOLDER_ID>` should list the
   target's subfolders. If it errors or is empty when files are expected, the remote's account
   lacks access — tell the user to share the folder (Editor) with that account or re-auth the
   remote to the owning account. Do not silently fail.
4. **Dry-run FIRST — always.** Preview change-detection and what would be archived, per
   subfolder. For each local subfolder `S` that maps to a Drive subfolder:
   ```bash
   rclone copy "<local_folder>/S" "gdrive:S" \
     --drive-root-folder-id <FOLDER_ID> \
     --checksum \
     --backup-dir "gdrive:S/archive/<YYYY-MM-DD>" \
     --dry-run -v
   ```
   `--checksum` compares MD5, so **identical files are skipped** (no upload, no archive) and
   only genuinely changed/new files are acted on. `--backup-dir` makes rclone **move** the
   superseded Drive copy into that subfolder's `archive/<date>/` before writing the new one,
   instead of overwriting or deleting. (If rclone reports a backup-dir/destination overlap for
   your rclone version, put the archive one level up as a sibling, e.g.
   `--backup-dir "gdrive:_archive/S/<YYYY-MM-DD>"`, which keeps it inside the target folder
   without overlapping the subfolder.) Show the user the plan.
5. **Real upload**: rerun each command without `--dry-run`. `rclone copy` (not `sync`) never
   removes Drive-only files, so unrelated existing files are left untouched; only changed files
   are archived-then-replaced, and new files are added. `archive/` is created on demand and
   only in subfolders that actually had a changed file.
6. **New subfolders / loose files**: for a local file whose subfolder doesn't exist on Drive
   yet, `rclone copyto "<local>/path" "gdrive:path" --drive-root-folder-id <FOLDER_ID>` (rclone
   creates parents). Nothing to archive for brand-new paths.
7. **Attach the assessment to the LMS** (WSQ courses): take the **WA (SAQ) question paper** and
   the **PP or Case Study question paper** — and nothing else — and set them on the course record
   at lms-tms.tertiaryinfotech.com. **Filter out every `Answer to …` / `Answers to …` file first.**
   See the HARD RULE above. If the course has no assessment folder, skip this step and say so.
8. **Report** per subfolder: what was uploaded (with links), what was archived and to where,
   and what was skipped as unchanged. To emit "anyone with the link can view" links, run
   `rclone link "gdrive:S/<file>" --drive-root-folder-id <FOLDER_ID>` per file and include them.
   State which two assessment papers were attached to the LMS, and that the answer keys were
   withheld as trainer-only.

## Drive `.DS_Store` note

Because step 2 strips junk locally and `rclone copy` never pushes files that aren't in the
source, no `.DS_Store`/lock files are uploaded. If pre-existing junk already sits on Drive from
an earlier browser upload, rclone can move it to `archive/` on request, but do **not** delete
it — this skill is upload/move-only. If the user specifically wants Drive junk *gone*, give the
manual route (Drive web UI → search within the folder → Remove) since nothing here deletes.

## Notes

- **Answer keys never leave the trainer's hands**: not to the LMS, not to a learner-visible field,
  not to GitHub. Drive (trainer folder) only.
- Never archive or re-upload an unchanged file — touch only what changed.
- Never delete anything on Drive; superseded files are **moved to `archive/`**, never removed.
  A file that cannot be moved is reported as a WARNING and skipped, never deleted.
- Match target subfolders case-insensitively (reuse an existing `Archive`/`archives` rather
  than creating a duplicate; prefer the canonical lowercase `archive`).
- If rclone isn't configured and can't be set up in a non-interactive session, tell the user to
  run the one-time `rclone config` sign-in first; the MCP connector cannot substitute because
  it can't move or delete.
