---
name: create-tms-ato
description: Create a client-branded Training Management System (TMS) document from the bundled base TMS .docx template — replace ALL of the template's company details (name, UEN, registered address, email, phone, WhatsApp, website), swap the cover logo, rewrite the footer copyright line, update hyperlink targets (mailto/website/WhatsApp), scrub authoring metadata, and export a matching PDF. Uses the SHORT brand mid-sentence and the FULL legal name only in the cover/footer, and leaves ZERO trace of the template's original company (or of Tertiary Infotech). Use when the user asks to "create a TMS", "rebrand the TMS", "make a TMS for <client>", or produce a client's Training Management / Training Administration Quality Management System doc for an ATO / SSG submission.
---

# Create TMS (ATO) — rebrand the base TMS to a client

Takes the **bundled base TMS template** (`reference/TMS_base_template.docx`, branded
"Skills Union Pte. Ltd.") and produces `<Client>_TMS.docx` + `<Client>_TMS.pdf` fully
rebranded to the client, in the house reference format, with **zero** residual references
to the base template's company — and no "Tertiary Infotech" / "Notion" /
template-author name anywhere.

> **The bundled base is already clean.** It is the Skills Union TMS after a full scrub:
> **no Notion**, no `notion.so` / `tertiarycourses.com.sg` links, no "Tertiary Infotech",
> and no template-author name (the Version Control Record "Who" column and the docProps
> creator/lastModifiedBy are the company, not a person). Systems are described
> vendor-neutrally as **"the internal document management system"** — never name Notion (or
> any third-party tool the client doesn't use). Keep it that way: any future edit to the
> base must preserve these invariants.

## Format reference (do not diverge from this)

The bundled `reference/TMS_base_template.docx` **is** the canonical format: logo cover →
`Training Administration Quality Management System` title → `Prepared by / <Full Legal
Name> / Version 1.0`; a Version Control Record table; a small-caps numbered Table of
Contents; numbered heading hierarchy (`1`, `1.1`, `1.1a`, `(i)`, `(a)`); multi-level
bullets (•, ○, ▪); a landscape org-chart diagram; boxed process flowcharts; full Annex
set; and a centred italic footer `Copyright <year> <Full Legal Name>. All rights
reserved.` + page number. Always rebrand this template — never rebuild the TMS from
scratch (that loses the flowcharts, org chart, and annexes).

## Brand rule (important — this is what keeps prose clean)

* **Full legal name** (`Acme Pte. Ltd.`) is used ONLY for the template's legal name →
  cover "Prepared by" and every footer copyright line.
* **Short brand** (`Acme`) is used for the standalone brand word that appears
  mid-sentence throughout the body, so it never reads "... the Acme Pte. Ltd. website ...".
  The script derives the short brand from `--name` by stripping a trailing legal suffix
  (Pte. Ltd. / LLP / Ltd. / Limited …); override with `--brand` if needed.

## Foreign links (stripped automatically — do not skip this)

Links the client does not own are handled on every run, with no extra flags:

| Kind | Example | Action |
|---|---|---|
| Base author's Drive/Docs | `drive.google.com/file/…`, `docs.google.com/document/…` | **unwrapped** — hyperlink removed, visible text kept (private files the client can't open) |
| Base author's site | `skillsunion.com` | **repointed** to the client's own website (via the `website_url` swap) |
| Stale WhatsApp | `wa.me/…` | unwrapped when `--no-whatsapp` |
| Official / client links | `tpgateway.gov.sg`, `myskillsfuture.gov.sg`, client site | **kept** |
| Notion / `tertiarycourses.com.sg` | — | **already gone from the base**; the `link_unwrap` + `text_fixes` rules remain as defence-in-depth in case a dirtier template is passed via `--template` |

Unwrapping also deletes the now-unused `Relationship` entry — leaving it behind would keep a
`drive.google.com` URL inside the `.docx` even though no visible link remains.

**Never name a third-party tool the client doesn't use.** Any residual "Notion" wording is
rewritten by `text_fixes` to **"the internal document management system"** (vendor-neutral) —
*not* to "Google Drive", which would assert a system the client may not have.

> **Hyperlink targets live in `.rels`, not `.xml`.** Any verification that scans only `*.xml`
> will report "clean" while `notion.so` / `tertiarycourses.com.sg` URLs are still embedded.
> Always scan `.xml` **and** `.rels`.

> **Author names hide in two places.** `docProps/core.xml` (creator / lastModifiedBy) **and**
> the body — the Version Control Record "Who" column. Scrubbing only docProps leaves the
> person's name visible on page 2. Scan the body text too.

## Inputs required (ask if missing)

| Field | Source | Example |
|---|---|---|
| Company name (full legal) | ACRA Bizfile | `Skills Union Pte. Ltd.` |
| UEN | ACRA Bizfile | `202021281D` |
| Registered address | ACRA Bizfile | `109 North Bridge Road, #07-22, Funan, Singapore 179097` |
| Email | client (ask) | `janice@skillsunion.com` |
| Phone | client (ask) | `8499 8618` (no `+65` — the template already prints `+65` before it) |
| WhatsApp number | client (ask) | `6584998618` (digits+country code) — or `--no-whatsapp` if none |
| Website | client (ask) | `skillsunion.com` (no `https://`) |
| Logo (PNG/JPG) | client folder / website | `SU_Logo.png` |

ACRA gives name/UEN/address only — **always ask the user for email, phone, WhatsApp and
website**; never invent contact details on a compliance document. If the client has no
WhatsApp, pass `--no-whatsapp` (removes the WhatsApp support line + any "/ WhatsApp"
channel mention) — confirm with the user first.

## Run

```bash
python3 create_tms.py \
  --template "reference/TMS_base_template.docx" \
  --out "<Client>_TMS.docx" \
  --logo "<Client>_Logo.png" \
  --name "<Full Legal Name>" --uen "<UEN>" \
  --address "<Registered Address>" \
  --email "<email>" --phone "<phone>" \
  --website "<domain>" \
  --whatsapp "<wa digits>"   # OR: --no-whatsapp \
  --pdf
```

Pass the bundled `reference/TMS_base_template.docx` (full path from the skill base dir) as
`--template`. `--brand "<Short Brand>"` overrides the auto-derived short brand. `--pdf`
exports a PDF via LibreOffice (`soffice`).

## How it works (why it's safe on split runs)

Word often splits one visible string across several `<w:t>` runs
(`info@chariot-learning` + `.com`, or `6785` + ` 0776`). The script matches each search
string against the **concatenated** run text of every part, then edits the underlying runs
in place (new text into the first run of the span, later runs emptied) so formatting and
hyperlinks survive. It processes `document.xml`, every `footer*.xml`/`header*.xml`, rewrites
`mailto:`/website/`wa.me` targets in `document.xml.rels`, and **scrubs `docProps/core.xml` +
`app.xml`** (creator / lastModifiedBy / Company) to the client name. The cover logo (first
image in document order) is fitted onto the original image's exact pixel canvas on a white
background, so the drawing extent stays valid and nothing is stretched. `brand_aliases`
(e.g. `Chariot Infotech`) collapse to the short brand so no stray foreign word survives.

## Different base template?

The default old-branding map (`OLD_DEFAULT`) matches the bundled **Skills Union** base
template. For a template with different existing branding (e.g. the older Chariot sample),
pass `--old old.json` overriding any of: `legal_names` (list), `brand`, `brand_aliases`
(list), `uen`, `email`, `address`, `website_url`, `website`, `whatsapp_url`, `whatsapp`,
`phones` (list of every number format), `link_unwrap`, `link_repoint`, `text_fixes`.

## Verify before reporting done

Re-open the output and confirm: (1) client name/UEN/address/email/phone/website all present;
(2) footer copyright = full legal name; cover "Prepared by" = full legal name; (3) body uses
the SHORT brand mid-sentence; (4) cover logo is the client's; (5) if `--no-whatsapp`, no
WhatsApp line or mention remains; (6) **no** occurrence anywhere of the base brand/UEN/
address/phone/website ("Skills Union", `skillsunion.com`, `202021281D`, …), no "Tertiary
Infotech" / "tertiarycourses" / "Notion", no template-author name, and no other client's
name; (7) the Version Control Record "Who" column is the **company**, not a person;
(8) the ONLY remaining hyperlinks are the client's own site + official `tpgateway.gov.sg` /
`myskillsfuture.gov.sg`.

Scan **`.xml` and `.rels`** (see the warnings above). One-liner:

```bash
python3 - <<'EOF'
import zipfile, re
z = zipfile.ZipFile("<Client>_TMS.docx")
blob = ''.join(z.read(n).decode('utf-8','ignore')
               for n in z.namelist() if n.endswith(('.xml', '.rels')))
for bad in ["Skills Union","skillsunion","202021281D","Chariot","Tertiary",
            "tertiarycourses","notion","Notion","Victor Ang","Jyoti",
            "Click here to access","drive.google.com","docs.google.com"]:
    print(("LEAK " if bad in blob else "ok   ") + bad)
for n in z.namelist():
    if n.endswith('.rels'):
        for m in re.finditer(r'Target="(https?://[^"]+)"',
                             z.read(n).decode('utf-8','ignore')):
            print("link:", m.group(1))
EOF
```

The script itself prints `WARNING residual forbidden text:` if any forbidden word survives —
treat that as a failure and investigate. Then eyeball the cover page of the PDF. For an
accurate TOC in the PDF, update fields (LibreOffice macro `getDocumentIndexes().update()` +
`refresh()`) before exporting — a plain `--convert-to pdf` leaves the TOC unpopulated.
