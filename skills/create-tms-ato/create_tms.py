#!/usr/bin/env python3
"""
create_tms.py — Rebrand a Training Management System (TMS) .docx template to a
client's company details, swap the cover logo, fix footers + hyperlinks, scrub
document metadata, and optionally export a PDF.

Works directly on the OOXML so it survives text that Word split across multiple
runs (e.g. "info@chariot-learning" + ".com", or "6785" + " 0776"). Every
replacement is applied *run-aware*: the search string is matched against the
concatenated visible text of the part, then the underlying <w:t> runs are edited
in place (new text goes into the first run of the span, the rest are emptied),
so formatting/hyperlinks are preserved.

Branding rule (matches the house reference format):
  * The FULL legal name ("Acme Pte. Ltd.") is used ONLY where the template's
    legal name / footer copyright / cover "Prepared by" appears.
  * The SHORT brand ("Acme") is used for the standalone brand word that occurs
    mid-sentence throughout the body, so the prose never reads
    "... the Acme Pte. Ltd. website ...".
  The short brand defaults to --name with a trailing legal suffix
  (Pte. Ltd. / LLP / Ltd. / Limited ...) stripped; override with --brand.

Usage:
  python3 create_tms.py \
      --template "Sample_TMS.docx" --out "Skills Union_TMS.docx" \
      --logo "SU_Logo.png" \
      --name "Skills Union Pte. Ltd." --uen 202021281D \
      --address "109 North Bridge Road, #07-22, Funan, Singapore 179097" \
      --email janice@skillsunion.com --phone "8499 8618" \
      --whatsapp 6584998618 --website skillsunion.com \
      [--brand "Skills Union"] [--no-whatsapp] [--old old_values.json] [--pdf]

The default --old map matches the house "Chariot Learning & Consultancy LLP"
sample template. Supply --old <json> to rebrand a template with different
existing branding (see OLD_DEFAULT below for the shape).
"""
import argparse, html, json, os, re, shutil, subprocess, sys, tempfile, zipfile

# ---- default "old" branding present in the bundled Skills Union base template ----
# The base template (reference/TMS_base_template.docx) is the CLEANED Skills Union
# TMS: it already has ZERO Notion / Tertiary Infotech / tertiarycourses.com.sg
# references and no template-author name. Rebranding therefore only has to replace
# the Skills Union identity below.
OLD_DEFAULT = {
    "legal_names": ["Skills Union Pte. Ltd.", "Skills Union Pte Ltd",
                    "Skills Union Pte. Ltd", "Skills Union Pte Ltd."],
    # phrases where the brand word is glued to a stray product word — collapse to
    # the SHORT brand so no foreign word survives. Processed before the brand.
    "brand_aliases": ["Skills Union Infotech"],
    "brand": "Skills Union",            # standalone brand word (done last)
    "uen": "202021281D",
    "email": "janice@skillsunion.com",
    "address": "109 North Bridge Road, #07-22, Funan, Singapore 179097",
    "website_url": "https://skillsunion.com",
    "website": "skillsunion.com",
    "whatsapp_url": "https://wa.me/6584998618",
    "whatsapp": "6584998618",
    "phones": ["84998618", "8499 8618"],   # all number formats used in the doc
    # ---- foreign hyperlinks baked into the template ----
    # UNWRAP: kill the hyperlink, keep the visible text (private 3rd-party docs
    # the new client can't open: the base author's Drive/Docs files).
    "link_unwrap": ["notion.so", "drive.google.com", "docs.google.com"],
    # REPOINT: nothing to repoint — the base's only commercial site is
    # skillsunion.com, which the `website_url`/`website` swap already rewrites to
    # the client's own site. Legacy Tertiary domains kept as defence-in-depth.
    "link_repoint": ["tertiarycourses.com.sg"],
    # dead "click here" phrasing + Notion wording left behind once links are gone
    "text_fixes": [
        ["Here are the Notion links to streamline our SOP reminder email process and ensure effective communication:",
         "Our SOPs for the reminder email process are documented internally to ensure effective communication:"],
        ["Here’s the link to our SOP for Rescheduling a Class, providing clear and structured guidelines for this process:",
         "Our SOP for Rescheduling a Class provides clear and structured guidelines for this process:"],
        ["Here’s the link to our SOP for Canceling a Class, offering transparent and step-by-step instructions for handling cancellations:",
         "Our SOP for Canceling a Class provides transparent and step-by-step instructions for handling cancellations:"],
        ["Here’s the link to our SOP Refund Policy, providing clear and transparent guidelines for all refund requests:",
         "Our SOP Refund Policy provides clear and transparent guidelines for all refund requests:"],
        ["Reschedule a Class – Click here to access", "SOP: Reschedule a Class"],
        ["Cancel a Class – Click here to access", "SOP: Cancel a Class"],
        ["Refund Policy – [Click here to access]", "SOP: Refund Policy"],
        ["Trainer credentials are securely maintained in both Google Drive and Notion for",
         "Trainer credentials are securely maintained in both Google Drive and the internal document management system for"],
        ["Notion is utilized to manage", "The internal document management system is utilized to manage"],
        ["maintained in Notion", "maintained in the internal document management system"],
        ["Jyoti - WSQ Certificate", "Trainer WSQ Certificate"],
        [" – [Click here to access]", ""],
        [" – Click here to access", ""],
        ["[Click here to access]", ""],
        ["Click here to access", ""],
        # catch-all: any remaining bare "Notion" -> vendor-neutral system name
        ["Notion", "the internal document management system"],
    ],
}

# words/urls that must NEVER survive into a client document (template authoring
# artifacts / third-party names). Checked at the end across .xml AND .rels.
# NOTE: "Skills Union"/"skillsunion" is the BASE template's own branding — it must
# be fully replaced by the new client's, so it is forbidden in the OUTPUT.
# Author names (docProps creator/lastModifiedBy AND the Version Control Record
# "Who" column) are a real leak source — scrub them, don't just trust docProps.
FORBIDDEN = ["Chariot", "chariot", "Tertiary Infotech", "tertiaryinfotech",
             "tertiarycourses", "Skills Union", "skillsunion",
             "Notion", "notion.so", "Click here to access",
             "Victor Ang", "Jyoti"]

LEGAL_SUFFIXES = [" Pte. Ltd.", " Pte Ltd.", " Pte. Ltd", " Pte Ltd",
                  " Private Limited", " Limited", " LLP", " LLC",
                  " Ltd.", " Ltd", " Inc.", " Inc"]

WT = re.compile(r'(<w:t\b[^>]*>)(.*?)(</w:t>)', re.DOTALL)
WP = re.compile(r'<w:p\b[^>]*>.*?</w:p>', re.DOTALL)


def derive_brand(name):
    """Short brand = legal name minus a trailing legal suffix."""
    b = name.strip()
    for suf in LEGAL_SUFFIXES:
        if b.endswith(suf):
            return b[: -len(suf)].strip()
        if b.lower().endswith(suf.lower()):
            return b[: -len(suf)].strip()
    return b


def run_aware_replace(xml, old, new):
    """Replace every occurrence of `old` in the concatenated <w:t> text of `xml`,
    editing the underlying runs so matches that span runs are handled."""
    if not old:
        return xml
    ms = list(WT.finditer(xml))
    if not ms:
        return xml
    texts = [html.unescape(m.group(2)) for m in ms]
    starts, pos = [], 0
    for t in texts:
        starts.append(pos)
        pos += len(t)
    full = ''.join(texts)
    if old not in full:
        return xml
    spans, i = [], 0
    while True:
        j = full.find(old, i)
        if j < 0:
            break
        spans.append((j, j + len(old)))
        i = j + len(old)
    new_text = list(texts)
    for (s, e) in reversed(spans):
        def run_of(p):
            lo, hi = 0, len(starts) - 1
            while lo < hi:
                mid = (lo + hi + 1) // 2
                if starts[mid] <= p:
                    lo = mid
                else:
                    hi = mid - 1
            return lo
        ri, rj = run_of(s), run_of(e - 1)
        s_off = s - starts[ri]
        e_off = e - starts[rj]
        head = new_text[ri][:s_off]
        tail = new_text[rj][e_off:]
        if ri == rj:
            new_text[ri] = head + new + tail
        else:
            new_text[ri] = head + new
            for k in range(ri + 1, rj):
                new_text[k] = ''
            new_text[rj] = tail
    out = xml
    for idx in range(len(ms) - 1, -1, -1):
        m = ms[idx]
        if new_text[idx] == texts[idx]:
            continue
        repl = m.group(1) + html.escape(new_text[idx], quote=False) + m.group(3)
        out = out[:m.start()] + repl + out[m.end():]
    return out


def apply_map(xml, old, v, keep_first_legal=False):
    """Rebrand one part. Legal name -> full name; standalone brand + aliases ->
    short brand. If keep_first_legal, the FIRST legal-name occurrence in this
    part is left as the full name and only later ones are shortened — used for
    the cover page where 'Prepared by <Full Legal Name>' must stay full while
    the standalone brand elsewhere becomes short (the brand word and the legal
    name are different strings, so this flag is normally unnecessary)."""
    for ln in old.get("legal_names", []):
        xml = run_aware_replace(xml, ln, v["name"])
    xml = run_aware_replace(xml, old["uen"], v["uen"])
    xml = run_aware_replace(xml, old["email"], v["email"])
    xml = run_aware_replace(xml, old["address"], v["address"])
    xml = run_aware_replace(xml, old["website_url"], "https://" + v["website"])
    xml = run_aware_replace(xml, old["website"], v["website"])
    if v.get("whatsapp"):
        xml = run_aware_replace(xml, old["whatsapp_url"], "https://wa.me/" + v["whatsapp"])
        xml = run_aware_replace(xml, old["whatsapp"], v["whatsapp"])
    for ph in old.get("phones", []):
        xml = run_aware_replace(xml, ph, v["phone"])
    # brand-glued phrases first (so a stray foreign word is dropped), then the
    # standalone brand word — both map to the SHORT brand.
    for alias in old.get("brand_aliases", []):
        xml = run_aware_replace(xml, alias, v["brand"])
    if old.get("brand"):
        xml = run_aware_replace(xml, old["brand"], v["brand"])
    return xml


def strip_whatsapp(xml):
    """Drop WhatsApp contact lines entirely and remove inline '/ WhatsApp'
    channel mentions — for clients that don't offer WhatsApp support."""
    def keep(m):
        block = m.group(0)
        txt = html.unescape(''.join(re.findall(r'<w:t\b[^>]*>(.*?)</w:t>', block, re.DOTALL)))
        low = txt.strip().lower()
        if 'wa.me' in low or low.startswith('whatsapp:'):
            return ''
        return block
    xml = WP.sub(keep, xml)
    for frag in (' / WhatsApp', 'WhatsApp / ', ' WhatsApp'):
        xml = run_aware_replace(xml, frag, '')
    return xml


def strip_foreign_links(doc_xml, rels, old, v, drop_whatsapp=False):
    """Remove every hyperlink that points somewhere the client doesn't own.

    * link_repoint domains (the template author's own site) are re-targeted to
      the client's website, so the link still works and reads correctly.
    * link_unwrap domains (Notion SOPs, the template author's private Google
      Drive/Docs files) are unwrapped: the <w:hyperlink> wrapper is removed, the
      visible text is kept as plain de-styled runs, and the now-unused
      Relationship entry is deleted. Leaving the rel behind would keep a
      notion.so / drive.google.com URL inside the .docx.
    """
    unwrap_domains = list(old.get("link_unwrap", []))
    if drop_whatsapp:
        unwrap_domains.append("wa.me")
    repoint_domains = old.get("link_repoint", [])
    relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="(https?://[^"]+)"', rels))

    unwrap_ids, repoint_ids = set(), set()
    for rid, url in relmap.items():
        u = url.lower()
        if any(d in u for d in unwrap_domains):
            unwrap_ids.add(rid)
        elif any(d in u for d in repoint_domains):
            repoint_ids.add(rid)

    # 1) repoint the author's site -> client site
    client_site = "https://" + v["website"]
    def _repoint(m):
        rid, url = m.group(1), m.group(2)
        if rid in repoint_ids:
            return m.group(0).replace('Target="%s"' % url, 'Target="%s"' % client_site)
        return m.group(0)
    rels = re.sub(r'<Relationship Id="(rId\d+)"[^>]*Target="(https?://[^"]+)"[^>]*/>',
                  _repoint, rels)

    # 2) unwrap foreign hyperlinks (hyperlinks don't nest, so non-greedy is safe)
    def _destyle(inner):
        inner = re.sub(r'<w:rStyle\b[^>]*/>', '', inner)
        inner = re.sub(r'<w:u\b[^>]*/>', '', inner)
        inner = re.sub(r'<w:color\b[^>]*/>', '', inner)
        return inner
    for rid in unwrap_ids:
        pat = re.compile(r'<w:hyperlink\b[^>]*r:id="%s"[^>]*>(.*?)</w:hyperlink>' % re.escape(rid),
                         re.DOTALL)
        doc_xml = pat.sub(lambda m: _destyle(m.group(1)), doc_xml)

    # 3) drop the now-unused Relationship entries
    for rid in unwrap_ids:
        rels = re.sub(r'<Relationship Id="%s"[^>]*/>' % re.escape(rid), '', rels)

    # 4) tidy the wording the dead links left behind (Notion, "click here", ...)
    for old_t, new_t in old.get("text_fixes", []):
        doc_xml = run_aware_replace(doc_xml, old_t, new_t)

    return doc_xml, rels, len(unwrap_ids), len(repoint_ids)


def rels_replace(rels, old, v):
    rels = rels.replace("mailto:" + old["email"], "mailto:" + v["email"])
    rels = rels.replace(old["website_url"], "https://" + v["website"])
    if v.get("whatsapp"):
        rels = rels.replace(old["whatsapp_url"], "https://wa.me/" + v["whatsapp"])
    return rels


def scrub_metadata(work, name):
    """Replace authoring metadata (creator / lastModifiedBy / Company) so the
    template author's org name never travels with the client document."""
    core = os.path.join(work, "docProps", "core.xml")
    if os.path.exists(core):
        c = open(core, encoding="utf-8").read()
        c = re.sub(r'<dc:creator>.*?</dc:creator>', f'<dc:creator>{name}</dc:creator>', c, flags=re.DOTALL)
        c = re.sub(r'<cp:lastModifiedBy>.*?</cp:lastModifiedBy>',
                   f'<cp:lastModifiedBy>{name}</cp:lastModifiedBy>', c, flags=re.DOTALL)
        open(core, "w", encoding="utf-8").write(c)
    app = os.path.join(work, "docProps", "app.xml")
    if os.path.exists(app):
        a = open(app, encoding="utf-8").read()
        if '<Company>' in a:
            a = re.sub(r'<Company>.*?</Company>', f'<Company>{name}</Company>', a, flags=re.DOTALL)
        elif '</Properties>' in a:
            a = a.replace('</Properties>', f'<Company>{name}</Company></Properties>')
        a = re.sub(r'<Manager>.*?</Manager>', '', a, flags=re.DOTALL)
        open(app, "w", encoding="utf-8").write(a)


def swap_logo(media_dir, embeds, logo_path):
    """Overwrite the first cover image with the client's logo, fitted onto the
    same pixel canvas so the drawing extent stays valid (no distortion)."""
    from PIL import Image
    target = os.path.join(media_dir, embeds[0])
    canvas_w, canvas_h = Image.open(target).size
    logo = Image.open(logo_path).convert("RGBA")
    scale = min(canvas_w / logo.width, canvas_h / logo.height)
    nw, nh = max(1, int(logo.width * scale)), max(1, int(logo.height * scale))
    logo_r = logo.resize((nw, nh), Image.LANCZOS)
    canvas = Image.new("RGB", (canvas_w, canvas_h), "white")
    canvas.paste(logo_r, ((canvas_w - nw) // 2, (canvas_h - nh) // 2), logo_r)
    ext = os.path.splitext(target)[1].lower()
    if ext in (".jpg", ".jpeg"):
        canvas.save(target, "JPEG", quality=92)
    else:
        canvas.save(target)
    return embeds[0]


def first_cover_images(doc_xml, rels):
    relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="media/([^"]+)"', rels))
    seen = []
    for e in re.findall(r'r:embed="(rId\d+)"', doc_xml):
        f = relmap.get(e)
        if f and f not in seen:
            seen.append(f)
    return seen


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--logo")
    ap.add_argument("--name", required=True, help="full legal name, e.g. 'Acme Pte. Ltd.'")
    ap.add_argument("--brand", help="short brand for mid-sentence use; defaults to --name minus legal suffix")
    ap.add_argument("--uen", required=True)
    ap.add_argument("--address", required=True)
    ap.add_argument("--email", required=True)
    ap.add_argument("--phone", required=True)
    ap.add_argument("--whatsapp", default="")
    ap.add_argument("--no-whatsapp", dest="no_whatsapp", action="store_true",
                    help="remove all WhatsApp contact lines/mentions (client has no WhatsApp)")
    ap.add_argument("--website", required=True)
    ap.add_argument("--old", help="JSON file overriding the default old-branding map")
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()

    old = dict(OLD_DEFAULT)
    if a.old:
        old.update(json.load(open(a.old)))
    brand = a.brand or derive_brand(a.name)
    whatsapp = "" if a.no_whatsapp else a.whatsapp
    v = dict(name=a.name, brand=brand, uen=a.uen, address=a.address, email=a.email,
             phone=a.phone, whatsapp=whatsapp, website=a.website)

    work = tempfile.mkdtemp(prefix="tms_")
    with zipfile.ZipFile(a.template) as z:
        z.extractall(work)
    wdir = os.path.join(work, "word")

    doc_p = os.path.join(wdir, "document.xml")
    doc = open(doc_p, encoding="utf-8").read()
    rels_p = os.path.join(wdir, "_rels", "document.xml.rels")
    rels = open(rels_p, encoding="utf-8").read() if os.path.exists(rels_p) else ""

    if a.logo:
        embeds = first_cover_images(doc, rels)
        if embeds:
            swap_logo(os.path.join(wdir, "media"), embeds, a.logo)

    doc = apply_map(doc, old, v)
    if a.no_whatsapp:
        doc = strip_whatsapp(doc)

    if rels:
        rels = rels_replace(rels, old, v)
        doc, rels, n_unwrap, n_repoint = strip_foreign_links(
            doc, rels, old, v, drop_whatsapp=a.no_whatsapp)
        print(f"foreign links: {n_unwrap} unwrapped, {n_repoint} repointed to {v['website']}")
        open(rels_p, "w", encoding="utf-8").write(rels)

    open(doc_p, "w", encoding="utf-8").write(doc)

    for name in os.listdir(wdir):
        if re.match(r'(footer|header)\d+\.xml$', name):
            p = os.path.join(wdir, name)
            src = open(p, encoding="utf-8").read()
            open(p, "w", encoding="utf-8").write(apply_map(src, old, v))

    scrub_metadata(work, a.name)

    # residual check — scan .xml AND .rels (hyperlink targets live in .rels, so
    # an .xml-only scan silently misses notion.so / drive.google.com URLs)
    leftover = {}
    for root, _, files in os.walk(work):
        for f in files:
            if f.endswith((".xml", ".rels")):
                t = open(os.path.join(root, f), encoding="utf-8", errors="ignore").read()
                hits = sorted({w for w in FORBIDDEN if w in t})
                if hits:
                    leftover[os.path.relpath(os.path.join(root, f), work)] = hits

    # repackage
    out = os.path.abspath(a.out)
    if os.path.exists(out):
        os.remove(out)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        ct = os.path.join(work, "[Content_Types].xml")
        z.write(ct, "[Content_Types].xml")
        for root, _, files in os.walk(work):
            for f in files:
                fp = os.path.join(root, f)
                arc = os.path.relpath(fp, work)
                if arc == "[Content_Types].xml":
                    continue
                z.write(fp, arc)
    shutil.rmtree(work, ignore_errors=True)

    print(f"Wrote {out}   (brand='{brand}', legal='{a.name}')")
    if leftover:
        print("WARNING residual forbidden text:")
        for f, hits in sorted(leftover.items()):
            print(f"  {f}: {', '.join(hits)}")

    if a.pdf:
        soffice = next((c for c in ("soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice")
                        if shutil.which(c) or os.path.exists(c)), None)
        if not soffice:
            print("PDF skipped: LibreOffice (soffice) not found")
        else:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                            "--outdir", os.path.dirname(out), out],
                           check=True, capture_output=True)
            print("Wrote", os.path.splitext(out)[0] + ".pdf")


if __name__ == "__main__":
    main()
