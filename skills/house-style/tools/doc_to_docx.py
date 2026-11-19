#!/usr/bin/env python3
"""Turn a kit document into a Word file (.docx) that people can edit, keeping its look.

    python3 tools/doc_to_docx.py samples/docs/garden-report.html            # builds it, then writes dist/garden-report.docx
    python3 tools/doc_to_docx.py samples/docs/bakery-proposal.md            # a Markdown source goes through tools/md_to_doc.py first
    python3 tools/doc_to_docx.py dist/report.html --built -o report.docx    # a file tools/build.py already built
    python3 tools/doc_to_docx.py report.html --check                        # also renders the Word file with LibreOffice and compares it with the PDF
    python3 tools/doc_to_docx.py report.html --check --word                 # and with Microsoft Word, on a Mac

Chrome lays the document out in print media at the width of the page's text, and tools/doc_to_docx.js reads every
paragraph, run, list, table, box and picture with its computed style. The writer turns them into Word's own parts:
headings and body text as paragraph styles, lists as numbering, boxes and blocks side by side as tables, figures and
charts as 300 dpi pictures, notes as Word footnotes, the contents as a contents field with the PDF's page numbers, and
each @page rule (size, margins, running heads, page numbers) as a section. The theme's fonts go inside the file, so it
looks the same on a computer that lacks them. The cover is a full-page picture of its art behind its text, with its
last lines (author, date) in the first page's footer, so it stays on one page in Word, LibreOffice and Google Docs.

--check writes "<name> check/" beside the Word file: the LibreOffice PDF (and Word's, with --word), a side-by-side PDF
against the kit's own PDF, and a report of page counts and any words missing. Word asks once for access to
tools/word-check: click Grant Access in Word's window. --gdocs prints the Google Drive connector steps that show how
Google Docs draws the file (tools/drive_qc.py), and --gdocs-compare <Google's PDF> compares it with the kit's PDF.
"""
import argparse
import base64
import html
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
import uuid
import zipfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import unquote

TOOLS = Path(__file__).resolve().parent
KIT = TOOLS.parent
sys.path.insert(0, str(TOOLS))
import chrome  # noqa: E402
import fontinfo  # noqa: E402

KIT_MARK = 'House Style'
WALKER = TOOLS / 'doc_to_docx.js'
CHECK_DIR = TOOLS / 'word-check'

# ---- lengths: the walker measures in CSS px, 96 to the inch

UNITS = {'px': 1.0, 'in': 96.0, 'cm': 96 / 2.54, 'mm': 96 / 25.4, 'q': 96 / 101.6, 'pt': 96 / 72, 'pc': 16.0}
SIZES = {'a3': (297, 420), 'a4': (210, 297), 'a5': (148, 210), 'b4': (250, 353), 'b5': (176, 250),
         'letter': (215.9, 279.4), 'legal': (215.9, 355.6), 'ledger': (279.4, 431.8)}   # mm


def tw(px):
    """Twentieths of a point."""
    return int(round(px * 15))


def emu(px):
    return int(round(px * 9525))


def half_pt(px):
    return max(1, int(round(px * 1.5)))


def eighths(px):
    """A border's width in eighths of a point (Word takes 2 to 96)."""
    return max(2, min(96, int(round(px * 6))))


def border_space(px):
    """A paragraph border's distance from the text, in points (Word takes 0 to 31)."""
    return max(0, min(31, int(round(px * 0.75))))


def css_len(v, default=None):
    m = re.match(r'\s*(-?[\d.]+)\s*([a-z]*)\s*$', v or '', re.I)
    if not m:
        return default
    n, u = float(m.group(1)), (m.group(2) or 'px').lower()
    return n * UNITS[u] if u in UNITS else default


# ---- the page set-up, from the @page rules the walker read

class Geo:
    """One named page as Word needs it: size and margins in px, and its margin boxes (running heads, page numbers)."""

    def __init__(self, w, h, margins, boxes):
        self.w, self.h = w, h
        self.mt, self.mr, self.mb, self.ml = margins
        self.boxes = boxes

    @property
    def cw(self):
        return self.w - self.ml - self.mr

    def box(self, name):
        b = self.boxes.get(name)
        return b if b and (b.get('parts') or []) else None


def page_size(v, data):
    words = (v or '').strip().lower().split()
    w = h = orient = None
    lens = [css_len(x) for x in words if css_len(x) is not None]
    for x in words:
        if x in SIZES:
            w, h = (n * UNITS['mm'] for n in SIZES[x])
        elif x in ('landscape', 'portrait'):
            orient = x
    if len(lens) >= 2:
        w, h = lens[:2]
    elif len(lens) == 1:
        w = h = lens[0]
    if w is None:
        w, h = (n * UNITS['mm'] for n in SIZES.get((data.get('page') or 'a4').lower(), SIZES['a4']))
        orient = orient or ('landscape' if (data.get('orient') or '').lower() == 'landscape' else None)
    if (orient == 'landscape' and w < h) or (orient == 'portrait' and w > h):
        w, h = h, w
    return w, h


def page_margins(rule, base):
    out = list(base)
    short = (rule.get('margin') or '').split()
    if short:
        v = [css_len(x) for x in short]
        if all(x is not None for x in v):
            out = (v * 4 if len(v) == 1 else [v[0], v[1], v[0], v[1]] if len(v) == 2
                   else [v[0], v[1], v[2], v[1]] if len(v) == 3 else v[:4])
    for i, k in enumerate(('margin-top', 'margin-right', 'margin-bottom', 'margin-left')):
        if rule.get(k):
            out[i] = css_len(rule[k], out[i])
    return out


def geometries(pages, data):
    """{page name: Geo}. A named page takes the default page's size, margins and boxes unless it sets its own; a box
    with content: none is empty there."""
    base = pages.get('') or {}
    w, h = page_size(base.get('size'), data)
    m = page_margins(base, [0, 0, 0, 0])
    geo = {'': Geo(w, h, m, dict(base.get('boxes') or {}))}
    for name, rule in pages.items():
        if not name or ':' in name:
            continue
        nw, nh = page_size(rule['size'], data) if rule.get('size') else (w, h)
        boxes = dict(base.get('boxes') or {})
        boxes.update(rule.get('boxes') or {})
        geo[name] = Geo(nw, nh, page_margins(rule, m), boxes)
    return geo


# ---- text

CTRL = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]')
COMPLEX = re.compile('[඀-෿஀-௿]')     # Sinhala, Tamil


def esc(s):
    return html.escape(CTRL.sub('', s or ''), quote=True)


def lang_tag(v):
    v = (v or '').strip()
    low, base = v.lower(), v.split('-')[0].lower()
    if low in ('', 'en'):
        return 'en-GB'
    return {'si': 'si-LK', 'ta': 'ta-LK'}.get(low) or ({'si': 'si-LK', 'ta': 'ta-LK'}.get(base) if '-' not in v else v)


def complex_share(t):
    letters = sum(1 for c in t if not c.isspace())
    return len(COMPLEX.findall(t)) / max(1, letters)


def css_families(ff):
    return [f.strip().strip('"\'') for f in (ff or '').split(',') if f.strip()]


# ---- fonts: CSS matching against the kit's faces, Word's names for them, and the files Word embeds

GENERIC = {'serif': 'Times New Roman', 'sans-serif': 'Arial', 'system-ui': 'Arial', '-apple-system': 'Arial',
           'blinkmacsystemfont': 'Arial', 'ui-sans-serif': 'Arial', 'ui-serif': 'Times New Roman',
           'monospace': 'Courier New', 'ui-monospace': 'Courier New', 'cursive': 'Comic Sans MS', 'fantasy': 'Impact'}
CLASSES = {'Times New Roman': 'roman', 'Arial': 'swiss', 'Courier New': 'modern', 'Comic Sans MS': 'script',
           'Impact': 'swiss'}


def pick_face(faces, weight, italic):
    """The face Chrome draws for a weight and style (CSS Fonts 4, 5.2): style first, then the nearest weight."""
    pool = [f for f in faces if (f.get('style') in ('italic', 'oblique')) == bool(italic)] or faces
    ws = sorted({int(f.get('weight') or 400) for f in pool})
    lighter = sorted([x for x in ws if x < weight], reverse=True)
    if weight in ws:
        best = weight
    elif 400 <= weight <= 500:
        best = ([x for x in ws if weight < x <= 500] + lighter + [x for x in ws if x > 500])[0]
    elif weight < 400:
        best = (lighter + [x for x in ws if x > weight])[0]
    else:
        best = ([x for x in ws if x > weight] + lighter)[0]
    return next(f for f in pool if int(f.get('weight') or 400) == best)


def word_name(F):
    """(name, bold, italic) as Word sees a face: a family of up to four faces (regular, bold, italic, bold italic)."""
    sub = (F.subfamily or 'Regular').strip().lower()
    if sub in ('regular', 'bold', 'italic', 'bold italic'):
        return F.family, sub.startswith('bold'), 'italic' in sub
    it = 'italic' in sub or 'oblique' in sub
    return (re.sub(r'\s+(Italic|Oblique)$', '', F.full or '') if it else F.full) or F.family, False, it


def odttf(data, guid):
    """Word's obfuscated font: the first 32 bytes XORed with the font key's bytes in reverse (ECMA-376 1, 17.8.1)."""
    key = bytes.fromhex(guid.replace('-', ''))[::-1]
    return bytes(b ^ key[i % 16] for i, b in enumerate(data[:32])) + data[32:]


def slot_of(bold, italic):
    return 'boldItalic' if bold and italic else 'bold' if bold else 'italic' if italic else 'regular'


class Fonts:
    """The fonts a document uses, by Word name, with the kit's files to embed (fonts/, fonts.json)."""

    def __init__(self):
        db = json.loads((KIT / 'fonts' / 'fonts.json').read_text(encoding='utf-8'))
        self.db = {k.lower(): v for k, v in db.items()}
        self.cache, self.faces, self.ttf = {}, {}, {}
        self.used = {}           # Word name -> {slot: font file}
        self.info = {}           # Word name -> fontinfo.Face (None for a font the kit lacks)
        self.families = {}       # Word name -> the CSS families it came from
        self.restricted = set()

    def face(self, path):
        if path not in self.faces:
            self.faces[path] = fontinfo.Face(path)
        return self.faces[path]

    def resolve(self, family, weight=400, italic=False):
        """(Word name, bold, italic) for a CSS family, weight and style. Bold or italic that Chrome synthesises
        (no such face) is asked of Word too."""
        key = ((family or '').strip().lower(), int(weight or 400), bool(italic))
        if key in self.cache:
            return self.cache[key]
        fam, w, it = key
        entry = self.db.get(fam)
        faces = [f for f in (entry or {}).get('faces') or [] if (KIT / 'fonts' / f['file']).is_file()]
        if not faces:
            name = GENERIC.get(fam, (family or '').strip() or 'Arial')
            self.info.setdefault(name, None)
            out = (name, w >= 600, it)
        else:
            f = pick_face(faces, w, it)
            path = KIT / 'fonts' / f['file']
            F = self.face(path)
            name, rb, ri = word_name(F)
            self.add(name, rb, ri, path, F, fam)
            fw, fit = int(f.get('weight') or 400), f.get('style') in ('italic', 'oblique')
            out = (name, rb or (w >= 600 and fw < 600), ri or (it and not fit))
        self.cache[key] = out
        return out

    def add(self, name, rb, ri, path, F, fam):
        if self.info.get(name) is None:
            self.info[name] = F
        self.families.setdefault(name, set()).add(fam)
        if (F.fstype & 0x000E) == 0x0002 or F.fstype & 0x0200:    # the font's licence forbids embedding
            self.restricted.add(name)
            return
        self.used.setdefault(name, {}).setdefault(slot_of(rb, ri), path)

    def complete(self):
        """Fill each used family's other slots (bold, italic) from the same kit family, so edits keep the font."""
        for name, fams in self.families.items():
            if name in self.restricted:
                continue
            for fam in fams:
                for f in self.db[fam]['faces']:
                    path = KIT / 'fonts' / f['file']
                    if path.is_file():
                        n, rb, ri = word_name(self.face(path))
                        if n == name:
                            self.used.setdefault(name, {}).setdefault(slot_of(rb, ri), path)

    def has(self, path, ch):
        import fitz
        if path not in self.ttf:
            self.ttf[path] = fitz.Font(fontfile=str(path))
        return self.ttf[path].has_glyph(ord(ch)) > 0

    def complex_font(self, ff, text, weight, italic):
        """The Word font for a run's Sinhala or Tamil letters: the first family in its CSS list whose face has them."""
        ch = next((c for c in text if COMPLEX.match(c)), None)
        if not ch:
            return None
        for fam in css_families(ff):
            entry = self.db.get(fam.lower())
            faces = [f for f in (entry or {}).get('faces') or [] if (KIT / 'fonts' / f['file']).is_file()]
            if faces and self.has(KIT / 'fonts' / pick_face(faces, int(weight or 400), italic)['file'], ch):
                return self.resolve(fam, weight, italic)
        return None

    def table(self):
        """word/fontTable.xml, its relationships and the embedded files [(name, bytes)]."""
        self.complete()
        out, rels, files = [], [], []
        for name in sorted(self.info, key=str.lower):
            F, kids = self.info[name], []
            if F is not None:
                kids.append('<w:panose1 w:val="%s"/>' % F.panose.hex().upper())
            kids.append('<w:charset w:val="00"/>')
            if F is not None:
                kind = 'modern' if F.panose[3] == 9 else 'swiss' if F.panose[1] in (11, 12, 13) else \
                    'roman' if 2 <= F.panose[1] <= 10 else 'auto'
            else:
                kind = CLASSES.get(name, 'auto')
            kids.append('<w:family w:val="%s"/>' % kind)
            kids.append('<w:pitch w:val="%s"/>' % ('fixed' if kind == 'modern' else 'variable'))
            if F is not None:
                kids.append('<w:sig w:usb0="%08X" w:usb1="%08X" w:usb2="%08X" w:usb3="%08X" w:csb0="%08X" w:csb1="%08X"/>'
                            % (tuple(F.ranges) + tuple(F.codepages)))
            for slot, tag in (('regular', 'embedRegular'), ('bold', 'embedBold'), ('italic', 'embedItalic'),
                              ('boldItalic', 'embedBoldItalic')):
                path = self.used.get(name, {}).get(slot)
                if path:
                    guid = str(uuid.uuid4()).upper()
                    rid, fname = 'rId%d' % (len(rels) + 1), 'font%d.odttf' % (len(files) + 1)
                    files.append((fname, odttf(path.read_bytes(), guid)))
                    rels.append((rid, fname))
                    kids.append('<w:%s r:id="%s" w:fontKey="{%s}"/>' % (tag, rid, guid))
            out.append('<w:font w:name="%s">%s</w:font>' % (esc(name), ''.join(kids)))
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:fonts xmlns:w="%s" xmlns:r="%s">%s</w:fonts>'
               % (NS['w'], NS['r'], ''.join(out)))
        rel = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="%s">%s</Relationships>'
               % (PKG_REL, ''.join('<Relationship Id="%s" Type="%s" Target="fonts/%s"/>' % (r, REL['font'], f)
                                   for r, f in rels)))
        return xml, rel, files


NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'}
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
REL = {k: _R + k for k in ('styles', 'settings', 'fontTable', 'numbering', 'footnotes', 'header', 'footer', 'image',
                           'hyperlink', 'font', 'officeDocument', 'extended-properties')}
REL['core-properties'] = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'

# ---- Chrome: lay the document out at the text width and read it

WAIT = ('new Promise(function (ok) { var done = function () { ok(location.href === %s); }; var go = function () {'
        ' Promise.resolve(window.PCDeck && PCDeck.ready ? PCDeck.ready() : 0).then(function () {'
        ' return document.fonts && document.fonts.ready; }).then(function () { requestAnimationFrame(function () {'
        ' requestAnimationFrame(done); }); }, function (e) { (window.__pcErrs = window.__pcErrs || []).push(String(e)); done(); }); };'
        ' document.readyState === "complete" ? go() : addEventListener("load", go); })')
FRAMES = 'new Promise(function (ok) { requestAnimationFrame(function () { requestAnimationFrame(function () { ok(1); }); }); })'
HIDE_TEXT = ('<style>.pc-cover, .pc-cover * { -webkit-text-fill-color: transparent !important; text-shadow: none !important;'
             ' text-decoration-color: transparent !important; caret-color: transparent !important }</style>')
COVER_TEXT = '''(function () {
  var c = document.querySelector('.pc-cover'), out = [], n, r = document.createRange();
  if (!c) return '[]';
  var w = document.createTreeWalker(c, NodeFilter.SHOW_TEXT);
  while ((n = w.nextNode())) {
    if (!n.nodeValue.trim()) continue;
    r.selectNodeContents(n);
    [].forEach.call(r.getClientRects(), function (q) {
      if (q.width > 0 && q.height > 0) out.push([q.left + scrollX, q.top + scrollY, q.width, q.height]);
    });
  }
  return JSON.stringify(out);
})()'''
NEEDS_REFS = '!!document.querySelector(\'.pc-toc a[href^="#"], a.pc-ref[href^="#"]\')'


class Browser:
    """One headless Chrome with a DevTools session, opened as chrome._print_pdf_wait opens it, in print media."""

    def __enter__(self):
        self.prof = tempfile.mkdtemp(prefix='pc-chrome-')
        args = [chrome.need_chrome()] + chrome.FLAGS + ['--window-size=1920,1080', '--remote-debugging-port=0',
                                                        '--user-data-dir=' + self.prof, 'about:blank']
        kw = {'start_new_session': True} if os.name == 'posix' else {}
        self.proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, **kw)
        try:
            port_file, t0 = Path(self.prof) / 'DevToolsActivePort', time.monotonic()
            while not (port_file.exists() and len(port_file.read_text().split()) >= 2):
                if time.monotonic() - t0 > 30 or self.proc.poll() is not None:
                    sys.exit('Chrome did not open its DevTools port')
                time.sleep(0.1)
            port, path = port_file.read_text().split()[:2]
            self.cdp = chrome._DevTools(int(port), path)
            target = self.cdp.call('Target.createTarget', url='about:blank')['targetId']
            self.sid = self.cdp.call('Target.attachToTarget', targetId=target, flatten=True)['sessionId']
            self.call('Page.enable')
            self.call('Emulation.setEmulatedMedia', media='print')
        except BaseException:
            self.__exit__()
            raise
        return self

    def __exit__(self, *exc):
        chrome._stop(self.proc)
        shutil.rmtree(self.prof, ignore_errors=True)

    def call(self, method, **params):
        return self.cdp.call(method, session=self.sid, **params)

    def eval(self, expr, wait=False):
        r = self.call('Runtime.evaluate', expression=expr, awaitPromise=wait, returnByValue=True)
        if r.get('exceptionDetails'):
            d = r['exceptionDetails']
            raise RuntimeError('page script failed: %s' % ((d.get('exception') or {}).get('description') or d.get('text')))
        return (r.get('result') or {}).get('value')

    def open(self, path, width, height=1400, timeout=180):
        """Load a page at a viewport width and wait for the kit's runtime, the fonts and two frames."""
        self.call('Emulation.setDeviceMetricsOverride', width=int(round(width)), height=height, deviceScaleFactor=1,
                  mobile=False)
        uri = path.as_uri()
        self.call('Page.navigate', url=uri)
        t0 = time.monotonic()
        while True:
            try:
                if self.eval(WAIT % json.dumps(uri), wait=True):
                    return
            except RuntimeError:
                pass                                     # the page was still navigating
            if time.monotonic() - t0 > timeout:
                sys.exit('%s did not finish loading in Chrome' % path.name)
            time.sleep(0.2)

    def shot(self, x, y, w, h, jpeg=False, scale=300 / 96):
        """A region of the page, in document px, at 300 dpi."""
        self.eval('scrollTo(0, 0)')
        kw = {'format': 'jpeg', 'quality': 90} if jpeg else {'format': 'png'}
        r = self.call('Page.captureScreenshot', clip={'x': x, 'y': y, 'width': max(1.0, w), 'height': max(1.0, h),
                                                      'scale': scale}, captureBeyondViewport=True, fromSurface=True, **kw)
        return base64.b64decode(r['data'])


def encode(im):
    """An opaque picture as PNG, or as JPEG when the PNG passes 2.5 MB (a photograph): (bytes, ext)."""
    buf = io.BytesIO()
    im.save(buf, 'PNG', optimize=True)
    if buf.tell() > 2.5e6:
        buf = io.BytesIO()
        im.save(buf, 'JPEG', quality=90)
        return buf.getvalue(), 'jpeg'
    return buf.getvalue(), 'png'


def cover_pieces(art, size, free, bg):
    """Cut the cover's opaque art into the fewest rectangles that hold every part of it differing from the page
    colour, none of them touching `free` (the box of the cover's bottom lines, which sit in the first page's footer:
    Word, LibreOffice and Google Docs all draw the body's behind-text pictures over the footer). Clear pixels would
    keep the footer visible too, but LibreOffice's PDF draws them black under the mask and greys thin art. size and
    free are in document px. Returns ([(x, y, w, h, (bytes, ext))], the free box's own piece or None when the art
    there is only the page colour)."""
    from PIL import Image, ImageChops
    im = Image.open(io.BytesIO(art)).convert('RGB')
    sw, sh = im.size
    k = sw / size[0]
    ink = ImageChops.difference(im, Image.new('RGB', im.size, '#' + (bg or 'ffffff'))).convert('L')
    ink = ink.point(lambda v: 255 if v > 8 else 0)
    R = (max(0, int(free[0] * k)), max(0, int(free[1] * k)), min(sw, int(-(-free[2] * k // 1))),
         min(sh, int(-(-free[3] * k // 1))))

    def trim(a):
        if a[2] <= a[0] or a[3] <= a[1]:
            return None
        bb = ink.crop(a).getbbox()
        return (a[0] + bb[0], a[1] + bb[1], a[0] + bb[2], a[1] + bb[3]) if bb else None

    def meets(a, b):
        return a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]

    # Each shape (pixels touching within f px) keeps to one picture, so no join crosses thin lines; only a shape
    # under the free box is cut into the bands around it.
    f = 8
    gw, gh = -(-sw // f), -(-sh // f)
    grid = list(ink.resize((gw, gh), Image.BOX).getdata())
    seen, boxes = bytearray(gw * gh), []
    for s in range(gw * gh):
        if not grid[s] or seen[s]:
            continue
        seen[s], stack, bx = 1, [s], [s % gw, s // gw, s % gw, s // gw]
        while stack:
            c = stack.pop()
            cx, cy = c % gw, c // gw
            bx = [min(bx[0], cx), min(bx[1], cy), max(bx[2], cx), max(bx[3], cy)]
            for ny in (cy - 1, cy, cy + 1):
                for nx in (cx - 1, cx, cx + 1):
                    if 0 <= nx < gw and 0 <= ny < gh and grid[ny * gw + nx] and not seen[ny * gw + nx]:
                        seen[ny * gw + nx] = 1
                        stack.append(ny * gw + nx)
        a = trim((max(0, (bx[0] - 1) * f), max(0, (bx[1] - 1) * f), min(sw, (bx[2] + 2) * f), min(sh, (bx[3] + 2) * f)))
        if a and meets(a, R):
            boxes += [x for x in (trim((a[0], a[1], a[2], min(a[3], R[1]))), trim((a[0], max(a[1], R[3]), a[2], a[3])),
                                  trim((a[0], max(a[1], R[1]), min(a[2], R[0]), min(a[3], R[3]))),
                                  trim((max(a[0], R[2]), max(a[1], R[1]), a[2], min(a[3], R[3])))) if x]
        elif a:
            boxes.append(a)
    merged = True
    while merged:
        merged = False
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = boxes[i], boxes[j]
                u = (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))
                if not meets(u, R):
                    boxes[i], merged = u, True
                    del boxes[j]
                    break
            if merged:
                break
    out = []
    for a in boxes:
        a = list(a)
        for side, step in ((0, -3), (1, -3), (2, 3), (3, 3)):   # overlap neighbours so no seam shows
            b = list(a)
            b[side] = min(max(0, b[side] + step), sw if side % 2 == 0 else sh)
            if not meets(b, R):
                a = b
        out.append((a[0] / k, a[1] / k, (a[2] - a[0]) / k, (a[3] - a[1]) / k, encode(im.crop(tuple(a)))))
    own = None
    if ink.crop(R).getbbox():
        own = (R[0] / k, R[1] / k, (R[2] - R[0]) / k, (R[3] - R[1]) / k, encode(im.crop(R)))
    return out, own


def layout(built, refs_of):
    """Read a built document in Chrome. The first pass reads the @page rules; the second lays the document out at the
    width of the page's text, with the PDF's page numbers in window.PC_PAGES, and walks it. Then the pictures, and
    last the cover's art with its text hidden. Returns (walk, {page name: Geo}, {picture id or 'cover': (bytes, ext)})."""
    js = WALKER.read_text(encoding='utf-8')
    text = built.read_text(encoding='utf-8')
    shots = {}
    with Browser() as b:
        t1 = chrome.temp_beside(built, chrome.inject(text, first=chrome.ERRORS, body=chrome.PRINT_READY), 'docx-setup')
        try:
            b.open(t1, 1200)
            first = json.loads(b.eval(js + ';JSON.stringify({set: PC_DOCX("pages"), refs: %s})' % NEEDS_REFS))
        finally:
            t1.unlink(missing_ok=True)
        geo = geometries(first['set']['pages'], first['set']['data'])
        g = geo['']
        refs = refs_of() if first['refs'] else {}
        head = ('<script>window.PC_PAGES = %s;</script><style>@media print { body { width: %.3fpx !important;'
                ' max-width: none !important; min-width: 0 !important; margin: 0 0 0 %.3fpx !important;'
                ' padding: 0 !important } }</style>' % (json.dumps(refs), g.cw, g.ml))
        t2 = chrome.temp_beside(built, chrome.inject(text, first=chrome.ERRORS, head=head, body=chrome.PRINT_READY), 'docx')
        try:
            b.open(t2, g.ml + max(g.w, max(x.w for x in geo.values())) + 40)
            data = json.loads(b.eval(js + ';JSON.stringify(PC_DOCX())'))
            data['errors'] = b.eval('window.__pcErrs || []') or []
            for p in data['pics']:
                if p['w'] >= 1 and p['h'] >= 1:
                    shots[p['id']] = (b.shot(p['x'], p['y'], p['w'], p['h'], jpeg=p.get('jpeg')),
                                      'jpeg' if p.get('jpeg') else 'png')
            cover = next((x for x in data['blocks'] if x.get('k') == 'cover'), None)
            if cover:
                # Where the cover's text sits, then its art with the text hidden; Writer.cover() cuts the art
                # around the bottom lines.
                data['cover_text'] = json.loads(b.eval(COVER_TEXT))
                b.eval('document.head.insertAdjacentHTML("beforeend", %s)' % json.dumps(HIDE_TEXT))
                b.eval(FRAMES, wait=True)
                shots['cover'] = (b.shot(cover['x'], cover['y'], cover['w'], cover['h']), 'png')
        finally:
            t2.unlink(missing_ok=True)
    data['refs'] = refs
    return data, geo, shots

# ---- the Word document: paragraphs and runs

PPR_ORDER = ('pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'numPr', 'pBdr', 'shd', 'tabs', 'spacing', 'ind', 'jc',
             'outlineLvl', 'rPr', 'sectPr')
BORDER = {'solid': 'single', 'dashed': 'dashed', 'dotted': 'dotted', 'double': 'double', 'groove': 'single',
          'ridge': 'single', 'inset': 'single', 'outset': 'single'}
BULLETS = {'disc': '•', 'circle': '◦', 'square': '▪', 'disclosure-closed': '▸', 'disclosure-open': '▾'}
FORMATS = {'decimal': 'decimal', 'decimal-leading-zero': 'decimalZero', 'lower-alpha': 'lowerLetter',
           'lower-latin': 'lowerLetter', 'upper-alpha': 'upperLetter', 'upper-latin': 'upperLetter',
           'lower-roman': 'lowerRoman', 'upper-roman': 'upperRoman'}
NOTE_MARK = re.compile(r'\s*[\d*†‡§¶a-z]{1,4}\s*$')
PAGE_NO = re.compile(r'\s*[\divxlcIVXLC]+\s*$')
TINY = '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'


class P:
    """A paragraph, its properties kept by name so they serialise in Word's order and a section can end on it."""

    def __init__(self, body='', **ppr):
        self.ppr = {k: v for k, v in ppr.items() if v}
        self.body = body

    def xml(self):
        pp = ''.join(self.ppr[k] for k in PPR_ORDER if self.ppr.get(k))
        return '<w:p>%s%s</w:p>' % ('<w:pPr>%s</w:pPr>' % pp if pp else '', self.body)


class T:
    def __init__(self, xml):
        self.text = xml

    def xml(self):
        return self.text


def spacing(before=0, line=None, rule='auto', after=0):
    line = 240 if line is None else line
    return '<w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="%s"/>' % (max(0, before), max(0, after), line, rule)


def ind(left=0, right=0, first=0):
    s = '<w:ind w:left="%d" w:right="%d"' % (max(0, left), max(0, right))
    if first > 0:
        s += ' w:firstLine="%d"' % first
    elif first < 0:
        s += ' w:hanging="%d"' % -first
    return s + '/>'


def jc(al):
    al = (al or '').lower()
    if 'center' in al:
        return '<w:jc w:val="center"/>'
    if al in ('right', 'end', '-webkit-right'):
        return '<w:jc w:val="right"/>'
    if al == 'justify':
        return '<w:jc w:val="both"/>'
    return ''


def border(tag, b, fill=None):
    """One side of a border, from the walker's {w, st, c}; with no border but a fill, a hairline in the fill colour so
    Word's shading reaches the padding."""
    if b:
        return '<w:%s w:val="%s" w:sz="%d" w:space="0" w:color="%s"/>' % (tag, BORDER.get(b['st'], 'single'), eighths(b['w']),
                                                                         (b.get('c') or '000000').upper())
    if fill:
        return '<w:%s w:val="single" w:sz="2" w:space="0" w:color="%s"/>' % (tag, fill.upper())
    return '<w:%s w:val="nil"/>' % tag


def shd(fill):
    return '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % fill.upper() if fill else ''


def rpr_xml(d, base, style=None):
    """A run's properties: only what differs from its paragraph style (base), switched off where the style has it on."""
    out = ['<w:rStyle w:val="%s"/>' % style] if style else []
    if d.get('font') != base.get('font') or d.get('cs') != base.get('cs') or d.get('hint'):
        cs = d.get('cs') or d.get('font')
        out.append('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"%s/>' % (
            esc(d['font']), esc(d['font']), esc(d['font']), esc(cs), ' w:hint="cs"' if d.get('hint') else ''))
    for k in ('b', 'i'):
        if bool(d.get(k)) != bool(base.get(k)):
            v = '' if d.get(k) else ' w:val="0"'
            out.append('<w:%s%s/><w:%sCs%s/>' % (k, v, k, v))
    for k in ('caps', 'smallCaps', 'strike'):
        if bool(d.get(k)) != bool(base.get(k)):
            out.append('<w:%s%s/>' % (k, '' if d.get(k) else ' w:val="0"'))
    if d.get('color') and d.get('color') != base.get('color'):
        out.append('<w:color w:val="%s"/>' % d['color'])
    if d.get('ls', 0) != base.get('ls', 0):
        out.append('<w:spacing w:val="%d"/>' % d.get('ls', 0))
    if d.get('sz') and d.get('sz') != base.get('sz'):
        out.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (d['sz'], d['sz']))
    if bool(d.get('u')) != bool(base.get('u')):
        out.append('<w:u w:val="%s"/>' % ('single' if d.get('u') else 'none'))
    if d.get('shd') and d.get('shd') != base.get('shd'):
        out.append(shd(d['shd']))
    if d.get('v') != base.get('v'):
        out.append('<w:vertAlign w:val="%s"/>' % (d.get('v') or 'baseline'))
    if (d.get('lang') and d.get('lang') != base.get('lang')) or d.get('bidi'):
        out.append('<w:lang w:val="%s"%s/>' % (d.get('lang') or base.get('lang') or 'en-GB',
                                              ' w:bidi="%s"' % d['bidi'] if d.get('bidi') else ''))
    return '<w:rPr>%s</w:rPr>' % ''.join(out) if out else ''


CS_LO, CS_HI = 0x0590, 0x1CFF       # Hebrew to the Indic, Sinhala and South-east Asian scripts


def cs_parts(t):
    """A run's text cut where a complex script starts or stops; spaces, punctuation and joiners stay in the piece
    they sit in: [(kind, text)], kind 'cs' or 'lat'."""
    out = []
    for c in t:
        k = 'cs' if CS_LO <= ord(c) <= CS_HI else 'lat' if c.isalnum() else None
        if out and (k is None or k == out[-1][0]):
            out[-1][1] += c
        else:
            out.append([k or 'lat', c])
    return [tuple(x) for x in out]


def text_xml(t):
    """A run's text; tabs and line ends (preformatted text) become Word's own."""
    out = []
    for i, line in enumerate(t.split('\n')):
        if i:
            out.append('<w:br/>')
        for j, part in enumerate(line.split('\t')):
            if j:
                out.append('<w:tab/>')
            if part:
                out.append('<w:t xml:space="preserve">%s</w:t>' % esc(part))
    return ''.join(out)


def is_pre(t):
    return '\n' in t or '\t' in t or '  ' in t


def is_fill(r):
    return bool(r.get('u')) and r.get('t') is not None and not r['t'].strip()


def tidy(runs, toc=False):
    """Collapse white space across runs as Chrome does, drop note numbers that a Word footnote mark replaces, and drop
    dot leaders from contents entries (the tab leader draws them)."""
    out = []
    for i, r in enumerate(runs):
        nxt = runs[i + 1] if i + 1 < len(runs) else {}
        if 'fn' in nxt and r.get('t') is not None and (r.get('fnref') or r.get('v') == 'super') and NOTE_MARK.match(r['t']):
            continue
        if toc and is_fill(r):
            continue
        out.append(dict(r))
    clean, prev_space = [], True                   # True at the start: leading spaces go
    for r in out:
        if r.get('br'):
            if clean and clean[-1].get('t') is not None and not is_pre(clean[-1]['t']) and not is_fill(clean[-1]):
                clean[-1]['t'] = clean[-1]['t'].rstrip(' ')
            clean.append(r)
            prev_space = True
            continue
        if r.get('t') is None:
            clean.append(r)
            prev_space = False
            continue
        t = r['t']
        if not is_pre(t) and not is_fill(r):
            if prev_space:
                t = t.lstrip(' ')
            if not t:
                continue
            prev_space = t.endswith(' ')
        else:
            prev_space = t.endswith(' ') or t.endswith('\n') or r.get('tab_at') is not None
        r['t'] = t
        clean.append(r)
    while clean and clean[-1].get('t') is not None and not is_fill(clean[-1]) and not is_pre(clean[-1]['t']):
        clean[-1]['t'] = clean[-1]['t'].rstrip(' ')
        if clean[-1]['t']:
            break
        clean.pop()
    while clean and clean[-1].get('br'):
        clean.pop()
    return clean

# ---- the Word document: the writer

STYLE_KEYS = ('font', 'b', 'i', 'smallCaps', 'color', 'sz', 'ls')
NAMES = {'Normal': 'Normal', 'Title': 'Title', 'TOCHeading': 'TOC Heading', 'FootnoteText': 'footnote text',
         'Header': 'header', 'Footer': 'footer'}
SINHALA = re.compile('[඀-෿]')
WEB = re.compile(r'(https?|mailto|tel|ftp):', re.I)
MARKS = '•·●▪■–-—◦○▸›»✓✔'
XMLNS = ' '.join('xmlns:%s="%s"' % kv for kv in NS.items())
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
FIELD_BEGIN = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "%s" \\h \\z'
               ' \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>')
FIELD_END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
DECORATIVE = ('<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative'
              ' xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>')
CELL_MAR0 = ('<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
             '<w:right w:w="0" w:type="dxa"/></w:tblCellMar>')
BREAK = '<w:pageBreakBefore/>'


def tiny(before=0, **ppr):
    """An empty paragraph a point high, with space before it in twips: after a table, in an empty cell, or where a
    section ends."""
    return P('', spacing=spacing(before, 20, 'exact'), rPr=TINY, **ppr)


def pstyle(sid):
    return '<w:pStyle w:val="%s"/>' % sid if sid and sid != 'Normal' else ''


def box_of(g, name):
    return g.box(name) or g.box('@' + name)


def bw(b):
    return b['w'] if b else 0


def top_of(b):
    if b.get('k') == 'p':
        return b['pf']['top']
    if b.get('k') == 'toc' and b.get('entries'):
        return b['entries'][0]['pf']['top']
    return b.get('top') or 0


def bottom_of(b):
    return b['pf']['bottom'] if b.get('k') == 'p' else (b.get('bottom') or 0)


def margin_of(b):
    return (b['pf'].get('mt') or 0) if b.get('k') == 'p' else 0


def break_of(b):
    return bool(b['pf'].get('pb')) if b.get('k') == 'p' else bool(b.get('pb'))


def line_rule(pf, sizes, pic=False, normal=0, k=1.2, nat=1.4):
    """Word's line spacing. Google Docs keeps at-least spacing, but turns exact X pt into a multiple of X / 12 pt
    whatever the paragraph's size or Normal's (tested 18 Nov 2026). So a line at least as tall as every app's own
    single line (nat times the largest run: Windows Word's win metrics, the hhea metrics of Word for Mac and
    LibreOffice, Google's 1.2 for a font it lacks) is written as at least the CSS line height, which every app keeps.
    A tighter paragraph whose text is larger or smaller than the body text (normal, px) takes a multiple of its
    largest run's single line (k times its size, hhea), which Word for Mac, LibreOffice and Google Docs keep (Windows
    Word draws it taller); body text keeps the CSS line height exactly, or at least it where Exactly would clip a
    larger run (Word clips glyph tops), tight lines or a picture."""
    lh = pf.get('lh') or 0
    if not lh:
        return 240, 'auto'
    big = max(sizes) if sizes else (pf.get('sz') or 0)
    if pic or (big and lh >= nat * big - 0.5):
        return tw(lh), 'atLeast'
    if normal and big and abs(big - normal) > 0.3:
        return max(120, int(round(240 * lh / (k * big)))), 'auto'
    if big > 1.15 * (pf.get('sz') or big) or lh < 1.25 * big:
        return tw(lh), 'atLeast'
    return tw(lh), 'exact'


def mark_rpr(sizes, base):
    """The paragraph mark at the size of the largest run. Google Docs has no exact or at-least line spacing: it turns
    X pt into a multiple of 1.2 times the mark's size, so a mark smaller than the text gives lines too tall."""
    if not sizes or half_pt(max(sizes)) == (base.get('sz') or 0):
        return ''
    return '<w:rPr><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (half_pt(max(sizes)), half_pt(max(sizes)))


def spaced(runs):
    """Keep a space between a marker drawn beside its text (a pre run first or last) and the text."""
    out = [dict(r) for r in runs]
    texts = [r for r in out if r.get('t') is not None]
    if len(texts) < 2:
        return out
    a, z = texts[0], texts[-1]
    if out[0] is a and a.get('pre') and not a['t'].endswith((' ', '\t')) and not texts[1]['t'].startswith(' '):
        a['t'] += ' '
    if z.get('pre') and not z['t'].startswith(' ') and not texts[-2]['t'].endswith(' '):
        z['t'] = ' ' + z['t']
    return out


def pbdr(pf):
    """A box round one paragraph: its borders at the padding's distance (a fill alone gets hairlines in its own
    colour, so the shading reaches the padding), and the fill."""
    bd, pad, fill = pf.get('bd') or {}, pf.get('pad') or [0, 0, 0, 0], pf.get('shade')
    out = []
    for tag, s, p in (('top', 't', pad[0]), ('left', 'l', pad[3]), ('bottom', 'b', pad[2]), ('right', 'r', pad[1])):
        if bd.get(s) or fill:
            out.append(border(tag, bd.get(s), fill).replace('w:space="0"', 'w:space="%d"' % border_space(p)))
    return ('<w:pBdr>%s</w:pBdr>' % ''.join(out) if out else ''), shd(fill)


class Writer:
    """The walker's blocks as Word's document body, with the styles, lists, notes, pictures, links, headers and
    footers they need."""

    def __init__(self, data, geo, shots, fonts):
        self.data, self.geo, self.shots, self.fonts = data, geo, shots, fonts
        self.lang0 = lang_tag(data.get('lang'))
        self.pics_in = {p['id']: p for p in data.get('pics') or []}
        self.body, self.final = [], ''
        self.rels, self.relmap = {}, {}              # part -> [(id, type, target, external)]
        self.media, self.pics, self.parts = [], {}, {}
        self.notes, self.abstract, self.numdefs, self.nums = [], [], [], {}
        self.bm, self.tags, self.names, self.placed = {}, {}, set(), set()
        self.ids = self.bmid = 0
        self.sd = {}                                  # paragraph style id -> the look of its runs
        self.pb, self.prev_page = False, None
        self.sec = {'page': '', 'type': 'nextPage', 'cols': None, 'first': None, 'start': 0}

    # -- relationships and pictures

    def rel(self, kind, target, part='document', external=False):
        key = (part, kind, target)
        if key not in self.relmap:
            lst = self.rels.setdefault(part, [])
            rid = 'rId%d' % (len(lst) + 1)
            lst.append((rid, REL[kind], target, external))
            self.relmap[key] = rid
        return self.relmap[key]

    def picture(self, key, part='document'):
        """The relationship id, in a part, of a captured picture (stored once in word/media), or None."""
        if key is None or key not in self.shots:
            return None
        if key not in self.pics:
            data, ext = self.shots[key]
            name = 'image%d.%s' % (len(self.media) + 1, 'jpeg' if ext == 'jpeg' else 'png')
            self.media.append((name, data))
            self.pics[key] = 'media/' + name
        return self.rel('image', self.pics[key], part)

    def drawing(self, rid, w, h, alt='', behind=False, x=0, y=0):
        """A picture in a run: inline, or behind the text x, y px from the page's top left corner (the cover's art)."""
        self.ids += 1
        n, cx, cy = self.ids, emu(w), emu(h)
        doc = ('<wp:docPr id="%d" name="Picture %d" descr="%s">%s</wp:docPr>' % (n, n, esc(alt), DECORATIVE) if behind
               else '<wp:docPr id="%d" name="Picture %d" descr="%s"/>' % (n, n, esc(alt)))
        lock = '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        graphic = ('<a:graphic><a:graphicData uri="%s"><pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="Picture %d"'
                   ' descr="%s"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/>'
                   '</a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
                   '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
                   % (NS['pic'], n, n, esc(alt), rid, cx, cy))
        if behind:
            body = ('<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="%d" behindDoc="1"'
                    ' locked="1" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
                    '<wp:positionH relativeFrom="page"><wp:posOffset>%d</wp:posOffset></wp:positionH>'
                    '<wp:positionV relativeFrom="page"><wp:posOffset>%d</wp:posOffset></wp:positionV>'
                    '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>%s%s%s'
                    '</wp:anchor>' % (n, emu(x), emu(y), cx, cy, doc, lock, graphic))
        else:
            body = ('<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="%d" cy="%d"/>'
                    '<wp:effectExtent l="0" t="0" r="0" b="0"/>%s%s%s</wp:inline>' % (cx, cy, doc, lock, graphic))
        return '<w:r><w:drawing>%s</w:drawing></w:r>' % body

    def inline_pic(self, pic, avail):
        rid = self.picture(pic.get('id'))
        if not rid:
            return ''
        w, h = pic.get('w') or 1, pic.get('h') or 1
        if avail and w > avail:
            w, h = avail, h * avail / w
        return self.drawing(rid, w, h, (self.pics_in.get(pic.get('id')) or {}).get('alt', ''))

    # -- bookmarks and styles

    def anchor(self, x, tag=''):
        """A bookmark name for an element id: letters, digits and _, at most 40 characters, unique in any case."""
        if x in self.bm:
            return self.bm[x]
        name = re.sub(r'[^A-Za-z0-9_]', '_', x)
        if not name[:1].isalpha():
            name = 'b_' + name
        name = base = name[:40]
        k = 2
        while name.lower() in self.names:
            name = '%s_%d' % (base[:36], k)
            k += 1
        self.names.add(name.lower())
        self.bm[x], self.tags[x] = name, tag
        return name

    def prescan(self):
        """Anchor every element id, then give each paragraph style the look of the runs that hold most of its text."""
        counts, reps = {}, {}

        def tally(sid, runs):
            c = counts.setdefault(sid, Counter())
            for r in runs:
                t = r.get('t')
                if t is None or r.get('v') or r.get('pre') or not t.strip():
                    continue
                key = (r.get('f'), r.get('w') or 400, bool(r.get('i')), r.get('sz'), r.get('c'), bool(r.get('sc')),
                       r.get('ls') or 0)
                c[key] += len(t.strip())
                reps.setdefault(key, r)

        def visit(blocks, where):
            for b in blocks:
                k = b.get('k')
                if k == 'p':
                    tag = b.get('tag') or ''
                    for x in [b.get('id')] + list(b.get('ids') or []):
                        if x:
                            self.anchor(x, tag)
                    if where == 'cover':
                        sid = 'Title' if tag == 'h1' else ''
                    elif b.get('toc_title'):
                        sid = 'TOCHeading'
                    elif re.match(r'h[1-6]$', tag):
                        sid = 'Heading' + tag[1]
                    else:
                        sid = 'Normal'
                    if sid:
                        tally(sid, b.get('runs') or [])
                    for r in b.get('runs') or []:
                        if 'fn' in r:
                            tally('FootnoteText', r['fn'])
                elif k == 'table':
                    for row in b.get('rows') or []:
                        for c in row.get('cells') or []:
                            visit(c.get('blocks') or [], where)
                elif k == 'cols':
                    visit(b.get('blocks') or [], where)
                elif k == 'cover':
                    visit(b.get('blocks') or [], 'cover')
                elif k == 'toc':
                    for e in b.get('entries') or []:
                        tally('TOC%d' % min(9, e.get('lvl') or 1), tidy(e.get('runs') or [], toc=True))

        visit(self.data.get('blocks') or [], 'body')

        def top(c):
            return reps[c.most_common(1)[0][0]] if c else None
        every = Counter()
        for c in counts.values():
            every.update(c)
        rep = top(counts.get('Normal')) or top(every)
        self.sd['Normal'] = self.style_d(rep) if rep else {'font': self.fonts.resolve('Arial')[0], 'sz': 21,
                                                            'color': '000000'}
        for sid, c in counts.items():
            if sid != 'Normal' and c:
                self.sd[sid] = self.style_d(top(c))
        self.sd['Header'] = dict(self.sd['Normal'])
        self.sd['Footer'] = dict(self.sd['Normal'])

    def style_d(self, r):
        d = self.fmt(r)
        return {k: d[k] for k in STYLE_KEYS if d.get(k)}

    def style_of(self, p, cover=False):
        tag = p.get('tag') or ''
        if cover:
            return 'Title' if tag == 'h1' and 'Title' in self.sd else 'Normal'
        if p.get('toc_title') and 'TOCHeading' in self.sd:
            return 'TOCHeading'
        if re.match(r'h[1-6]$', tag) and 'Heading' + tag[1] in self.sd:
            return 'Heading' + tag[1]
        return 'Normal'

    # -- runs

    def fmt(self, r):
        """Word's properties for a walker run, as rpr_xml takes them. CSS capitals are written as capital letters,
        since Google Docs drops Word's all caps."""
        name, b, i = self.fonts.resolve(r.get('f'), r.get('w') or 400, bool(r.get('i')))
        d = {'font': name, 'b': b, 'i': i, 'sz': half_pt(r.get('sz') or 14), 'color': (r.get('c') or '000000').upper()}
        if r.get('sc'):
            d['smallCaps'] = True
        if r.get('s'):
            d['strike'] = True
        if r.get('u'):
            d['u'] = True
        if r.get('ls'):
            d['ls'] = tw(r['ls'])
        if r.get('v') in ('super', 'sub'):
            d['v'] = 'superscript' if r['v'] == 'super' else 'subscript'
        if r.get('bg'):
            d['shd'] = r['bg'].upper()
        lang = lang_tag(r.get('lang')) if r.get('lang') else None
        if lang and lang != self.lang0:
            d['lang'] = lang
        t = r.get('t') or ''
        if COMPLEX.search(t):
            cf = self.fonts.complex_font(r.get('ff') or r.get('f'), t, r.get('w') or 400, bool(r.get('i')))
            if cf:
                d['cs'] = cf[0]
            d['bidi'] = 'si-LK' if SINHALA.search(t) else 'ta-LK'
            if complex_share(t) > 0.5:
                d['hint'] = True
        return d

    def link(self, href, part):
        """The w:hyperlink tag a linked run opens with, or None where the link cannot go (no such bookmark, or a
        part without relationships for it)."""
        if not href or part is None:
            return None
        if href.startswith('#'):
            x = unquote(href[1:])
            return '<w:hyperlink w:anchor="%s" w:history="1">' % self.bm[x] if x in self.bm else None
        if WEB.match(href) and part in ('document', 'footnotes'):
            return '<w:hyperlink r:id="%s" w:history="1">' % self.rel('hyperlink', href, part, external=True)
        return None

    def runs_xml(self, runs, base, avail, part='document', links=True):
        """A paragraph's runs as Word XML: (xml, text sizes in px, whether it holds a picture). Notes and pictures go
        only in the document's body."""
        seq, sizes, pic = [], [], False
        for r in runs:
            if r.get('br'):
                seq.append(('br', None, None))
            elif 'fn' in r:
                if part == 'document':
                    seq.append(('fn', r, None))
            elif 'pic' in r:
                if part == 'document':
                    seq.append(('pic', r['pic'], None))
                    pic = True
            elif r.get('t') is not None:
                t = r['t'].replace('\r', '')
                if r.get('caps'):
                    t = t.upper()
                if not t:
                    continue
                if not r.get('v'):
                    sizes.append(r.get('sz') or 0)
                href = r.get('href') if links and part else None
                d = self.fmt(dict(r, t=t))
                if seq and seq[-1][0] == 't' and seq[-1][1]['d'] == d and seq[-1][2] == href:
                    seq[-1][1]['t'] += t
                else:
                    seq.append(('t', {'d': d, 't': t}, href))
        out, i = [], 0
        while i < len(seq):
            kind, _, href = seq[i]
            tag = self.link(href, part) if kind == 't' and href else None
            if tag:
                j = i
                while j < len(seq) and seq[j][0] == 't' and seq[j][2] == href:
                    j += 1
                out.append(tag + ''.join(self.run_xml(x, base, avail) for x in seq[i:j]) + '</w:hyperlink>')
                i = j
            else:
                out.append(self.run_xml(seq[i], base, avail))
                i += 1
        return ''.join(out), sizes, pic

    def run_xml(self, x, base, avail):
        kind, v, _ = x
        if kind == 'br':
            return '<w:r><w:br/></w:r>'
        if kind == 'fn':
            return self.footnote(v['fn'])
        if kind == 'pic':
            return self.inline_pic(v, avail)
        d = v['d']
        if d.get('cs') and d['cs'] != d.get('font') and any(CS_LO <= ord(c) <= CS_HI for c in v['t']):
            # Google Docs draws a complex script (Sinhala, Tamil) in the run's ascii font and ignores w:cs, so the
            # script's own pieces name its font in every slot (probed 18 Nov 2026: then it shows Noto Sans Sinhala)
            plain = {k: x for k, x in d.items() if k != 'hint'}
            return ''.join('<w:r>%s%s</w:r>' % (rpr_xml(dict(d, font=d['cs']) if k == 'cs' else plain, base), text_xml(t))
                           for k, t in cs_parts(v['t']))
        return '<w:r>%s%s</w:r>' % (rpr_xml(d, base), text_xml(v['t']))

    def footnote(self, runs):
        """A Word footnote from a note's runs; returns the reference run that marks it in the text."""
        base = self.sd.setdefault('FootnoteText', dict(self.sd['Normal']))
        n = len(self.notes) + 1
        body = self.runs_xml(tidy(runs), base, 0, part='footnotes')[0]
        self.notes.append('<w:footnote w:id="%d"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr>'
                          '<w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r>'
                          '<w:t xml:space="preserve"> </w:t></w:r>%s</w:p></w:footnote>' % (n, body))
        return '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="%d"/></w:r>' % n

    def marks(self, p):
        """Bookmarks round a paragraph for its own id and the ids of the blocks it opens, each placed once."""
        start, end = [], []
        for x in [p.get('id')] + list(p.get('ids') or []):
            if x and x in self.bm and x not in self.placed:
                self.placed.add(x)
                self.bmid += 1
                start.append('<w:bookmarkStart w:id="%d" w:name="%s"/>' % (self.bmid, self.bm[x]))
                end.append('<w:bookmarkEnd w:id="%d"/>' % self.bmid)
        return ''.join(start), ''.join(end)

    # -- lists

    def marker_rpr(self, f, w, color, sz):
        if f:
            name, b, _ = self.fonts.resolve(f, w or 400, False)
        else:
            name, b = self.sd['Normal']['font'], (w or 400) >= 600
        bold = '<w:b/><w:bCs/>' if b else '<w:b w:val="0"/><w:bCs w:val="0"/>'
        return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>%s<w:i w:val="0"/><w:iCs w:val="0"/>'
                '<w:color w:val="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
                % (esc(name), esc(name), esc(name), bold, (color or '000000').upper(), half_pt(sz or 14), half_pt(sz or 14)))

    def num(self, L, numfmt, text, rpr, left, hang):
        """The numId of a list (an abstract numbering and an instance for each list on the page); left and hang in
        twips."""
        key = L.get('id')
        if key in self.nums:
            return self.nums[key]
        n = len(self.nums) + 1
        lvl0 = min(8, L.get('lvl') or 0)
        levels = []
        for k in range(9):
            start = (L.get('start') or 1) if k == lvl0 else 1
            lt = text if numfmt == 'bullet' else '%%%d.' % (k + 1)
            levels.append('<w:lvl w:ilvl="%d"><w:start w:val="%d"/><w:numFmt w:val="%s"/><w:lvlText w:val="%s"/>'
                          '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="%d" w:hanging="%d"/></w:pPr>%s</w:lvl>'
                          % (k, start, numfmt, esc(lt), max(hang, left + (k - lvl0) * 360), hang, rpr))
        self.abstract.append('<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="hybridMultilevel"/>%s'
                             '</w:abstractNum>' % (n, ''.join(levels)))
        self.numdefs.append('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>' % (n, n))
        self.nums[key] = n
        return n

    def list_marker(self, p, runs, left):
        """A list item's marker: (numId or None, the runs without a marker that Word now draws, hanging indent in px
        or None)."""
        L = p.get('list') or {}
        hang = L.get('hang') or 0
        if hang < 2:
            hang = min(left, 1.5 * (L.get('ms') or p['pf'].get('sz') or 14))
        t = (L.get('type') or '').strip()
        lt, ht = tw(left), tw(min(hang, left))
        if t == 'none':
            texts = [r for r in runs if r.get('t') is not None]
            first = texts[0] if texts else None
            mark = first['t'].strip() if first is not None else ''
            if len(mark) == 1 and mark in MARKS:
                rp = self.marker_rpr(first.get('f'), first.get('w'), first.get('c'), first.get('sz'))
                return self.num(L, 'bullet', mark, rp, lt, ht), [r for r in runs if r is not first], hang
            if runs and runs[0].get('pre') and runs[0].get('t') is not None:
                r0 = dict(runs[0], t=runs[0]['t'].rstrip(' \t\n'))   # keep a filled marker's padding
                r0.pop('pre', None)
                tab = dict(r0, t='\t')                               # the tab to the text, without the marker's fill
                tab.pop('bg', None)
                return None, [r0, tab] + list(runs[1:]), hang
            return None, runs, None
        rp = self.marker_rpr(L.get('mf'), L.get('mw'), L.get('mc'), L.get('ms'))
        if t in FORMATS:
            return self.num(L, FORMATS[t], '', rp, lt, ht), runs, hang
        m = re.match(r'^["\'](.*)["\']$', t)
        mark = (m.group(1).strip() or '•') if m else BULLETS.get(t, '•')
        return self.num(L, 'bullet', mark, rp, lt, ht), runs, hang

    # -- paragraphs

    def normal_px(self):
        return (self.sd['Normal'].get('sz') or 21) / 1.5

    def ratio(self, runs):
        """A single line's height over the font size, for the paragraph's largest run: the font's hhea ascent,
        descent and line gap, as Word for Mac, LibreOffice and Google Docs lay lines out."""
        F = self.face(runs)
        if F:
            return F.asc + F.desc + F.gap
        return 1.15 if any(r.get('t') and r.get('f') and r.get('sz') for r in runs) else 1.2

    def natural(self, runs):
        """The tallest single line any app draws for the paragraph's largest run, over its size: Windows Word from
        the win metrics, Word for Mac and LibreOffice from hhea, Google Docs 1.2 for a font it lacks."""
        F = self.face(runs)
        return max(1.2, F.win_asc + F.win_desc, F.asc + F.desc + F.gap) if F else 1.4

    def face(self, runs):
        rs = [r for r in runs if r.get('t') and r.get('f') and r.get('sz')]
        if not rs:
            return None
        r = max(rs, key=lambda x: x['sz'])
        return self.fonts.info.get(self.fonts.resolve(r['f'], r.get('w') or 400, bool(r.get('i')))[0])

    def para(self, p, C, before, after=0, pb=False, style=None, cover=False, auto=False, part='document'):
        """One walker paragraph as a Word paragraph. C holds the shift from the page Chrome laid it out on (dx, dxr)
        and the width it had; before and after are px."""
        pf = p['pf']
        sid = style or self.style_of(p, cover)
        base = self.sd.get(sid) or self.sd['Normal']
        deco = not cover and bool(pf.get('shade') or any((pf.get('bd') or {}).values()))
        pad = pf.get('pad') or [0, 0, 0, 0]
        left = (pf.get('il') or 0) + C['dx'] + (pad[3] if deco else 0)
        right = (pf.get('ir') or 0) + C['dxr'] + (pad[1] if deco else 0)
        first = pf.get('ti') or 0
        runs = spaced(tidy(p.get('runs') or []))
        numpr = tabs = ''
        if p.get('list') and part == 'document':
            nid, runs, hang = self.list_marker(p, runs, left)
            if hang:
                first = -min(hang, left)
                tabs = '<w:tabs><w:tab w:val="left" w:pos="%d"/></w:tabs>' % tw(left)
            if nid:
                numpr = '<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>' % (
                    min(8, p['list'].get('lvl') or 0), nid)
        avail = C['width'] - (pf.get('il') or 0) - (pf.get('ir') or 0)
        body, sizes, pic = self.runs_xml(runs, base, avail, part)
        start, end = self.marks(p) if part == 'document' else ('', '')
        line, rule = (240, 'auto') if auto else line_rule(pf, sizes, pic, self.normal_px(), self.ratio(runs), self.natural(runs))
        mark = mark_rpr(sizes, base)
        bdr, fill = pbdr(pf) if deco else ('', '')
        return P(start + body + end, pStyle=pstyle(sid),
                 keepNext='<w:keepNext/>' if pf.get('kn') and not cover else '',
                 keepLines='<w:keepLines/>' if pf.get('kl') and not cover else '',
                 pageBreakBefore=BREAK if pb else '', numPr=numpr, pBdr=bdr, shd=fill, tabs=tabs,
                 spacing=spacing(tw(before), line, rule, tw(after)), ind=ind(tw(left), tw(right), tw(first)),
                 jc=jc(pf.get('al')), rPr=mark)

    # -- blocks (part 3c)

    def block(self, b, C, before, pb=False):
        k = b.get('k')
        if k == 'p':
            return [self.para(b, C, before, pb=pb)]
        if k == 'table':
            t = self.table(b, C)
            return [tiny(max(0, tw(before) - 20), pageBreakBefore=BREAK if pb else ''), t] if t else []
        if k == 'rule':
            return [self.rule(b, C, before, pb)]
        if k == 'pic':
            p = self.pic_block(b, C, before, pb)
            return [p] if p else []
        if k == 'toc':
            return self.toc(b, C, before, pb)
        return []

    def flow(self, blocks, C, prev, cols_top=None, fresh=False):
        """Blocks one after another in a cell, a column or the cover: (items, bottom of the last block)."""
        items, last = [], 0
        for b in blocks:
            k = b.get('k')
            if k in ('sect', 'cover'):
                continue
            if k == 'cols':
                colw = max(1, (C['width'] - b['gap'] * (b['n'] - 1)) / b['n'])
                sub, prev = self.flow(b.get('blocks') or [], dict(C, width=colw, dxr=C['dxr'] + C['width'] - colw),
                                      prev, cols_top=b.get('top'), fresh=fresh and not items)
                items += sub
                continue
            top = top_of(b)
            if prev is None or (fresh and not items):
                before = margin_of(b)
            elif cols_top is not None and top < prev - 1:
                before = margin_of(b) or last
            else:
                before = top - prev
            got = self.block(b, C, before)
            if got:
                items += got
                last = max(0, before)
            prev = bottom_of(b)
        return items, prev

    def table(self, t, C):
        rows = t.get('rows') or []
        xs = sorted([t['x'], t['x'] + t['w']] + [v for r in rows for c in r.get('cells') or [] for v in (c['x0'], c['x1'])])
        edges = [xs[0]]
        for x in xs[1:]:
            if x > edges[-1] + 1.5:
                edges.append(x)
        if len(edges) < 2:
            return None
        span = edges[-1] - edges[0]
        k = C['width'] / span if span > C['width'] + 2 else 1.0
        offs = [tw((e - edges[0]) * k) for e in edges]
        ws = [offs[i + 1] - offs[i] for i in range(len(offs) - 1)]

        def idx(x):
            return min(range(len(edges)), key=lambda i: abs(edges[i] - x))
        page_h = (self.geo.get(self.sec['page']) or self.geo['']).h
        out, pending = [], {}                       # start column -> [end column, cell, rows left]

        def fill(cur, stop, cells):
            while cur < stop:
                if cur in pending:
                    e, c, n = pending[cur]
                    cells.append(self.cell(c, ws[cur:e], merge='continue'))
                    pending[cur][2] -= 1
                    if pending[cur][2] <= 0:
                        del pending[cur]
                    cur = e
                else:
                    nxt = min([s for s in pending if s > cur] + [stop])
                    cells.append(self.cell(None, ws[cur:nxt]))
                    cur = nxt
            return cur
        for row in rows:
            if 'gap' in row:
                out.append('<w:tr><w:trPr><w:trHeight w:val="%d" w:hRule="exact"/></w:trPr>%s</w:tr>'
                           % (max(20, tw(row['gap'])), self.cell(None, ws)))
                continue
            cells, cur, hts = [], 0, []
            for c in sorted(row.get('cells') or [], key=lambda c: c['x0']):
                s, e = idx(c['x0']), idx(c['x1'])
                if e <= s or s < cur:
                    continue
                cur = fill(cur, s, cells)
                rs = c.get('rspan') or 1
                cells.append(self.cell(c, ws[s:e], merge='restart' if rs > 1 else None))
                if rs > 1:
                    pending[s] = [e, c, rs - 1]
                else:
                    hts.append((c['bot0'] - c['top0'] + c['pad'][0] + c['pad'][2] + bw(c['bd'].get('t')) + bw(c['bd'].get('b')),
                                c['bot0'] - c['top0']))
                cur = e
            fill(cur, len(ws), cells)
            h = max(x[0] for x in hts) if hts else 0
            tr = ''
            if t.get('kind') in ('grid', 'box') and h < 0.6 * page_h:
                tr += '<w:cantSplit/>'
            # the least height leaves out the cells' padding: Word counts padding inside it, LibreOffice adds it on top,
            # so a minimum of the content alone gives both the kit's height (the text and padding make it up)
            hc = max(x[1] for x in hts) if hts else 0
            if hc > 1:
                tr += '<w:trHeight w:val="%d" w:hRule="atLeast"/>' % tw(hc)
            if row.get('head'):
                tr += '<w:tblHeader/>'
            out.append('<w:tr>%s%s</w:tr>' % ('<w:trPr>%s</w:trPr>' % tr if tr else '', ''.join(cells)))
        nil = ''.join('<w:%s w:val="nil"/>' % s for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
        pr = ('<w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblInd w:w="%d" w:type="dxa"/><w:tblBorders>%s</w:tblBorders>'
              '<w:tblLayout w:type="fixed"/>%s<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0"'
              ' w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>'
              % (offs[-1], tw(max(0, (t.get('il') or 0) + (edges[0] - t['x']) + C['dx'])), nil, CELL_MAR0))
        grid = '<w:tblGrid>%s</w:tblGrid>' % ''.join('<w:gridCol w:w="%d"/>' % w for w in ws)
        return T('<w:tbl>%s%s%s</w:tbl>' % (pr, grid, ''.join(out)))

    def cell(self, c, ws, merge=None):
        """A table cell over len(ws) grid columns; c None is an empty spacer."""
        span = '<w:gridSpan w:val="%d"/>' % len(ws) if len(ws) > 1 else ''
        width = '<w:tcW w:w="%d" w:type="dxa"/>' % sum(ws)
        if c is None:
            nil = ''.join('<w:%s w:val="nil"/>' % s for s in ('top', 'left', 'bottom', 'right'))
            return '<w:tc><w:tcPr>%s%s<w:tcBorders>%s</w:tcBorders></w:tcPr>%s</w:tc>' % (width, span, nil, tiny().xml())
        bd, pad = c.get('bd') or {}, c.get('pad') or [0, 0, 0, 0]
        m = [pad[0] + bw(bd.get('t')), pad[1] + bw(bd.get('r')), pad[2] + bw(bd.get('b')), pad[3] + bw(bd.get('l'))]
        vm = '<w:vMerge w:val="restart"/>' if merge == 'restart' else '<w:vMerge/>' if merge == 'continue' else ''
        borders = '<w:tcBorders>%s%s%s%s</w:tcBorders>' % (border('top', bd.get('t')), border('left', bd.get('l')),
                                                           border('bottom', bd.get('b')), border('right', bd.get('r')))
        mar = '<w:tcMar>%s</w:tcMar>' % ''.join('<w:%s w:w="%d" w:type="dxa"/>' % (s, tw(v)) for s, v in
                                                 (('top', m[0]), ('left', m[3]), ('bottom', m[2]), ('right', m[1])))
        va = {'center': '<w:vAlign w:val="center"/>', 'bottom': '<w:vAlign w:val="bottom"/>'}.get(c.get('va'), '')
        pr = '<w:tcPr>%s%s%s%s%s%s%s</w:tcPr>' % (width, span, vm, borders, shd(c.get('shade')), mar, va)
        if merge == 'continue':
            return '<w:tc>%s%s</w:tc>' % (pr, tiny().xml())
        C = {'dx': 0, 'dxr': 0, 'width': max(1, c['x1'] - c['x0'] - m[1] - m[3])}
        items, _ = self.flow(c.get('blocks') or [], C, c.get('top0'), fresh=c.get('va') in ('center', 'bottom'))
        if not items or not isinstance(items[-1], P):
            items.append(tiny())
        return '<w:tc>%s%s</w:tc>' % (pr, ''.join(x.xml() for x in items))

    def rule(self, b, C, before, brk):
        side = 'bottom' if b.get('side') == 'b' else 'top'
        return P('', pageBreakBefore=BREAK if brk else '', pBdr='<w:pBdr>%s</w:pBdr>' % border(side, b.get('bd')),
                 spacing=spacing(tw(before), 20, 'exact'),
                 ind=ind(tw((b.get('il') or 0) + C['dx']), tw((b.get('ir') or 0) + C['dxr'])), rPr=TINY)

    def pic_block(self, b, C, before, pb):
        al, il = b.get('al') or 'left', max(0, b.get('il') or 0)
        body = self.inline_pic(b.get('pic') or {}, C['width'] - (il if al == 'left' else 0))
        if not body:
            return None
        return P(body, pageBreakBefore=BREAK if pb else '', spacing=spacing(tw(before)),
                 ind=ind(tw((il if al == 'left' else 0) + C['dx']), tw(C['dxr'])), jc=jc(al))

    def toc(self, b, C, before, pb):
        entries = b.get('entries') or []
        lv = sorted({int(self.tags[x][1]) for x in (unquote((e.get('href') or '')[1:]) for e in entries
                                                    if (e.get('href') or '').startswith('#'))
                     if re.match(r'h[1-6]$', self.tags.get(x) or '')})
        rng = '%d-%d' % (lv[0], lv[-1]) if lv else '1-3'
        out, prev = [], None
        for i, e in enumerate(entries):
            pf, runs = e['pf'], list(e.get('runs') or [])
            page, src = (e.get('page') or '').strip(), e.get('page_run')
            if not page and runs and runs[-1].get('t') is not None and PAGE_NO.match(runs[-1]['t']) and runs[-1]['t'].strip():
                src, page = runs[-1], runs[-1]['t'].strip()
                runs = runs[:-1]
                while runs and runs[-1].get('br'):
                    runs.pop()
            runs = tidy(runs, toc=True)
            if src is None:
                src = next((r for r in reversed(runs) if r.get('t')), None) or {}
            sid = 'TOC%d' % min(9, e.get('lvl') or 1)
            sid = sid if sid in self.sd else 'Normal'
            base = self.sd[sid]
            body, sizes, _ = self.runs_xml(runs, base, C['width'], links=False)
            leader = ' w:leader="dot"' if e.get('leader') == 'dot' else ''
            right = C['dx'] + C['width'] - (pf.get('ir') or 0)
            end = C['dx'] + e['dots_end'] if leader and e.get('dots_end') is not None else right
            short = end < right - 2                    # the kit's dots stop short of the page number: two right tabs
            stops = ['<w:tab w:val="left" w:pos="%d"/>' % tw(C['dx'] + r['tab_at']) for r in runs if r.get('tab_at') is not None]
            if short:
                stops += ['<w:tab w:val="right"%s w:pos="%d"/>' % (leader, tw(end)), '<w:tab w:val="right" w:pos="%d"/>' % tw(right)]
            else:
                stops.append('<w:tab w:val="right"%s w:pos="%d"/>' % (leader, tw(right)))
            if page:
                rp = rpr_xml(self.fmt(dict(src, t=page)), base)
                body += '<w:r>%s<w:tab/>%s</w:r><w:r>%s%s</w:r>' % (rp, '<w:tab/>' if short else '', rp, text_xml(page))
            href = e.get('href') or ''
            x = unquote(href[1:]) if href.startswith('#') else None
            if x in self.bm:
                body = '<w:hyperlink w:anchor="%s" w:history="1">%s</w:hyperlink>' % (self.bm[x], body)
            if i == 0:
                body = FIELD_BEGIN % rng + body
            if i == len(entries) - 1:
                body += FIELD_END
            gap = before if prev is None else pf['top'] - prev
            prev = pf['bottom']
            line, rule = line_rule(pf, sizes, normal=self.normal_px(), k=self.ratio(runs), nat=self.natural(runs))
            out.append(P(body, pStyle=pstyle(sid), pageBreakBefore=BREAK if pb and i == 0 else '',
                         tabs='<w:tabs>%s</w:tabs>' % ''.join(stops),
                         spacing=spacing(tw(gap), line, rule),
                         ind=ind(tw((pf.get('il') or 0) + C['dx']), tw((pf.get('ir') or 0) + C['dxr']), tw(pf.get('ti') or 0)),
                         jc=jc(pf.get('al')), rPr=mark_rpr(sizes, base)))
        return out

    def cover(self, b):
        """The cover: its art behind the text, its top lines as paragraphs kept well inside the page, and its bottom
        lines (author, date) in the first page's footer, so nothing spills to a second page in any app."""
        g = self.geo.get(self.sec['page']) or self.geo['']
        C, H, ay = {'dx': -g.ml, 'dxr': -g.mr, 'width': g.w}, b['h'], b['y']
        ps = [p for p in b.get('blocks') or [] if p.get('k') == 'p']
        cut = next((i for i, p in enumerate(ps) if p['pf'].get('meta') and p['pf']['top'] - ay > 0.6 * H), None)
        if cut is None:
            i = len(ps) - 1
            while i > 0 and ps[i]['pf']['top'] - ay >= 0.6 * H:
                if ps[i]['pf']['top'] - ps[i - 1]['pf']['bottom'] > 0.1 * H:
                    cut = i
                    break
                i -= 1
        top, meta = (ps, []) if cut is None else (ps[:cut], ps[cut:])
        fd = self.hf_dist(g, 'bottom')
        foot = []
        for j, p in enumerate(meta):
            gap = 0 if j == 0 else p['pf']['top'] - meta[j - 1]['pf']['bottom']
            after = max(0, H - (p['pf']['bottom'] - ay) - fd) if j == len(meta) - 1 else 0
            foot.append(self.para(p, C, gap, after=after, style='Footer', cover=True, auto=True, part=None))
        limit = max(g.mb, H - (meta[0]['pf']['top'] - ay)) if meta else g.mb
        room = 0.8 * (H - g.mt - limit)
        prev, gaps, heights = ay + g.mt, [], []
        for p in top:
            gaps.append(max(0, p['pf']['top'] - prev))
            heights.append(max(0, p['pf']['bottom'] - p['pf']['top']))
            prev = p['pf']['bottom']
        if gaps and sum(gaps) + sum(heights) > room:
            k = max(0, (room - sum(heights)) / sum(gaps)) if sum(gaps) else 0
            gaps = [x * k for x in gaps]
        items = [self.para(p, C, gaps[j], cover=True) for j, p in enumerate(top)] or [tiny()]
        foot = foot or [tiny(pStyle=pstyle('Footer'))]
        fname = self.part_name('footer')
        if 'cover' in self.shots:
            size = (min(b['w'], g.w), min(H, g.h))
            free = (size[0], size[1], size[0], size[1])
            if meta:
                y0, y1 = meta[0]['pf']['top'], meta[-1]['pf']['bottom']
                rs = [r for r in self.data.get('cover_text') or [] if y0 - 2 <= r[1] + r[3] / 2 <= y1 + 2]
                x0 = min([r[0] for r in rs] or [b['x'] + g.ml]) - b['x']
                x1 = max([r[0] + r[2] for r in rs] or [b['x'] + size[0] - g.mr]) - b['x']
                y0, y1 = min([r[1] for r in rs] or [y0]) - ay, max([r[1] + r[3] for r in rs] or [y1]) - ay
                # room for the lines to land a little apart in Word or Google Docs, or in a wider font
                free = (max(0, x0 - 12), max(0, y0 - 16), min(size[0], x1 + max(24, 0.25 * (x1 - x0))),
                        min(size[1], y1 + 16))
            pieces, own = cover_pieces(self.shots['cover'][0], size, free, self.data.get('bg'))
            for j, (x, y, w, h, shot) in enumerate(pieces):
                self.shots['cover:%d' % j] = shot
                items[0].body = self.drawing(self.picture('cover:%d' % j), w, h, behind=True, x=x, y=y) + items[0].body
            if own:
                self.shots['cover:foot'] = own[4]
                foot[0].body = self.drawing(self.picture('cover:foot', fname), own[2], own[3], behind=True,
                                            x=own[0], y=own[1]) + foot[0].body
        self.body += items
        self.sec['first'] = (self.hf_part('header', [tiny(pStyle=pstyle('Header'))]),
                             self.hf_part('footer', foot, fname))

    # -- headers, footers and sections

    def hf_dist(self, g, where):
        """Word's header or footer distance from the page edge, where the kit's margin box puts its text: the box fills
        the page margin less its own margin, and its text sits at the top, middle or bottom inside its border and
        padding."""
        bx = next((x for x in (box_of(g, '%s-%s' % (where, s)) for s in ('left', 'center', 'right')) if x), None)
        M = g.mt if where == 'top' else g.mb
        if not bx:
            return max(0, min(M / 2, 48))
        lh, va = 1.2 * (bx.get('sz') or 10), bx.get('va') or 'middle'
        near = (bx.get('mb') or 0) + bw(bx.get('bb')) + (bx.get('pb') or 0)    # bottom side, margin to text
        far = (bx.get('mt') or 0) + bw(bx.get('bt')) + (bx.get('pt') or 0)     # top side
        if where == 'top':
            d = M - near - lh if va == 'bottom' else far if va == 'top' else (M - lh) / 2
        else:
            d = M - far - lh if va == 'top' else near if va == 'bottom' else (M - lh) / 2
        return max(0, d)

    def box_runs(self, bx, base):
        rp = rpr_xml(self.fmt({'f': bx.get('f'), 'w': bx.get('w'), 'sz': bx.get('sz'), 'c': bx.get('c'),
                               'ls': bx.get('ls'), 't': 'x'}), base)
        out = []
        for part in bx.get('parts') or []:
            if part.get('f'):
                out.append('<w:fldSimple w:instr=" %s "><w:r>%s<w:t>1</w:t></w:r></w:fldSimple>' % (part['f'], rp))
            elif part.get('t'):
                out.append('<w:r>%s%s</w:r>' % (rp, text_xml(part['t'].upper() if bx.get('caps') else part['t'])))
        return ''.join(out)

    def box_para(self, g, where):
        sid = 'Header' if where == 'top' else 'Footer'
        bxs = [box_of(g, '%s-%s' % (where, s)) for s in ('left', 'center', 'right')]
        if not any(bxs):
            return tiny(pStyle=pstyle(sid))
        L, M, R = (self.box_runs(x, self.sd[sid]) if x else '' for x in bxs)
        tabs, body = '', L
        if M:
            tabs += '<w:tab w:val="center" w:pos="%d"/>' % tw(g.cw / 2)
            body += '<w:r><w:tab/></w:r>' + M
        if R:
            tabs += '<w:tab w:val="right" w:pos="%d"/>' % tw(g.cw)
            body += '<w:r><w:tab/></w:r>' + R
        line = ''
        for x in bxs:
            b = x and (x.get('bb') if where == 'top' else x.get('bt'))
            if b and (x.get('parts') or x.get('text')):
                tag, pad = ('bottom', x.get('pb')) if where == 'top' else ('top', x.get('pt'))
                line = '<w:pBdr>%s</w:pBdr>' % border(tag, b).replace('w:space="0"', 'w:space="%d"' % border_space(pad or 0))
                break
        return P(body, pStyle=pstyle(sid), pBdr=line, tabs='<w:tabs>%s</w:tabs>' % tabs if tabs else '', spacing=spacing())

    def part_name(self, kind):
        return '%s%d.xml' % (kind, sum(1 for n in self.parts if n.startswith(kind)) + 1)

    def hf_part(self, kind, paras, name=None):
        name = name or self.part_name(kind)
        tag = 'hdr' if kind == 'header' else 'ftr'
        self.parts[name] = '%s<w:%s %s>%s</w:%s>' % (XML_HEAD, tag, XMLNS, ''.join(p.xml() for p in paras), tag)
        return self.rel(kind, name)

    def sect_xml(self, s):
        g = self.geo.get(s['page']) or self.geo['']
        refs = ''
        if s['page'] != self.prev_page or s.get('first'):
            if s['page'] not in self.hf:
                self.hf[s['page']] = (self.hf_part('header', [self.box_para(g, 'top')]),
                                      self.hf_part('footer', [self.box_para(g, 'bottom')]))
            h, f = self.hf[s['page']]
            refs = '<w:headerReference w:type="default" r:id="%s"/>' % h
            if s.get('first'):
                refs += '<w:headerReference w:type="first" r:id="%s"/>' % s['first'][0]
            refs += '<w:footerReference w:type="default" r:id="%s"/>' % f
            if s.get('first'):
                refs += '<w:footerReference w:type="first" r:id="%s"/>' % s['first'][1]
        self.prev_page = s['page']
        cols = ('<w:cols w:num="%d" w:space="%d"/>' % (s['cols'][0], tw(s['cols'][1])) if s.get('cols')
                else '<w:cols w:space="720"/>')
        return ('<w:sectPr>%s<w:type w:val="%s"/><w:pgSz w:w="%d" w:h="%d"%s/><w:pgMar w:top="%d" w:right="%d"'
                ' w:bottom="%d" w:left="%d" w:header="%d" w:footer="%d" w:gutter="0"/>%s%s</w:sectPr>'
                % (refs, s['type'], tw(g.w), tw(g.h), ' w:orient="landscape"' if g.w > g.h else '', tw(g.mt), tw(g.mr),
                   tw(g.mb), tw(g.ml), tw(self.hf_dist(g, 'top')), tw(self.hf_dist(g, 'bottom')), cols,
                   '<w:titlePg/>' if s.get('first') else ''))

    def ctx(self):
        g0, g = self.geo[''], self.geo.get(self.sec['page']) or self.geo['']
        return {'dx': g0.ml - g.ml, 'dxr': g0.mr - g.mr, 'width': g0.cw}

    def begin(self, page, kind):
        if len(self.body) == self.sec['start'] and not self.sec.get('first'):
            self.sec['page'] = page
            if kind == 'nextPage':
                self.sec['type'] = kind
            return
        self.close()
        self.sec = {'page': page, 'type': kind, 'cols': None, 'first': None, 'start': len(self.body)}

    def close(self):
        # A section never ends on the paragraph that closes a field: Google Docs turns a TOC field into its own
        # contents element and drops a section break inside it, which pulled the next chapter onto the contents page.
        if len(self.body) == self.sec['start'] or not isinstance(self.body[-1], P) or 'sectPr' in self.body[-1].ppr \
                or FIELD_END in self.body[-1].body:
            self.body.append(tiny())
        self.body[-1].ppr['sectPr'] = self.sect_xml(self.sec)

    def write(self):
        self.prescan()
        self.hf = {}
        C, prev, fresh, after_cover = self.ctx(), None, True, False
        for b in self.data.get('blocks') or []:
            k = b.get('k')
            if k == 'sect':
                page = b.get('page') or ''
                if page == 'pc-cover' or (after_cover and not page):
                    continue
                self.begin(page if page in self.geo else '', 'nextPage')
                C, prev, fresh, self.pb = self.ctx(), None, True, False
                continue
            after_cover = False
            if k == 'cover':
                if len(self.body) > self.sec['start'] or self.sec.get('first'):
                    self.begin(self.sec['page'], 'nextPage')
                self.cover(b)
                C, prev, fresh, after_cover, self.pb = self.ctx(), None, False, True, True
                continue
            if k == 'cols':
                brk = (self.pb or break_of(b)) and not fresh
                gap = b['top'] - prev if prev is not None and not brk else 0
                last = self.body[-1] if len(self.body) > self.sec['start'] else None
                if gap > 0.5 and isinstance(last, P):        # the gap above the columns, kept after the block before them
                    for key, v in last.ppr.items():
                        if isinstance(v, str) and v.startswith('<w:spacing') and 'w:after="0"' in v:
                            last.ppr[key] = v.replace('w:after="0"', 'w:after="%d"' % tw(gap), 1)
                self.begin(self.sec['page'], 'nextPage' if brk else 'continuous')
                self.sec['cols'] = (b['n'], b['gap'])
                colw = max(1, (C['width'] - b['gap'] * (b['n'] - 1)) / b['n'])
                items, prev = self.flow(b.get('blocks') or [], dict(C, width=colw), b['top'] if prev is None else prev,
                                        cols_top=b['top'], fresh=True)
                # The next block's gap counts from the foot of the taller column, as the section's end does in Word,
                # not from the last block, which may sit in a shorter last column.
                prev = max(prev or 0, b.get('bottom') or 0)
                self.body += items
                self.begin(self.sec['page'], 'continuous')
                fresh, self.pb = False, False
                continue
            pb = (self.pb or break_of(b)) and not fresh
            before = margin_of(b) if prev is None or pb else top_of(b) - prev
            items = self.block(b, C, before, pb)
            if items:
                self.body += items
                fresh, self.pb = False, False
            elif pb:
                self.pb = True
            prev = bottom_of(b)
        if not self.body or not isinstance(self.body[-1], P) or len(self.body) == self.sec['start'] \
                or 'sectPr' in self.body[-1].ppr:
            self.body.append(tiny())
        self.final = self.sect_xml(self.sec)

    # -- the other XML parts the body needs

    def styles_xml(self):
        N = self.sd['Normal']
        font, sz = esc(N.get('font') or 'Arial'), N.get('sz') or 21
        lang = ('<w:lang w:val="en-GB" w:eastAsia="en-GB" w:bidi="%s"/>' % self.lang0 if self.lang0 in ('si-LK', 'ta-LK')
                else '<w:lang w:val="%s" w:eastAsia="%s"/>' % (self.lang0, self.lang0))
        out = ['<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
               '<w:sz w:val="%d"/><w:szCs w:val="%d"/>%s</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0"'
               ' w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
               % (font, font, font, font, sz, sz, lang),
               '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>%s'
               '</w:style>' % rpr_xml(N, {})]
        order = (['Title'] + ['Heading%d' % k for k in range(1, 7)] + ['TOCHeading'] + ['TOC%d' % k for k in range(1, 10)]
                 + ['FootnoteText', 'Header', 'Footer'])
        for sid in order:
            if sid not in self.sd:
                continue
            m = re.match(r'(Heading|TOC)(\d)$', sid)
            head = bool(m and m.group(1) == 'Heading')
            name = NAMES.get(sid) or ('heading %s' % m.group(2) if head else 'toc %s' % m.group(2))
            pri = 9 if head else 10 if sid == 'Title' else 39 if sid.startswith('TOC') else 99
            ppr = ('<w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="%d"/></w:pPr>' % (int(m.group(2)) - 1)
                   if head else '')
            out.append('<w:style w:type="paragraph" w:styleId="%s"><w:name w:val="%s"/><w:basedOn w:val="Normal"/>%s'
                       '<w:uiPriority w:val="%d"/>%s%s%s%s</w:style>'
                       % (sid, name, '<w:next w:val="Normal"/>' if head or sid in ('Title', 'TOCHeading') else '', pri,
                          '' if head or sid == 'Title' else '<w:unhideWhenUsed/>',
                          '<w:qFormat/>' if head or sid in ('Title', 'TOCHeading') else '', ppr,
                          rpr_xml(self.sd[sid], N)))
        out.append('<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default'
                   ' Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>')
        if self.notes:
            out.append('<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>'
                       '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr>'
                       '<w:vertAlign w:val="superscript"/></w:rPr></w:style>')
        out.append('<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
                   '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
                   '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0"'
                   ' w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>')
        out.append('<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>'
                   '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>')
        return XML_HEAD + '<w:styles %s>%s</w:styles>' % (XMLNS, ''.join(out))

    def numbering_xml(self):
        return XML_HEAD + '<w:numbering %s>%s%s</w:numbering>' % (XMLNS, ''.join(self.abstract), ''.join(self.numdefs))

    def footnotes_xml(self):
        sep = ''.join('<w:footnote w:type="%s" w:id="%d"><w:p><w:pPr><w:spacing w:after="0" w:line="240"'
                      ' w:lineRule="auto"/></w:pPr><w:r><w:%s/></w:r></w:p></w:footnote>' % (t, i, t)
                      for t, i in (('separator', -1), ('continuationSeparator', 0)))
        return XML_HEAD + '<w:footnotes %s>%s%s</w:footnotes>' % (XMLNS, sep, ''.join(self.notes))

    def background(self):
        bg = self.data.get('bg')
        bg = bg.lstrip('#').upper() if isinstance(bg, str) and re.match(r'#?[0-9a-fA-F]{6}$', bg) else 'FFFFFF'
        return '' if bg == 'FFFFFF' else bg

    def document_xml(self):
        bg = self.background()
        return XML_HEAD + '<w:document %s>%s<w:body>%s%s</w:body></w:document>' % (
            XMLNS, '<w:background w:color="%s"/>' % bg if bg else '', ''.join(x.xml() for x in self.body), self.final)

# ---- the package: the other parts and the zip

WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
TYPES = {'document.xml': WML + 'document.main+xml', 'styles.xml': WML + 'styles+xml', 'settings.xml': WML + 'settings+xml',
         'fontTable.xml': WML + 'fontTable+xml', 'numbering.xml': WML + 'numbering+xml',
         'footnotes.xml': WML + 'footnotes+xml', 'header': WML + 'header+xml', 'footer': WML + 'footer+xml',
         'core.xml': 'application/vnd.openxmlformats-package.core-properties+xml',
         'app.xml': 'application/vnd.openxmlformats-officedocument.extended-properties+xml'}
DEFAULTS = (('rels', 'application/vnd.openxmlformats-package.relationships+xml'), ('xml', 'application/xml'),
            ('png', 'image/png'), ('jpeg', 'image/jpeg'),
            ('odttf', 'application/vnd.openxmlformats-officedocument.obfuscatedFont'))


def settings_xml(bg, notes, embed):
    return (XML_HEAD + '<w:settings %s>%s%s<w:defaultTabStop w:val="720"/><w:characterSpacingControl'
            ' w:val="doNotCompress"/>%s<w:compat><w:compatSetting w:name="compatibilityMode"'
            ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
            % (XMLNS, '<w:displayBackgroundShape/>' if bg else '', '<w:embedTrueTypeFonts/>' if embed else '',
               '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' if notes else ''))


def core_xml(title, author):
    now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    return (XML_HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/'
            'core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"'
            ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>%s</dc:title><dc:creator>%s</dc:creator>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">'
            '%s</dcterms:modified></cp:coreProperties>' % (esc(title), esc(author), now, now))


def app_xml():
    return (XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
            ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>%s'
            '</Application></Properties>' % KIT_MARK)


def rels_xml(lst):
    return XML_HEAD + '<Relationships xmlns="%s">%s</Relationships>' % (PKG_REL, ''.join(
        '<Relationship Id="%s" Type="%s" Target="%s"%s/>' % (rid, typ, esc(target), ' TargetMode="External"' if ext else '')
        for rid, typ, target, ext in lst))


def package(w, fonts, out, title, author):
    """Write the .docx."""
    font_xml, font_rels, font_files = fonts.table()
    doc = w.document_xml()
    parts = {'word/document.xml': doc, 'word/styles.xml': w.styles_xml(),
             'word/settings.xml': settings_xml(bool(w.background()), bool(w.notes), True),
             'word/fontTable.xml': font_xml, 'docProps/core.xml': core_xml(title, author), 'docProps/app.xml': app_xml()}
    if font_files:
        parts['word/_rels/fontTable.xml.rels'] = font_rels
    fixed = [('styles', 'styles.xml'), ('settings', 'settings.xml'), ('fontTable', 'fontTable.xml')]
    if w.numdefs:
        parts['word/numbering.xml'] = w.numbering_xml()
        fixed.append(('numbering', 'numbering.xml'))
    if w.notes:
        parts['word/footnotes.xml'] = w.footnotes_xml()
        fixed.append(('footnotes', 'footnotes.xml'))
        if w.rels.get('footnotes'):
            parts['word/_rels/footnotes.xml.rels'] = rels_xml(w.rels['footnotes'])
    for name, xml in w.parts.items():
        parts['word/' + name] = xml
        if w.rels.get(name):
            parts['word/_rels/%s.rels' % name] = rels_xml(w.rels[name])
    parts['word/_rels/document.xml.rels'] = rels_xml(list(w.rels.get('document') or []) + [
        ('rIdX%d' % (i + 1), REL[k], t, False) for i, (k, t) in enumerate(fixed)])
    parts['_rels/.rels'] = rels_xml([('rId1', REL['officeDocument'], 'word/document.xml', False),
                                     ('rId2', REL['core-properties'], 'docProps/core.xml', False),
                                     ('rId3', REL['extended-properties'], 'docProps/app.xml', False)])
    over = []
    for name in parts:
        if name.endswith('.rels'):
            continue
        base = name.rsplit('/', 1)[-1]
        kind = TYPES.get(base) or TYPES.get(re.sub(r'\d+\.xml$', '', base))
        over.append('<Override PartName="/%s" ContentType="%s"/>' % (name, kind))
    ct = (XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">%s%s</Types>'
          % (''.join('<Default Extension="%s" ContentType="%s"/>' % d for d in DEFAULTS), ''.join(over)))
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_name('.%s.tmp' % out.name)
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        for name in sorted(parts, key=lambda n: (not n.startswith('_rels'), not n.startswith('docProps'), n)):
            z.writestr(name, parts[name])
        for name, data in w.media:
            z.writestr('word/media/' + name, data)
        for name, data in font_files:
            z.writestr('word/fonts/' + name, data)
    os.replace(tmp, out)
    return out

# ---- checks: the Word file rendered by LibreOffice, Word or Google Docs, against the kit's PDF


def find_soffice():
    for c in (os.environ.get('SOFFICE_PATH', ''), shutil.which('soffice') or '', shutil.which('libreoffice') or '',
              '/Applications/LibreOffice.app/Contents/MacOS/soffice',
              os.path.expanduser('~/Applications/LibreOffice.app/Contents/MacOS/soffice'),
              r'C:\Program Files\LibreOffice\program\soffice.exe'):
        if c and os.path.isfile(c):
            return c
    return None


def lo_pdf(docx, folder):
    exe = find_soffice()
    if not exe:
        return None, 'LibreOffice was not found (set SOFFICE_PATH to its soffice)'
    prof = tempfile.mkdtemp(prefix='pc-lo-')
    try:
        r = subprocess.run([exe, '-env:UserInstallation=' + Path(prof).as_uri(), '--headless', '--convert-to', 'pdf',
                            '--outdir', prof, str(docx)], capture_output=True, text=True, timeout=300)
        made = Path(prof) / (Path(docx).stem + '.pdf')
        if not made.exists():
            return None, (r.stderr or r.stdout or 'no PDF').strip()[-300:]
        keep = Path(folder) / 'libreoffice.pdf'
        shutil.move(str(made), keep)
        return keep, ''
    finally:
        shutil.rmtree(prof, ignore_errors=True)


def word_pdf(docx, folder):
    """Have Microsoft Word (macOS) save the file as PDF. Word opens files only where the user has let it: the first
    run shows Word's Grant File Access window for CHECK_DIR; later runs need nothing. If the check started Word, it
    quits Word afterwards; if Word was open, it closes only its own copy."""
    if sys.platform != 'darwin' or not Path('/Applications/Microsoft Word.app').exists():
        return None, 'it needs Microsoft Word on a Mac'
    box = CHECK_DIR
    src, pdf = box / 'check.docx', box / 'check.pdf'
    try:
        box.mkdir(exist_ok=True)
        for f in (src, pdf):
            f.unlink(missing_ok=True)
        shutil.copy(docx, src)
    except OSError as e:
        return None, 'could not write to %s (%s)' % (box, e.strerror)
    running = lambda: subprocess.run(['pgrep', '-x', 'Microsoft Word'], capture_output=True).returncode == 0
    was_open = running()
    script = ('with timeout of 300 seconds\n tell application "Microsoft Word"\n  set e to ""\n  set n to 560\n'
              '  try\n   open POSIX file "%s"\n  on error m number k\n   set e to " (" & m & ", " & k & ")"\n'
              '   set n to 40\n  end try\n  set d to missing value\n  repeat n times\n'
              '   try\n    repeat with i from 1 to (count of documents)\n'
              '     if name of document i is "check.docx" then set d to i\n    end repeat\n'
              '    if d is not missing value then exit repeat\n'
              '   end try\n   delay 0.5\n  end repeat\n'
              '  if d is missing value then error "Word did not open the file" & e\n  delay 2\n'
              '  save as document d file name ((POSIX file "%s") as text) file format format PDF\n'
              '  close document d saving no\n end tell\nend timeout'
              % (src, pdf))
    err = ''
    try:
        r = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=330)
        lines = (r.stderr or '').strip().splitlines()
        err = lines[-1] if lines else ''
    except Exception as e:
        err = 'Word did not respond: %s' % e
    finally:
        if not was_open:
            for _ in range(3):
                try:
                    subprocess.run(['osascript', '-e', 'tell application "Microsoft Word" to quit saving no'],
                                   capture_output=True, timeout=60)
                except Exception:
                    pass
                for _ in range(10):
                    if not running():
                        break
                    time.sleep(1)
                else:
                    continue
                break
    if not pdf.exists():
        src.unlink(missing_ok=True)
        hint = (' Word opened the file but handles no save command, as it does when it is not activated (read-only '
                'mode): sign in to Word, or check the file in LibreOffice and Google Docs only.') if '-1708' in (err or '') else ''
        return None, ('Word did not save a PDF%s.%s The first run asks for access to %s: click Grant Access in Word\'s '
                      'window (and allow control of Word if macOS asks), then run it again.' % (' (%s)' % err if err else '', hint, box))
    keep = Path(folder) / 'word.pdf'
    shutil.copy(pdf, keep)
    for f in (src, pdf):                               # the folder stays: Word's access is tied to it
        f.unlink(missing_ok=True)
    return keep, ''


def page_words(page, band=True):
    """A page's words as (text, top, line key), read in order; band leaves out the running heads and page numbers.
    Words come from the characters: they also end at a gap wider than a third of the type size (a tab, a table's
    next cell) and at a run of leader dots, which the plain word list glues to their neighbours. A word Chrome
    hyphenated at a line's end (U+2010 or a soft hyphen) is joined to the next line's first word."""
    H = page.rect.height
    found = []
    for bi, blk in enumerate(page.get_text('rawdict').get('blocks') or []):
        for li, line in enumerate(blk.get('lines') or []):
            cur, end = [], None

            def flush():
                if cur:
                    for part in re.split(r'\.{3,}|\u2026+', ''.join(c for c, _ in cur)):
                        found.append((part, cur[0][1], (bi, li), min(x for _, x in cur)))
                del cur[:]
            for span in line.get('spans') or []:
                for ch in span.get('chars') or []:
                    c, (x0, y0, x1, _) = ch['c'], ch['bbox']
                    if c.isspace() or (end is not None and x0 - end > max(1.5, span['size'] / 3)):
                        flush()
                    if not c.isspace():
                        cur.append((c, (x0, y0)))
                    end = x1
            flush()
    lines = {}
    for i, w in enumerate(found):
        lines.setdefault(w[2], []).append(i)
    gone = set()
    for (bi, li), idx in lines.items():
        last, nxt = found[idx[-1]], lines.get((bi, li + 1))
        if nxt and last[0][-1:] in ('\u2010', '\u00ad') and found[nxt[0]][0][:1].islower():
            found[idx[-1]] = (last[0][:-1] + found[nxt[0]][0],) + last[1:]
            gone.add(nxt[0])
    found = [w for i, w in enumerate(found) if i not in gone]
    out = []
    for t, (x0, y0), key, _ in sorted(found, key=lambda w: (round(w[1][1] / 3), w[1][0])):
        t = re.sub(r'^\W+|\W+$', '', unicodedata.normalize('NFKC', t)).casefold()
        if t and (not band or 0.085 * H < y0 < 0.915 * H):
            out.append((t, y0, key))
    return out


def compare(ref_pdf, test_pdf, label, sheet=None):
    """How a render of the Word file differs from the kit's PDF: pages, where each page starts, the cover's lines,
    words. Writes a side-by-side PDF (the kit's page left, the render right) when sheet is a path."""
    import fitz
    A, B = fitz.open(ref_pdf), fitz.open(test_pdf)
    out = ['%s: %d pages, the kit PDF %d%s' % (label, B.page_count, A.page_count,
                                                '' if A.page_count == B.page_count else '  ! page counts differ')]
    if A.page_count and B.page_count:
        miss = Counter(t for t, _, _ in page_words(A[0], False)) - Counter(t for t, _, _ in page_words(B[0], False))
        out.append('  page 1: ' + ('every word of the kit\'s page 1 is on it' if not miss else
                                   '! not on page 1: ' + ' '.join(list(miss.elements())[:15])))
        lines_a, lines_b = {}, {}
        for words, lines in ((page_words(A[0], False), lines_a), (page_words(B[0], False), lines_b)):
            for t, y, key in words:
                lines.setdefault(key, [y, []])[1].append(t)
        at = {' '.join(v[1]): v[0] for v in lines_b.values()}
        moves = [(abs(at[' '.join(v[1])] - v[0]), ' '.join(v[1])) for v in lines_a.values() if ' '.join(v[1]) in at]
        far = [m for m in moves if m[0] > 12]
        out.append('  page 1 lines: %d of %d found, %d moved over 12 pt%s' % (
            len(moves), len(lines_a), len(far), ' (%s)' % '; '.join('%s: %.0f pt' % (t[:30], d) for d, t in far[:5])
            if far else ''))
    starts = [i + 1 for i in range(min(A.page_count, B.page_count))
              if [t for t, _, _ in page_words(A[i])][:4] != [t for t, _, _ in page_words(B[i])][:4]]
    out.append('  pages that start differently: %s' % (', '.join(map(str, starts[:25])) or 'none'))
    ca = Counter(t for p in A for t, _, _ in page_words(p, False))
    cb = Counter(t for p in B for t, _, _ in page_words(p, False))
    for name, c in (('missing', ca - cb), ('added', cb - ca)):
        out.append('  words %s: %d%s' % (name, sum(c.values()), ' (%s)' % ', '.join(
            '%s x%d' % kv for kv in c.most_common(10)) if c else ''))
    if sheet:
        doc = fitz.open()
        for i in range(max(A.page_count, B.page_count)):
            ra = A[min(i, A.page_count - 1)].rect
            rb = B[min(i, B.page_count - 1)].rect
            page = doc.new_page(width=ra.width + rb.width + 24, height=max(ra.height, rb.height))
            if i < A.page_count:
                page.show_pdf_page(fitz.Rect(0, 0, ra.width, ra.height), A, i)
            if i < B.page_count:
                page.show_pdf_page(fitz.Rect(ra.width + 24, 0, ra.width + 24 + rb.width, rb.height), B, i)
        doc.save(sheet, garbage=3, deflate=True)
        out.append('  side by side (kit left, %s right): %s' % (label, sheet))
    return out

# ---- the command line


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('doc', help='a document source (.html or .md), or a built file with --built')
    ap.add_argument('-o', '--out', help='the Word file (default: <name>.docx beside the built document)')
    ap.add_argument('--built', action='store_true', help='the file is already built by tools/build.py')
    ap.add_argument('--theme')
    ap.add_argument('--palette')
    ap.add_argument('--mode')
    ap.add_argument('--check', action='store_true', help='render the Word file with LibreOffice and compare it with the kit PDF')
    ap.add_argument('--word', action='store_true', help='also have Microsoft Word save it as PDF and compare (Mac)')
    ap.add_argument('--gdocs', action='store_true', help='also print the Google Drive steps for the Google Docs check '
                    '(tools/drive_qc.py)')
    ap.add_argument('--gdocs-folder', metavar='ID', help='the id of the Drive folder kept for tests, for --gdocs')
    ap.add_argument('--gdocs-compare', metavar='PDF', help="compare Google Docs' PDF of the Word file with the kit PDF")
    a = ap.parse_args()
    src = Path(a.doc).resolve()
    if a.built:
        built = src
    else:
        if src.suffix.lower() in ('.md', '.markdown'):
            import md_to_doc
            src = Path(md_to_doc.convert(src)).resolve()
        import build
        res = build.build(src, theme_id=a.theme, palette=a.palette, mode=a.mode)
        built = Path(res['path']).resolve()
        if res.get('problems'):
            print('Build problems (above); converting anyway.', file=sys.stderr)
    out = Path(a.out).resolve() if a.out else built.with_suffix('.docx')
    check = out.with_name(out.stem + ' check')
    import doc_pdf
    ref = built.with_suffix('.pdf')

    def kit_pdf():
        if not (ref.exists() and ref.stat().st_mtime >= built.stat().st_mtime):
            doc_pdf.make_pdf(built, ref, quiet=True)
        return ref
    if a.gdocs_compare:
        import drive_qc
        drive_qc.compare(out, Path(a.gdocs_compare).resolve(), kit_pdf(), check)
        return
    data, geo, shots = layout(built, lambda: doc_pdf.pages_of(kit_pdf()))
    fonts = Fonts()
    w = Writer(data, geo, shots, fonts)
    w.write()
    title, author = data.get('title') or '', doc_pdf.meta(built.read_text(encoding='utf-8'), 'author')
    package(w, fonts, out, title, author)
    secs = sum(1 for x in w.body if isinstance(x, P) and 'sectPr' in x.ppr) + 1
    print('Word %s: %d sections, %d pictures, %d footnotes, %d lists, %.0f KB'
          % (out, secs, len(w.media), len(w.notes), len(w.nums), out.stat().st_size / 1024))
    fonts.complete()
    inside = sorted(n for n in fonts.used if n not in fonts.restricted and fonts.info.get(n) is not None)
    print('  fonts inside: %s' % (', '.join(inside) or 'none'))
    if fonts.restricted:
        print('  ! named only, their licence forbids embedding: %s' % ', '.join(sorted(fonts.restricted)))
    lacking = sorted(n for n, F in fonts.info.items() if F is None)
    if lacking:
        print('  ! not in the kit, so Word uses its own: %s' % ', '.join(lacking))
    if data.get('errors'):
        print('  ! script errors: %s' % '; '.join(str(e) for e in data['errors'][:5]))
    if a.gdocs:
        import drive_qc
        drive_qc.prepare(out, folder=a.gdocs_folder, ref=kit_pdf())
    if a.check or a.word:
        check.mkdir(exist_ok=True)
        refp = kit_pdf()
        runs = [('LibreOffice', lo_pdf)] if a.check else []
        runs += [('Word', word_pdf)] if a.word else []
        for label, fn in runs:
            pdf, err = fn(out, check)
            if pdf:
                print('\n'.join(compare(refp, pdf, label, check / ('%s side by side.pdf' % label.lower()))))
            else:
                print('%s check skipped: %s' % (label, err))


if __name__ == '__main__':
    main()
