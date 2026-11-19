#!/usr/bin/env python3
"""Check how Google Docs, Slides or Sheets shows a Word, PowerPoint or Excel file, through the Google Drive connector.

    python3 tools/drive_qc.py prepare "Report.docx" --folder <id>             # what to upload, and the connector steps
    python3 tools/drive_qc.py compare "Report.docx" google.pdf               # Google's PDF against a LibreOffice render
    python3 tools/drive_qc.py compare "Report.docx" google.pdf --ref Report.pdf   # or against the PDF it was made from

Google draws the file with its own engine and its own fonts, so a file that looks right in Word or PowerPoint can
change in Drive: fonts swapped (Google ignores fonts embedded in the file), lines taller or shorter, pages that start
elsewhere. The check uploads the full file to the Drive folder kept for tests, has Google export it as a PDF, deletes
the upload and compares that PDF with a reference: page or slide counts, how each page starts and ends, lines that
moved, the words, every font Google swapped and for what, pictures per page, and a side-by-side sheet with each pair
at its own size.

A script cannot reach the connector, so Claude runs the three connector steps that `prepare` prints (create_file,
download_file_content, trash_file), then `compare`. The full file goes up by reference, so its bytes never pass
through the chat. --slim uploads a copy instead, for when Drive refuses the full file: its pictures are shrunk to a
few pixels (each keeps its place and size on the page) and a Word file loses the fonts embedded in it, which Google
Docs never uses; the theme, styles and everything else stay, so Google lays the copy out as it would the full file.

Reading the report: a PDF drawn from HTML (a kit document) keeps charts and logos as vectors, so it counts fewer
pictures than Google's PDF of the Word file. Google's PDF draws some text, often the largest, in an unnamed Type 3
font, which hides its face: look at those words in the side-by-side sheet.

Without --folder, prepare uses the id saved in `.drive-check-folder` in the kit folder (the one holding SKILL.md),
written when the kit was first set up. Where two Claude accounts share one kit folder, pass --folder instead.
"""
import argparse
import hashlib
import io
import re
import sys
import zipfile
from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from doc_to_docx import lo_pdf, page_words  # noqa: E402

KIT_MARK = 'Decks, Docs and Graphics'
TYPES = {'.docx': ('application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Google Docs', 'page'),
         '.pptx': ('application/vnd.openxmlformats-officedocument.presentationml.presentation', 'Google Slides', 'slide'),
         '.xlsx': ('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Google Sheets', 'page')}
PICTURE = re.compile(r'\.(png|jpe?g|gif|bmp|tiff?)$', re.I)
T3 = 'unnamed Type 3'
MOVED = 12          # pt: a line that lands further than this from its place in the reference is reported


def words(n):
    return '%d word%s' % (n, '' if n == 1 else 's')


def check_dir(path):
    d = path.parent / (path.stem + ' check')
    d.mkdir(exist_ok=True)
    return d


def slim(path):
    """A copy for when Drive refuses the full file: pictures shrunk to 16 px across (place and size on the page
    unchanged) and, in a Word file, no embedded fonts. Theme, styles and the rest are kept as they are."""
    from PIL import Image
    out = check_dir(path) / (path.stem + ' (slim)' + path.suffix)
    word = path.suffix.lower() == '.docx'
    with zipfile.ZipFile(path) as zi, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            name, data = item.filename, zi.read(item.filename)
            if word and (name.startswith('word/fonts/') or name == 'word/_rels/fontTable.xml.rels'):
                continue
            if word and name == 'word/fontTable.xml':
                data = re.sub(rb'<w:embed(?:Regular|Bold|Italic|BoldItalic)\b[^>]*/>', b'', data)
            elif PICTURE.search(name) and not name.endswith('/'):
                try:
                    im = Image.open(io.BytesIO(data))
                    fmt = im.format or 'PNG'
                    im.thumbnail((16, 16))
                    buf = io.BytesIO()
                    im.save(buf, format=fmt)
                    data = buf.getvalue()
                except Exception:
                    pass
            zo.writestr(item, data)
    return out


def saved_folder():
    f = Path(__file__).resolve().parents[1] / '.drive-check-folder'
    return f.read_text(encoding='utf-8').strip() or None if f.exists() else None


def prepare(path, use_slim=False, folder=None, ref=None):
    folder = folder or saved_folder()
    mime, app, _ = TYPES[path.suffix.lower()]
    up = slim(path) if use_slim else path
    sha = hashlib.sha256(up.read_bytes()).hexdigest()
    print('Upload %s (%d KB) to %s through the Google Drive connector:' % (up.name, up.stat().st_size // 1024, app))
    print('  1. create_file  title "%s (Claude QC, delete)"' % path.stem)
    print('                  base64Content "sha256:%s:%s"' % (sha, up.resolve().as_posix()))
    print('                  contentMimeType "%s"' % mime)
    print('                  parentId "%s"' % folder if folder else
          '                  parentId: the id of the Drive folder kept for tests (never another folder)')
    print('  2. download_file_content  fileId from step 1, exportMimeType "application/pdf"; note the PDF it saves')
    print('  3. trash_file  fileId from step 1, every time, also when a step fails')
    print('  4. python3 "%s" compare "%s" "<that PDF>"%s' % (Path(__file__).resolve(), path.resolve(),
                                                            ' --ref "%s"' % Path(ref).resolve() if ref else ''))
    if not use_slim:
        print('If create_file refuses the reference, run prepare again with --slim. If it refuses that too, stop and '
              'tell the user: never paste the file into the chat.')


def spans(page):
    """(word, font) for every word on the page in reading order, the font without its subset prefix."""
    out = []
    for b in page.get_text('dict')['blocks']:
        for line in b.get('lines', []):
            for s in line['spans']:
                font = s['font'].split('+')[-1]
                out += [(w.casefold(), font, s['size']) for w in s['text'].split() if any(c.isalnum() for c in w)]
    return out


def family(font):
    if not font or font.startswith('Unnamed'):
        return T3
    f = re.sub(r'(MT|PS)$', '', font.split(',')[0])
    return re.sub(r'[-_ ]?(Regular|Roman|Book)$', '', f, flags=re.I) or font


def cap(items, n=8):
    return ', '.join(items[:n]) + (' and %d more' % (len(items) - n) if len(items) > n else '') if items else 'none'


def norm(name):
    return re.sub(r'[^a-z0-9]', '', name.lower())


def named(path):
    """The fonts the file asks for, as far as its XML shows: a Word file's run fonts and those of the styles it uses,
    a deck's slide typefaces, a workbook's cell fonts, and the theme's fonts where text inherits them."""
    z = zipfile.ZipFile(path)
    names = set(z.namelist())

    def read(n):
        return z.read(n).decode('utf-8', 'ignore') if n in names else ''
    theme = next((read(n) for n in sorted(names) if re.match(r'(word|ppt|xl)/theme/theme\d*\.xml$', n)), '')

    def theme_font(kind):
        m = re.search(r'<a:%sFont>.*?<a:latin typeface="([^"]*)"' % kind, theme, re.S)
        return m.group(1) if m else ''
    found, inherit = set(), set()

    def typeface(v):
        if v.startswith('+mj'):
            inherit.add('major')
        elif v.startswith('+mn'):
            inherit.add('minor')
        elif v:
            found.add(v)
    ext = path.suffix.lower()
    if ext == '.docx':
        text = ''.join(read(n) for n in sorted(names)
                       if re.match(r'word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$', n))
        styles = read('word/styles.xml')
        blocks, todo = {}, set(re.findall(r'<w:(?:pStyle|rStyle|tblStyle) w:val="([^"]+)"', text))
        for m in re.finditer(r'<w:style\b([^>]*)>(.*?)</w:style>', styles, re.S):
            sid = re.search(r'w:styleId="([^"]+)"', m.group(1))
            if sid:
                based = re.search(r'<w:basedOn w:val="([^"]+)"', m.group(2))
                blocks[sid.group(1)] = (based.group(1) if based else None, m.group(2))
                if 'w:default="1"' in m.group(1) and 'w:type="paragraph"' in m.group(1):
                    todo.add(sid.group(1))
        bodies, seen = [text] + re.findall(r'<w:docDefaults>.*?</w:docDefaults>', styles, re.S), set()
        while todo:
            sid = todo.pop()
            if sid in blocks and sid not in seen:
                seen.add(sid)
                bodies.append(blocks[sid][1])
                if blocks[sid][0]:
                    todo.add(blocks[sid][0])
        for b in bodies:
            for k, v in re.findall(r'w:(ascii|hAnsi|asciiTheme|hAnsiTheme)="([^"]+)"', b):
                if k.endswith('Theme'):
                    inherit.add('major' if v.startswith('major') else 'minor')
                else:
                    found.add(v)
    elif ext == '.pptx':
        bare = False
        for n in sorted(names):
            if re.match(r'ppt/slides/slide\d+\.xml$', n):
                x = read(n)
                for v in re.findall(r'<a:latin typeface="([^"]*)"', x):
                    typeface(v)
                bare = bare or any('<a:latin' not in r for r in re.findall(r'<a:r>.*?</a:r>', x, re.S))
        if bare:        # runs without a typeface of their own take the master's text styles
            for n in sorted(names):
                if re.match(r'ppt/slideMasters/slideMaster\d+\.xml$', n):
                    for v in re.findall(r'<a:latin typeface="([^"]*)"', ''.join(
                            re.findall(r'<p:txStyles>.*?</p:txStyles>', read(n), re.S))):
                        typeface(v)
    elif ext == '.xlsx':
        st = read('xl/styles.xml')
        box = re.search(r'<fonts\b[^>]*>(.*?)</fonts>', st, re.S)
        fonts = [re.search(r'<name val="([^"]+)"', f or '') for f in
                 re.findall(r'<font\b[^>]*?(?:/>|>(.*?)</font>)', box.group(1), re.S)] if box else []
        xbox = re.search(r'<cellXfs\b[^>]*>(.*?)</cellXfs>', st, re.S)
        xfs = []
        for a in re.findall(r'<xf\b([^>]*)', xbox.group(1)) if xbox else []:
            m = re.search(r'fontId="(\d+)"', a)
            xfs.append(int(m.group(1)) if m else 0)
        used = set()
        for n in sorted(names):
            if re.match(r'xl/worksheets/sheet\d+\.xml$', n):
                cells = re.findall(r'<c\b([^>]*)', read(n))
                used |= {int(m.group(1)) for m in (re.search(r'\bs="(\d+)"', c) for c in cells) if m}
                if any(not re.search(r'\bs="', c) for c in cells):
                    used.add(0)
        for i in used:
            f = fonts[xfs[i]] if i < len(xfs) and xfs[i] < len(fonts) else None
            if f:
                found.add(f.group(1))
    for kind in inherit:
        if theme_font(kind):
            found.add(theme_font(kind))
    return found


def lines(page, scale=1.0):
    """Each line's text and its top, in the reference page's points: only texts found once on the page, so a chart
    label never pairs with the same word in a table."""
    got = {}
    for t, y, key in page_words(page, False):
        got.setdefault(key, [y, []])[1].append(t)
    seen = Counter(' '.join(v[1]) for v in got.values())
    return {' '.join(v[1]): v[0] * scale for v in got.values() if seen[' '.join(v[1])] == 1}


def compare(path, google, ref=None, out=None):
    import fitz
    unit = TYPES.get(path.suffix.lower(), (None, 'Google', 'page'))[2]
    folder = Path(out) if out else check_dir(path)
    folder.mkdir(parents=True, exist_ok=True)
    ref_given = bool(ref)
    if ref:
        ref, label = Path(ref), Path(ref).name
    else:
        ref, err = lo_pdf(path, folder)
        if not ref:
            sys.exit('No reference PDF: %s. Pass --ref with a PDF of the file.' % err)
        label = 'LibreOffice'
    A, B = fitz.open(str(ref)), fitz.open(str(google))
    print('Google: %d %ss, %s %d%s' % (B.page_count, unit, label, A.page_count,
                                       '' if A.page_count == B.page_count else '  ! counts differ'))
    band = path.suffix.lower() == '.docx'       # leave out running heads and page numbers; slides and sheets have none
    starts, ends, moved, pics = [], [], [], []
    swaps, wa, wb, t3 = Counter(), Counter(), Counter(), []
    for i in range(max(A.page_count, B.page_count)):
        pa = A[i] if i < A.page_count else None
        pb = B[i] if i < B.page_count else None
        ta = [t for t, _, _ in page_words(pa, band)] if pa else []
        tb = [t for t, _, _ in page_words(pb, band)] if pb else []
        same = ta[:4] == tb[:4]
        if not same:
            starts.append('%d ("%s" / "%s")' % (i + 1, ' '.join(ta[:4]), ' '.join(tb[:4])))
        if ta[-4:] != tb[-4:]:
            ends.append('%d ("%s" / "%s")' % (i + 1, ' '.join(ta[-4:]), ' '.join(tb[-4:])))
        if pa:
            wa.update(t for t, _, _ in page_words(pa, False))
        if pb:
            wb.update(t for t, _, _ in page_words(pb, False))
            t3 += [s for _, f, s in spans(pb) if family(f) == T3]
        if pa and pb:
            sa, sb = spans(pa), spans(pb)
            sm = SequenceMatcher(None, [w for w, _, _ in sa], [w for w, _, _ in sb], autojunk=False)
            for a0, b0, n in sm.get_matching_blocks():
                for k in range(n):
                    fa, fb = family(sa[a0 + k][1]), family(sb[b0 + k][1])
                    if fa != fb:
                        swaps[(fa, fb)] += 1
            if same or unit == 'slide':      # a page that starts elsewhere has reflowed: its lines all move
                la, lb = lines(pa), lines(pb, pa.rect.height / pb.rect.height)
                common = [t for t in la if t in lb]
                far = sorted((abs(lb[t] - la[t]) for t in common if abs(lb[t] - la[t]) > MOVED), reverse=True)
                if far:
                    moved.append((far[0], '%d (%d of %d lines, up to %.0f pt)' % (i + 1, len(far), len(common), far[0])))
        na = len(pa.get_image_info()) if pa else 0
        nb = len(pb.get_image_info()) if pb else 0
        if na != nb:
            pics.append('%d (%d / %d)' % (i + 1, na, nb))
    print('  %ss that start differently: %s' % (unit, cap(starts)))
    print('  %ss that end differently: %s' % (unit, cap(ends)))
    print('  lines that moved over %d pt, on %ss that start alike (worst first): %s'
          % (MOVED, unit, cap([m for _, m in sorted(moved, key=lambda m: -m[0])])))
    fa = Counter(family(f) for p in A for _, f, _ in spans(p))
    fb = Counter(family(f) for p in B for _, f, _ in spans(p))
    want = sorted(named(path), key=str.lower)
    print('  fonts the file names: %s' % (', '.join(want) or 'none found'))
    for side, fonts in (("Google's PDF", fb), (label, fa)):
        lack = [n for n in want if not any(norm(f).startswith(norm(n)) for f in fonts if f != T3)]
        if lack:
            print('  ! %s never shows %s%s' % (side, ', '.join(lack), '' if side != label or ref_given else
                                              ': LibreOffice may lack them, so for a true comparison pass --ref with '
                                              'a PDF from the app that made the file'))
    print('  fonts, %s: %s' % (label, ', '.join('%s (%s)' % (f, words(n)) for f, n in fa.most_common())))
    print('  fonts, Google: %s' % ', '.join('%s (%s)' % (f, words(n)) for f, n in fb.most_common()))
    print('  fonts Google swapped, word by word against %s: %s' % (label, ', '.join('%s for %s (%s)' % (g, f, words(n))
                                                    for (f, g), n in swaps.most_common()) or 'none'))
    if t3:
        print('  ! %s in an %s font (%g to %g pt): its face shows only in the side-by-side sheet'
              % (words(len(t3)), T3, min(t3), max(t3)))
    print('  pictures that differ (%s / Google): %s' % (label, ', '.join(pics) or 'none'))
    miss, add = wa - wb, wb - wa
    for name, c in (('missing', miss), ('added', add)):
        n = sum(c.values())
        print('  words %s: %d%s' % (name, n, (' (' + ', '.join('%s x%d' % kv for kv in c.most_common(12)) + ')') if n else ''))
    sheet = folder / 'google side by side.pdf'
    S = fitz.open()
    gap = 18
    for i in range(max(A.page_count, B.page_count)):
        ra = A[i].rect if i < A.page_count else fitz.Rect(0, 0, 0, 0)
        rb = B[i].rect if i < B.page_count else fitz.Rect(0, 0, 0, 0)
        pg = S.new_page(width=ra.width + gap + rb.width, height=max(ra.height, rb.height))
        if i < A.page_count:
            pg.show_pdf_page(fitz.Rect(0, 0, ra.width, ra.height), A, i)
        if i < B.page_count:
            pg.show_pdf_page(fitz.Rect(ra.width + gap, 0, ra.width + gap + rb.width, rb.height), B, i)
    S.set_metadata({'creator': KIT_MARK, 'producer': KIT_MARK})
    S.save(str(sheet), garbage=3, deflate=True)
    print('  side by side (%s left, Google right, each pair at its own size): %s' % (label, sheet))


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n\n')[0])
    sub = ap.add_subparsers(dest='cmd', required=True)
    p = sub.add_parser('prepare', help='print what to upload and the connector steps')
    p.add_argument('file')
    p.add_argument('--slim', action='store_true', help='upload a copy with tiny pictures and no embedded fonts')
    p.add_argument('--folder', help='the id of the Drive folder kept for tests, printed into the steps')
    p.add_argument('--ref', help='a PDF of the file to compare against, printed into the last step')
    c = sub.add_parser('compare', help="compare Google's PDF export with a reference PDF")
    c.add_argument('file')
    c.add_argument('google_pdf')
    c.add_argument('--ref', help='reference PDF (default: a LibreOffice render of the file)')
    c.add_argument('-o', '--out', help="folder for the reports (default: '<name> check' beside the file)")
    a = ap.parse_args()
    path = Path(a.file).expanduser()
    if not path.exists():
        sys.exit('No such file: %s' % path)
    if path.suffix.lower() not in TYPES:
        sys.exit('Word (.docx), PowerPoint (.pptx) or Excel (.xlsx) only')
    if a.cmd == 'prepare':
        prepare(path, a.slim, a.folder, a.ref)
    else:
        compare(path, Path(a.google_pdf).expanduser(), a.ref, a.out)


if __name__ == '__main__':
    main()
