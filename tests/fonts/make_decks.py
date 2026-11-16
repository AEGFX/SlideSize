"""Draws the test decks for Font Extractor.

    python3 tests/fonts/make_decks.py OUT_DIR

Writes
    full.pptx      two whole fonts, one plain EOT and one XOR-scrambled EOT
    subset.pptx    regular  cut down to a few characters, flagged in the EOT header
                   bold     every glyph slot kept but unused outlines emptied, no flag
                   italic   a whole font with its lowercase taken away
                   bold italic  points at the same file as regular
                   plus a raw .ttf part the font list does not mention
    none.pptx      no embedded fonts

PowerPoint's own fonts are MTX compressed, which nothing here can write, so
that path is checked by hand against decks saved from PowerPoint.

Needs python-pptx and fontTools, and DejaVu Sans on the machine.
"""
import io, os, struct, sys, zipfile
from pptx import Presentation
from pptx.util import Cm
from fontTools.ttLib import TTFont
from fontTools import subset as ftsubset
from fontTools.ttLib.tables._g_l_y_f import Glyph
import logging; logging.getLogger('fontTools').setLevel(logging.ERROR)

out = sys.argv[1] if len(sys.argv) > 1 else 'decks'
os.makedirs(out, exist_ok=True)

def find(name):
    for root in ('/usr/share/fonts', '/Library/Fonts', '/System/Library/Fonts', os.path.expanduser('~/Library/Fonts'), 'C:/Windows/Fonts'):
        for d, _, files in os.walk(root):
            if name in files:
                return os.path.join(d, name)
    raise SystemExit('Cannot find ' + name)

REG, BOLD, OBL = find('DejaVuSans.ttf'), find('DejaVuSans-Bold.ttf'), find('DejaVuSans-Oblique.ttf')
KEEP = 'Helo Wrd 26'


def dump(font):
    b = io.BytesIO(); font.save(b); return b.getvalue()


def cut(path):
    """A true subset: only the glyphs for KEEP survive."""
    opts = ftsubset.Options(); opts.notdef_outline = True; opts.name_IDs = ['*']
    f = TTFont(path); s = ftsubset.Subsetter(opts); s.populate(text=KEEP); s.subset(f)
    return dump(f)


def hollow(path):
    """PowerPoint style: glyph count and cmap untouched, unused outlines emptied."""
    f = TTFont(path); cmap = f.getBestCmap(); glyf = f['glyf']
    keep = {cmap[ord(c)] for c in KEEP if ord(c) in cmap} | {'.notdef'}
    more = set()
    for g in keep:
        if glyf[g].isComposite():
            more.update(c.glyphName for c in glyf[g].components)
    keep |= more
    for g in f.getGlyphOrder():
        if g not in keep:
            glyf[g] = Glyph()
    return dump(f)


def no_lowercase(path):
    f = TTFont(path)
    for t in f['cmap'].tables:
        for cp in range(ord('a'), ord('z') + 1):
            t.cmap.pop(cp, None)
    return dump(f)


def eot(ttf, subset=False, xor=False):
    f = TTFont(io.BytesIO(ttf)); name = f['name']; os2 = f['OS/2']; head = f['head']
    def s(nid):
        return (name.getDebugName(nid) or '').encode('utf-16-le')
    data = bytes(b ^ 0x50 for b in ttf) if xor else ttf
    flags = (1 if subset else 0) | (0x10000000 if xor else 0)
    names = b''
    for i, nid in enumerate((1, 2, 5, 4)):
        v = s(nid)
        names += (b'\0\0' if i else b'') + struct.pack('<H', len(v)) + v
    names += b'\0\0' + struct.pack('<H', 0)          # no root string
    fixed = struct.pack('<IIII', 0, len(data), 0x00020001, flags)
    fixed += bytes(10)                               # PANOSE
    fixed += struct.pack('<BBIHH', 1, 1 if os2.fsSelection & 1 else 0, os2.usWeightClass, os2.fsType, 0x504C)
    fixed += struct.pack('<IIII', os2.ulUnicodeRange1, os2.ulUnicodeRange2, os2.ulUnicodeRange3, os2.ulUnicodeRange4)
    fixed += struct.pack('<II', getattr(os2, 'ulCodePageRange1', 0), getattr(os2, 'ulCodePageRange2', 0))
    fixed += struct.pack('<I', head.checkSumAdjustment) + bytes(16) + b'\0\0'
    blob = fixed + names + data
    return struct.pack('<I', len(blob)) + blob[4:]


def deck(path, typeface, slots, extra=None):
    """slots: list of (slot, part name, bytes). The same part name may repeat."""
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = 'Font Extractor check'
    tb = s.shapes.add_textbox(Cm(2), Cm(5), Cm(20), Cm(3))
    tb.text_frame.text = 'Hello World 2026'
    b = io.BytesIO(); prs.save(b)
    zin = zipfile.ZipFile(io.BytesIO(b.getvalue()))
    zout = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)
    parts, rels, lst = {}, '', ''
    for i, (slot, part, data) in enumerate(slots):
        parts[part] = data
        rels += '<Relationship Id="rIdF%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/%s"/>' % (i, part)
        lst += '<p:%s r:id="rIdF%d"/>' % (slot, i)
    for n in zin.namelist():
        d = zin.read(n)
        if slots and n == 'ppt/presentation.xml':
            x = d.decode('utf-8')
            x = x.replace('<p:presentation ', '<p:presentation embedTrueTypeFonts="1" ', 1)
            block = '<p:embeddedFontLst><p:embeddedFont><p:font typeface="%s"/>%s</p:embeddedFont></p:embeddedFontLst>' % (typeface, lst)
            x = x.replace('<p:defaultTextStyle', block + '<p:defaultTextStyle', 1)
            assert 'embeddedFontLst' in x
            d = x.encode('utf-8')
        elif slots and n == 'ppt/_rels/presentation.xml.rels':
            d = d.decode('utf-8').replace('</Relationships>', rels + '</Relationships>').encode('utf-8')
        elif slots and n == '[Content_Types].xml':
            d = d.decode('utf-8').replace('<Default ', '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1).encode('utf-8')
        zout.writestr(n, d)
    for part, data in parts.items():
        zout.writestr('ppt/fonts/' + part, data)
    for part, data in (extra or {}).items():
        zout.writestr('ppt/fonts/' + part, data)
    zout.close()


reg, bold = open(REG, 'rb').read(), open(BOLD, 'rb').read()
deck(os.path.join(out, 'full.pptx'), 'DejaVu Sans',
     [('regular', 'font1.fntdata', eot(reg)), ('bold', 'font2.fntdata', eot(bold, xor=True))])
deck(os.path.join(out, 'subset.pptx'), 'DejaVu Sans',
     [('regular', 'font1.fntdata', eot(cut(REG), subset=True)),
      ('bold', 'font2.fntdata', eot(hollow(BOLD))),
      ('italic', 'font3.fntdata', eot(no_lowercase(OBL))),
      ('boldItalic', 'font1.fntdata', eot(cut(REG), subset=True))],
     extra={'stray.fntdata': reg})
deck(os.path.join(out, 'none.pptx'), '', [])
print('decks in', out)
