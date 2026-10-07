"""Builds the test presentations for the PowerPoint to PDF checks.

Needs python-pptx, Pillow and ffmpeg. msoffcrypto-tool and LibreOffice are
optional. Without them the password protected deck and the legacy .ppt are
left as they are in the repository.

    python3 tests/pptpdf/make_decks.py            writes into tests/pptpdf/decks
    python3 tests/pptpdf/make_decks.py some/dir   writes somewhere else

The decks are small on purpose. Between them they hold custom fonts, hidden
slides, video with and without a usable poster frame, speaker notes, mixed
slide sizes, complex graphics, animations, transitions, a link to a picture
on the internet and a slide number field.
"""
import io
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Pt

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), 'decks')
os.makedirs(OUT, exist_ok=True)
TMP = tempfile.mkdtemp(prefix='slidesize-decks-')

NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def picture(path, size, kind):
    w, h = size
    im = Image.new('RGB', size, (18, 18, 18))
    d = ImageDraw.Draw(im)
    if kind == 'poster':
        for i in range(0, w, 8):
            d.line([(i, 0), (w - i, h)], fill=(40 + i % 200, 90, 160), width=3)
        d.ellipse([w * 0.3, h * 0.2, w * 0.7, h * 0.8], outline=(240, 240, 240), width=6)
    elif kind == 'photo':
        for y in range(h):
            d.line([(0, y), (w, y)], fill=(int(255 * y / h), 80, int(255 * (1 - y / h))))
        d.rectangle([w * 0.1, h * 0.1, w * 0.45, h * 0.55], outline=(255, 255, 255), width=5)
    elif kind == 'black':
        pass
    im.save(path)
    return path


def video(path):
    subprocess.run(['ffmpeg', '-y', '-loglevel', 'error', '-f', 'lavfi', '-i', 'testsrc=size=320x180:rate=25:duration=1',
                    '-pix_fmt', 'yuv420p', '-c:v', 'mpeg4', '-q:v', '12', path], check=True)
    return path


def text(slide, x, y, w, h, value, font=None, size=24, colour=(20, 20, 20)):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = value
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(*colour)
    if font:
        run.font.name = font
    return box


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def slide_number_field(slide):
    box = slide.shapes.add_textbox(Inches(0.3), Inches(6.9), Inches(1), Inches(0.4))
    p = box.text_frame.paragraphs[0]._p
    fld = p.makeelement('{%s}fld' % NS_A, {'id': '{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}', 'type': 'slidenum'})
    t = fld.makeelement('{%s}t' % NS_A, {})
    t.text = '‹#›'
    fld.append(t)
    p.append(fld)


def timing(slide, xml):
    from lxml import etree
    slide._element.append(etree.fromstring(xml))


def effect(cls, preset, spid):
    return ('<p:par><p:cTn id="%d" presetID="%d" presetClass="%s" presetSubtype="0" fill="hold" nodeType="clickEffect">'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold"/>'
            '<p:tgtEl><p:spTgt spid="%s"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
            '</p:cBhvr><p:to><p:strVal val="%s"/></p:to></p:set></p:childTnLst></p:cTn></p:par>'
            % (100 + int(spid) * 3 + (0 if cls == 'entr' else 1), preset, cls, 200 + int(spid) * 3 + (0 if cls == 'entr' else 1),
               spid, 'visible' if cls == 'entr' else 'hidden'))


def kitchen():
    """16:9, seven slides, most of what a PDF conversion has to cope with."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    poster = picture(os.path.join(TMP, 'poster.png'), (640, 360), 'poster')
    black = picture(os.path.join(TMP, 'black.png'), (640, 360), 'black')
    photo = picture(os.path.join(TMP, 'photo.jpg'), (1200, 800), 'photo')
    clip = video(os.path.join(TMP, 'clip.mp4'))

    s = blank(prs)                                               # 1 title, custom fonts
    text(s, 0.8, 1.0, 11, 1.5, 'Annual Conference 2026', font='Gotham Light', size=54)
    text(s, 0.8, 2.8, 11, 1.0, 'Opening plenary', font='Montserrat', size=28)
    text(s, 0.8, 4.2, 11, 1.0, 'Body text in an everyday font', font='Carlito', size=20)
    slide_number_field(s)

    s = blank(prs)                                               # 2 notes and a transition
    text(s, 0.8, 0.8, 11, 1.2, 'Agenda', font='Gotham Light', size=44)
    text(s, 0.8, 2.2, 11, 3.0, 'Welcome\nKeynote\nPanel\nClose', font='Carlito', size=24)
    link = text(s, 0.8, 5.6, 11, 0.6, 'slidesize.com', font='Carlito', size=18)
    link.text_frame.paragraphs[0].runs[0].hyperlink.address = 'https://slidesize.com/'
    s.notes_slide.notes_text_frame.text = 'Walk on from stage left. Pause for the sting. Introduce the first speaker by name.'
    from lxml import etree
    s._element.append(etree.fromstring('<p:transition xmlns:p="%s" spd="med"><p:fade/></p:transition>' % NS_P))
    slide_number_field(s)

    s = blank(prs)                                               # 3 hidden
    text(s, 0.8, 0.8, 11, 1.2, 'Backup slide, hidden', font='Carlito', size=40)
    s._element.set('show', '0')
    slide_number_field(s)

    s = blank(prs)                                               # 4 video with a real poster frame
    text(s, 0.8, 0.5, 11, 1.0, 'Opening film', font='Carlito', size=36)
    s.shapes.add_movie(clip, Inches(2.5), Inches(1.7), Inches(8), Inches(4.5), poster_frame_image=poster, mime_type='video/mp4')
    slide_number_field(s)

    s = blank(prs)                                               # 5 video whose poster frame is black
    text(s, 0.8, 0.5, 11, 1.0, 'Sting', font='Carlito', size=36)
    s.shapes.add_movie(clip, Inches(2.5), Inches(1.7), Inches(8), Inches(4.5), poster_frame_image=black, mime_type='video/mp4')

    s = blank(prs)                                               # 6 stacked shapes that swap on a click
    text(s, 0.8, 0.5, 11, 1.0, 'Before and after', font='Carlito', size=36)
    a = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3), Inches(2), Inches(6), Inches(3.5))
    a.text_frame.text = 'BEFORE'
    b = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.2), Inches(2.2), Inches(6), Inches(3.5))
    b.text_frame.text = 'AFTER'
    b.fill.solid()
    b.fill.fore_color.rgb = RGBColor(200, 40, 40)
    timing(s, '<p:timing xmlns:p="%s"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
              '<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
              '%s%s</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
              % (NS_P, effect('exit', 1, str(a.shape_id)), effect('entr', 1, str(b.shape_id))))

    s = blank(prs)                                               # 7 complex graphics and a linked web picture
    text(s, 0.8, 0.4, 11, 1.0, 'Graphics', font='Carlito', size=36)
    s.shapes.add_picture(photo, Inches(0.8), Inches(1.5), height=Inches(4))
    for i, shape in enumerate([MSO_SHAPE.STAR_5_POINT, MSO_SHAPE.CHEVRON, MSO_SHAPE.CLOUD, MSO_SHAPE.DONUT]):
        sh = s.shapes.add_shape(shape, Inches(7.4 + (i % 2) * 2.7), Inches(1.5 + (i // 2) * 2.2), Inches(2.3), Inches(1.9))
        sh.fill.gradient()
        sh.fill.gradient_angle = 45 * i
        sh.rotation = 7 * i
        sh.shadow.inherit = False
    linked = s.shapes.add_picture(photo, Inches(0.8), Inches(5.8), height=Inches(1.2))
    rid = s.part.relate_to('https://example.invalid/logo.png',
                           'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', is_external=True)
    blip = linked._element.xpath('.//a:blip')[0]
    blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link', rid)
    slide_number_field(s)

    prs.save(os.path.join(OUT, 'Kitchen sink 16x9.pptx'))


def plain(name, width, height, slides, hidden=()):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(width), Emu(height)
    wi, hi = width / 914400.0, height / 914400.0
    for i in range(slides):
        s = blank(prs)
        text(s, wi * 0.08, hi * 0.1, wi * 0.84, hi * 0.2, '%s, slide %d' % (name, i + 1), font='Carlito', size=30)
        sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(wi * 0.1), Inches(hi * 0.4), Inches(wi * 0.3), Inches(hi * 0.45))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor(31, 182, 166)
        sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(wi * 0.55), Inches(hi * 0.4), Inches(wi * 0.35), Inches(hi * 0.45))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor(231, 48, 42)
        if (i + 1) in hidden:
            s._element.set('show', '0')
    path = os.path.join(OUT, name + '.pptx')
    prs.save(path)
    return path


def main():
    kitchen()
    plain('Classic 4x3', 9144000, 6858000, 3)
    plain('Poster A4 portrait', 7560000, 10692000, 2)
    plain('Wide blend 32x9', 24384000, 6858000, 2, hidden=(2,))
    # the same file name in a second folder, to exercise duplicate output names
    os.makedirs(os.path.join(OUT, 'second'), exist_ok=True)
    plain('second/Classic 4x3', 9144000, 6858000, 1)

    base = plain('Password protected', 12192000, 6858000, 2)
    try:
        import msoffcrypto
        from msoffcrypto.format.ooxml import OOXMLFile
        with open(base, 'rb') as f:
            enc = io.BytesIO()
            OOXMLFile(f).encrypt('slidesize', enc)
        with open(base, 'wb') as f:
            f.write(enc.getvalue())
    except Exception as e:                                       # pragma: no cover
        print('password protected deck not rebuilt:', e)
        os.remove(base)

    whole = open(os.path.join(OUT, 'Classic 4x3.pptx'), 'rb').read()
    open(os.path.join(OUT, 'Cut short.pptx'), 'wb').write(whole[:len(whole) * 2 // 3])
    open(os.path.join(OUT, 'Not a deck.pptx'), 'wb').write(b'This is a text file with the wrong extension.\n' * 20)
    with zipfile.ZipFile(os.path.join(OUT, 'A zip, not a deck.pptx'), 'w') as z:
        z.writestr('readme.txt', 'nothing to see')
    open(os.path.join(OUT, 'Empty.pptx'), 'wb').close()

    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if soffice:
        subprocess.run([soffice, '--headless', '--convert-to', 'ppt', '--outdir', TMP, os.path.join(OUT, 'Classic 4x3.pptx')],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        shutil.copy(os.path.join(TMP, 'Classic 4x3.ppt'), os.path.join(OUT, 'Legacy 97-2003.ppt'))
        # PDFs for the checks of the PDF reader and editor. Made by LibreOffice, so they stand in for
        # PowerPoint's output only as far as being real PDFs with real fonts, tags and links.
        pdfs = os.path.join(OUT, '..', 'pdfs')
        os.makedirs(pdfs, exist_ok=True)
        for src, opt, dst in [('Classic 4x3.pptx', '{}', 'classic.pdf'),
                              ('Kitchen sink 16x9.pptx', '{"UseTaggedPDF":{"type":"boolean","value":"true"}}', 'kitchen.pdf'),
                              ('Classic 4x3.pptx', '{"SelectPdfVersion":{"type":"long","value":"1"}}', 'classic-pdfa.pdf')]:
            subprocess.run([soffice, '--headless', '--convert-to', 'pdf:impress_pdf_Export:' + opt, '--outdir', TMP, os.path.join(OUT, src)],
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            shutil.copy(os.path.join(TMP, os.path.splitext(src)[0] + '.pdf'), os.path.join(pdfs, dst))
    shutil.rmtree(TMP, ignore_errors=True)
    for n in sorted(os.listdir(OUT)):
        p = os.path.join(OUT, n)
        if os.path.isfile(p):
            print('%8d  %s' % (os.path.getsize(p), n))


if __name__ == '__main__':
    main()
