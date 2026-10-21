"""Draws the test decks for Deck Compare.

    python3 tests/compare/make_decks.py OUT_DIR

Writes
    v1.pptx           the starting deck
    v2.pptx           v1 with a known set of edits (see EXPECT below)
    v2-pasted.pptx    v2 where one slide has a new slide ID, as if it was
                      copy-pasted back in from another deck
    other.pptx        an unrelated deck whose slide IDs collide with v1's

Needs python-pptx and Pillow.
"""
import io, os, re, sys, zipfile
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from PIL import Image

out = sys.argv[1] if len(sys.argv) > 1 else 'decks'
os.makedirs(out, exist_ok=True)


def png(colour, size=(320, 180)):
    b = io.BytesIO()
    Image.new('RGB', size, colour).save(b, 'PNG')
    b.seek(0)
    return b


def mp4_bytes():
    # Any bytes will do, the engine only hashes the file
    return io.BytesIO(b'\x00\x00\x00\x18ftypmp42' + os.urandom(64))


VIDEO_A = mp4_bytes()
VIDEO_B = mp4_bytes()


def build(version):
    p = Presentation()
    p.slide_width, p.slide_height = Cm(33.867), Cm(19.05)
    L = p.slide_layouts

    # 1 title
    s = p.slides.add_slide(L[0])
    s.shapes.title.text = 'Annual Kick-off 2026'
    s.placeholders[1].text = 'London, 14 January' if version == 1 else 'London, 15 January'

    # 2 agenda with a typo fixed in v2
    s = p.slides.add_slide(L[1])
    s.shapes.title.text = 'Agenda'
    body = s.placeholders[1].text_frame
    items = ['Welcome and introductions', 'Results for the yaer', 'Product roadmap', 'Questions'] if version == 1 else \
            ['Welcome and introductions', 'Results for the year', 'Product roadmap', 'Questions']
    body.text = items[0]
    for t in items[1:]:
        body.add_paragraph().text = t

    # 3 picture slide, picture swapped and moved in v2
    s = p.slides.add_slide(L[5])
    s.shapes.title.text = 'Our new office'
    left = Cm(4) if version == 1 else Cm(6)
    s.shapes.add_picture(png((200, 30, 30) if version == 1 else (30, 30, 200)), left, Cm(5), Cm(16), Cm(9))

    # 4 video slide
    s = p.slides.add_slide(L[5])
    s.shapes.title.text = 'Sizzle reel'
    VIDEO_A.seek(0); VIDEO_B.seek(0)
    s.shapes.add_movie(VIDEO_A if version == 1 else VIDEO_B, Cm(4), Cm(4), Cm(24), Cm(13.5),
                       poster_frame_image=png((0, 0, 0)), mime_type='video/mp4')

    # 5 table, one figure changed in v2
    s = p.slides.add_slide(L[5])
    s.shapes.title.text = 'Numbers'
    t = s.shapes.add_table(3, 2, Cm(4), Cm(5), Cm(20), Cm(6)).table
    vals = [['Region', 'Revenue'], ['EMEA', '12.4m'], ['APAC', '8.1m' if version == 1 else '8.3m']]
    for r, row in enumerate(vals):
        for c, v in enumerate(row):
            t.cell(r, c).text = v

    # 6 speaker slide, only in v1 (deleted in v2)
    if version == 1:
        s = p.slides.add_slide(L[1])
        s.shapes.title.text = 'Guest speaker'
        s.placeholders[1].text = 'Dr Jane Smith, Head of Research'

    # 7 closing, font size bumped and notes added in v2
    s = p.slides.add_slide(L[1])
    s.shapes.title.text = 'Thank you'
    tf = s.placeholders[1].text_frame
    tf.text = 'See you at the drinks reception'
    if version == 2:
        tf.paragraphs[0].runs[0].font.size = Pt(36)
        s.notes_slide.notes_text_frame.text = 'Remind everyone about the cloakroom.'

    # 8 new slide in v2
    if version == 2:
        s = p.slides.add_slide(L[1])
        s.shapes.title.text = 'Housekeeping'
        s.placeholders[1].text = 'Fire exits are at the back of the room'

    return p


def save(p, name):
    path = os.path.join(out, name)
    p.save(path)
    return path


v1 = save(build(1), 'v1.pptx')
p2 = build(2)
# v2 moves the closing slide in front of the table slide and hides the agenda
lst = p2.slides._sldIdLst
ids = list(lst)
# order now: title, agenda, picture, video, table, thanks, housekeeping
lst.remove(ids[5]); lst.insert(4, ids[5])
p2.slides[1]._element.set('show', '0')
# python-pptx numbers slide ids from 256 in add order. Give v2 the same
# ids as v1 for the slides they share, as if v2 was saved from v1.
id_map = {'Annual Kick-off 2026': 256, 'Agenda': 257, 'Our new office': 258, 'Sizzle reel': 259,
          'Numbers': 260, 'Thank you': 262, 'Housekeeping': 263}
for sld_id, slide in zip(list(lst), p2.slides):
    sld_id.set('id', str(id_map[slide.shapes.title.text]))
v2 = save(p2, 'v2.pptx')

# v2-pasted: the picture slide gets a brand new id
with zipfile.ZipFile(v2) as zi, zipfile.ZipFile(os.path.join(out, 'v2-pasted.pptx'), 'w', zipfile.ZIP_DEFLATED) as zo:
    for item in zi.infolist():
        data = zi.read(item.filename)
        if item.filename == 'ppt/presentation.xml':
            data = re.sub(rb'id="258"', b'id="400"', data)
        zo.writestr(item, data)

# an unrelated deck, ids 256.. collide with v1
o = Presentation()
for t in ['Quarterly safety briefing', 'Ladders and lifting', 'Hi-vis policy', 'Incident log', 'Contacts']:
    s = o.slides.add_slide(o.slide_layouts[1])
    s.shapes.title.text = t
    s.placeholders[1].text = 'Site rules for ' + t.lower()
save(o, 'other.pptx')
print('decks written to', out)
