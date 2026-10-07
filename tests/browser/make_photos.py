import sys, os, math, random
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import numpy as np
out = sys.argv[1]
os.makedirs(out, exist_ok=True)
random.seed(7)
def font(sz):
    for p in ['/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf','/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf']:
        if os.path.exists(p): return ImageFont.truetype(p, sz)
    return ImageFont.load_default()
def photo(w, h, label, hue):
    # smooth "scenery" from low-frequency noise so motion and fades have real detail to chew on
    rng = np.random.default_rng(hash(label) % (2**32))
    small = rng.random((max(4, h // 160), max(4, w // 160), 3))
    img = Image.fromarray((small * 255).astype('uint8')).resize((w, h), Image.BICUBIC)
    a = np.asarray(img).astype('float32') / 255
    yy, xx = np.mgrid[0:h, 0:w]
    grad = (yy / h)[..., None]
    tint = np.array([0.5 + 0.5 * math.cos(hue), 0.5 + 0.5 * math.cos(hue + 2.1), 0.5 + 0.5 * math.cos(hue + 4.2)])
    a = 0.55 * a + 0.45 * (tint * (1 - 0.6 * grad))
    a += rng.normal(0, 0.02, a.shape)       # a little grain
    img = Image.fromarray((np.clip(a, 0, 1) * 255).astype('uint8'))
    d = ImageDraw.Draw(img)
    m = min(w, h)
    # border and thirds so cropping is easy to read
    d.rectangle([4, 4, w - 5, h - 5], outline=(255, 255, 255), width=max(4, m // 200))
    for i in (1, 2):
        d.line([w * i // 3, 0, w * i // 3, h], fill=(255, 255, 255, 80), width=max(1, m // 600))
        d.line([0, h * i // 3, w, h * i // 3], fill=(255, 255, 255, 80), width=max(1, m // 600))
    # a "face" at the default focus area
    cx, cy, r = w // 2, int(h * 0.45), m // 8
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(245, 222, 190), outline=(40, 30, 20), width=max(3, m // 300))
    d.ellipse([cx - r // 2, cy - r // 4, cx - r // 4, cy], fill=(40, 30, 20)); d.ellipse([cx + r // 4, cy - r // 4, cx + r // 2, cy], fill=(40, 30, 20))
    d.arc([cx - r // 2, cy, cx + r // 2, cy + r // 2], 0, 180, fill=(40, 30, 20), width=max(3, m // 300))
    f = font(m // 9)
    d.text((w // 2, h * 0.8), label, font=f, fill=(255, 255, 255), anchor='mm', stroke_width=max(2, m // 300), stroke_fill=(0, 0, 0))
    f2 = font(m // 22)
    d.text((w // 2, h * 0.9), f'{w} x {h}', font=f2, fill=(255, 255, 255), anchor='mm', stroke_width=2, stroke_fill=(0, 0, 0))
    for (x, y, t) in [(60, 40, 'TL'), (w - 60, 40, 'TR'), (60, h - 40, 'BL'), (w - 60, h - 40, 'BR')]:
        d.text((x, y), t, font=f2, fill=(255, 255, 0), anchor='mm', stroke_width=2, stroke_fill=(0, 0, 0))
    return img
sizes = [('land32', 3000, 2000), ('land43', 2800, 2100), ('land169', 3200, 1800), ('port23', 2000, 3000), ('port34', 2100, 2800),
         ('square', 2400, 2400), ('pano', 4200, 1400), ('land54', 2500, 2000), ('small', 800, 533)]
n = int(sys.argv[2]) if len(sys.argv) > 2 else 12
k = 0
for i in range(n):
    name, w, h = sizes[i % len(sizes)]
    k += 1
    photo(w, h, f'{k:02d} {name}', i * 0.7).save(os.path.join(out, f'IMG_{k:03d}_{name}.jpg'), quality=88)
if len(sys.argv) > 3 and sys.argv[3] == 'extras':
    # EXIF orientation 6: stored landscape, must display portrait
    im = photo(3000, 2000, 'EXIF rot6', 1.0)
    ex = im.getexif(); ex[0x0112] = 6
    im.save(os.path.join(out, 'exif_rot6_stored3000x2000.jpg'), quality=88, exif=ex)
    photo(1600, 1200, 'WEBP', 2.0).save(os.path.join(out, 'pic.webp'), quality=85)
    p = photo(1600, 1067, 'PNG alpha', 3.0).convert('RGBA')
    al = Image.new('L', p.size, 0); ImageDraw.Draw(al).ellipse([100, 60, 1500, 1007], fill=255); p.putalpha(al)
    p.save(os.path.join(out, 'alpha.png'))
    open(os.path.join(out, 'broken.jpg'), 'wb').write(b'\xff\xd8\xff\xe0 this is not really a jpeg' * 20)
    open(os.path.join(out, 'notes.txt'), 'w').write('not an image')
    open(os.path.join(out, 'phone.heic'), 'wb').write(os.urandom(2000))
print(len(os.listdir(out)), 'files')
