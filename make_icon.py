"""Рисует иконку Video2Doc: видеокамера с объективом на синем скруглённом квадрате.

    python make_icon.py <путь к .ico>
"""
import sys

from PIL import Image, ImageDraw

S = 1024                      # рисуем крупно, потом уменьшаем — сглаживание
ACCENT, DARK, WHITE = (79, 140, 255), (18, 19, 22), (255, 255, 255)

img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)
d.rounded_rectangle((40, 40, S - 40, S - 40), radius=210, fill=ACCENT)

# корпус камеры
d.rounded_rectangle((150, 330, 690, 730), radius=80, fill=WHITE)
# видоискатель-«рупор» справа
d.polygon([(700, 470), (880, 360), (880, 700), (700, 590)], fill=WHITE)
# объектив: кольца
cx, cy = 420, 530
for r, color in ((150, DARK), (118, ACCENT), (70, DARK), (34, WHITE)):
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=color)
# блик и индикатор записи
d.ellipse((cx - 58, cy - 62, cx - 30, cy - 34), fill=(255, 255, 255, 200))
d.ellipse((585, 380, 635, 430), fill=(224, 85, 79))
# катушки сверху — узнаваемый силуэт кинокамеры
for x in (270, 500):
    d.ellipse((x - 95, 120, x + 95, 310), fill=WHITE)
    d.ellipse((x - 35, 180, x + 35, 250), fill=ACCENT)

img.save(sys.argv[1], sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
img.resize((256, 256), Image.LANCZOS).save(sys.argv[1].replace(".ico", "_preview.png"))
