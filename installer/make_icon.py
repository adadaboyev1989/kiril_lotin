"""installer/KirillLotin.ico ni yasaydi (Pillow kerak): pip install pillow"""
from PIL import Image, ImageDraw, ImageFont

FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
S, M, R = 256, 8, 44

bg = Image.new("RGBA", (S, S), (30, 94, 170, 255))
ImageDraw.Draw(bg).polygon([(S, 0), (S, S), (0, S)], fill=(22, 140, 90, 255))
d = ImageDraw.Draw(bg)
f = ImageFont.truetype(FONT, 104)
d.text((84, 84), "Ж", font=f, fill="white", anchor="mm")   # kirill Zh
d.text((174, 172), "J", font=f, fill="white", anchor="mm")

mask = Image.new("L", (S, S), 0)
ImageDraw.Draw(mask).rounded_rectangle((M, M, S - M, S - M), radius=R, fill=255)
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
img.paste(bg, (0, 0), mask)

img.save("KirillLotin.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
