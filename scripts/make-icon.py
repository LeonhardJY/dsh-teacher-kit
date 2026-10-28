
from PIL import Image, ImageDraw, ImageFont
import base64, io, os

S = 512
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# 圆角底：深蓝渐变
bg = Image.new("RGB", (S, S), (22, 33, 58))
for y in range(S):
    t = y / (S - 1)
    r = int(22 + (28 - 22) * t)
    g = int(33 + (56 - 33) * t)
    b = int(58 + (96 - 58) * t)
    d.line([(0, y), (S, y)], fill=(r, g, b))

mask = Image.new("L", (S, S), 0)
ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], radius=112, fill=255)
out = Image.new("RGBA", (S, S), (0, 0, 0, 0))
out.paste(img, (0, 0), mask)

d = ImageDraw.Draw(out)
# 顶部一条强调色（粉笔黄）
d.rounded_rectangle([S*0.30, S*0.155, S*0.70, S*0.185], radius=8, fill=(245, 196, 82, 255))

font = ImageFont.truetype("C:/Windows/Fonts/msyhbd.ttc", 268)
text = "教"
bb = d.textbbox((0, 0), text, font=font)
tw, th = bb[2] - bb[0], bb[3] - bb[1]
d.text(((S - tw) / 2 - bb[0], S * 0.50 - th / 2 - bb[1] + 18), text, font=font, fill=(255, 255, 255, 255))

os.makedirs("market", exist_ok=True)
out.resize((256, 256), Image.LANCZOS).save("market/icon-256.png")
small = out.resize((128, 128), Image.LANCZOS)
buf = io.BytesIO()
small.save(buf, format="PNG", optimize=True)
b = buf.getvalue()
print("png128 bytes:", len(b))
data_url = "data:image/png;base64," + base64.b64encode(b).decode()
print("data_url chars:", len(data_url), "limit 32768 ->", "OK" if len(data_url) <= 32768 else "TOO BIG")
open("market/icon-dataurl.txt", "w").write(data_url)
