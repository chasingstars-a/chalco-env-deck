# -*- coding: utf-8 -*-
"""生成占位素材（脱敏版）：全部由代码合成，不含任何真实公司标识或版权图片。

    python tools/make_placeholders.py        # 覆盖写入 ../assets/

想换成自己公司的素材时，用同名文件覆盖 assets/ 即可（尺寸比例保持一致）。
"""
from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont

ASSETS = Path(__file__).resolve().parent.parent / "assets"
SS = 2                      # 超采样倍数

GREEN_DEEP = (33, 75, 13)
GREEN = (51, 172, 72)
GREEN_LIGHT = (80, 197, 99)
MINT = (47, 207, 157)


# --------------------------------------------------------------------------- #
# 基础工具
# --------------------------------------------------------------------------- #
def sample_stops(stops, t):
    t = min(1.0, max(0.0, t))
    for i in range(len(stops) - 1):
        p0, c0 = stops[i]
        p1, c1 = stops[i + 1]
        if p0 <= t <= p1:
            f = 0 if p1 == p0 else (t - p0) / (p1 - p0)
            return tuple(round(a + (b - a) * f) for a, b in zip(c0, c1))
    return stops[-1][1]


def vgrad(size, stops):
    """竖向渐变。"""
    w, h = size
    strip = Image.new("RGB", (1, h))
    px = strip.load()
    for y in range(h):
        px[0, y] = sample_stops(stops, y / max(1, h - 1))
    return strip.resize((w, h), Image.BILINEAR)


def hgrad(size, stops):
    """横向渐变。"""
    w, h = size
    strip = Image.new("RGB", (1, w))
    px = strip.load()
    for x in range(w):
        px[0, x] = sample_stops(stops, x / max(1, w - 1))
    return strip.resize((w, h), Image.BILINEAR)


def radial_mask(size, center, radius, power=1.8, rings=90):
    """柔和的圆形 alpha 遮罩：由外向内叠加同心圆，再高斯模糊，圆心不必在画布中间。"""
    w, h = size
    mask = Image.new("L", size, 0)
    d = ImageDraw.Draw(mask)
    cx, cy = center
    for i in range(rings, 0, -1):
        r = radius * i / rings
        v = int(255 * (1 - i / rings) ** power)
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=v)
    return mask.filter(ImageFilter.GaussianBlur(max(1, radius * 0.03)))


def paste_color(base, color, mask):
    """用遮罩把纯色/渐变层贴到 base 上（正确 alpha 混合）。"""
    layer = Image.new("RGB", base.size, color) if isinstance(color, tuple) else color
    base.paste(layer, (0, 0), mask)
    return base


def mountains(size, baseline, amplitude, phase, blur=2.0, segments=260):
    """起伏山丘/草坡遮罩：baseline 以下不透明。"""
    w, h = size
    mask = Image.new("L", size, 0)
    d = ImageDraw.Draw(mask)
    pts = [(0, h)]
    for i in range(segments + 1):
        x = w * i / segments
        y = (baseline
             + amplitude * math.sin(i / segments * math.pi * 2 + phase)
             + amplitude * 0.42 * math.sin(i / segments * math.pi * 5.1 + phase * 1.6))
        pts.append((x, y))
    pts.append((w, h))
    d.polygon(pts, fill=255)
    return mask.filter(ImageFilter.GaussianBlur(blur))


def lens_mask(size, center, length, width, angle_deg):
    """叶片/花瓣形状（两个圆的交集），返回 alpha 遮罩。"""
    w, h = size
    big = Image.new("L", (w * 2, h * 2), 0)
    d = ImageDraw.Draw(big)
    cx, cy = w, h
    r = length * 1.05
    dx = length * 0.42
    d.ellipse([cx - dx - r, cy - r, cx - dx + r, cy + r], fill=255)
    other = Image.new("L", (w * 2, h * 2), 0)
    d2 = ImageDraw.Draw(other)
    d2.ellipse([cx + dx - r, cy - r, cx + dx + r, cy + r], fill=255)
    lens = ImageChops.darker(big, other)
    # 压扁成叶形
    lens = lens.resize((int(w * 2 * width), int(h * 2)), Image.LANCZOS)
    canvas = Image.new("L", (w * 2, h * 2), 0)
    canvas.paste(lens, (int(w - lens.width / 2), 0))
    canvas = canvas.rotate(angle_deg, center=(w, h), resample=Image.BICUBIC)
    canvas = canvas.crop((w // 2, h // 2, w // 2 + w, h // 2 + h))
    return canvas.resize(size, Image.LANCZOS)


# --------------------------------------------------------------------------- #
# 徽标 / 字标 / 二维码占位
# --------------------------------------------------------------------------- #
def make_emblem(path, white=False):
    """占位徽标：一片渐变叶子 + 叶脉，透明底，不放任何文字或真实标识。"""
    W, H = 205 * SS, 209 * SS
    canvas = Image.new("RGBA", (W, H), (0, 0, 0, 0))

    # 后面一片浅色叶子增加层次
    back = lens_mask((W, H), (W * 0.58, H * 0.46), W * 0.26, 0.92, 34)
    back = Image.eval(back, lambda v: int(v * 0.55))
    back = back.filter(ImageFilter.GaussianBlur(SS * 1.0))
    back_color = (255, 255, 255) if white else GREEN_LIGHT
    canvas.paste(Image.new("RGBA", (W, H), back_color + (255,)), (0, 0), back)

    # 主叶
    leaf = lens_mask((W, H), (W * 0.47, H * 0.50), W * 0.32, 0.95, -30)
    leaf = leaf.filter(ImageFilter.GaussianBlur(SS * 0.7))
    if white:
        main = Image.new("RGBA", (W, H), (255, 255, 255, 255))
    else:
        main = vgrad((W, H), [(0.0, MINT), (0.45, GREEN), (1.0, GREEN_DEEP)]).convert("RGBA")
    canvas.paste(main, (0, 0), leaf)

    # 叶脉
    vein = Image.new("L", (W, H), 0)
    ImageDraw.Draw(vein).line([(W * 0.36, H * 0.68), (W * 0.56, H * 0.38)], fill=120,
                              width=int(SS * 2.2))
    vein = vein.filter(ImageFilter.GaussianBlur(SS * 0.4))
    canvas.paste(Image.new("RGBA", (W, H), (255, 255, 255, 210)), (0, 0), vein)

    canvas.resize((205, 209), Image.LANCZOS).save(path, "PNG", optimize=True)
    return path


def make_wordmark(path, text="示例环保集团"):
    W, H = 1283 * 2, 151 * 2
    canvas = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    font = ImageFont.truetype(r"C:\Windows\Fonts\msyhbd.ttc", int(H * 0.72))
    mask = Image.new("L", (W, H), 0)
    md = ImageDraw.Draw(mask)
    box = md.textbbox((0, 0), text, font=font)
    md.text(((W - (box[2] - box[0])) / 2 - box[0], (H - (box[3] - box[1])) / 2 - box[1]),
            text, font=font, fill=255)
    grad = hgrad((W, H), [(0.0, MINT), (0.45, GREEN), (1.0, GREEN_DEEP)])
    canvas.paste(grad, (0, 0), mask)
    canvas.resize((1283, 151), Image.LANCZOS).save(path, "PNG", optimize=True)
    return path


def make_text_wordmark(path, text, size, ratio=0.72):
    """用微软雅黑渲染的渐变字标（替代模板里的艺术字图片）。"""
    W, H = size[0] * 2, size[1] * 2
    canvas = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    font = ImageFont.truetype(r"C:\Windows\Fonts\msyhbd.ttc", int(H * ratio))
    mask = Image.new("L", (W, H), 0)
    md = ImageDraw.Draw(mask)
    box = md.textbbox((0, 0), text, font=font)
    md.text(((W - (box[2] - box[0])) / 2 - box[0], (H - (box[3] - box[1])) / 2 - box[1]),
            text, font=font, fill=255)
    grad = hgrad((W, H), [(0.0, MINT), (0.45, GREEN), (1.0, GREEN_DEEP)])
    canvas.paste(grad, (0, 0), mask)
    canvas.resize(size, Image.LANCZOS).save(path, "PNG", optimize=True)
    return path


def make_qr_placeholder(path, modules=25, seed=20261021):
    """不是真二维码：固定种子的方块阵列 + 三个定位角，纯占位。"""
    cell, quiet = 12, 4
    n = modules + quiet * 2
    img = Image.new("RGB", (n * cell, n * cell), (255, 255, 255))
    d = ImageDraw.Draw(img)
    rnd = __import__("random").Random(seed)
    dark = (74, 86, 80)
    for y in range(modules):
        for x in range(modules):
            if rnd.random() < 0.45:
                px, py = (x + quiet) * cell, (y + quiet) * cell
                d.rounded_rectangle([px, py, px + cell - 2, py + cell - 2], radius=2, fill=dark)
    for ox, oy in ((0, 0), (modules - 7, 0), (0, modules - 7)):
        x0, y0 = (ox + quiet) * cell, (oy + quiet) * cell
        d.rectangle([x0, y0, x0 + 7 * cell - 2, y0 + 7 * cell - 2], outline=dark, width=cell)
        d.rectangle([x0 + 2 * cell, y0 + 2 * cell, x0 + 5 * cell - 2, y0 + 5 * cell - 2],
                    fill=dark)
    img.resize((311, 311), Image.LANCZOS).save(path, "PNG", optimize=True)
    return path


# --------------------------------------------------------------------------- #
# 底图
# --------------------------------------------------------------------------- #
def scene(size, palette, hills, sun=(0.70, 0.14, 0.30), haze=0.30):
    W, H = size[0] * SS, size[1] * SS
    img = vgrad((W, H), palette["sky"])
    # 太阳/柔光
    if sun:
        mask = radial_mask((W, H), (W * sun[0], H * sun[1]), W * sun[2], power=2.2)
        mask = Image.eval(mask, lambda v: int(v * 0.85))
        img = paste_color(img, (255, 252, 232), mask)
    # 山丘分层
    for baseline, color, amp, phase, blur in hills:
        mask = mountains((W, H), H * baseline, H * amp, phase, blur * SS)
        layer = vgrad((W, H), [
            (0.0, tuple(min(255, c + 16) for c in color)),
            (1.0, tuple(max(0, c - 18) for c in color)),
        ])
        img = paste_color(img, layer, mask)
    # 底部薄雾，让文字区更透气
    if haze:
        mask = Image.new("L", (W, H), 0)
        md = ImageDraw.Draw(mask)
        for y in range(int(H * (1 - haze)), H):
            t = max(0.0, min(1.0, (y - H * (1 - haze)) / (H * haze)))
            md.line([(0, y), (W, y)], fill=int(255 * t ** 1.5))
        img = paste_color(img, (255, 255, 255), mask.filter(ImageFilter.GaussianBlur(6 * SS)))
    return img


def make_cover(path, size=(2200, 1238)):
    palette = {"sky": [(0.0, (243, 250, 255)), (0.34, (208, 232, 250)),
                       (0.62, (188, 221, 245)), (0.80, (226, 241, 251)),
                       (1.0, (247, 252, 255))]}
    hills = [
        (0.64, (186, 220, 182), 0.028, 0.6, 1.2),
        (0.72, (138, 202, 128), 0.024, 2.2, 1.6),
        (0.81, (84, 176, 92), 0.020, 4.1, 2.0),
    ]
    img = scene(size, palette, hills)
    img.resize(size, Image.LANCZOS).save(path, "JPEG", quality=90, optimize=True,
                                         progressive=True)
    return path


def make_closing(path, size=(2200, 1238)):
    palette = {"sky": [(0.0, (232, 246, 250)), (0.30, (186, 226, 240)),
                       (0.58, (150, 209, 224)), (0.78, (196, 232, 238)),
                       (1.0, (236, 248, 249))]}
    hills = [
        (0.70, (150, 200, 178), 0.030, 1.1, 1.4),
        (0.78, (96, 172, 132), 0.026, 3.0, 1.8),
        (0.86, (44, 128, 92), 0.022, 5.2, 2.2),
    ]
    img = scene(size, palette, hills, sun=(0.30, 0.18, 0.26), haze=0.22)
    img.resize(size, Image.LANCZOS).save(path, "JPEG", quality=90, optimize=True,
                                         progressive=True)
    return path


def make_section(path, size=(2200, 734)):
    palette = {"sky": [(0.0, (246, 251, 255)), (0.5, (206, 233, 226)), (1.0, (150, 206, 170))]}
    hills = [
        (0.52, (176, 218, 172), 0.05, 0.4, 1.2),
        (0.66, (112, 190, 128), 0.05, 2.6, 1.6),
        (0.82, (52, 156, 88), 0.04, 4.6, 2.0),
    ]
    img = scene(size, palette, hills, sun=(0.62, 0.10, 0.26), haze=0.18)
    img.resize(size, Image.LANCZOS).save(path, "JPEG", quality=90, optimize=True,
                                         progressive=True)
    return path


def make_light_bg(path, size=(2000, 1127)):
    W, H = size[0] * SS, size[1] * SS
    img = vgrad((W, H), [(0.0, (231, 241, 252)), (0.45, (245, 249, 254)),
                         (1.0, (228, 239, 250))])
    glow = radial_mask((W, H), (W * 0.70, H * 0.26), W * 0.46, power=2.4)
    img = paste_color(img, (255, 255, 255), Image.eval(glow, lambda v: int(v * 0.9)))
    # 两条柔光弧线
    for i, (cx, cy, rx, ry, a) in enumerate((
            (0.55, 1.02, 0.75, 0.55, 46), (0.78, 1.15, 0.95, 0.62, 30))):
        arc = Image.new("L", (W, H), 0)
        ImageDraw.Draw(arc).arc([W * (cx - rx), H * (cy - ry), W * (cx + rx), H * (cy + ry)],
                                start=200, end=340, fill=a, width=int(3 * SS))
        arc = arc.filter(ImageFilter.GaussianBlur(5 * SS))
        img = paste_color(img, (255, 255, 255), arc)
    img.resize(size, Image.LANCZOS).save(path, "JPEG", quality=88, optimize=True,
                                         progressive=True)
    return path


def make_leaf_photo(path, size=(1600, 1067)):
    """抽象「绿植虚化」配图：深浅绿斑 + 高光，不指向任何真实场景。"""
    import random as _random

    W, H = size[0] * SS, size[1] * SS
    img = vgrad((W, H), [(0.0, (232, 247, 236)), (0.42, (176, 224, 186)),
                         (1.0, (52, 140, 84))])
    rnd = _random.Random(20261021)
    palette = [(46, 132, 78), (66, 158, 96), (28, 108, 62), (140, 205, 150), (250, 253, 245)]
    for _ in range(52):
        cx, cy = rnd.uniform(0, W), rnd.uniform(0, H)
        r = rnd.uniform(0.05, 0.22) * W
        color = rnd.choice(palette)
        alpha = rnd.randint(28, 110)
        mask = Image.new("L", (W, H), 0)
        ImageDraw.Draw(mask).ellipse([cx - r, cy - r * 0.8, cx + r, cy + r * 0.8],
                                     fill=alpha)
        mask = mask.filter(ImageFilter.GaussianBlur(r * 0.55))
        img = paste_color(img, color, mask)
    # 一道柔光
    streak = radial_mask((W, H), (W * 0.72, H * 0.18), W * 0.30, power=2.0)
    img = paste_color(img, (255, 255, 250), Image.eval(streak, lambda v: int(v * 0.75)))
    img = img.filter(ImageFilter.GaussianBlur(1.6 * SS))
    img.resize(size, Image.LANCZOS).save(path, "JPEG", quality=90, optimize=True,
                                         progressive=True)
    return path


def main():
    ASSETS.mkdir(exist_ok=True)
    make_emblem(ASSETS / "logo_emblem.png")
    make_emblem(ASSETS / "logo_emblem_white.png", white=True)
    make_wordmark(ASSETS / "logo_wordmark.png")
    make_text_wordmark(ASSETS / "wordmark_mulu.png", "目录", (239, 151), 0.80)
    make_text_wordmark(ASSETS / "wordmark_thanks.png", "分享结束 感谢聆听", (933, 151), 0.72)
    make_qr_placeholder(ASSETS / "qr_wechat.png")
    make_cover(ASSETS / "bg_cover_full.jpg")
    make_closing(ASSETS / "bg_closing.jpg")
    make_section(ASSETS / "bg_section.jpg")
    make_light_bg(ASSETS / "bg_light.jpg")
    make_leaf_photo(ASSETS / "photo_leaf.jpg")
    for p in sorted(ASSETS.iterdir()):
        if p.suffix.lower() in (".png", ".jpg", ".jpeg"):
            print(f"{p.name:24s} {Image.open(p).size} {p.stat().st_size // 1024} KB")


if __name__ == "__main__":
    main()
