#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
generar_logo.py
Crea assets/logo_tienda.png (logo cuadrado) y assets/portada.jpg (banner de la hoja
Inicio, con el logo). Para usar el logo real de la empresa, reemplace
assets/logo_tienda.png y vuelva a ejecutar este script con --solo-portada,
o simplemente reemplace assets/portada.jpg por su propia imagen (1200x420).
"""
import math
import os
import sys

from PIL import Image, ImageDraw, ImageFilter, ImageFont

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
ASSETS = os.path.join(BASE, "assets")
NAVY, EMERALD, WHITE, SLATE = (15, 23, 42), (16, 185, 129), (255, 255, 255), (148, 163, 184)


def fuente(size, bold=True):
    for p in (
        "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
        "C:/Windows/Fonts/georgiab.ttf",
    ):
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def sans(size, bold=True):
    for p in (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/segoeuib.ttf",
    ):
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def pent(cx, cy, r, rot=-90):
    return [(cx + r * math.cos(math.radians(rot + 72 * i)), cy + r * math.sin(math.radians(rot + 72 * i))) for i in range(5)]


def logo(size=800):
    S = 2
    n = size * S
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    c = n / 2
    d.ellipse([n * 0.03, n * 0.03, n * 0.97, n * 0.97], fill=NAVY)
    d.ellipse([n * 0.07, n * 0.07, n * 0.93, n * 0.93], outline=EMERALD, width=int(n * 0.018))
    # balon
    R = n * 0.22
    cy = n * 0.40
    d.ellipse([c - R, cy - R, c + R, cy + R], fill=WHITE)
    d.polygon(pent(c, cy, R * 0.36), fill=NAVY)
    for i in range(5):
        a = math.radians(-90 + 72 * i)
        d.line([(c + R * 0.36 * math.cos(a), cy + R * 0.36 * math.sin(a)), (c + R * 0.8 * math.cos(a), cy + R * 0.8 * math.sin(a))], fill=NAVY, width=int(n * 0.008))
        ox, oy = c + R * 0.97 * math.cos(a + math.radians(36)), cy + R * 0.97 * math.sin(a + math.radians(36))
        d.polygon(pent(ox, oy, R * 0.3, rot=math.degrees(a) + 216), fill=EMERALD)
    d.ellipse([c - R, cy - R, c + R, cy + R], outline=EMERALD, width=int(n * 0.012))
    # texto
    f1 = fuente(int(n * 0.105))
    f2 = sans(int(n * 0.045))
    t1, t2 = "TIENDA DE", "FÚTBOL"
    b = d.textbbox((0, 0), t1, font=f2)
    d.text((c - (b[2] - b[0]) / 2, n * 0.68), t1, font=f2, fill=SLATE)
    b = d.textbbox((0, 0), t2, font=f1)
    d.text((c - (b[2] - b[0]) / 2 - b[0], n * 0.735), t2, font=f1, fill=WHITE)
    d.rectangle([c - n * 0.12, n * 0.87, c + n * 0.12, n * 0.878], fill=EMERALD)
    return img.resize((size, size), Image.LANCZOS)


def portada(lg):
    S = 1
    w, h = 1200, 420
    img = Image.new("RGB", (w, h), NAVY)
    d = ImageDraw.Draw(img)
    for y in range(h):
        t = y / h
        col = tuple(int(a * (1 - t) + b * t) for a, b in zip((15, 23, 42), (30, 41, 59)))
        d.line([(0, y), (w, y)], fill=col)
    # lineas de cancha tenues
    ln = (30, 58, 73)
    d.rectangle([40, 30, w - 40, h - 30], outline=ln, width=3)
    d.line([(w / 2, 30), (w / 2, h - 30)], fill=ln, width=3)
    d.ellipse([w / 2 - 90, h / 2 - 90, w / 2 + 90, h / 2 + 90], outline=ln, width=3)
    d.rectangle([40, h / 2 - 90, 190, h / 2 + 90], outline=ln, width=3)
    d.rectangle([w - 190, h / 2 - 90, w - 40, h / 2 + 90], outline=ln, width=3)
    d.rectangle([0, h - 14, w, h], fill=EMERALD)
    sz = 340
    l2 = lg.resize((sz, sz), Image.LANCZOS)
    img.paste(l2, (70, (h - sz) // 2 - 6), l2)
    f_t = fuente(54)
    f_s = sans(21, False)
    d.text((450, 130), "TIENDA DE FÚTBOL", font=f_t, fill=WHITE)
    d.rectangle([452, 210, 620, 217], fill=EMERALD)
    d.text((452, 240), "Camisetas · Botines · Balones · Guantes · Accesorios", font=f_s, fill=(203, 213, 225))
    d.text((452, 284), "Compra, vende y controla tu inventario en un solo lugar", font=sans(18, False), fill=SLATE)
    return img


def main():
    os.makedirs(ASSETS, exist_ok=True)
    ruta_logo = os.path.join(ASSETS, "logo_tienda.png")
    if "--solo-portada" in sys.argv and os.path.exists(ruta_logo):
        lg = Image.open(ruta_logo).convert("RGBA")
    else:
        lg = logo()
        lg.save(ruta_logo)
    portada(lg).save(os.path.join(ASSETS, "portada.jpg"), quality=93)
    # Logo para los formularios (fondo azul marino del banner; MSForms no lee PNG con transparencia)
    fondo = Image.new("RGB", (300, 300), NAVY)
    mini = lg.resize((300, 300), Image.LANCZOS)
    fondo.paste(mini, (0, 0), mini)
    fondo.save(os.path.join(ASSETS, "logo_form.jpg"), quality=95)
    print("Logo y portada generados en", ASSETS)


if __name__ == "__main__":
    main()
