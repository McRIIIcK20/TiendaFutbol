#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
generar_imagenes.py
Genera las 60 imagenes de producto (imagenes_productos/<SKU>.jpg) y el logo de la
tienda (assets/). Cada imagen se dibuja segun la ESPECIFICACION del producto:
  - Camisetas: colores oficiales de cada seleccion y variante (Local / Visitante).
  - Botines: cafe con tacos / sala de suela plana / cesped sintetico con tacos cortos.
  - Balones: tamano (N5, N4, N3, mini) y estilo segun el modelo.
  - Canilleras, guantes y accesorios: forma y color segun el modelo y la talla.
Uso:  python scripts/generar_imagenes.py   (requiere Pillow)
Las imagenes son ilustraciones propias; reemplazables por fotos reales con el mismo
nombre de archivo (SKU.jpg).
"""
import csv
import math
import os
import re

from PIL import Image, ImageDraw, ImageFont, ImageFilter

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUT = os.path.join(BASE, "imagenes_productos")
ASSETS = os.path.join(BASE, "assets")
CSV = os.path.join(BASE, "powerbi", "datos_csv", "vw_rentabilidad_productos.csv")
W = 1000          # lienzo de trabajo (se reduce a 500)
K = 2             # factor: coordenadas en base 500


def rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def shade(c, f):
    return tuple(max(0, min(255, int(v * f))) for v in c)


def font(size, bold=True):
    for p in (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/segoeuib.ttf",
    ):
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def P(pts):
    return [(x * K, y * K) for x, y in pts]


def fondo(acento):
    """Fondo claro con degradado suave y un circulo del color del producto."""
    img = Image.new("RGB", (W, W), rgb("#F8FAFC"))
    d = ImageDraw.Draw(img)
    for y in range(W):
        t = y / W
        c = tuple(int(a * (1 - t) + b * t) for a, b in zip(rgb("#F8FAFC"), rgb("#E2E8F0")))
        d.line([(0, y), (W, y)], fill=c)
    halo = Image.new("RGB", (W, W), (0, 0, 0))
    m = Image.new("L", (W, W), 0)
    ImageDraw.Draw(m).ellipse([120, 120, 880, 880], fill=70)
    m = m.filter(ImageFilter.GaussianBlur(60))
    img.paste(Image.new("RGB", (W, W), acento), (0, 0), m)
    return img


def sombra(img, caja):
    m = Image.new("L", (W, W), 0)
    ImageDraw.Draw(m).ellipse([c * K for c in caja], fill=110)
    m = m.filter(ImageFilter.GaussianBlur(16))
    img.paste(Image.new("RGB", (W, W), (30, 41, 59)), (0, 0), m)


def etiqueta(img, texto, color):
    """Insignia de talla/tamano en la esquina inferior derecha."""
    d = ImageDraw.Draw(img)
    f = font(64 if len(texto) < 6 else 46)
    bb = d.textbbox((0, 0), texto, font=f)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    pad = 30
    x1, y1 = W - 44, W - 44
    x0, y0 = x1 - tw - 2 * pad, y1 - th - 2 * pad
    d.rounded_rectangle([x0, y0, x1, y1], radius=26, fill=rgb("#0F172A"))
    d.rounded_rectangle([x0, y0, x1, y0 + 8], radius=4, fill=color)
    d.text((x0 + pad - bb[0], y0 + pad - bb[1] + 2), texto, font=f, fill=(255, 255, 255))


# ----------------------------------------------------------------------------
# CAMISETAS
# ----------------------------------------------------------------------------
EQUIPOS = {
    # clave: (cuerpo, cuello/puños, detalle, color texto escudo)
    "chile local": ("#D52B1E", "#0B3C9E", "#FFFFFF"),
    "argentina visitante": ("#1B2A5C", "#75AADB", "#75AADB"),
    "brasil local": ("#FFDF00", "#009C3B", "#1D4ED8"),
    "colombia local": ("#FCD116", "#003893", "#CE1126"),
    "espana local": ("#C60B1E", "#FFC400", "#FFC400"),
    "alemania local": ("#F5F5F5", "#111111", "#DD0000"),
}


def camiseta(equipo, talla):
    cuerpo, trim, det = [rgb(c) for c in EQUIPOS[equipo]]
    img = fondo(cuerpo if equipo != "alemania local" else rgb("#94A3B8"))
    sombra(img, (120, 392, 380, 432))
    d = ImageDraw.Draw(img)
    cuerpo_pts = [(215, 100), (150, 120), (66, 205), (116, 242), (152, 212), (158, 400), (342, 400),
                  (348, 212), (384, 242), (434, 205), (350, 120), (285, 100), (250, 142)]
    d.polygon(P(cuerpo_pts), fill=cuerpo)
    # sombreado lateral
    d.polygon(P([(152, 212), (158, 400), (180, 400), (176, 215)]), fill=shade(cuerpo, 0.9))
    d.polygon(P([(348, 212), (342, 400), (320, 400), (324, 215)]), fill=shade(cuerpo, 0.92))
    # puños
    d.polygon(P([(66, 205), (82, 222), (132, 258), (116, 242)]), fill=trim)
    d.polygon(P([(434, 205), (418, 222), (368, 258), (384, 242)]), fill=trim)
    # franja segun seleccion
    if equipo == "chile local":
        d.polygon(P([(150, 120), (164, 128), (88, 214), (66, 205)]), fill=rgb("#FFFFFF"))
        d.polygon(P([(350, 120), (336, 128), (412, 214), (434, 205)]), fill=rgb("#FFFFFF"))
    elif equipo == "argentina visitante":
        for x in (200, 232, 264):
            d.rectangle([x * K, 150 * K, (x + 14) * K, 400 * K], fill=shade(cuerpo, 1.25))
    elif equipo == "brasil local":
        d.polygon(P([(158, 372), (342, 372), (342, 400), (158, 400)]), fill=trim)
    elif equipo == "colombia local":
        d.polygon(P([(156, 300), (344, 300), (344, 326), (156, 326)]), fill=trim)
        d.polygon(P([(157, 326), (343, 326), (343, 344), (157, 344)]), fill=det)
    elif equipo == "espana local":
        d.polygon(P([(150, 120), (164, 128), (88, 214), (66, 205)]), fill=trim)
        d.polygon(P([(350, 120), (336, 128), (412, 214), (434, 205)]), fill=trim)
    elif equipo == "alemania local":
        d.polygon(P([(157, 285), (343, 285), (343, 300), (157, 300)]), fill=det)
        d.polygon(P([(157, 300), (343, 300), (343, 312), (157, 312)]), fill=rgb("#FFCE00"))
        d.polygon(P([(157, 312), (343, 312), (343, 324), (157, 324)]), fill=rgb("#111111"))
    # cuello en V redondeado
    d.polygon(P([(215, 100), (250, 142), (285, 100), (298, 104), (250, 160), (202, 104)]), fill=trim)
    # escudo
    ex, ey = 292, 196
    d.ellipse([(ex - 16) * K, (ey - 16) * K, (ex + 16) * K, (ey + 16) * K], fill=rgb("#FFFFFF") if equipo != "alemania local" else rgb("#111111"))
    d.ellipse([(ex - 10) * K, (ey - 10) * K, (ex + 10) * K, (ey + 10) * K], fill=det if equipo != "brasil local" else rgb("#009C3B"))
    # pliegues
    d.line(P([(250, 200), (252, 392)]), fill=shade(cuerpo, 0.94), width=3 * K)
    img = img.resize((500, 500), Image.LANCZOS)
    etiqueta_img(img, talla, trim if equipo != "alemania local" else rgb("#DD0000"))
    return img


def etiqueta_img(img500, texto, color):
    big = img500.resize((W, W), Image.LANCZOS)
    etiqueta(big, texto, color)
    img500.paste(big.resize((500, 500), Image.LANCZOS))


# ----------------------------------------------------------------------------
# BOTINES
# ----------------------------------------------------------------------------
def botin(tipo, talla):
    if tipo == "cafe":
        cuero, acento, suela = rgb("#7A4A2B"), rgb("#F5E6CC"), rgb("#2B2B2B")
        halo = rgb("#A16207")
    elif tipo == "sala":
        cuero, acento, suela = rgb("#1D4ED8"), rgb("#FFFFFF"), rgb("#D6B98C")
        halo = rgb("#1D4ED8")
    else:
        cuero, acento, suela = rgb("#161616"), rgb("#F97316"), rgb("#2B2B2B")
        halo = rgb("#F97316")
    img = fondo(halo)
    sombra(img, (70, 372, 450, 420))
    d = ImageDraw.Draw(img)
    # tacos / suela
    if tipo == "cafe":
        for x in (135, 190, 285, 340, 395):
            d.polygon(P([(x - 12, 385), (x + 12, 385), (x + 8, 408), (x - 8, 408)]), fill=rgb("#E5E7EB"))
    elif tipo == "cesped":
        for i in range(11):
            for j in range(2):
                x = 120 + i * 30
                d.ellipse([(x - 6) * K, (392 + j * 0) * K, (x + 6) * K, (404) * K], fill=rgb("#E5E7EB"))
    d.polygon(P([(95, 360), (435, 360), (440, 388), (100, 388)]), fill=suela)
    if tipo == "sala":
        d.polygon(P([(95, 372), (438, 372), (440, 388), (100, 388)]), fill=rgb("#B08B57"))
    # capellada
    upper = [(105, 232), (118, 168), (238, 158), (252, 214), (335, 252), (412, 298),
             (428, 338), (404, 362), (104, 362)]
    d.polygon(P(upper), fill=cuero)
    # talon y refuerzo
    d.polygon(P([(105, 232), (118, 168), (142, 166), (132, 362), (104, 362)]), fill=shade(cuero, 0.78))
    # lengueta
    d.polygon(P([(190, 160), (250, 154), (262, 212), (214, 222)]), fill=shade(cuero, 1.18))
    # franjas decorativas
    for off in (0, 22, 44):
        d.polygon(P([(170 + off, 362), (200 + off, 362), (262 + off, 290), (240 + off, 282)]), fill=acento)
    # cordones
    for i in range(5):
        x, y = 226 + i * 26, 214 + i * 12
        d.line(P([(x, y), (x + 34, y - 20)]), fill=acento, width=5 * K)
    # punta
    d.polygon(P([(380, 288), (412, 298), (428, 338), (404, 362), (372, 362), (388, 322)]), fill=shade(cuero, 0.88))
    # brillo
    d.line(P([(150, 190), (230, 182)]), fill=shade(cuero, 1.45), width=4 * K)
    img = img.resize((500, 500), Image.LANCZOS)
    etiqueta_img(img, "Talla " + talla, halo)
    return img


# ----------------------------------------------------------------------------
# BALONES
# ----------------------------------------------------------------------------
def pentagono(cx, cy, r, rot=-90):
    return [(cx + r * math.cos(math.radians(rot + 72 * i)), cy + r * math.sin(math.radians(rot + 72 * i))) for i in range(5)]


def dibujar_balon(img, cx, cy, R, base, parche, linea):
    """Dibuja un balon en el lienzo grande (coordenadas ya multiplicadas por K)."""
    capa = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    d = ImageDraw.Draw(capa)
    d.ellipse([cx - R, cy - R, cx + R, cy + R], fill=base)
    p = pentagono(cx, cy, R * 0.34)
    d.polygon(p, fill=parche)
    for i in range(5):
        a = math.radians(-90 + 72 * i)
        x1, y1 = cx + R * 0.34 * math.cos(a), cy + R * 0.34 * math.sin(a)
        x2, y2 = cx + R * 0.62 * math.cos(a), cy + R * 0.62 * math.sin(a)
        d.line([(x1, y1), (x2, y2)], fill=linea, width=max(2, int(R * 0.03)))
        ox, oy = cx + R * 0.98 * math.cos(a + math.radians(36)), cy + R * 0.98 * math.sin(a + math.radians(36))
        d.polygon(pentagono(ox, oy, R * 0.3, rot=math.degrees(a) + 36 + 180), fill=parche)
        bx, by = cx + R * 0.62 * math.cos(a), cy + R * 0.62 * math.sin(a)
        for s in (-1, 1):
            a2 = a + s * math.radians(36)
            d.line([(bx, by), (cx + R * 0.86 * math.cos(a2), cy + R * 0.86 * math.sin(a2))], fill=linea, width=max(2, int(R * 0.03)))
    # sombreado esferico
    sh = Image.new("L", (W, W), 0)
    ImageDraw.Draw(sh).ellipse([cx - R * 0.2, cy - R * 0.2, cx + R * 1.25, cy + R * 1.25], fill=0)
    mask = Image.new("L", (W, W), 0)
    ImageDraw.Draw(mask).ellipse([cx - R, cy - R, cx + R, cy + R], fill=255)
    cap = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    cap.paste(capa, (0, 0), mask)
    dd = ImageDraw.Draw(cap)
    shadow = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).ellipse([cx - R * 0.9, cy - R * 0.9, cx + R * 1.15, cy + R * 1.15], fill=(0, 0, 0, 70))
    shadow = shadow.filter(ImageFilter.GaussianBlur(R * 0.25))
    inv = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    inv.paste(shadow, (0, 0), mask)
    base_img = img.convert("RGBA")
    base_img.alpha_composite(cap)
    # luz: sombra interior abajo-derecha
    lum = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    ImageDraw.Draw(lum).ellipse([cx - R, cy - R, cx + R, cy + R], outline=(15, 23, 42, 160), width=max(3, int(R * 0.02)))
    base_img.alpha_composite(lum)
    hl = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    ImageDraw.Draw(hl).ellipse([cx - R * 0.55, cy - R * 0.75, cx - R * 0.15, cy - R * 0.5], fill=(255, 255, 255, 110))
    hl = hl.filter(ImageFilter.GaussianBlur(R * 0.06))
    base_img.alpha_composite(hl)
    return base_img.convert("RGB")


BALONES = {
    # clave: (radio base 500, color base, parche, linea, halo, etiqueta)
    "n5 professional": (150, "#FFFFFF", "#111827", "#9CA3AF", "#0F172A", "N5 Pro"),
    "n5 entrenamiento": (150, "#F97316", "#111827", "#7C2D12", "#F97316", "N5 Entr."),
    "n4 juvenil": (128, "#FFFFFF", "#1D4ED8", "#93C5FD", "#1D4ED8", "N4"),
    "futbolito n4": (128, "#FDE047", "#15803D", "#A16207", "#EAB308", "Futbolito N4"),
    "n3 infantil": (108, "#DC2626", "#FFFFFF", "#7F1D1D", "#DC2626", "N3"),
    "mini souvenir": (70, "#FFFFFF", "#7C3AED", "#C4B5FD", "#7C3AED", "Mini"),
}


def balon(clave):
    r, base, parche, linea, halo, et = BALONES[clave]
    img = fondo(rgb(halo))
    cx, cy = 250 * K, (235 + (150 - r) * 0.55) * K
    d = ImageDraw.Draw(img)
    sh = Image.new("L", (W, W), 0)
    ImageDraw.Draw(sh).ellipse([cx - r * K * 0.95, 400 * K - 18 * K + (150 - r) * 0.2 * K, cx + r * K * 0.95, 400 * K + 14 * K + (150 - r) * 0.2 * K], fill=120)
    sh = sh.filter(ImageFilter.GaussianBlur(14))
    img.paste(Image.new("RGB", (W, W), (30, 41, 59)), (0, 0), sh)
    img = dibujar_balon(img, cx, cy + (150 - r) * 0.6 * K * 0, r * K, rgb(base), rgb(parche), rgb(linea))
    img = img.resize((500, 500), Image.LANCZOS)
    etiqueta_img(img, et, rgb(halo))
    return img


# ----------------------------------------------------------------------------
# CANILLERAS
# ----------------------------------------------------------------------------
def canillera(modelo, talla):
    esc = {"chica": 0.78, "mediana": 0.89, "grande": 1.00}[talla.lower()]
    if modelo == "pro":
        cuerpo, borde, halo = rgb("#111827"), rgb("#E5E7EB"), rgb("#475569")
    elif modelo == "tobillera":
        cuerpo, borde, halo = rgb("#1D4ED8"), rgb("#FFFFFF"), rgb("#1D4ED8")
    else:
        cuerpo, borde, halo = rgb("#DC2626"), rgb("#111827"), rgb("#DC2626")
    img = fondo(halo)
    sombra(img, (90, 392, 410, 428))
    d = ImageDraw.Draw(img)
    for cx in (165, 335):
        h = 290 * esc
        top = 395 - h
        w_top, w_bot = 118 * esc, 80 * esc
        pts = [(cx - w_top / 2, top), (cx + w_top / 2, top), (cx + w_bot / 2, 395 - 14), (cx + w_bot / 2 - 6, 395), (cx - w_bot / 2 + 6, 395), (cx - w_bot / 2, 395 - 14)]
        d.polygon(P(pts), fill=cuerpo)
        d.line(P(pts + [pts[0]]), fill=borde, width=4 * K)
        # nervaduras
        for i in range(1, 6):
            y = top + h * i / 7
            wi = w_top + (w_bot - w_top) * (i / 7)
            d.line(P([(cx - wi / 2 + 12, y), (cx + wi / 2 - 12, y)]), fill=shade(cuerpo, 1.5 if modelo != "pro" else 2.2), width=3 * K)
        if modelo == "tobillera":
            d.ellipse(P([(cx - 40 * esc, 395 - 70 * esc), (cx + 40 * esc, 395 + 6)]), fill=shade(cuerpo, 0.85), outline=borde, width=3 * K)
            d.line(P([(cx - 42, top + 20 * esc), (cx + 42, top + 20 * esc)]), fill=rgb("#FFFFFF"), width=6 * K)
        if modelo == "escolar":
            for yy in (top + h * 0.28, top + h * 0.68):
                d.rectangle(P([(cx - w_top / 2 - 8, yy), (cx + w_top / 2 + 8, yy + 16)]), fill=rgb("#111827"))
                d.rectangle(P([(cx - 6, yy + 2), (cx + 6, yy + 14)]), fill=rgb("#F8FAFC"))
        d.rounded_rectangle(P([(cx - 18, top + 10), (cx + 18, top + 26)]), radius=6 * K, fill=borde)
    img = img.resize((500, 500), Image.LANCZOS)
    etiqueta_img(img, talla.capitalize(), halo)
    return img


# ----------------------------------------------------------------------------
# GUANTES
# ----------------------------------------------------------------------------
def guante(modelo, talla):
    if modelo == "pro":
        palma, dorso, acento, halo = rgb("#F8FAFC"), rgb("#111827"), rgb("#84CC16"), rgb("#84CC16")
    elif modelo == "entrenamiento":
        palma, dorso, acento, halo = rgb("#F8FAFC"), rgb("#1D4ED8"), rgb("#F97316"), rgb("#1D4ED8")
    else:
        palma, dorso, acento, halo = rgb("#FDF2F8"), rgb("#EC4899"), rgb("#7C3AED"), rgb("#EC4899")
    img = fondo(halo)
    sombra(img, (110, 396, 390, 430))
    d = ImageDraw.Draw(img)
    # dedos
    dedos = [(150, 130, 62, 190), (214, 98, 62, 222), (278, 108, 62, 212), (340, 146, 58, 176)]
    for x, y, w, h in dedos:
        d.rounded_rectangle(P([(x, y), (x + w, y + h)]), radius=31 * K, fill=dorso)
        d.rounded_rectangle(P([(x + 8, y + 8), (x + w - 8, y + h - 20)]), radius=24 * K, fill=shade(dorso, 1.12))
        d.line(P([(x + w / 2, y + 40), (x + w / 2, y + h - 40)]), fill=shade(dorso, 0.85), width=3 * K)
    # palma y pulgar
    d.rounded_rectangle(P([(140, 270), (396, 372)]), radius=40 * K, fill=dorso)
    d.polygon(P([(140, 300), (86, 250), (68, 276), (110, 340), (150, 356)]), fill=dorso)
    d.rounded_rectangle(P([(60, 232), (112, 290)]), radius=24 * K, fill=dorso)
    d.rounded_rectangle(P([(160, 282), (372, 340)]), radius=26 * K, fill=palma)
    for i in range(6):
        d.line(P([(176 + i * 34, 290), (176 + i * 34, 332)]), fill=shade(palma, 0.8), width=2 * K)
    # muñeca
    d.rounded_rectangle(P([(150, 366), (386, 432)]), radius=14 * K, fill=shade(dorso, 0.8))
    d.rectangle(P([(150, 388), (386, 408)]), fill=acento)
    d.rounded_rectangle(P([(230, 370), (310, 428)]), radius=8 * K, fill=rgb("#F8FAFC"))
    # franjas de acento en dedos
    for x, y, w, h in dedos:
        d.rounded_rectangle(P([(x + 8, y + 64), (x + w - 8, y + 76)]), radius=4 * K, fill=acento)
    img = img.resize((500, 500), Image.LANCZOS)
    etiqueta_img(img, "Talla " + talla, halo)
    return img


# ----------------------------------------------------------------------------
# ACCESORIOS
# ----------------------------------------------------------------------------
def acc_rodillera():
    halo = rgb("#475569")
    img = fondo(halo)
    sombra(img, (110, 396, 390, 428))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle(P([(140, 120), (360, 400)]), radius=60 * K, fill=rgb("#1F2937"))
    d.ellipse(P([(185, 205), (315, 335)]), fill=rgb("#0F172A"), outline=rgb("#94A3B8"), width=5 * K)
    d.ellipse(P([(212, 232), (288, 308)]), fill=rgb("#334155"))
    for y in range(140, 400, 22):
        d.line(P([(150, y), (350, y)]), fill=rgb("#374151"), width=3 * K)
    d.rectangle(P([(140, 130), (360, 150)]), fill=rgb("#94A3B8"))
    d.rectangle(P([(140, 372), (360, 392)]), fill=rgb("#94A3B8"))
    return finalizar(img, "Rodillera", halo)


def acc_cinta():
    halo = rgb("#EAB308")
    img = fondo(halo)
    sombra(img, (60, 360, 440, 410))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle(P([(80, 170), (420, 330)]), radius=22 * K, fill=rgb("#FACC15"))
    d.rectangle(P([(80, 170), (420, 196)]), fill=rgb("#CA8A04"))
    d.rectangle(P([(80, 304), (420, 330)]), fill=rgb("#CA8A04"))
    d.ellipse(P([(205, 188), (295, 312)]), fill=rgb("#111827"))
    d.text((250 * K, 250 * K), "C", font=font(120), fill=rgb("#FACC15"), anchor="mm")
    d.rectangle(P([(350, 196), (410, 304)]), fill=rgb("#EAB308"))
    for y in range(206, 300, 14):
        d.line(P([(356, y), (404, y)]), fill=rgb("#A16207"), width=2 * K)
    return finalizar(img, "Capitan", halo)


def acc_bomba():
    halo = rgb("#0EA5E9")
    img = fondo(halo)
    sombra(img, (110, 396, 390, 430))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle(P([(210, 150), (290, 380)]), radius=16 * K, fill=rgb("#94A3B8"))
    d.rounded_rectangle(P([(222, 160), (278, 370)]), radius=12 * K, fill=rgb("#CBD5E1"))
    d.rounded_rectangle(P([(150, 108), (350, 140)]), radius=14 * K, fill=rgb("#0F172A"))
    d.rectangle(P([(240, 138), (260, 160)]), fill=rgb("#0F172A"))
    d.rounded_rectangle(P([(160, 372), (340, 410)]), radius=10 * K, fill=rgb("#0369A1"))
    d.line(P([(340, 392), (410, 330), (410, 250)]), fill=rgb("#111827"), width=8 * K)
    d.polygon(P([(402, 250), (418, 250), (414, 205), (406, 205)]), fill=rgb("#CBD5E1"))
    return finalizar(img, "Bomba", halo)


def acc_morral():
    halo = rgb("#10B981")
    img = fondo(halo)
    sombra(img, (100, 396, 400, 430))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle(P([(140, 120), (360, 400)]), radius=50 * K, fill=rgb("#1E3A8A"))
    d.rounded_rectangle(P([(170, 270), (330, 370)]), radius=24 * K, fill=rgb("#1E40AF"), outline=rgb("#10B981"), width=4 * K)
    d.line(P([(250, 285), (250, 285)]), fill=rgb("#10B981"), width=4 * K)
    d.arc(P([(190, 70), (310, 170)]), 180, 360, fill=rgb("#0F172A"), width=12 * K)
    d.rounded_rectangle(P([(200, 150), (300, 190)]), radius=14 * K, fill=rgb("#1E40AF"))
    d.line(P([(200, 150), (300, 150)]), fill=rgb("#10B981"), width=5 * K)
    d.rounded_rectangle(P([(122, 220), (150, 360)]), radius=12 * K, fill=rgb("#172554"))
    d.rounded_rectangle(P([(350, 220), (378, 360)]), radius=12 * K, fill=rgb("#172554"))
    d.text((250 * K, 225 * K), "30L", font=font(46), fill=rgb("#10B981"), anchor="mm")
    return finalizar(img, "30 L", halo)


def acc_medias(talla):
    halo = rgb("#DC2626")
    img = fondo(halo)
    sombra(img, (110, 392, 400, 428))
    d = ImageDraw.Draw(img)
    for cx, col in ((190, "#F8FAFC"), (310, "#F8FAFC")):
        pts = [(cx - 42, 90), (cx + 42, 90), (cx + 42, 280), (cx + 90, 330), (cx + 80, 388), (cx - 20, 388), (cx - 42, 280)]
        d.polygon(P(pts), fill=rgb(col), outline=rgb("#CBD5E1"))
        for y, c in ((100, "#DC2626"), (140, "#1D4ED8"), (180, "#DC2626")):
            d.polygon(P([(cx - 42, y), (cx + 42, y), (cx + 42, y + 22), (cx - 42, y + 22)]), fill=rgb(c))
        d.polygon(P([(cx + 42, 330), (cx + 84, 336), (cx + 80, 388), (cx + 30, 388)]), fill=rgb("#111827"))
    return finalizar(img, talla, halo)


def acc_pelota_mini():
    halo = rgb("#F59E0B")
    img = fondo(halo)
    cx, cy, R = 250 * K, 245 * K, 78 * K
    sh = Image.new("L", (W, W), 0)
    ImageDraw.Draw(sh).ellipse([cx - R * 0.95, 380 * K, cx + R * 0.95, 400 * K], fill=120)
    sh = sh.filter(ImageFilter.GaussianBlur(14))
    img.paste(Image.new("RGB", (W, W), (30, 41, 59)), (0, 0), sh)
    d = ImageDraw.Draw(img)
    d.rounded_rectangle(P([(170, 342), (330, 366)]), radius=10 * K, fill=rgb("#0F172A"))
    d.polygon(P([(200, 366), (300, 366), (330, 392), (170, 392)]), fill=rgb("#1E293B"))
    img = dibujar_balon(img, cx, cy - 6 * K, R, rgb("#FACC15"), rgb("#DC2626"), rgb("#B45309"))
    return finalizar(img, "Pausa Mini", halo)


def finalizar(img, texto, halo):
    img = img.resize((500, 500), Image.LANCZOS)
    etiqueta_img(img, texto, halo)
    return img


# ----------------------------------------------------------------------------
# DESPACHO SEGUN EL NOMBRE DEL PRODUCTO
# ----------------------------------------------------------------------------
def norm(s):
    s = s.lower()
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ñ", "n")):
        s = s.replace(a, b)
    return s


def generar(sku, nombre, talla):
    n = norm(nombre)
    if sku.startswith("CAM"):
        for clave in EQUIPOS:
            if clave in n:
                return camiseta(clave, talla)
    if sku.startswith("BOT"):
        if "cafe" in n:
            return botin("cafe", talla)
        if "sala" in n or "futbolito" in n:
            return botin("sala", talla)
        return botin("cesped", talla)
    if sku.startswith("BAL"):
        for clave in BALONES:
            if clave in n:
                return balon(clave)
    if sku.startswith("CAN"):
        mod = "tobillera" if "tobillera" in n else ("escolar" if "escolar" in n else "pro")
        t = "chica" if "chica" in n else ("mediana" if "mediana" in n else "grande")
        return canillera(mod, t)
    if sku.startswith("GUA"):
        mod = "pro" if "pro grip" in n else ("entrenamiento" if "entrenamiento" in n else "juvenil")
        return guante(mod, talla)
    if sku.startswith("ACC"):
        if "rodillera" in n:
            return acc_rodillera()
        if "cinta" in n:
            return acc_cinta()
        if "bomba" in n:
            return acc_bomba()
        if "morral" in n:
            return acc_morral()
        if "medias" in n:
            return acc_medias(talla)
        if "pelota" in n:
            return acc_pelota_mini()
    raise ValueError("Sin plantilla de imagen para " + sku + " " + nombre)


def main():
    os.makedirs(OUT, exist_ok=True)
    n = 0
    with open(CSV, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            sku, nombre, talla = row["sku"], row["producto"], (row.get("talla") or "").strip()
            if not talla:
                m = re.search(r"Talla\s+(\S+)", nombre)
                talla = m.group(1) if m else ""
            if sku.startswith("ACC") and "medias" in norm(nombre):
                m = re.search(r"Talla\s+(\S+)", nombre)
                talla = m.group(1) if m else talla
            img = generar(sku, nombre, talla)
            img.convert("RGB").save(os.path.join(OUT, sku + ".jpg"), quality=92)
            n += 1
    print("Imagenes de producto generadas:", n, "->", OUT)


if __name__ == "__main__":
    main()
