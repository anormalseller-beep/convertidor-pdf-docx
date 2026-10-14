# Arma imágenes lado a lado: PDF original vs. DOCX abierto en Microsoft Word (exportado a PDF).
# Requiere: pip install pypdfium2 pillow   (solo herramienta de QA local)
# Uso: python scripts/comparar.py original.pdf word.pdf carpeta [etiqueta]
import sys, os
import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont

orig, word, carpeta = sys.argv[1:4]
etiqueta = sys.argv[4] if len(sys.argv) > 4 else "DOCX abierto en Microsoft Word"
os.makedirs(carpeta, exist_ok=True)
a, b = pdfium.PdfDocument(orig), pdfium.PdfDocument(word)
print(f"Páginas: original {len(a)}, Word {len(b)}")
try:
    fuente = ImageFont.truetype("arial.ttf", 18)
except OSError:
    fuente = ImageFont.load_default()
ESC = 1.4
for i in range(len(a)):
    ia = a[i].render(scale=ESC).to_pil()
    ib = b[i].render(scale=ESC).to_pil() if i < len(b) else Image.new("RGB", ia.size, "#eee")
    pad, sep = 34, 24
    w = ia.width + ib.width + sep
    h = max(ia.height, ib.height) + pad
    out = Image.new("RGB", (w, h), "#cccccc")
    d = ImageDraw.Draw(out)
    d.text((8, 8), f"PDF original - página {i + 1} ({a[i].get_width():.0f} x {a[i].get_height():.0f} pt)", fill="#000", font=fuente)
    if i < len(b):
        d.text((ia.width + sep + 8, 8), f"{etiqueta} - página {i + 1} ({b[i].get_width():.0f} x {b[i].get_height():.0f} pt)", fill="#000", font=fuente)
    out.paste(ia, (0, pad))
    out.paste(ib, (ia.width + sep, pad))
    out.save(os.path.join(carpeta, f"pagina-{i + 1:02d}.png"))
print("OK", carpeta)
