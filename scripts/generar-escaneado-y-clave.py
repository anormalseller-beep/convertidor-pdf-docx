# Genera pruebas/escaneado-3-paginas.pdf (solo imágenes, como un escáner) y
# pruebas/protegido-clave-prueba123.pdf (cifrado con la contraseña "prueba123").
# Requiere: pip install pypdfium2 pillow pypdf   (solo para pruebas)
import pypdfium2 as pdfium
from pypdf import PdfReader, PdfWriter

src = pdfium.PdfDocument("pruebas/muestra-12-paginas.pdf")
imgs = [src[i].render(scale=200 / 72).to_pil().convert("L").convert("RGB") for i in (0, 1, 2)]
imgs[0].save("pruebas/escaneado-3-paginas.pdf", save_all=True, append_images=imgs[1:], resolution=200)

w = PdfWriter()
for p in PdfReader("pruebas/muestra-12-paginas.pdf").pages[:3]:
    w.add_page(p)
w.encrypt("prueba123", algorithm="AES-128")
with open("pruebas/protegido-clave-prueba123.pdf", "wb") as f:
    w.write(f)
print("OK")
