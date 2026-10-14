# Convertidor de PDF a DOC y DOCX

Aplicación web que convierte PDF (y también DOCX, DOC, ODT, RTF, TXT, JPG y PNG) a DOCX o DOC.
Toda la conversión ocurre en el navegador del usuario, dentro de un Web Worker. Ningún archivo
se sube a un servidor.

- Producción: https://convertidor-pdf-docx.vercel.app
- Repositorio: https://github.com/anormalseller-beep/convertidor-pdf-docx

## Uso

1. Elegir o arrastrar un archivo.
2. Elegir el formato de salida: DOCX (por defecto) o DOC.
3. Opcional: OCR para páginas escaneadas.
4. Convertir. La barra muestra el porcentaje y la página actual sobre el total.
5. Descargar. Debajo aparece el informe de verificación (QA) y el registro de avisos.

Con sesión de Google iniciada se guarda un historial (solo metadatos en Firestore). El archivo
convertido queda en IndexedDB del mismo navegador para volver a descargarlo.

## Arquitectura

```
Navegador
 ├─ Interfaz (Next.js App Router, React, Tailwind)         src/components/Convertidor.tsx
 └─ Web Worker de conversión                              src/workers/conversor.worker.ts
     ├─ pdf.js (pdfjs-dist 4.10.38) sin sub-worker; PDFs de más de 64 MB se leen por rangos
     │   desde el disco (PDFDataRangeTransport) en lugar de cargarlos enteros en memoria
     ├─ Intérprete de operadores                          src/lib/motor/interprete.ts
     │   reproduce el estado gráfico de pdf.js (CTM, matriz de texto, colores) y produce:
     │     - texto real: fuente, tamaño, negrita, cursiva, color y coordenadas exactas
     │     - imágenes con su caja real (ancho, alto, posición, rotación, volteo)
     │     - rectángulos y líneas (bordes de tablas, recuadros, subrayados) como formas
     │     - curvas complejas como imagen transparente recortada a su caja
     ├─ Escritor DOCX en streaming                         src/lib/motor/escritor-docx.ts
     ├─ Escritor DOC (Word 2003 XML)                       src/lib/motor/escritor-doc.ts
     ├─ QA: relee el archivo generado por partes           src/lib/motor/qa.ts
     ├─ OCR opcional con Tesseract.js (spa + eng)
     └─ Formatos de flujo (TXT, RTF, ODT, DOCX, DOC)       src/lib/flujo/
Vercel (plan Hobby): solo sirve la aplicación estática. No hay funciones de servidor.
Firebase (plan Spark): Auth con Google + Firestore para el historial de metadatos.
```

### Cómo se conserva el diseño del PDF

| Problema típico | Solución aplicada |
| --- | --- |
| Imágenes apiladas al final | Cada imagen se ancla a la página (`wp:anchor`, relativo a la página) en su posición X/Y exacta. |
| Imágenes deformadas | El tamaño sale de la matriz de transformación del PDF (escala 1.0) y se escribe en EMU (1 pt = 12 700 EMU) con `noChangeAspect`. Nunca `width: 100%`. |
| Texto que se reacomoda o se sale del margen | Cada línea es un cuadro de texto anclado a la página con `wrap="none"`: Word no puede partirla ni moverla. El ancho se ajusta con espaciado de caracteres (y escala horizontal si hace falta) medido contra la fuente sustituta. |
| Imágenes que tapan el texto | El texto siempre va en una capa superior (`relativeHeight` mayor); imágenes y formas mantienen entre sí el orden de pintado del PDF. |
| Saltos de página mal puestos | Una sección por página con salto de sección "página siguiente". El flujo de cada página es un único párrafo de 1 pt, así nada puede desbordar a otra página. |
| Tamaño u orientación distintos | Cada sección usa el tamaño exacto de su página (Letter, A4 o personalizado) y su orientación. Páginas mixtas verticales/horizontales quedan cada una con la suya. |
| Tablas | Las líneas y rectángulos del PDF se escriben como formas nativas de Word (editables y nítidas), en su posición exacta. |

Fuentes: cada familia del PDF se mapea a una fuente estándar con métricas parecidas
(Arial, Times New Roman, Calibri, Cambria, Georgia, Courier New, Verdana, Tahoma, etc.;
ver `src/lib/motor/fuentes.ts`). Si la fuente embebida no se reconoce por nombre, se usa la
familia genérica que informa pdf.js (serif, sans-serif o monospace). No hay una fuente global.

### Verificación (QA)

Tras cada conversión de PDF o imagen, el Worker relee el archivo generado en streaming y compara:
cantidad de páginas (secciones), tamaño de cada página y área total, y cantidad de imágenes
colocadas. Cualquier diferencia se muestra como `ERROR` en el informe.

### DOC (Word 97-2003)

No existe una forma gratuita de generar el formato binario de Word 97 en el navegador (haría falta
LibreOffice en un servidor). La salida `.doc` es un documento XML de Word 2003 (WordprocessingML)
con extensión `.doc`; Word y LibreOffice lo abren como documento de Word. Usa cuadros de texto e
imágenes VML con posición absoluta. La interfaz avisa que DOCX es la opción recomendada.

### Otros formatos de entrada

TXT, RTF, ODT y DOCX se convierten como documento de flujo con la librería `docx` (DOCX) o como
Word 2003 XML (DOC): párrafos, negrita, cursiva, subrayado, tamaños, colores, títulos, listas,
tablas e imágenes. Un DOCX pedido como DOCX se entrega sin cambios. Del DOC binario se extrae el
texto (tabla de piezas del formato Word 97). JPG y PNG se colocan como una página del tamaño de
la imagen (96 ppp), con OCR opcional.

### Librerías

| Librería | Uso |
| --- | --- |
| pdfjs-dist 4.10.38 | Lectura del PDF y lista de operadores |
| fflate | ZIP del DOCX en streaming, PNG, relectura del QA |
| docx | Escritura de los formatos de flujo (TXT, RTF, ODT, DOCX) |
| mammoth | Lectura de DOCX de entrada |
| cfb | Lectura del DOC binario de entrada |
| tesseract.js | OCR en el navegador |
| firebase | Auth (Google) y Firestore |

`jszip` y `html-docx-js` no se usan: JSZip necesita todo el contenido en memoria y no sirve para
10 000 páginas (fflate escribe el ZIP por partes), y `html-docx-js` genera DOCX con `altChunk`
HTML, que pierde el posicionamiento. `docx-preview` no se incluyó: no dibuja cuadros de texto
anclados y la vista previa sería engañosa; el QA visual se hizo con Microsoft Word (ver Pruebas).

## Límites del plan gratuito y cómo se respetan

El proyecto no usa ningún servicio de pago y no requiere tarjeta.

| Servicio | Límite gratuito (consultar valores vigentes en la web del proveedor) | Cómo se respeta |
| --- | --- | --- |
| Vercel Hobby | 100 GB/mes de transferencia, 1 M de solicitudes de borde/mes, uso no comercial | El sitio es estático (~3 MB incluyendo las fuentes estándar de pdf.js). Los archivos de los usuarios nunca pasan por Vercel. |
| Firebase Auth (Spark) | Inicio de sesión con Google sin costo | Solo inicio de sesión con Google. |
| Cloud Firestore (Spark) | 1 GiB almacenado, 50 000 lecturas, 20 000 escrituras y 20 000 borrados por día | Una escritura por conversión con sesión iniciada; el historial lee como máximo 50 documentos. Sin archivos, solo metadatos (nombre, fecha, páginas, tamaño). |
| Cloud Storage | Desde octubre de 2024 los buckets nuevos de Firebase exigen el plan Blaze | No se usa Storage. Los archivos se guardan en IndexedDB del navegador (máximo 20). |
| Tesseract.js | Gratuito (Apache 2.0) | El motor y el modelo de idioma se descargan del CDN público jsDelivr solo cuando el usuario activa OCR. |

## Limitaciones conocidas

- Memoria del navegador: el PDF de entrada debe caber en la memoria del navegador si pesa menos
  de 64 MB; por encima se lee por rangos desde el disco. El DOCX de salida se arma en bloques
  (Blob), que el navegador puede guardar en disco. Por encima de 200 MB la interfaz pide
  confirmación. Navegadores móviles tienen bastante menos memoria.
- Rendimiento medido (PC de escritorio, Chrome): 10 000 páginas con texto, una imagen y un
  recuadro por página en 176 s, DOCX de 15,2 MB, QA sin diferencias.
- Word abre documentos con decenas de miles de cuadros de texto, pero tarda en paginarlos.
- La fuente sustituta no es idéntica a la del PDF; el ancho de cada línea se iguala con
  espaciado, por lo que el texto ocupa el mismo lugar, aunque las letras puedan diferir un poco.
- Texto vertical (CJK) e imágenes inclinadas se colocan sin inclinación; degradados y patrones
  del PDF no se reproducen; el texto invisible (capa OCR de PDFs ya escaneados) no se copia.
- El OCR reemplaza la imagen escaneada por texto editable; las fotos dentro de una página
  escaneada se pierden en ese modo (sin OCR la página se conserva como imagen).
- DOC: ver la sección DOC. DOC binario de entrada: solo texto.
- RTF de entrada: las tablas se convierten en texto separado por tabulaciones.
- Formatos de flujo: no conservan el diseño exacto (no lo tienen definido).

## Desarrollo

```bash
npm install
npm run dev
```

Variables de entorno (ver `.env.example`): `NEXT_PUBLIC_FIREBASE_*`. Sin ellas la app funciona
igual, sin historial.

### Pruebas

```bash
npx tsx scripts/generar-pdf-prueba.ts            # PDF de 12 páginas con imágenes, tablas, texto
npx tsx scripts/probar-motor.ts pruebas/muestra-12-paginas.pdf docx
npx tsx scripts/generar-pdf-grande.ts 10000      # PDF de 10 000 páginas
npx tsx scripts/probar-flujo.ts                  # TXT, RTF, ODT, DOCX y DOC de entrada
```

Comparación visual con Microsoft Word (Windows con Word instalado, solo QA local):

```bash
powershell -File scripts/word-a-pdf.ps1 pruebas/muestra-12-paginas.docx pruebas/word-docx.pdf
python scripts/comparar.py pruebas/muestra-12-paginas.pdf pruebas/word-docx.pdf pruebas/comparacion-docx
node scripts/medir-desvio.mjs pruebas/muestra-12-paginas.pdf pruebas/word-docx.pdf
```

Resultados en `pruebas/`:

- `comparacion-docx/` y `comparacion-doc/`: PDF original y DOCX/DOC abierto en Microsoft Word,
  lado a lado, las 12 páginas.
- Desvío medido sobre 270 líneas: posición horizontal 0,02 pt, vertical 0,02 pt, ancho 0,05 pt
  (medianas). Páginas 12 de 12, imágenes 11 de 11, tamaño de página idéntico.

## Registro de cambios por pedido

### 2026-10-13 — Pedido inicial
- Proyecto Next.js 16 (App Router, TypeScript, Tailwind 4) con motor de conversión propio en Web Worker.
- Intérprete de operadores de pdf.js, escritores DOCX (streaming) y DOC (Word 2003 XML), QA por relectura.
- Tablas y recuadros como formas nativas; imágenes repetidas guardadas una sola vez.
- OCR con Tesseract.js, PDFs con contraseña, cancelación, aviso de archivos grandes.
- Entradas TXT, RTF, ODT, DOCX, DOC, JPG, PNG.
- Historial opcional con Firebase Auth (Google) + Firestore; archivos en IndexedDB.
- Pruebas: muestra de 12 páginas comparada en Microsoft Word; 10 000 páginas en el navegador.
