"""Dibuja la figura del ejemplo: la proporcion de uso del color.

    python ejemplo/figura.py

Deja ejemplo/proporcion.png a 300 ppp. Es lo unico del repositorio que quiere
matplotlib; ooxmlkit no lo usa ni lo necesita, porque recibe el PNG ya compuesto.

Se dibuja a 15,5 cm y se inserta a 15,0 cm. Los cuerpos de texto de aqui estan
calibrados para esa reduccion: cambiar el ancho de insercion sin recalibrarlos
deja los rotulos ilegibles.
"""
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager as fm  # noqa: E402

# Anclado al archivo y no al directorio de trabajo, para que corra desde
# cualquier sitio y no solo desde la raiz del repositorio.
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from ignacia import PALETA as P, PROPORCION  # noqa: E402  todo en un solo sitio

ANCHO_CM, ALTO_CM = 15.5, 4.2

# Las familias de Ignacia, si estan instaladas. matplotlib no lee el registro de
# fuentes de Windows, asi que hay que pasarle los archivos por ruta. Sin ellas la
# figura sale igual, con la tipografia por defecto, y el aviso lo dice en vez de
# dejar que el cambio pase inadvertido.
for carpeta in (os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft",
                             "Windows", "Fonts"),
                os.path.join(os.environ.get("WINDIR", ""), "Fonts")):
    if os.path.isdir(carpeta):
        for archivo in os.listdir(carpeta):
            if archivo.lower().endswith((".ttf", ".otf")):
                try:
                    fm.fontManager.addfont(os.path.join(carpeta, archivo))
                except Exception:
                    pass

hay = {f.name for f in fm.fontManager.ttflist}
SANS = "Nunito Sans" if "Nunito Sans" in hay else None
MONO = "Space Mono" if "Space Mono" in hay else None
if SANS is None or MONO is None:
    print("aviso: faltan las familias de Ignacia, la figura sale con la "
          "tipografia por defecto de matplotlib")

RELLENO = {"Neutros": P["superficie"], "Primario": P["primario"],
           "Señal": P["senal"]}
TRAMOS = [(nombre, valor, RELLENO[nombre]) for nombre, valor in PROPORCION]

fig, ax = plt.subplots(figsize=(ANCHO_CM / 2.54, ALTO_CM / 2.54))
fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis("off")

# La barra va a sangre, alineada con la caja de texto. Los rotulos NO van sobre
# cada tramo: con 2 % y 8 % se pisan entre si y se salen del lienzo. Van en una
# fila de leyenda debajo, que ademas se lee en el mismo orden.
x = 0
for _, valor, relleno in TRAMOS:
    # El tramo de neutros es casi del color de la hoja: lleva filete para no
    # desaparecer, que es la misma regla que rige las muestras de color.
    ax.add_patch(plt.Rectangle(
        (x, 58), valor, 30, facecolor=relleno, linewidth=0.8,
        edgecolor=P["borde"] if relleno == P["superficie"] else "none"))
    x += valor

for i, (nombre, valor, relleno) in enumerate(TRAMOS):
    xl = i * 33.0
    ax.add_patch(plt.Rectangle(
        (xl, 33), 3.2, 7.5, facecolor=relleno, linewidth=0.8,
        edgecolor=P["borde"] if relleno == P["superficie"] else "none"))
    ax.text(xl + 5.2, 36.5, nombre.upper(), ha="left", va="center", fontsize=8,
            color=P["texto_sec"], fontfamily=SANS)
    # La cifra va a una distancia fija del inicio de su columna, no pegada al
    # rotulo: asi las tres quedan alineadas entre si y no siguen al largo de la
    # palabra que las precede.
    ax.text(xl + 23.5, 36.5, f"{valor} %", ha="left", va="center", fontsize=9,
            color=P["texto"], fontfamily=MONO, fontweight="bold")

ax.text(0, 12, "La proporción es la regla. El color de señal se raciona, y por "
               "eso significa algo cuando aparece.",
        ha="left", va="center", fontsize=8, color=P["texto_sec"],
        fontfamily=SANS)

salida = os.path.join(AQUI, "proporcion.png")
fig.savefig(salida, dpi=300, facecolor="#FFFFFF")
print(salida)
