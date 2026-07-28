# Genera el PNG que inserta informe.py. Requiere matplotlib, que la libreria
# no usa: ooxmlkit inserta imagenes, no las dibuja.
#
#     python ejemplo/figura.py
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(AQUI))
from ooxmlkit import PALETA

CM = 1 / 2.54
anios = ["2021", "2022", "2023", "2024", "2025"]
valor = [38, 44, 51, 49, 63]

# 15,5 cm de diseno, 15,0 cm de insercion: los cuerpos de texto estan
# calibrados para esa reduccion.
fig, ax = plt.subplots(figsize=(15.5 * CM, 7.5 * CM), dpi=300)
barras = ax.bar(anios, valor, width=0.62, color=PALETA["azul"])
barras[-1].set_color(PALETA["naranja_prof"])

for x, v in zip(anios, valor):
    ax.text(x, v + 1.5, str(v), ha="center", va="bottom", size=7.5,
            color=PALETA["grafito"])

ax.set_ylim(0, 75)
ax.tick_params(axis="x", length=0, labelsize=7.5, colors=PALETA["grafito_med"])
ax.set_yticks([])
for lado in ("top", "right", "left"):
    ax.spines[lado].set_visible(False)
ax.spines["bottom"].set_color(PALETA["borde"])

fig.tight_layout(pad=0.3)
salida = os.path.join(AQUI, "serie.png")
fig.savefig(salida, dpi=300, facecolor="white")
print(salida)
