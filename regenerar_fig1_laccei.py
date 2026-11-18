#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera la Figura 1 (arquitectura metodologica CRISP-DM) al TAMANO FINAL de
impresion (ancho de columna = 3.45 in), con tipografia legible (>= 6.5 pt
reales sobre el papel). El detalle de cada fase se documenta en el pie de
figura y en las subsecciones de la Metodologia, por lo que el recuadro
solo rotula las seis fases.
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import numpy as np

FIG_DIR = "figuras_laccei"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 7.0,
})

FASES = [
    "1. Comprensión\ndel dominio",
    "2. Comprensión\nde los datos",
    "3. Preparación y\ncaracterísticas",
    "4. Modelamiento\nmultialgorítmico",
    "5. Evaluación y\nvalidación",
    "6. Despliegue y\nreproducibilidad",
]

AZUL = "#1f4e79"
FONDO = "#eef4fb"

fig, ax = plt.subplots(figsize=(3.45, 2.20))
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")

w, h = 0.432, 0.235
xs = [0.025, 0.543]
ys = [0.72, 0.385, 0.05]
pos = [(xs[0], ys[0]), (xs[1], ys[0]),
       (xs[1], ys[1]), (xs[0], ys[1]),
       (xs[0], ys[2]), (xs[1], ys[2])]

for texto, (x, y) in zip(FASES, pos):
    ax.add_patch(FancyBboxPatch((x, y), w, h,
                                boxstyle="round,pad=0.006,rounding_size=0.02",
                                fc=FONDO, ec=AZUL, lw=0.9))
    ax.text(x + w / 2, y + h / 2, texto, ha="center", va="center",
            fontsize=6.6, weight="bold", color=AZUL, linespacing=1.25)

def flecha(p0, p1):
    ax.annotate("", xy=p1, xytext=p0,
                arrowprops=dict(arrowstyle="-|>", color=AZUL, lw=0.9,
                                shrinkA=0, shrinkB=0,
                                mutation_scale=6))

y_c1 = ys[0] + h / 2
y_c2 = ys[1] + h / 2
y_c3 = ys[2] + h / 2
flecha((xs[0] + w, y_c1), (xs[1], y_c1))          # 1 -> 2
flecha((xs[1] + w / 2, ys[0]), (xs[1] + w / 2, ys[1] + h))   # 2 -> 3
flecha((xs[1], y_c2), (xs[0] + w, y_c2))          # 3 -> 4
flecha((xs[0] + w / 2, ys[1]), (xs[0] + w / 2, ys[2] + h))   # 4 -> 5
flecha((xs[0] + w, y_c3), (xs[1], y_c3))          # 5 -> 6

fig.subplots_adjust(0, 0, 1, 1)
fig.savefig(f"{FIG_DIR}/fig1_crispdm.png", dpi=400)
plt.close(fig)
print("[*] Fig. 1 regenerada a 3.45 x 2.20 in")
