#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Regenera la Figura 3 (espacio latente PCA + firmas de centroides) y la Figura 4
(mapa coropletico distrital) al TAMANO FINAL de impresion, con tipografias
legibles (>= 6.5 pt reales sobre el papel).

Fuente de datos: resultados_clusters_laccei_2026.xlsx (1,890 distritos) y
peru_distrital_simple.geojson. No recalcula clusters: usa CLUSTER_K3 ya
validado, por lo que las cifras del paper permanecen inalteradas.

Verificacion incluida: la varianza explicada del PCA debe reproducir
43.71 % / 22.19 % (acumulada 65.90 %), tal como se reporta en el articulo.
"""
import json
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from matplotlib.patches import Patch

import openpyxl

FIG_DIR = "figuras_laccei"
XLSX = "resultados_clusters_laccei_2026.xlsx"
GEOJSON = "peru_distrital_simple.geojson"

PARTIES = [
    "PCT_JUNTOS_POR_EL_PERÚ",
    "PCT_FUERZA_POPULAR",
    "PCT_PARTIDO_CÍVICO_OBRAS",
    "PCT_AHORA_NACIÓN_-_AN",
    "PCT_RENOVACIÓN_POPULAR",
    "PCT_PARTIDO_DEL_BUEN_GOBIERNO",
]
VARS_B = PARTIES + ["PCT_OTROS", "NEP", "MARGEN_VICTORIA"]
COLORES = ["#1f77b4", "#2ca02c", "#d62728"]

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 7.5,
    "axes.titlesize": 7.8,
    "axes.labelsize": 7.5,
    "xtick.labelsize": 7.0,
    "ytick.labelsize": 7.0,
    "legend.fontsize": 6.6,
    "axes.linewidth": 0.6,
})

# ── 1. Carga de datos ────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(XLSX, read_only=True)
ws = wb[wb.sheetnames[0]]
rows = ws.iter_rows(values_only=True)
hdr = list(next(rows))
idx = {h: i for i, h in enumerate(hdr)}
data = [r for r in rows if r[0] is not None]
print(f"[*] Distritos cargados: {len(data)}")

X = np.array([[float(r[idx[c]]) for c in VARS_B] for r in data])
cluster = np.array([int(r[idx["CLUSTER_K3"]]) for r in data])
ubigeo = [str(r[idx["UBIGEO_INEI"]]).zfill(6) for r in data]


def pca_2d(mat):
    """PCA por SVD sobre matriz estandarizada (equivalente a sklearn PCA)."""
    Z = (mat - mat.mean(0)) / mat.std(0, ddof=0)
    U, S, Vt = np.linalg.svd(Z - Z.mean(0), full_matrices=False)
    var = S ** 2 / (S ** 2).sum() * 100
    return Z @ Vt[:2].T, var, Vt[:2]


P, var, comps = pca_2d(X)
print(f"[*] Varianza explicada: CP1 = {var[0]:.2f} %, CP2 = {var[1]:.2f} %, "
      f"acumulada = {var[0] + var[1]:.2f} %")
assert abs((var[0] + var[1]) - 65.9) < 0.2, "La varianza no reproduce la del articulo"

# ── 2. Figura 3: 7.0 x 2.55 in (se inserta a \textwidth = 7.16 in) ───────────
fig, (ax_pca, ax_bar) = plt.subplots(
    1, 2, figsize=(7.0, 2.75), gridspec_kw={"width_ratios": [1.0, 1.05]}
)

nombres = [
    "Clúster 0: Hegemonía JPP",
    "Clúster 1: Hiperfragmentación",
    "Clúster 2: Competencia FP",
]
for c in range(3):
    m = cluster == c
    ax_pca.scatter(P[m, 0], P[m, 1], s=2.2, alpha=0.55, color=COLORES[c],
                   label=f"{nombres[c]} (n = {m.sum()})", linewidths=0)
ax_pca.set_xlabel(f"Componente principal 1 ({var[0]:.1f} % var.)")
ax_pca.set_ylabel(f"Componente principal 2 ({var[1]:.1f} % var.)")
ax_pca.set_ylim(-5.4, 5.0)
ax_pca.set_title(f"(a) Proyección PCA de los 1,890 distritos ($k$ = 3)\n"
                 f"Varianza acumulada: {var[0] + var[1]:.1f} %", weight="bold")
ax_pca.grid(True, ls=":", lw=0.4, alpha=0.5)
ax_pca.tick_params(length=2, width=0.5)
ax_pca.legend(loc="upper left", frameon=True, framealpha=0.92, handletextpad=0.25,
              borderpad=0.35, labelspacing=0.3, borderaxespad=0.4)

vars_plot = [PARTIES[0], PARTIES[1], PARTIES[2], PARTIES[3], "PCT_OTROS", "MARGEN_VICTORIA"]
nombres_plot = ["Juntos por el Perú", "Fuerza Popular", "Cívico Obras",
                "Ahora Nación", "Otros partidos", "Margen victoria"]
y = np.arange(len(vars_plot))
h = 0.26
for c in range(3):
    sub = X[cluster == c]
    vals = [sub[:, VARS_B.index(v)].mean() for v in vars_plot]
    ax_bar.barh(y + (c - 1) * h, vals, height=h, color=COLORES[c],
                edgecolor="white", linewidth=0.4, label=f"Clúster {c}")
ax_bar.set_yticks(y)
ax_bar.set_yticklabels(nombres_plot)
ax_bar.invert_yaxis()
ax_bar.set_xlabel("Promedio distrital (%)")
ax_bar.set_title("(b) Firma electoral de los centroides", weight="bold")
ax_bar.grid(True, axis="x", ls=":", lw=0.4, alpha=0.5)
ax_bar.tick_params(length=2, width=0.5)
ax_bar.legend(loc="lower right", frameon=True, framealpha=0.9, handletextpad=0.3,
              borderpad=0.3, labelspacing=0.25)

fig.tight_layout(pad=0.4, w_pad=1.2)
fig.savefig(f"{FIG_DIR}/fig3_espacio_latente.png", dpi=400)
plt.close(fig)
print("[*] Fig. 3 regenerada a 7.0 x 2.75 in")

# ── 3. Figura 4: mapa coropletico, 3.4 x 5.3 in (a \columnwidth) ─────────────
with open(GEOJSON, "r", encoding="utf-8") as f:
    geo = json.load(f)

cmap = dict(zip(ubigeo, cluster))
verts = {0: [], 1: [], 2: [], "nd": []}
distritos_emparejados = 0
sin_match = 0
for feat in geo["features"]:
    iddist = str((feat.get("properties") or {}).get("IDDIST", "")).zfill(6)
    if iddist in cmap:
        distritos_emparejados += 1
    else:
        sin_match += 1
    geom = feat.get("geometry") or {}
    coords = geom.get("coordinates")
    if not coords:
        continue
    c = cmap.get(iddist, "nd")
    polys = [coords] if geom.get("type") == "Polygon" else coords
    for poly in polys:
        if poly:
            verts[c].append(np.asarray(poly[0], dtype=float))

emparejados = distritos_emparejados
print(f"[*] Distritos emparejados: {emparejados} de {len(geo['features'])} features")

fig, ax = plt.subplots(figsize=(3.4, 5.3))
ax.set_aspect("equal")
estilos = [("nd", "#e8e8e8", "#c4c4c4", 0.10),
           (0, COLORES[0], "#0d3b66", 0.12),
           (1, COLORES[1], "#134f13", 0.12),
           (2, COLORES[2], "#7a1212", 0.12)]
for key, face, edge, lw in estilos:
    if verts[key]:
        ax.add_collection(PolyCollection(verts[key], facecolors=face,
                                         edgecolors=edge, linewidths=lw))
ax.set_xlim(-81.6, -68.4)
ax.set_ylim(-18.6, 0.4)
ax.set_xlabel("Longitud (°O)")
ax.set_ylabel("Latitud (°S)")
ax.tick_params(length=2, width=0.5)
ax.set_xticks([-80, -76, -72, -69])
ax.set_yticks([-18, -14, -10, -6, -2])
ax.grid(True, ls=":", lw=0.35, alpha=0.45)
ax.set_title(f"Clústeres electorales por distrito\n(K-Means, $k$ = 3; "
             f"{emparejados:,} de {len(data):,} distritos)", weight="bold", pad=6)

handles = [
    Patch(facecolor=COLORES[0], edgecolor="#0d3b66", label="Clúster 0: Hegemonía JPP"),
    Patch(facecolor=COLORES[1], edgecolor="#134f13", label="Clúster 1: Hiperfragmentación"),
    Patch(facecolor=COLORES[2], edgecolor="#7a1212", label="Clúster 2: Competencia FP"),
    Patch(facecolor="#e8e8e8", edgecolor="#c4c4c4",
          label=f"Sin correspondencia UBIGEO ({sin_match} dist.)"),
]
ax.legend(handles=handles, loc="lower left", frameon=True, framealpha=0.92,
          handlelength=0.9, borderpad=0.35, labelspacing=0.3, fontsize=6.3)

fig.tight_layout(pad=0.3)
fig.savefig(f"{FIG_DIR}/fig4_mapa_clusters.png", dpi=400)
plt.close(fig)
print("[*] Fig. 4 regenerada a 3.4 x 5.3 in")
