# elecciones-peru-2026-ml

[![Conference](https://img.shields.io/badge/LACCEI-2026%20Student%20Paper-003366.svg)](https://laccei.org/)
[![Python](https://img.shields.io/badge/Python-3.11%2B-3776AB.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Reproducibility](https://img.shields.io/badge/CRISP--DM-random__state%3D42-orange.svg)](#reproducibilidad-y-reglas-laccei)

Repositorio oficial de datos homologados, código reproducible y manuscrito del artículo científico:

> **Clusterización territorial de las preferencias electorales en la primera vuelta presidencial del Perú 2026 aplicando algoritmos de Machine Learning no supervisado sobre datos de la ONPE**  
> *(Territorial Clustering of Electoral Preferences in the 2026 Peruvian Presidential First Round Using Unsupervised Machine Learning on Official ONPE Data)*

- **Autor:** Luis Miguel Tirabantes Córdova (`u17307151@utp.edu.pe`)
- **Docente Asesor (*Faculty Mentor*):** Omar Santillán Aching (`c25611@utp.edu.pe`)
- **Afiliación:** Department of Engineering, Universidad Tecnológica del Perú (UTP), Lima, Perú
- **Conferencia:** *24th LACCEI International Multi-Conference for Engineering, Education, and Technology (LACCEI 2026)* — Modalidad Estudiantil

---

## 📌 Descripción del Proyecto

Este repositorio implementa un *pipeline* computacional modular, auditable y reproducible bajo el estándar **CRISP-DM** para caracterizar la estructura territorial del voto en la primera vuelta presidencial del Perú 2026 sobre la totalidad de las **1,890 circunscripciones distritales** del país, utilizando datos abiertos oficiales de la **ONPE** y **JNE**.

El proyecto aborda tres desafíos metodológicos centrales:
1. **Homologación geográfica INEI vs. RENIEC/ONPE:** conciliación algorítmica determinista de los códigos UBIGEO (incluyendo el desplazamiento de prefijos departamentales `08`–`25` y los 23 distritos históricamente conflictivos), logrando el **100.00 %** de integridad analítica (`N = 1,890` distritos) y emparejando **1,832 polígonos** (`96.93 %`) en la capa GeoJSON para cartografía coroplética.
2. **Reducción de dispersión dimensional:** transición desde la matriz cruda de 35 candidaturas válidas (*Esquema A*) hacia una representación compacta (*Esquema B*) con las 6 fuerzas de representatividad sustantiva ($\ge 4\,\%$ medio o victorias distritales), masa agregada (`PCT_OTROS`), **Número Efectivo de Partidos (NEP)** y **Margen de Victoria Distrital (MV)**, además de contraste composicional (*Esquema C*, CLR + PCA).
3. **Evaluación multialgorítmica sin inferencia circular:** comparación sistemática de cinco familias no supervisadas (**K-Means**, **K-Medoids / PAM**, **HAC con enlace Ward**, **GMM Full** y **Spectral Clustering**) con métricas intrínsecas (*Silhouette*, *Calinski-Harabasz*, *Davies-Bouldin*) y concordancia inter-modelo (*ARI* y *NMI*).

---

## 📁 Estructura del Repositorio

```text
elecciones-peru-2026-ml/
├── dataset_onpe_1890_distritos.xlsx     # Matriz porcentual distrital homologada (1,890 × 35)
├── peru_distrital_simple.geojson        # Capa vectorial distrital para cartografía coroplética
├── pipeline_electoral_peru_2026.py      # Pipeline maestro CRISP-DM (clustering, métricas y JSON)
├── resultados_paper_laccei_2026.json    # Métricas reproducibles de los 28 modelos evaluados
├── resultados_clusters_laccei_2026.xlsx # Asignación distrital de clústeres y centroides
├── regenerar_fig1_laccei.py             # Script reproducible de la Fig. 1 (arquitectura CRISP-DM)
├── regenerar_fig3_fig4_laccei.py        # Script reproducible de Fig. 3 (PCA/firmas) y Fig. 4 (mapa)
├── generar_docx_laccei_2026.py          # Generador automático del manuscrito Word desde LaTeX
├── Paper_LACCEI_2026.tex                # Código fuente completo del artículo (formato IEEE/LACCEI)
├── Paper_LACCEI_2026.pdf                # Manuscrito compilado (10 páginas, límite oficial 6–10)
├── Paper_LACCEI_2026.docx               # Versión Word camera-ready generada desde el .tex
├── IEEEtran.cls                         # Clase oficial IEEEtran para compilación LaTeX
├── figuras_laccei/                      # Figuras finales en PNG de alta resolución (Fig. 1–4)
├── requirements.txt                     # Dependencias exactas de Python
├── CITATION.cff                         # Metadatos estándar de citación académica
├── LICENSE                              # Licencia MIT
└── README.md                            # Documentación técnica y guía de reproducción
```

---

## ⚙️ Instalación y Requisitos

Se recomienda **Python 3.11+** en un entorno virtual aislado:

```bash
# 1. Clonar el repositorio institucional
git clone https://github.com/Lumitico/elecciones-peru-2026-ml.git
cd elecciones-peru-2026-ml

# 2. Crear y activar entorno virtual
python -m venv venv
source venv/bin/activate        # En Linux/macOS
# venv\Scripts\activate         # En Windows

# 3. Instalar dependencias
pip install -r requirements.txt
```

---

## 🚀 Reproducción del Estudio

### 1. Ejecutar el pipeline principal de Machine Learning
Genera las métricas comparativas (`resultados_paper_laccei_2026.json`), asignaciones por distrito (`resultados_clusters_laccei_2026.xlsx`) y curvas de validación:

```bash
python pipeline_electoral_peru_2026.py
```

### 2. Regenerar las figuras finales del artículo (`Fig. 1`, `Fig. 3`, `Fig. 4`)
```bash
python regenerar_fig1_laccei.py
python regenerar_fig3_fig4_laccei.py
```

### 3. Compilar el manuscrito LACCEI (PDF y DOCX)
```bash
# Compilar PDF (10 páginas en formato IEEE/LACCEI)
tectonic Paper_LACCEI_2026.tex

# Regenerar versión Word (.docx) directamente desde Paper_LACCEI_2026.tex
python generar_docx_laccei_2026.py
```

---

## 📊 Resumen de Resultados Verificados

- **Universo analizado:** `1,890` circunscripciones distritales (`100.00 %` de cobertura nacional).
- **Mejora del Esquema B frente a Línea Base (35 variables crudas):**
  - *Silhouette Score:* de `0.1409` ($k=2$) a **`0.2841`** ($k=2$) y **`0.2650`** ($k=3$).
  - *Calinski-Harabasz:* de `208.39` ($k=2$) a **`824.39`** ($k=2$) y **`702.34`** ($k=3$).
  - *Davies-Bouldin:* reducción de `2.5742` ($k=2$) a **`1.3214`** ($k=2$) y **`1.4868`** ($k=3$).
- **Concordancia inter-modelo ($k=3$):** K-Means vs. HAC Ward alcanza `ARI = 0.5971` y `NMI = 0.6133`.
- **Tipología territorial seleccionada (`K-Means`, $k=3$, PCA 2D = `65.90 %` varianza acumulada):**
  1. **Clúster 0 — Hegemonía Andino-Rural (`n = 672`, `35.56 %`, azul):** JPP `47.70 %`, MV `32.70 pp`, NEP `3.89`.
  2. **Clúster 1 — Hiperfragmentación y Dispersión (`n = 604`, `31.96 %`, verde):** NEP `8.07`, MV `6.80 pp` (incluye los 5 empates por redondeo al tercer decimal).
  3. **Clúster 2 — Competitividad Urbano-Costera (`n = 614`, `32.49 %`, rojo):** FP `27.99 %`, MV `13.69 pp`, NEP `6.81`.

---

## 🔬 Reproducibilidad y Reglas LACCEI

- **Determinismo computacional:** todas las ejecuciones fijan semilla pseudoaleatoria (`random_state = 42`, `n_init = 10`, `k-means++`).
- **Trazabilidad de empates:** el conteo algorítmico (`1,107` JPP / `583` FP) frente al cómputo nominal de ONPE (`1,108` / `583`) se documenta íntegramente en la Tabla III del manuscrito (empate al tercer decimal en el distrito de Ingenio, Junín).
- **Alcance inferencial y ética de datos:** los datos provienen exclusivamente de fuentes públicas agregadas (ONPE/JNE/INEI); las conclusiones describen **circunscripciones distritales** y no comportamientos individuales, evitando explícitamente la **falacia ecológica**.

---

## 📖 Cómo Citar

```bibtex
@inproceedings{tirabantes2026clustering,
  author    = {Tirabantes C{\'o}rdova, Luis Miguel},
  title     = {Territorial Clustering of Electoral Preferences in the 2026 Peruvian Presidential First Round Using Unsupervised Machine Learning on Official ONPE Data},
  booktitle = {Proceedings of the 24th LACCEI International Multi-Conference for Engineering, Education, and Technology (LACCEI 2026)},
  year      = {2026},
  address   = {Santiago, Chile},
  note      = {Faculty Mentor: Omar Santill{\'a}n Aching}
}
```

---

## 📜 Licencia

Este proyecto se distribuye bajo la Licencia **MIT**. Consulta el archivo [LICENSE](LICENSE) para más detalles.
