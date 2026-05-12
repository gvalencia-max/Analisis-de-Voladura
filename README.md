# 💥 SPC Blasting Analysis — Streamlit Web App

**Kuz-Ram Fragmentation · Sadovsky PPV · Statistical Process Control · Monte Carlo**

> Universidad Nacional del Altiplano — Ingeniería de Minas  
> Autor: Giovany Valencia

---

## 📋 Descripción

Aplicación web interactiva para el análisis estadístico de voladuras mineras. Integra cuatro modelos en una sola interfaz:

| Módulo | Descripción |
|--------|-------------|
| **Kuz-Ram** | Distribución granulométrica (Rosin-Rammler) |
| **Sadovsky PPV** | Atenuación de vibración de partículas (PPV) |
| **SPC (X̄-R Charts)** | Control estadístico de proceso + Reglas de Nelson |
| **Monte Carlo** | Análisis probabilístico con N simulaciones |

---

## 🚀 Instalación y ejecución

### 1. Requisitos previos
- Python 3.10 o superior
- pip

### 2. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 3. Ejecutar la aplicación

```bash
py -m streamlit run app.py
```

La aplicación se abrirá automáticamente en el navegador en `http://localhost:8501`

---

## 📁 Estructura de archivos

```
├── app.py              ← Aplicación principal Streamlit
├── requirements.txt    ← Bibliotecas necesarias
└── README.md           ← Este archivo
```

---

## 🖥️ Uso de la aplicación

### Panel lateral (Sidebar)
Ingrese todos los parámetros de entrada:

**Sección 1 — Modelo Kuz-Ram**
- `A` — Rock Factor (Blastability)
- `K` — Powder Factor (kg/m³)
- `Q` — Carga explosiva por tiro (kg)
- `RWS` — Resistencia relativa (ANFO=100)
- `d` — Diámetro de tiro (mm)
- `B` — Burden (m)
- `S` — Espaciamiento (m)
- `H` — Altura de banco (m)
- `W` — Desviación estándar de perforación (m)
- `BCL` — Longitud de carga de fondo (m)
- `CCL` — Longitud de carga de columna (m)

**Sección 2 — Modelo Sadovsky PPV**
- `K_s` — Coeficiente del sitio
- `α` — Exponente de atenuación
- `R` — Distancia al punto de control (m)
- `Qd` — Carga máxima por retardo (kg)
- `PPV_max` — Límite máximo admisible (mm/s)

**Sección 3 — Especificaciones SPC**
- `X50_target` — X50 objetivo (mm)
- `Tolerancia ±` — Ventana de especificación (mm)
- `n subgrupos` — Tamaño de subgrupo

**Sección 4 — Monte Carlo**
- Número de simulaciones: 1,000 – 50,000

### Datos históricos X50
- Use el dataset de ejemplo predeterminado (5 subgrupos), o
- Desactive "Use default example dataset" e ingrese sus propios datos

### Botón ▶ Run Analysis
Ejecuta todos los modelos y muestra los resultados organizados en 5 pestañas:

| Pestaña | Contenido |
|---------|-----------|
| 🪨 Kuz-Ram | Curva granulométrica + interpretación |
| 📳 Sadovsky PPV | Curva de atenuación + semáforo |
| 📈 SPC Charts | Cartas X̄-R + Capacidad de proceso |
| 🎲 Monte Carlo | Histogramas + Tornado + Estadísticas |
| 📥 Export | Descarga del reporte Excel (.xlsx) |

---

## 📊 Outputs

### Gráficas generadas
1. Distribución Rosin-Rammler (Kuz-Ram)
2. Curva de atenuación PPV (Sadovsky)
3. Carta X̄ (medias de subgrupos)
4. Carta R (rangos)
5. Capacidad de proceso (Cp, Cpk)
6. Histograma Monte Carlo — X50
7. Histograma Monte Carlo — PPV
8. Tornado chart — Sensibilidad X50
9. Tornado chart — Sensibilidad PPV

### Reporte Excel (4 hojas)
- `1_Input_Parameters` — Todos los parámetros ingresados
- `2_Summary_Results` — Resultados determinísticos + probabilísticos
- `3_SPC_Calculations` — Subgrupos, límites de control, estatus
- `4_MonteCarlo_Stats` — Estadísticas percentilares

---

## 📚 Referencias técnicas

- **Kuznetsov (1973)** & **Cunningham (1983, 1987)** — Modelo Kuz-Ram
- **Sadovsky (1945)** — Modelo de atenuación PPV
- **Shewhart / Nelson** — Reglas de control estadístico
- **NTP Perú / USBM / DIN 4150-3 / ISO 4866** — Límites de vibración

---

## 🛠️ Dependencias principales

| Biblioteca | Versión mínima | Uso |
|-----------|---------------|-----|
| `streamlit` | 1.35.0 | Framework web |
| `numpy` | 1.26.0 | Cálculos numéricos |
| `matplotlib` | 3.8.0 | Generación de gráficas |
| `scipy` | 1.13.0 | Distribuciones estadísticas |
| `openpyxl` | 3.1.0 | Exportación Excel |
| `pandas` | 2.2.0 | Tablas de resultados |

---

## ⚠️ Notas importantes

- La semilla aleatoria del Monte Carlo es fija (`seed=42`) para reproducibilidad.
- Para subgrupos, el mínimo requerido es **n ≥ 3** subgrupos.
- Los tamaños de subgrupo válidos para SPC son: 2, 3, 4, 5, 6, 7, 8, 10.
