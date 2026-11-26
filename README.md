# Análisis térmico de PCR microfluídico

Scripts en Python (Google Colab / Jupyter) para caracterizar el termociclado de los dispositivos PCR
microfluídicos del laboratorio (Newn1 a Newn6). Leen los archivos de cada corrida, separan cada ciclo
en sus fases, filtran la señal y calculan, ciclo por ciclo:

- **Rampas:** tasa de calentamiento y enfriamiento (°C/s), retardo y duración.
- **Holds:** duración frente a la programada, temperatura media frente al setpoint, desviación estándar,
  máximo, mínimo y sobrepaso.
- **Repetibilidad:** media ± desviación estándar ciclo tras ciclo, ciclos atípicos y ciclos fallidos.

Con esos resultados generan gráficas y un Excel resumen para comparar corridas, dispositivos,
controladores (PID estándar o adaptativo) y sensores (termopar o IR).

---

## Contenido del repositorio

```
├── v1/                         Primera iteración (dispositivo Newn1, controlador PID estándar)
│   ├── pcr_analisis_v1.py
│   └── analisis_pcr_v1.ipynb
├── v2/                         Segunda iteración (Newn1 a Newn6, PID y adaptativo) ← versión actual
│   ├── pcr_analisis_v2.py
│   └── analisis_pcr_v2.ipynb
├── herramientas/
│   └── construir_notebook.py   Genera el .ipynb de Colab a partir del .py
├── ejemplos/
│   ├── v1_Newn1/               Resultados de v1 con 3 corridas de Newn1
│   └── v2_Newn2-6/             Resultados de v2 con 10 corridas de Newn2 a Newn6
└── docs/
    ├── reporte_v1_Newn1.html   Explicación del método y resultados de Newn1
    └── reporte_v2_Newn2-6.html Resultados de Newn2 a Newn6
```

Los archivos de datos crudos (`*_pcr_data_*.xlsx`) **no se incluyen** en el repositorio. Consulta con el
laboratorio cómo conseguirlos.

---

## Uso rápido

### En Google Colab (recomendado)

1. Sube los archivos `.xlsx` de las corridas a una carpeta de Google Drive, por ejemplo
   `Mi unidad/Datos Microfluídicos`.
2. Abre `v2/analisis_pcr_v2.ipynb` en Colab.
3. En la celda **CONFIGURACIÓN**, revisa `CARPETA_DATOS` (la ruta de la carpeta en Drive) y el resto
   de las opciones.
4. Ejecuta todas las celdas (*Entorno de ejecución → Ejecutar todas*). Colab te pedirá permiso para
   acceder a Drive.
5. Al final se descarga un `.zip` con las gráficas y `resumen_pcr.xlsx`.

Si prefieres no usar Drive, pon `USAR_DRIVE = False` en la celda de carga y Colab te pedirá subir los archivos.

### En tu computadora

```bash
pip install -r requirements.txt
python v2/pcr_analisis_v2.py <carpeta_con_xlsx> <carpeta_de_salida>
```

---

## Formato de los datos de entrada

Cada corrida es un `.xlsx` con estas hojas:

| Hoja | Contenido |
|---|---|
| `metadata` | Protocolo (setpoints, tiempos, ciclos, `ts`), ciclos fallidos, timeouts y configuración del controlador |
| `settings` | *(desde Newn2)* Parámetros del controlador (`pidControllerRPM.*`) |
| `temperature` | `Thermocouple`, `IR Object`, `t_s`, `cycle`, `phase` y, según la versión, `heater_on_s` y `t_amb` |
| `photodetector`, `photodetector_raw` | Fluorescencia (este análisis no la usa) |

Hay **dos formatos de `metadata`**, y la v2 lee los dos:

- **Newn1:** una sola celda `prefix` con todo en texto: `project: Newn1-high_temp: 90.0-low_temp: 50.0-…`.
- **Newn2 en adelante:** una fila por parámetro (`key` / `value`). Si existe la clave `adaptive` con
  valor `on …`, el script considera que la corrida usó el controlador adaptativo.

Fases de un ciclo: `high_ramp → high_hold → cooling → low_hold → ext_ramp → ext_hold → fluor_read`.
Fuera de los ciclos (`cycle = 0`) están `initial`, `denat_ramp`, `denat_hold`, `final_ext`, `final_read` y `done`.

---

## Método

1. **Carga:** se leen `metadata` (setpoints, tiempos programados, ciclos fallidos) y `temperature`.
   Se usa siempre la columna `t_s`, porque el muestreo es irregular (dt de 0.08 a 0.57 s) y hay
   marcas de tiempo repetidas.
2. **Filtrado:** se promedian los tiempos duplicados y se remuestrea a `ts`. Luego se aplica un filtro
   de mediana de 1 s, que quita picos, y uno Savitzky-Golay de 1 s, que suaviza los rizos del control.
   Al final se vuelve a los tiempos originales.
3. **Segmentación:** cada cambio de `phase` o `cycle` abre un segmento nuevo, por ejemplo
   "cooling del ciclo 12".
4. **Métricas por segmento:**
   - `rate_10_90`: 0.8·ΔT / (t₉₀ − t₁₀), usando el **último** cruce de cada nivel. Es la métrica
     principal, porque no la distorsionan el arranque lento ni el sobrepaso.
   - `rate_avg`: ΔT/Δt de toda la fase. `rate_max`: pendiente máxima. `retardo_s`: tiempo hasta cruzar el 10 %.
   - Holds: `T_media`, `T_desv`, `T_min`, `T_max`, `error_vs_setpoint`, `sobrepaso` (T máx − setpoint)
     y `error_duracion_s` (duración real − programada).
5. **Exclusiones:** los ciclos fallidos del metadata y el `high_ramp` del ciclo 1, que arranca desde la
   desnaturalización, quedan fuera de las estadísticas. Los ciclos atípicos (más de 3.5·MAD de la mediana
   y más que un mínimo práctico) solo se marcan.
6. **Resumen:** media, DE y coeficiente de variación ciclo tras ciclo, por corrida, fase y sensor.
   En las fases únicas (`denat_hold`, `final_ext`) solo hay un valor por corrida, así que en lugar de
   la DE entre ciclos se reporta la DE **dentro** de la fase.
7. **Ciclo ideal (v2):** los holds van en el setpoint con la duración programada. Las rampas son lineales
   entre setpoints, con la duración media medida, porque el protocolo no programa la duración de las rampas.

---

## Configuración principal (v2)

| Variable | Por defecto | Para qué |
|---|---|---|
| `CARPETA_DATOS` / `CARPETA_SALIDA` | rutas de Colab | Carpeta de entrada y carpeta de resultados |
| `SENSOR` | `'ambos'` | `'Thermocouple'`, `'IR Object'` o `'ambos'` (gráficas separadas por sensor) |
| `SENSOR_REFERENCIA` | `'Thermocouple'` | Sensor que se usa para la carpeta `seleccion/` |
| `FILTRAR`, `VENTANA_MEDIANA_S`, `VENTANA_SAVGOL_S` | `True`, 1.0, 1.0 | Filtrado de la señal (ventanas en segundos) |
| `TASA_PRINCIPAL` | `'rate_10_90'` | Tasa que aparece en gráficas y atípicos. Alternativas: `'rate_avg'`, `'rate_max'` |
| `EXCLUIR_CICLOS_FALLIDOS` | `True` | Quita de las estadísticas los ciclos fallidos del metadata |
| `EXCLUIR_CICLO_1_HIGH_RAMP` | `True` | Quita el `high_ramp` del ciclo 1 |
| `CONTROLADOR_MANUAL` | `{}` | Fija el controlador por proyecto si el archivo no lo indica |
| `MIN_DESVIO` | 1 s · 25 % · 1 °C | Desviación mínima para marcar un ciclo como atípico |

---

## Salidas

`resumen_pcr.xlsx`

| Hoja | Contenido |
|---|---|
| `resumen` | Media, DE y CV por corrida, sensor y fase |
| `por_ciclo` | Una fila por ciclo y fase con todas las métricas (útil para Machine Learning) |
| `metadata` | Protocolo y configuración de cada corrida |
| `settings_controlador` | *(v2)* Parámetros del controlador de cada corrida |

Gráficas (una versión con el termopar, `TC`, y otra con el IR):

| Archivo | Qué muestra |
|---|---|
| `00_filtro_*` | Un ciclo con la señal cruda frente a la filtrada |
| `01_panorama_*` | La corrida completa (en v2, con los puntos máximo, mínimo y de extensión) |
| `02_superpuestas_*` | Todos los ciclos alineados en t = 0, por fase, con la curva promedio |
| `03_ciclo_a_ciclo_*` | Cada métrica frente al número de ciclo |
| `04_comparacion_*` | Comparación entre corridas (en v2, frente al setpoint y coloreada por controlador o dispositivo) |
| `05_TC_vs_IR` | Termopar frente a IR |
| `06_ideal_vs_real_*` | *(v2)* Ciclo real promedio sobre el ciclo ideal |
| `07_max_min_ext_*` | *(v2)* Máximo, mínimo y extensión por ciclo frente al setpoint |
| `seleccion/` | *(v2)* La gráfica más útil de cada tipo |

**Cómo leer `07_max_min_ext`:** en el panel de extensión, el punto es la temperatura media del hold en
ese ciclo, la banda horizontal es la media ± DE entre ciclos y la línea vertical es el **rango mín–máx
dentro del hold**, no una barra de error. Ese rango es real: refleja el sobrepaso al entrar al hold y la
corrección posterior del controlador.

---

## Historial de versiones

### v1 (oct 2026), Newn1, PID estándar
- Segmentación con las columnas `cycle` y `phase` (antes se detectaban los picos con `find_peaks`).
- Filtrado de mediana y Savitzky-Golay, tasa 10–90 %, retardo y detección de ciclos atípicos.
- Gráficas 00 a 05 y Excel resumen.
- Solo lee el formato `metadata` de Newn1 (celda `prefix`).

### v2 (nov 2026), Newn1 a Newn6, PID y adaptativo
- Lee los dos formatos de `metadata` y la hoja `settings`. Corrige la lectura de ciclos fallidos cuando
  hay un solo ciclo (Excel lo guardaba como `16.0`).
- Detecta el controlador (PID o adaptativo) y compara frente al setpoint y el tiempo programado,
  porque cada dispositivo usa un protocolo distinto.
- Fases únicas: la barra de error es la DE dentro de la fase, y así lo indica la gráfica.
- Unidades en todos los ejes.
- Gráficas nuevas: ciclo real frente al ideal (06) y puntos máximo, mínimo y de extensión (07).
- Carpeta `seleccion/` con las gráficas principales.
- Etiquetas con fecha y hora de inicio, y colores por dispositivo cuando todas las corridas usan el mismo controlador.

---

## Resultados de ejemplo

- **v1, Newn1 (PID):** los holds de 50 y 90 °C son muy repetibles, pero el extension hold se pasa
  del setpoint hasta +6.8 °C. Detalle en `docs/reporte_v1_Newn1.html`.
- **v2, Newn2 a Newn6 (adaptativo):** el low hold queda entre +0.07 y +0.18 °C del setpoint y el
  extension hold a ≤ 0.2 °C en las versiones recientes. Lo que falta mejorar es el sobrepaso del high hold
  (0.5–3.5 °C). Newn3 calienta más lento y es el que tiene más fallos. El sensor IR no coincide con el
  termopar. Detalle en `docs/reporte_v2_Newn2-6.html`.

Los reportes `.html` se abren en cualquier navegador y traen las gráficas incluidas.

---

## Requisitos

Python ≥ 3.10, `numpy`, `pandas`, `scipy`, `matplotlib`, `openpyxl` (ver `requirements.txt`).
Google Colab ya trae todo instalado.
