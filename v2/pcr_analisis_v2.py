# %% [markdown]
# # Análisis térmico PCR microfluídico — v2
# Segmenta cada experimento con las columnas `cycle` y `phase`, filtra la señal y calcula ciclo por ciclo
# rampas (heating/cooling rate) y holds (duración, T media, desv. est.). Compara experimentos y controladores
# (PID estándar vs adaptativo).
#
# **Cambios respecto a v1**
# * Lee los dos formatos de `metadata`: texto en una celda (`prefix`, Newn1) o filas clave/valor (Newn4). Lee también `settings`.
# * Detecta el controlador (PID estándar / adaptativo) y compara contra el **setpoint** (los protocolos son distintos).
# * Fases únicas (`denat_hold`, `final_ext`): ya no se muestra una DE "entre ciclos" (solo hay 1 por experimento);
#   la barra de error es la DE **dentro** de la fase y se indica en la gráfica.
# * Unidades en todos los ejes.
# * Nuevas gráficas: ciclo real vs ciclo ideal, y puntos máximo / mínimo / extensión por ciclo.
# * Carpeta `seleccion/` con la mejor gráfica de cada tipo.
#
# | Fase | Métricas |
# |---|---|
# | `cooling` | cooling rate, duración |
# | `low_hold`, `ext_hold`, `high_hold` | duración vs programada, T media vs setpoint, DE ciclo tras ciclo, máx/mín |
# | `ext_ramp`, `high_ramp` | heating rate (1ª y 2ª rampa de calentamiento) |
# | `denat_hold`, `final_ext` | duración y T media (1 por experimento) |
# | `fluor_read`, `initial`, `denat_ramp`, `final_read`, `done` | se omiten |

# %%
# ===================== CONFIGURACIÓN =====================
CARPETA_DATOS = '/content/datos'          # carpeta con los .xlsx / .csv
CARPETA_SALIDA = '/content/resultados'    # aquí se guardan gráficas y Excel resumen

SENSOR = 'ambos'                  # 'Thermocouple', 'IR Object' o 'ambos' (gráficas separadas por sensor)
SENSOR_REFERENCIA = 'Thermocouple'  # sensor usado para la carpeta seleccion/

FILTRAR = True              # filtrar la señal antes de calcular (quita picos y rizos)
VENTANA_MEDIANA_S = 1.0     # filtro de mediana: elimina picos aislados (s)
VENTANA_SAVGOL_S = 1.0      # Savitzky-Golay: suaviza rizos del control (s)

EXCLUIR_CICLOS_FALLIDOS = True    # quitar de las estadísticas los ciclos marcados como fallidos en metadata
EXCLUIR_CICLO_1_HIGH_RAMP = True  # el high_ramp del ciclo 1 viene de la desnaturalización (no es una rampa real)

TASA_PRINCIPAL = 'rate_10_90'   # 'rate_10_90' (recomendada), 'rate_avg' (ΔT/Δt de toda la fase) o 'rate_max'

# Si un archivo no dice qué controlador usó, se puede forzar por proyecto: {'Newn1': 'PID estándar'}
CONTROLADOR_MANUAL = {}

# umbral mínimo para marcar un ciclo como atípico (además de 3.5·MAD)
MIN_DESVIO = {'duracion_s': lambda med: max(1.0, 0.15 * abs(med)),   # s
              'rate':       lambda med: 0.25 * abs(med),             # aplica a TASA_PRINCIPAL
              'T_media':    lambda med: 1.0}                         # °C

FASES_RAMPA = ['ext_ramp', 'high_ramp', 'cooling']
FASES_HOLD = ['low_hold', 'ext_hold', 'high_hold']
FASES_UNICAS = ['denat_hold', 'final_ext']
ORDEN_CICLO = ['high_ramp', 'high_hold', 'cooling', 'low_hold', 'ext_ramp', 'ext_hold', 'fluor_read']
# ==========================================================

# %%
import re, glob, os, shutil
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.signal import savgol_filter, medfilt

COLORES = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6']
COLOR_CTRL = {'PID estándar': '#2a78d6', 'Adaptativo': '#eb6834'}
COLOR_FALLIDO = '#e34948'
TINTA, TINTA2 = '#0b0b0b', '#52514e'
plt.rcParams.update({'axes.spines.top': False, 'axes.spines.right': False,
                     'axes.grid': True, 'grid.alpha': 0.3, 'grid.linestyle': ':',
                     'figure.dpi': 110, 'font.size': 9})

NOMBRE_FASE = {'ext_ramp': 'Extension ramp (1er calentamiento)', 'high_ramp': 'High ramp (2º calentamiento)',
               'cooling': 'Cooling', 'low_hold': 'Low hold', 'ext_hold': 'Extension hold',
               'high_hold': 'High hold', 'denat_hold': 'Denat hold', 'final_ext': 'Final extension',
               'fluor_read': 'Fluor read'}
SETPOINT_KEY = {'low_hold': 'low_temp', 'ext_hold': 'ext_temp', 'high_hold': 'high_temp',
                'denat_hold': 'denat_temp', 'final_ext': 'ext_temp'}
DURACION_KEY = {'low_hold': 'time_low', 'ext_hold': 'ext_time', 'high_hold': 'time_high',
                'denat_hold': 'denat_time', 'final_ext': 'ext_time_final'}
UNIDAD = {'rate_avg': '°C/s', 'rate_10_90': '°C/s', 'rate_max': '°C/s', 'retardo_s': 's', 'duracion_s': 's',
          'T_media': '°C', 'T_desv': '°C', 'T_min': '°C', 'T_max': '°C', 'error_vs_setpoint': '°C',
          'error_duracion_s': 's', 'sobrepaso': '°C'}


# %% [markdown]
# ## 1. Lectura de archivos (Excel o CSV) y metadatos

# %%
def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return v.strip() if isinstance(v, str) else v


def parsear_prefijo(prefijo):
    """Formato Newn1: 'project: Newn1-high_temp: 90.0-...' en una sola celda. Separa solo en '-' seguido
    de 'clave:' (así no rompe la fecha 'start:2026-10-21T...')."""
    meta = {}
    for parte in re.split(r'-(?=[A-Za-z_]+\s*:)', prefijo):
        if ':' in parte:
            k, v = parte.split(':', 1)
            meta[k.strip()] = _num(v)
    return meta


def completar_meta(meta):
    fc = str(meta.get('failed_cycles', 'none'))
    meta['failed_list'] = [] if fc in ('none', 'nan', '') else [int(float(x)) for x in fc.split('|') if x.strip()]
    proyecto = str(meta.get('project', 'exp'))
    if proyecto in CONTROLADOR_MANUAL:
        meta['controlador'] = CONTROLADOR_MANUAL[proyecto]
    else:
        meta['controlador'] = 'Adaptativo' if str(meta.get('adaptive', '')).lower().startswith('on') else 'PID estándar'
    m = re.search(r'\d{4}-(\d{2})-(\d{2})T(\d{2}:\d{2})', str(meta.get('start', '')))
    meta['hora_inicio'] = f'{m.group(2)}/{m.group(1)} {m.group(3)}' if m else ''
    return meta


def cargar_experimento(ruta):
    """Devuelve (df_temperatura, metadatos, nombre). Acepta .xlsx/.xls (hojas metadata + temperature [+ settings])
    o .csv (1ª línea = prefijo de metadatos, 2ª = encabezados)."""
    ruta = Path(ruta)
    if ruta.suffix.lower() in ('.xlsx', '.xls'):
        hojas = pd.read_excel(ruta, sheet_name=None)
        md = hojas['metadata']
        if 'prefix' in md['key'].astype(str).values:            # formato Newn1
            meta = parsear_prefijo(md.loc[md['key'] == 'prefix', 'value'].iloc[0])
        else:                                                   # formato Newn4: una fila por clave
            meta = {str(k).strip(): _num(v) for k, v in zip(md['key'], md['value'])}
        if 'settings' in hojas:
            meta['settings'] = dict(zip(hojas['settings']['key'], hojas['settings']['value']))
        df = hojas['temperature']
    else:
        with open(ruta) as f:
            meta = parsear_prefijo(f.readline().strip())
        df = pd.read_csv(ruta, skiprows=1)
    faltan = {'Thermocouple', 'IR Object', 't_s', 'cycle', 'phase'} - set(df.columns)
    if faltan:
        raise ValueError(f'{ruta.name}: faltan columnas {faltan}')
    meta = completar_meta(meta)
    m = re.search(r'(\d{8}_\d{6})', ruta.stem)
    nombre = f"{meta.get('project', 'exp')}_{m.group(1)}" if m else ruta.stem
    meta['etiqueta'] = f"{meta.get('project', '')} {meta['hora_inicio']}\n{meta['controlador']}"
    df = df.sort_values('t_s', kind='stable').reset_index(drop=True)
    return df, meta, nombre


# %% [markdown]
# ## 2. Filtrado de la señal
# Hay marcas de tiempo repetidas y muestreo irregular (dt de 0.08 a 0.57 s), así que:
# promedio de duplicados → remuestreo uniforme a `ts` → mediana (picos) → Savitzky-Golay (rizos)
# → se interpola de regreso a los tiempos originales para conservar las etiquetas de fase.

# %%
def filtrar_senal(t, y, ts=0.1):
    s = pd.Series(y).groupby(np.asarray(t)).mean()
    tu, yu = s.index.values, s.values
    grid = np.arange(tu[0], tu[-1] + ts, ts)
    yg = np.interp(grid, tu, yu)
    yg = medfilt(yg, max(3, int(round(VENTANA_MEDIANA_S / ts)) | 1))
    yg = savgol_filter(yg, max(5, int(round(VENTANA_SAVGOL_S / ts)) | 1), 2)
    return np.interp(t, grid, yg)


def preparar(df, meta):
    ts = meta.get('ts', 0.1)
    for s in ('Thermocouple', 'IR Object'):
        df[s + '_f'] = filtrar_senal(df['t_s'].values, df[s].values, ts) if FILTRAR else df[s]
    df['seg'] = (df['phase'].ne(df['phase'].shift()) | df['cycle'].ne(df['cycle'].shift())).cumsum()
    return df


# %% [markdown]
# ## 3. Métricas por segmento
# * **Rampas:** `rate_10_90` = 0.8·ΔT / tiempo entre cruzar el 10 % y el 90 % del cambio (no incluye arranque
#   lento ni cola); `rate_avg` = ΔT/Δt de toda la fase; `rate_max` = dT/dt pico; `retardo_s` = tiempo hasta cruzar el 10 %.
# * **Holds:** duración y su error vs la programada, T media y su error vs setpoint, DE dentro del hold,
#   mín., máx. y `sobrepaso` = T máx − setpoint.
# * **Atípicos:** se alejan de la mediana > 3.5·MAD y más que `MIN_DESVIO`. Solo se marcan; los **fallidos**
#   del metadata sí se excluyen.

# %%
def tasa_10_90(t, y):
    """Usa el ÚLTIMO cruce de cada nivel para que un pico falso no adelante el cruce.
    Devuelve (tasa °C/s, retardo s)."""
    y0, y1 = y[0], y[-1]
    dT = y1 - y0
    if abs(dT) < 1:
        return np.nan, np.nan
    frac = (y - y0) / dT

    def ultimo_cruce(nivel):
        debajo = np.where(frac < nivel)[0]
        if len(debajo) == 0:
            return t[0]
        i = debajo[-1]
        if i + 1 >= len(frac):
            return None
        f0, f1 = frac[i], frac[i + 1]
        return t[i] + (nivel - f0) / (f1 - f0) * (t[i + 1] - t[i]) if f1 != f0 else t[i + 1]

    t10, t90 = ultimo_cruce(0.1), ultimo_cruce(0.9)
    if t10 is None or t90 is None or t90 <= t10:
        return np.nan, np.nan
    return 0.8 * dT / (t90 - t10), t10 - t[0]


def metricas_experimento(df, meta, nombre, sensor):
    col = sensor + '_f'
    filas = []
    for _, g in df.groupby('seg', sort=True):
        fase, ciclo = g['phase'].iloc[0], int(g['cycle'].iloc[0])
        if fase not in FASES_RAMPA + FASES_HOLD + FASES_UNICAS or len(g) < 3:
            continue
        t, y = g['t_s'].values, g[col].values
        fila = {'experimento': nombre, 'proyecto': meta.get('project'), 'controlador': meta['controlador'],
                'sensor': sensor, 'ciclo': ciclo, 'fase': fase, 'fallido': ciclo in meta['failed_list'],
                't_inicio_s': t[0], 'duracion_s': t[-1] - t[0], 'T_inicio': y[0], 'T_fin': y[-1]}
        if fase in FASES_RAMPA:
            dt, dy = np.diff(t), np.diff(y)
            ok = dt > 1e-3
            deriv = dy[ok] / dt[ok]
            r1090, retardo = tasa_10_90(t, y)
            fila.update(rate_avg=(y[-1] - y[0]) / (t[-1] - t[0]), rate_10_90=r1090, retardo_s=retardo,
                        rate_max=(deriv.max() if fase != 'cooling' else deriv.min()) if ok.any() else np.nan)
        else:
            sp = meta.get(SETPOINT_KEY[fase], np.nan)
            dprog = meta.get(DURACION_KEY[fase], np.nan)
            fila.update(setpoint=sp, T_media=y.mean(), T_desv=y.std(ddof=1), T_min=y.min(), T_max=y.max(),
                        error_vs_setpoint=y.mean() - sp, sobrepaso=y.max() - sp,
                        duracion_prog_s=dprog, error_duracion_s=(t[-1] - t[0]) - dprog)
        fila['excluido'] = (fase == 'high_ramp' and ciclo == 1 and EXCLUIR_CICLO_1_HIGH_RAMP) or \
                           (EXCLUIR_CICLOS_FALLIDOS and fila['fallido'])
        filas.append(fila)
    m = pd.DataFrame(filas)
    m['atipico'] = False
    for fase, g in m[m.fase.isin(FASES_RAMPA + FASES_HOLD) & ~m.excluido].groupby('fase'):
        for v, minimo in MIN_DESVIO.items():
            v = TASA_PRINCIPAL if v == 'rate' else v
            if v in g and g[v].notna().sum() > 5:
                x = g[v]
                dev = (x - x.median()).abs()
                lim = max(3.5 * 1.4826 * dev.median(), minimo(x.median()))
                m.loc[g.index[(dev > lim).values], 'atipico'] = True
    return m


# %%
def resumen(metricas):
    """Media, DE, CV % ciclo tras ciclo por experimento / sensor / fase.
    En las fases únicas (n = 1) no hay DE entre ciclos: se reporta `T_desv` (DE dentro de la fase)."""
    vals = ['duracion_s', 'error_duracion_s', 'rate_avg', 'rate_10_90', 'rate_max', 'retardo_s',
            'T_media', 'error_vs_setpoint', 'T_desv', 'T_min', 'T_max', 'sobrepaso']
    d = metricas[~metricas['excluido']]
    filas = []
    for (exp, sen, fase), g in d.groupby(['experimento', 'sensor', 'fase']):
        fila = {'experimento': exp, 'proyecto': g.proyecto.iloc[0], 'controlador': g.controlador.iloc[0],
                'sensor': sen, 'fase': fase, 'n_ciclos': len(g),
                'ciclos_atipicos': ','.join(map(str, g.loc[g.atipico, 'ciclo']))}
        if 'setpoint' in g:
            fila['setpoint'] = g['setpoint'].iloc[0]
            fila['duracion_prog_s'] = g['duracion_prog_s'].iloc[0]
        for v in vals:
            if v in g and g[v].notna().any():
                x = g[v].dropna()
                fila[f'{v}_media'] = x.mean()
                fila[f'{v}_de'] = x.std(ddof=1) if len(x) > 1 else np.nan
                if v in ('duracion_s', 'rate_avg', 'rate_10_90', 'T_media') and len(x) > 1:
                    fila[f'{v}_cv%'] = 100 * fila[f'{v}_de'] / abs(fila[f'{v}_media'])
        filas.append(fila)
    orden = {f: i for i, f in enumerate(FASES_RAMPA + FASES_HOLD + FASES_UNICAS)}
    r = pd.DataFrame(filas)
    return r.sort_values(['sensor', 'fase', 'experimento'],
                         key=lambda s: s.map(orden) if s.name == 'fase' else s).reset_index(drop=True)


# %% [markdown]
# ## 4. Ciclo promedio real y ciclo ideal
# * **Real:** cada fase de cada ciclo se re-muestrea en tiempo normalizado (0–1) y se promedia entre ciclos;
#   cada fase dura su duración media real.
# * **Ideal:** holds en el setpoint con la duración **programada** (metadata); rampas lineales entre setpoints
#   con la duración **media medida** (el protocolo no programa la duración de las rampas).

# %%
def ciclo_promedio(df, metr, meta, sensor, n_pts=60):
    col = sensor + '_f'
    validos = metr[(metr.ciclo >= 2) & ~metr.excluido].ciclo.unique()
    real, t0 = [], 0.0
    for fase in ORDEN_CICLO:
        curvas, durs = [], []
        for c in validos:
            g = df[(df.cycle == c) & (df.phase == fase)]
            if len(g) < 3:
                continue
            tt = g.t_s.values
            u = (tt - tt[0]) / max(tt[-1] - tt[0], 1e-6)
            curvas.append(np.interp(np.linspace(0, 1, n_pts), u, g[col].values))
            durs.append(tt[-1] - tt[0])
        if not curvas:
            continue
        C, D = np.array(curvas), float(np.mean(durs))
        real.append(dict(fase=fase, t=t0 + np.linspace(0, D, n_pts), media=C.mean(0), de=C.std(0, ddof=1),
                         mn=C.min(0), mx=C.max(0), dur=D, t0=t0))
        t0 += D
    # ideal
    dur_real = {r['fase']: r['dur'] for r in real}
    hi, lo, ex = meta.get('high_temp'), meta.get('low_temp'), meta.get('ext_temp')
    perfil = {'high_ramp': (ex, hi), 'high_hold': (hi, hi), 'cooling': (hi, lo), 'low_hold': (lo, lo),
              'ext_ramp': (lo, ex), 'ext_hold': (ex, ex), 'fluor_read': (ex, ex)}
    ideal, t0 = [], 0.0
    for fase in ORDEN_CICLO:
        if fase not in dur_real:
            continue
        D = meta.get(DURACION_KEY[fase], dur_real[fase]) if fase in DURACION_KEY else dur_real[fase]
        a, b = perfil[fase]
        ideal.append(dict(fase=fase, t=[t0, t0 + D], T=[a, b], t0=t0, dur=D))
        t0 += D
    return real, ideal


# %% [markdown]
# ## 5. Gráficas

# %%
def _etiq_eje(var, fase=None):
    nombres = {'rate_10_90': 'tasa 10–90 %', 'rate_avg': 'tasa ΔT/Δt', 'rate_max': 'tasa máx.',
               'duracion_s': 'duración', 'T_media': 'T media', 'error_vs_setpoint': 'T media − setpoint',
               'T_desv': 'DE dentro de la fase', 'error_duracion_s': 'duración − programada',
               'sobrepaso': 'T máx − setpoint', 'retardo_s': 'retardo'}
    return f"{nombres.get(var, var)} ({UNIDAD.get(var, '')})"


def fig_filtro(df, nombre, sensor, ciclo=10):
    """Zoom a un ciclo: señal cruda vs filtrada."""
    g = df[df.cycle == ciclo]
    fig, ax = plt.subplots(figsize=(13, 3.6))
    ax.plot(g.t_s, g[sensor], color='#b8b7b0', lw=1.2, label='cruda')
    ax.plot(g.t_s, g[sensor + '_f'], color=COLORES[0], lw=1.6, label='filtrada')
    for _, gg in g.groupby('seg'):
        ax.axvline(gg.t_s.iloc[0], color=TINTA2, lw=0.5, ls=':')
        ax.text(gg.t_s.iloc[0] + 0.3, ax.get_ylim()[1], gg.phase.iloc[0], rotation=90, va='top', fontsize=7, color=TINTA2)
    ax.set_title(f'{nombre} — {sensor}: ciclo {ciclo}, cruda vs filtrada', loc='left', fontsize=10)
    ax.set_xlabel('Tiempo (s)'); ax.set_ylabel('Temperatura (°C)'); ax.legend(frameon=False, loc='lower left')
    fig.tight_layout()
    return fig


def fig_panorama(datos, metricas, sensor):
    """Señal completa con puntos máximo (high hold), mínimo (low hold) y extensión (T media ext hold) de cada ciclo."""
    fig, axs = plt.subplots(len(datos), 1, figsize=(13, 2.9 * len(datos)), squeeze=False)
    for ax, (nombre, (df, meta)) in zip(axs[:, 0], datos.items()):
        ax.plot(df['t_s'], df[sensor], color='#b8b7b0', lw=0.6, label='cruda')
        ax.plot(df['t_s'], df[sensor + '_f'], color=COLORES[0], lw=0.8, label='filtrada')
        m = metricas[(metricas.experimento == nombre) & (metricas.sensor == sensor)]
        for fase, var, mk, col, lab in [('high_hold', 'T_max', '^', COLORES[1], 'máximo'),
                                        ('low_hold', 'T_min', 'v', TINTA, 'mínimo'),
                                        ('ext_hold', 'T_media', 'o', COLORES[2], 'extensión')]:
            d = m[m.fase == fase]
            tpos = d.t_inicio_s + d.duracion_s / 2
            ax.plot(tpos, d[var], mk, color=col, ms=4, ls='none', label=lab)
        for c in meta['failed_list']:
            g = df[(df.cycle == c) & (df.phase == 'ext_ramp')]
            if len(g):
                ax.axvspan(g.t_s.iloc[0], g.t_s.iloc[-1], color=COLOR_FALLIDO, alpha=0.25, lw=0)
        ax.set_title(f"{nombre} — {sensor} — {meta['controlador']}" +
                     (f"  (ciclos fallidos: {meta['failed_list']})" if meta['failed_list'] else ''), loc='left', fontsize=10)
        ax.set_ylabel('Temperatura (°C)'); ax.set_xlabel('Tiempo (s)')
    axs[0, 0].legend(loc='lower right', ncol=5, frameon=False, fontsize=8)
    fig.tight_layout()
    return fig


def fig_superpuestas(df, meta, metr, nombre, sensor):
    """Todos los ciclos alineados en t = 0 por fase, con curva promedio."""
    fig, axs = plt.subplots(2, 3, figsize=(14, 7.5))
    for ax, fase in zip(axs.ravel(), FASES_RAMPA + FASES_HOLD):
        mf = metr[metr.fase == fase]
        curvas = []
        for _, r in mf.iterrows():
            g = df[(df.cycle == r.ciclo) & (df.phase == fase)]
            t = g['t_s'].values - g['t_s'].values[0]
            y = g[sensor + '_f'].values
            if r.fallido:
                ax.plot(t, y, color=COLOR_FALLIDO, lw=1, alpha=0.8)
            elif not r.excluido:
                ax.plot(t, y, color=COLORES[0], lw=0.8, alpha=0.25)
                curvas.append((t, y))
        if curvas:
            grid = np.linspace(0, np.median([c[0][-1] for c in curvas]), 200)
            ax.plot(grid, np.mean([np.interp(grid, t, y) for t, y in curvas], axis=0), color=TINTA, lw=2, label='promedio')
        d = mf[~mf.excluido]
        if fase in FASES_RAMPA:
            txt = (f"10–90 %: {d.rate_10_90.mean():+.2f} ± {d.rate_10_90.std():.2f} °C/s\n"
                   f"ΔT/Δt fase: {d.rate_avg.mean():+.2f} ± {d.rate_avg.std():.2f} °C/s\n"
                   f"retardo = {d.retardo_s.mean():.2f} ± {d.retardo_s.std():.2f} s")
        else:
            sp = meta.get(SETPOINT_KEY[fase])
            ax.axhline(sp, color=TINTA2, ls='--', lw=1, label=f'setpoint {sp:.0f} °C')
            txt = f"T = {d.T_media.mean():.2f} ± {d.T_media.std():.2f} °C\ndur = {d.duracion_s.mean():.2f} ± {d.duracion_s.std():.2f} s"
        ax.text(0.98, 0.04 if fase != 'cooling' else 0.96, txt, transform=ax.transAxes, ha='right',
                va='bottom' if fase != 'cooling' else 'top', fontsize=8, bbox=dict(fc='white', ec='none', alpha=0.85))
        ax.set_title(NOMBRE_FASE[fase] + f'  (n = {len(d)})', loc='left', fontsize=10)
        ax.set_xlabel('Tiempo desde inicio de fase (s)'); ax.set_ylabel('Temperatura (°C)')
    if metr.fallido.any():
        axs[0, 0].plot([], [], color=COLOR_FALLIDO, label='ciclo fallido')
    axs[0, 0].legend(frameon=False, fontsize=8, loc='upper left')
    axs[1, 0].legend(frameon=False, fontsize=8, loc='upper left')
    fig.suptitle(f"{nombre} — {sensor} — {meta['controlador']}: ciclos superpuestos", x=0.01, ha='left', fontsize=12, fontweight='bold')
    fig.tight_layout()
    return fig


def fig_ciclo_a_ciclo(metr, meta, nombre, sensor):
    """Cada métrica contra el número de ciclo (repetibilidad)."""
    paneles = [(f, TASA_PRINCIPAL) for f in FASES_RAMPA] + [(f, 'T_media') for f in FASES_HOLD] + \
              [(f, 'duracion_s') for f in FASES_RAMPA + FASES_HOLD]
    fig, axs = plt.subplots(4, 3, figsize=(14, 11.5), sharex=True)
    for ax, (fase, var) in zip(axs.ravel(), paneles):
        d = metr[metr.fase == fase].sort_values('ciclo')
        ok = d[~d.excluido]
        if var == 'T_media':
            ax.fill_between(d.ciclo, d.T_min, d.T_max, color=COLORES[0], alpha=0.15, lw=0, label='mín–máx')
            ax.axhline(meta.get(SETPOINT_KEY[fase]), color=TINTA2, lw=1, ls=':', label='setpoint')
        if var == 'duracion_s' and fase in DURACION_KEY:
            ax.axhline(meta.get(DURACION_KEY[fase]), color=TINTA2, lw=1, ls=':', label='programada')
        mu, sd = ok[var].mean(), ok[var].std()
        ax.axhspan(mu - sd, mu + sd, color=TINTA2, alpha=0.1, lw=0)
        ax.axhline(mu, color=TINTA2, lw=1, ls='--', label=f'media ± DE: {mu:.2f} ± {sd:.2f}')
        ax.plot(ok.ciclo, ok[var], 'o-', color=COLORES[0], ms=3.5, lw=1)
        a = ok[ok.atipico]
        if len(a):
            ax.plot(a.ciclo, a[var], 'o', mfc='none', mec=TINTA, ms=8, mew=1.2, label='atípico')
        f = d[d.fallido]
        if len(f):
            ax.plot(f.ciclo, f[var], 'X', color=COLOR_FALLIDO, ms=7, label='fallido (excluido)' if EXCLUIR_CICLOS_FALLIDOS else 'fallido')
        ax.set_title(NOMBRE_FASE[fase], loc='left', fontsize=9)
        ax.set_ylabel(_etiq_eje(var))
        ax.legend(frameon=False, fontsize=7, loc='best')
    for ax in axs[-1]:
        ax.set_xlabel('Ciclo (n.º)')
        ax.tick_params(labelbottom=True)
    fig.suptitle(f"{nombre} — {sensor} — {meta['controlador']}: métricas ciclo tras ciclo", x=0.01, ha='left', fontsize=12, fontweight='bold')
    fig.tight_layout()
    return fig


def fig_ideal_vs_real(df, meta, metr, nombre, sensor):
    """Ciclo promedio real sobre el ciclo ideal, con puntos máximo, mínimo y de extensión."""
    real, ideal = ciclo_promedio(df, metr, meta, sensor)
    fig, ax = plt.subplots(figsize=(13, 5.2))
    tonos = {'high_ramp': '#f8e3d8', 'ext_ramp': '#f8e3d8', 'cooling': '#dce9f6', 'fluor_read': '#ffffff'}
    for r in real:
        ax.axvspan(r['t0'], r['t0'] + r['dur'], color=tonos.get(r['fase'], '#eceef1'), lw=0, zorder=0)
        ax.text(r['t0'] + r['dur'] / 2, 1.01, r['fase'], transform=ax.get_xaxis_transform(), ha='center',
                va='bottom', fontsize=7.5, color=TINTA2, rotation=0 if r['dur'] > 6 else 90)
    ti = np.concatenate([i['t'] for i in ideal]); Ti = np.concatenate([i['T'] for i in ideal])
    ax.plot(ti, Ti, color=TINTA, lw=1.6, ls='--', label='ciclo ideal (setpoints + tiempos programados)', zorder=3)
    tr = np.concatenate([r['t'] for r in real])
    ax.fill_between(tr, np.concatenate([r['mn'] for r in real]), np.concatenate([r['mx'] for r in real]),
                    color=COLORES[0], alpha=0.12, lw=0, label='mín–máx entre ciclos', zorder=1)
    Tm, De = np.concatenate([r['media'] for r in real]), np.concatenate([r['de'] for r in real])
    ax.fill_between(tr, Tm - De, Tm + De, color=COLORES[0], alpha=0.25, lw=0, label='± DE entre ciclos', zorder=2)
    ax.plot(tr, Tm, color=COLORES[0], lw=2.2, label='ciclo real promedio', zorder=4)
    # puntos clave (estadística por ciclo, ubicados sobre la curva promedio)
    ok = metr[~metr.excluido]
    puntos = [('high_hold', 'T_max', 'máx', COLORES[1], '^', np.argmax),
              ('low_hold', 'T_min', 'mín', TINTA, 'v', np.argmin),
              ('ext_hold', 'T_media', 'ext', COLORES[2], 'o', None)]
    for fase, var, lab, col, mk, fn in puntos:
        r = next((x for x in real if x['fase'] == fase), None)
        if r is None:
            continue
        i = fn(r['media']) if fn else len(r['media']) // 2
        st = ok[ok.fase == fase][var]
        sp = meta.get(SETPOINT_KEY[fase])
        ax.plot(r['t'][i], r['media'][i], mk, color=col, ms=10, mec='white', mew=1.5, zorder=6)
        dy = 14 if lab != 'mín' else -26
        ax.annotate(f'{lab}: {st.mean():.1f} ± {st.std():.1f} °C\n(setpoint {sp:.0f} °C)', (r['t'][i], r['media'][i]),
                    xytext=(8, dy), textcoords='offset points', fontsize=8, color=TINTA,
                    bbox=dict(fc='white', ec=col, lw=1, alpha=0.95), zorder=7)
    dur_real, dur_ideal = sum(r['dur'] for r in real), ti[-1]
    ax.set_title(f"{nombre} — {sensor} — {meta['controlador']}: ciclo real vs ideal  "
                 f"(duración ciclo real {dur_real:.1f} s · ideal {dur_ideal:.1f} s · n = {metr[(metr.ciclo >= 2) & ~metr.excluido].ciclo.nunique()} ciclos)",
                 loc='left', fontsize=10, pad=18)
    ax.set_xlabel('Tiempo desde inicio del ciclo (s)'); ax.set_ylabel('Temperatura (°C)')
    ax.set_xlim(0, max(dur_real, dur_ideal))
    ax.legend(frameon=False, fontsize=8, loc='upper right', ncol=2)
    lo, hi = np.nanmin(np.concatenate([r['mn'] for r in real])), np.nanmax(np.concatenate([r['mx'] for r in real]))
    ax.set_ylim(lo - 6, hi + 8)
    fig.tight_layout()
    return fig


def fig_puntos(metr, meta, nombre, sensor):
    """Puntos máximo, mínimo y de extensión de cada ciclo, contra el setpoint."""
    defs = [('high_hold', 'T_max', 'Máximo por ciclo (T máx en high hold)', COLORES[1], '^'),
            ('low_hold', 'T_min', 'Mínimo por ciclo (T mín en low hold)', TINTA, 'v'),
            ('ext_hold', 'T_media', 'Extensión por ciclo (T media del extension hold)', COLORES[2], 'o')]
    fig, axs = plt.subplots(3, 1, figsize=(12, 9), sharex=True)
    for ax, (fase, var, tit, col, mk) in zip(axs, defs):
        d = metr[metr.fase == fase].sort_values('ciclo')
        ok = d[~d.excluido]
        sp = meta.get(SETPOINT_KEY[fase])
        mu, sd = ok[var].mean(), ok[var].std()
        ax.axhspan(mu - sd, mu + sd, color=col, alpha=0.12, lw=0)
        ax.axhline(mu, color=col, lw=1, ls='--', label=f'media ± DE: {mu:.2f} ± {sd:.2f} °C')
        ax.axhline(sp, color=TINTA2, lw=1.2, ls=':', label=f'setpoint {sp:.0f} °C')
        if fase == 'ext_hold':
            ax.vlines(ok.ciclo, ok.T_min, ok.T_max, color=col, alpha=0.4, lw=2, label='rango mín–máx del hold')
        ax.plot(ok.ciclo, ok[var], mk, color=col, ms=6, mec='white', mew=0.8, ls='-', lw=0.8)
        f = d[d.fallido]
        if len(f):
            ax.plot(f.ciclo, f[var], 'X', color=COLOR_FALLIDO, ms=8, ls='none', label='fallido (excluido)')
        ax.set_title(tit, loc='left', fontsize=10)
        ax.set_ylabel('Temperatura (°C)')
        ax.legend(frameon=False, fontsize=8, loc='best', ncol=2)
    axs[-1].set_xlabel('Ciclo (n.º)')
    fig.suptitle(f"{nombre} — {sensor} — {meta['controlador']}: puntos máximo, mínimo y de extensión",
                 x=0.01, ha='left', fontsize=12, fontweight='bold')
    fig.tight_layout()
    return fig


def _barras(ax, r, exps, etiquetas, ctrls, fase, var, unica=False):
    d = r[r.fase == fase].set_index('experimento').reindex(exps)
    mu = d.get(f'{var}_media', pd.Series(np.nan, index=d.index)).values.astype(float)
    if unica and var in ('error_vs_setpoint', 'T_media'):
        sd = d.get('T_desv_media', pd.Series(np.nan, index=d.index)).values.astype(float)
    elif unica:
        sd = np.full(len(exps), np.nan)
    else:
        sd = d.get(f'{var}_de', pd.Series(np.nan, index=d.index)).values.astype(float)
    x = np.arange(len(exps))
    ax.bar(x, mu, yerr=np.where(np.isfinite(sd), sd, 0), color=ctrls,
           width=0.62, capsize=4 if not unica or var != 'error_duracion_s' else 0, edgecolor='white', linewidth=2,
           error_kw=dict(lw=1, ecolor=TINTA2))
    ax.axhline(0, color=TINTA2, lw=0.8)
    lo = np.nanmin(np.minimum(mu - np.nan_to_num(sd), 0)); hi = np.nanmax(np.maximum(mu + np.nan_to_num(sd), 0))
    pad = (hi - lo) * 0.28 + 1e-3
    ax.set_ylim(lo - (pad if lo < 0 else 0), hi + (pad if hi > 0 else 0))
    for xi, m_, s_ in zip(x, mu, sd):
        if not np.isfinite(m_):
            continue
        arriba = m_ >= 0
        ax.annotate(f'{m_:.2f}' + (f'\n±{s_:.2f}' if np.isfinite(s_) else ''),
                    (xi, m_ + (np.nan_to_num(s_) if arriba else -np.nan_to_num(s_))), ha='center',
                    va='bottom' if arriba else 'top', fontsize=6.5 if len(exps) <= 6 else 5.5, xytext=(0, 2 if arriba else -2), textcoords='offset points')
    ax.set_xticks(x, etiquetas, fontsize=6.5, rotation=90 if len(exps) > 6 else 0)
    ax.set_ylabel(_etiq_eje(var), fontsize=8)
    nota = ''
    if unica:
        nota = '\n(1 por experimento; barra = DE dentro de la fase)' if var in ('error_vs_setpoint', 'T_media') else '\n(1 por experimento)'
    ax.set_title(f'{NOMBRE_FASE[fase]}{nota}', loc='left', fontsize=9)


def fig_comparacion(res, datos, sensor):
    """Comparación entre experimentos, normalizada contra setpoint / tiempo programado
    (los protocolos de cada dispositivo son distintos). Color = controlador."""
    r = res[res.sensor == sensor]
    exps = sorted(r.experimento.unique())
    etiquetas = [datos[e][1]['etiqueta'] for e in exps]
    # color por controlador si hay más de uno; si todos usan el mismo, color por dispositivo (proyecto)
    ctrl_exp = [datos[e][1]['controlador'] for e in exps]
    if len(set(ctrl_exp)) > 1:
        grupo, titulo_ley = ctrl_exp, 'Controlador'
        paleta = {g: COLOR_CTRL.get(g, COLORES[2]) for g in grupo}
    else:
        grupo, titulo_ley = [str(datos[e][1].get('project')) for e in exps], f'Dispositivo ({ctrl_exp[0]})'
        paleta = {g: COLORES[i % len(COLORES)] for i, g in enumerate(dict.fromkeys(grupo))}
    ctrls = [paleta[g] for g in grupo]
    filas = [
        [(f, TASA_PRINCIPAL, False) for f in FASES_RAMPA] + [('denat_hold', 'error_vs_setpoint', True)],
        [(f, 'duracion_s', False) for f in FASES_RAMPA] + [('final_ext', 'error_vs_setpoint', True)],
        [(f, 'error_vs_setpoint', False) for f in FASES_HOLD] + [('denat_hold', 'error_duracion_s', True)],
        [(f, 'sobrepaso', False) for f in FASES_HOLD] + [('final_ext', 'error_duracion_s', True)],
        [(f, 'error_duracion_s', False) for f in FASES_HOLD] + [None],
    ]
    ancho = max(17, 4 * (1.3 + 0.45 * len(exps)))
    fig, axs = plt.subplots(len(filas), 4, figsize=(ancho, 4.4 * len(filas)))
    for fila, axr in zip(filas, axs):
        for p, ax in zip(fila, axr):
            if p is None:
                ax.axis('off')
                from matplotlib.patches import Patch
                ax.legend(handles=[Patch(color=col, label=g) for g, col in paleta.items()],
                          title=titulo_ley, frameon=False, loc='center', fontsize=10, title_fontsize=10)
                ax.text(0.5, 0.15, 'Barras de error: DE ciclo tras ciclo\n(salvo fases únicas, ver título)',
                        transform=ax.transAxes, ha='center', fontsize=8, color=TINTA2)
                continue
            _barras(ax, r, exps, etiquetas, ctrls, *p)
    fig.suptitle(f'Comparación entre experimentos — {sensor} (media ± DE)', x=0.01, ha='left', fontsize=13, fontweight='bold')
    fig.tight_layout()
    return fig


def fig_sensores(res, datos):
    """Termopar vs IR (mismo experimento, mismo setpoint)."""
    exps = sorted(res.experimento.unique())
    etiquetas = [datos[e][1]['etiqueta'] for e in exps]
    paneles = [(f, TASA_PRINCIPAL) for f in FASES_RAMPA] + [(f, 'error_vs_setpoint') for f in FASES_HOLD]
    fig, axs = plt.subplots(2, 3, figsize=(max(15, 3 * (1.3 + 0.5 * len(exps))), 8.5))
    for ax, (fase, var) in zip(axs.ravel(), paneles):
        x = np.arange(len(exps)); w = 0.36
        for i, sen in enumerate(['Thermocouple', 'IR Object']):
            d = res[(res.fase == fase) & (res.sensor == sen)].set_index('experimento').reindex(exps)
            ax.bar(x + (i - 0.5) * w, d[f'{var}_media'], w, yerr=d[f'{var}_de'], color=[COLORES[6], COLORES[2]][i], capsize=3,
                   edgecolor='white', linewidth=2, error_kw=dict(lw=1, ecolor=TINTA2), label=sen)
        ax.axhline(0, color=TINTA2, lw=0.8)
        ax.set_xticks(x, etiquetas, fontsize=7, rotation=90 if len(exps) > 6 else 0)
        ax.set_ylabel(_etiq_eje(var), fontsize=8)
        ax.set_title(NOMBRE_FASE[fase], loc='left', fontsize=9)
    axs[0, 0].legend(frameon=False, fontsize=8)
    fig.suptitle('Termopar vs sensor IR (media ± DE ciclo tras ciclo)', x=0.01, ha='left', fontsize=12, fontweight='bold')
    fig.tight_layout()
    return fig


# %% [markdown]
# ## 6. Ejecutar

# %%
def guardar(fig, ruta):
    fig.savefig(ruta, bbox_inches='tight')
    plt.close(fig)


def correr():
    os.makedirs(CARPETA_SALIDA, exist_ok=True)
    sel = os.path.join(CARPETA_SALIDA, 'seleccion')
    os.makedirs(sel, exist_ok=True)
    rutas = sorted(glob.glob(os.path.join(CARPETA_DATOS, '*.xlsx')) + glob.glob(os.path.join(CARPETA_DATOS, '*.xls'))
                   + glob.glob(os.path.join(CARPETA_DATOS, '*.csv')))
    print(f'{len(rutas)} archivos encontrados')
    sensores = ['Thermocouple', 'IR Object'] if SENSOR == 'ambos' else [SENSOR]
    datos, todas = {}, []
    for r in rutas:
        df, meta, nombre = cargar_experimento(r)
        df = preparar(df, meta)
        datos[nombre] = (df, meta)
        print(f"  {nombre}: {meta['controlador']}, protocolo {meta.get('high_temp'):.0f}/{meta.get('low_temp'):.0f}/"
              f"{meta.get('ext_temp'):.0f} °C, {len(df)} puntos, {df.t_s.iloc[-1]:.0f} s, "
              f"ciclos fallidos: {meta['failed_list'] or 'ninguno'}")
        for s in sensores:
            todas.append(metricas_experimento(df, meta, nombre, s))
    metricas = pd.concat(todas, ignore_index=True)
    res = resumen(metricas)

    for s in sensores:
        tag = 'TC' if s == 'Thermocouple' else 'IR'
        guardar(fig_panorama(datos, metricas, s), f'{CARPETA_SALIDA}/01_panorama_{tag}.png')
        for nombre, (df, meta) in datos.items():
            m = metricas[(metricas.experimento == nombre) & (metricas.sensor == s)]
            guardar(fig_filtro(df, nombre, s), f'{CARPETA_SALIDA}/00_filtro_{nombre}_{tag}.png')
            guardar(fig_superpuestas(df, meta, m, nombre, s), f'{CARPETA_SALIDA}/02_superpuestas_{nombre}_{tag}.png')
            guardar(fig_ciclo_a_ciclo(m, meta, nombre, s), f'{CARPETA_SALIDA}/03_ciclo_a_ciclo_{nombre}_{tag}.png')
            guardar(fig_ideal_vs_real(df, meta, m, nombre, s), f'{CARPETA_SALIDA}/06_ideal_vs_real_{nombre}_{tag}.png')
            guardar(fig_puntos(m, meta, nombre, s), f'{CARPETA_SALIDA}/07_max_min_ext_{nombre}_{tag}.png')
        guardar(fig_comparacion(res, datos, s), f'{CARPETA_SALIDA}/04_comparacion_{tag}.png')
    if len(sensores) == 2:
        guardar(fig_sensores(res, datos), f'{CARPETA_SALIDA}/05_TC_vs_IR.png')

    # selección: la gráfica más útil de cada tipo, con el sensor de referencia
    tag = 'TC' if SENSOR_REFERENCIA == 'Thermocouple' else 'IR'
    elegidas = [f'04_comparacion_{tag}.png', f'01_panorama_{tag}.png', '05_TC_vs_IR.png'] + \
               [f'{p}_{n}_{tag}.png' for n in datos for p in ('06_ideal_vs_real', '07_max_min_ext', '02_superpuestas')]
    for f in elegidas:
        if os.path.exists(os.path.join(CARPETA_SALIDA, f)):
            shutil.copy(os.path.join(CARPETA_SALIDA, f), sel)

    with pd.ExcelWriter(f'{CARPETA_SALIDA}/resumen_pcr.xlsx') as xw:
        res.round(4).to_excel(xw, sheet_name='resumen', index=False)
        metricas.round(4).to_excel(xw, sheet_name='por_ciclo', index=False)
        pd.DataFrame([{'experimento': n} | {k: v for k, v in m.items() if k not in ('failed_list', 'settings')}
                      for n, (_, m) in datos.items()]).to_excel(xw, sheet_name='metadata', index=False)
        st = {n: m['settings'] for n, (_, m) in datos.items() if 'settings' in m}
        if st:
            pd.DataFrame(st).rename_axis('parametro').to_excel(xw, sheet_name='settings_controlador')
    print('Listo →', CARPETA_SALIDA, '(selección en', sel + ')')
    return datos, metricas, res


if __name__ == '__main__':
    import sys
    if len(sys.argv) == 3 and not sys.argv[1].startswith('-'):
        CARPETA_DATOS, CARPETA_SALIDA = sys.argv[1], sys.argv[2]
        import matplotlib; matplotlib.use('Agg')
    datos, metricas, res = correr()
