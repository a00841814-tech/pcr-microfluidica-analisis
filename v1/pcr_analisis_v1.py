# %% [markdown]
# # Análisis térmico PCR microfluídico — dispositivo Newn1
# Segmenta cada experimento con las columnas `cycle` y `phase` del archivo, filtra la señal
# y calcula, ciclo por ciclo, rampas (heating/cooling rate) y holds (duración, T media, desv. est.).
#
# **Qué se analiza (según notas de la reunión):**
# | Fase | Métricas |
# |---|---|
# | `cooling` | cooling rate, duración |
# | `low_hold` | duración (media, DE), T media, DE ciclo tras ciclo |
# | `ext_ramp` | heating rate (1ª rampa de calentamiento) |
# | `ext_hold` | duración, T media |
# | `high_ramp` | heating rate (2ª rampa de calentamiento) |
# | `high_hold` | duración (media, DE), T media |
# | `denat_hold`, `final_ext` | duración y T media (1 por experimento → se compara entre experimentos) |
# | `fluor_read`, `initial`, `denat_ramp`, `final_read`, `done` | se omiten |

# %%
# ===================== CONFIGURACIÓN =====================
CARPETA_DATOS = '/content/datos'          # carpeta con los .xlsx / .csv (en Colab)
CARPETA_SALIDA = '/content/resultados'    # aquí se guardan gráficas y Excel resumen

SENSOR = 'ambos'            # 'Thermocouple', 'IR Object' o 'ambos' (gráficas separadas por sensor)

FILTRAR = True              # filtrar la señal antes de calcular (quita picos y rizos)
VENTANA_MEDIANA_S = 1.0     # filtro de mediana: elimina picos aislados (s)
VENTANA_SAVGOL_S = 1.0      # Savitzky-Golay: suaviza rizos del control (s)

EXCLUIR_CICLOS_FALLIDOS = True   # quitar de las estadísticas los ciclos marcados como fallidos en metadata
EXCLUIR_CICLO_1_HIGH_RAMP = True # el high_ramp del ciclo 1 viene de la desnaturalización (no es una rampa real)

# umbral mínimo para marcar un ciclo como atípico (además de 3.5·MAD)
MIN_DESVIO = {'duracion_s': lambda med: max(1.0, 0.15 * abs(med)),   # s
              'rate':       lambda med: 0.25 * abs(med),   # aplica a TASA_PRINCIPAL               # 25 % de la tasa típica
              'T_media':    lambda med: 1.0}                           # °C

TASA_PRINCIPAL = 'rate_10_90'   # 'rate_10_90' (recomendada) o 'rate_avg' (ΔT/Δt de toda la fase)

FASES_RAMPA = ['ext_ramp', 'high_ramp', 'cooling']
FASES_HOLD = ['low_hold', 'ext_hold', 'high_hold']
FASES_UNICAS = ['denat_hold', 'final_ext']
# ==========================================================

# %%
import re, glob, os
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.signal import savgol_filter, medfilt

COLORES = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6']
COLOR_FALLIDO = '#e34948'
plt.rcParams.update({'axes.spines.top': False, 'axes.spines.right': False,
                     'axes.grid': True, 'grid.alpha': 0.3, 'grid.linestyle': ':',
                     'figure.dpi': 110, 'font.size': 9})

NOMBRE_FASE = {'ext_ramp': 'Extension ramp (1er calentamiento)', 'high_ramp': 'High ramp (2º calentamiento)',
               'cooling': 'Cooling', 'low_hold': 'Low hold', 'ext_hold': 'Extension hold',
               'high_hold': 'High hold', 'denat_hold': 'Denat hold', 'final_ext': 'Final extension'}


# %% [markdown]
# ## 1. Lectura de archivos (Excel o CSV) y metadatos

# %%
def parsear_metadatos(prefijo):
    """Convierte 'project: Newn1-high_temp: 90.0-...' en dict. Separa solo en '-' seguido de 'clave:'
    (así no rompe la fecha 'start:2026-10-21T...')."""
    meta = {}
    for parte in re.split(r'-(?=[A-Za-z_]+\s*:)', prefijo):
        if ':' not in parte:
            continue
        k, v = parte.split(':', 1)
        k, v = k.strip(), v.strip()
        try:
            v = float(v)
        except ValueError:
            pass
        meta[k] = v
    fc = str(meta.get('failed_cycles', 'none'))
    meta['failed_list'] = [] if fc == 'none' else [int(x) for x in fc.split('|') if x.strip()]
    return meta


def cargar_experimento(ruta):
    """Devuelve (df_temperatura, metadatos, nombre). Acepta .xlsx/.xls (hojas metadata + temperature)
    o .csv (1ª línea = prefijo de metadatos, 2ª = encabezados)."""
    ruta = Path(ruta)
    if ruta.suffix.lower() in ('.xlsx', '.xls'):
        hojas = pd.read_excel(ruta, sheet_name=None)
        md = hojas['metadata']
        prefijo = md.loc[md['key'] == 'prefix', 'value'].iloc[0]
        df = hojas['temperature']
    else:
        with open(ruta) as f:
            prefijo = f.readline().strip()
        df = pd.read_csv(ruta, skiprows=1)
    meta = parsear_metadatos(prefijo)
    m = re.search(r'(\d{8}_\d{6})', ruta.stem)
    nombre = f"{meta.get('project', 'exp')}_{m.group(1)}" if m else ruta.stem
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
    k = max(3, int(round(VENTANA_MEDIANA_S / ts)) | 1)
    yg = medfilt(yg, k)
    w = max(5, int(round(VENTANA_SAVGOL_S / ts)) | 1)
    yg = savgol_filter(yg, w, 2)
    return np.interp(t, grid, yg)


def preparar(df, meta):
    ts = meta.get('ts', 0.1)
    for s in ('Thermocouple', 'IR Object'):
        df[s + '_f'] = filtrar_senal(df['t_s'].values, df[s].values, ts) if FILTRAR else df[s]
    # id de segmento: cambia cada vez que cambia la fase o el ciclo
    df['seg'] = (df['phase'].ne(df['phase'].shift()) | df['cycle'].ne(df['cycle'].shift())).cumsum()
    return df


# %% [markdown]
# ## 3. Métricas por segmento
# * **Rampas:** `rate_avg` = ΔT/Δt entre inicio y fin de la fase; `rate_10_90` = 0.8·ΔT / tiempo entre
#   cruzar el 10 % y el 90 % del cambio (no incluye el arranque lento ni la cola); `retardo_s` = tiempo
#   desde que empieza la fase hasta que la T cruza el 10 % (respuesta del calefactor / sensor);
#   `rate_max` = dT/dt pico.
# * **Atípicos:** ciclos que se alejan de la mediana > 3.5·MAD y más de un mínimo práctico (`MIN_DESVIO`).
#   Solo se marcan (círculo negro en las gráficas); los **fallidos** de metadata sí se excluyen.
# * **Holds:** duración, T media, DE dentro del hold, mín., máx. y error contra el setpoint.

# %%
def setpoint(fase, meta):
    return {'low_hold': meta.get('low_temp'), 'ext_hold': meta.get('ext_temp'),
            'high_hold': meta.get('high_temp'), 'denat_hold': meta.get('denat_temp'),
            'final_ext': meta.get('ext_temp')}.get(fase, np.nan)


def tasa_10_90(t, y, objetivo=None):
    """Tasa entre el 10 % y el 90 % del cambio T_inicio → T_fin de la fase (o `objetivo` si se da).
    Usa el ÚLTIMO cruce de cada nivel, así un pico o bajón falso no adelanta el cruce.
    Devuelve (tasa °C/s, retardo_s = tiempo desde inicio de fase hasta cruzar el 10 %)."""
    y0 = y[0]
    y1 = objetivo if objetivo is not None and np.isfinite(objetivo) else y[-1]
    dT = y1 - y0
    if abs(dT) < 1:
        return np.nan, np.nan
    frac = (y - y0) / dT
    def ultimo_cruce(nivel):
        debajo = np.where(frac < nivel)[0]
        if len(debajo) == 0:
            return 0
        i = debajo[-1]
        if i + 1 >= len(frac):
            return None                      # nunca llegó al nivel
        # interpolación lineal entre i e i+1
        f0, f1 = frac[i], frac[i + 1]
        return t[i] + (nivel - f0) / (f1 - f0) * (t[i + 1] - t[i]) if f1 != f0 else t[i + 1]
    t10, t90 = ultimo_cruce(0.1), ultimo_cruce(0.9)
    if t10 is None or t90 is None or t90 <= t10:
        return np.nan, np.nan
    return 0.8 * dT / (t90 - t10), t10 - t[0]



def metricas_experimento(df, meta, nombre, sensor):
    col = sensor + '_f'
    filas = []
    for seg, g in df.groupby('seg', sort=True):
        fase, ciclo = g['phase'].iloc[0], int(g['cycle'].iloc[0])
        if fase not in FASES_RAMPA + FASES_HOLD + FASES_UNICAS or len(g) < 3:
            continue
        t, y = g['t_s'].values, g[col].values
        fila = {'experimento': nombre, 'sensor': sensor, 'ciclo': ciclo, 'fase': fase,
                'fallido': ciclo in meta['failed_list'],
                't_inicio_s': t[0], 'duracion_s': t[-1] - t[0], 'T_inicio': y[0], 'T_fin': y[-1]}
        if fase in FASES_RAMPA:
            dt = np.diff(t); dy = np.diff(y)
            ok = dt > 1e-3
            deriv = dy[ok] / dt[ok]
            r1090, retardo = tasa_10_90(t, y)
            fila.update(rate_avg=(y[-1] - y[0]) / (t[-1] - t[0]),
                        rate_10_90=r1090, retardo_s=retardo,
                        rate_max=(deriv.max() if fase != 'cooling' else deriv.min()) if ok.any() else np.nan)
        else:
            sp = setpoint(fase, meta)
            fila.update(setpoint=sp, T_media=y.mean(), T_desv=y.std(ddof=1), T_min=y.min(), T_max=y.max(),
                        error_vs_setpoint=y.mean() - sp)
        if fase == 'high_ramp' and ciclo == 1 and EXCLUIR_CICLO_1_HIGH_RAMP:
            fila['excluido'] = True
        filas.append(fila)
    m = pd.DataFrame(filas)
    m['excluido'] = m.get('excluido', False)
    m['excluido'] = m['excluido'].fillna(False).astype(bool)
    if EXCLUIR_CICLOS_FALLIDOS:
        m.loc[m['fallido'], 'excluido'] = True
    # ciclos atípicos (solo se marcan, no se excluyen): se aleja de la mediana más de 3.5·MAD robusto
    # Y más de un mínimo práctico (MIN_DESVIO), para no marcar diferencias de centésimas
    m['atipico'] = False
    for fase, g in m[m.fase.isin(FASES_RAMPA + FASES_HOLD) & ~m.excluido].groupby('fase'):
        for v, minimo in MIN_DESVIO.items():
            v = TASA_PRINCIPAL if v == 'rate' else v
            if v in g and g[v].notna().sum() > 5:
                x = g[v]; dev = (x - x.median()).abs()
                lim = max(3.5 * 1.4826 * dev.median(), minimo(x.median()))
                m.loc[g.index[(dev > lim).values], 'atipico'] = True
    return m


# %%
def resumen(metricas):
    """Media, DE, CV %, mín y máx ciclo tras ciclo por experimento / sensor / fase."""
    vals = ['duracion_s', 'rate_avg', 'rate_10_90', 'rate_max', 'retardo_s', 'T_media', 'T_desv', 'T_min', 'T_max',
            'error_vs_setpoint']
    d = metricas[~metricas['excluido']]
    filas = []
    for (exp, sen, fase), g in d.groupby(['experimento', 'sensor', 'fase']):
        fila = {'experimento': exp, 'sensor': sen, 'fase': fase, 'n_ciclos': len(g),
                'ciclos_atipicos': ','.join(map(str, g.loc[g.atipico, 'ciclo'])) if 'atipico' in g else ''}
        for v in vals:
            if v in g and g[v].notna().any():
                x = g[v].dropna()
                fila[f'{v}_media'] = x.mean()
                fila[f'{v}_de'] = x.std(ddof=1) if len(x) > 1 else np.nan
                if v in ('duracion_s', 'rate_avg', 'rate_10_90', 'T_media'):
                    fila[f'{v}_cv%'] = 100 * fila[f'{v}_de'] / abs(fila[f'{v}_media']) if len(x) > 1 else np.nan
        filas.append(fila)
    orden = {f: i for i, f in enumerate(FASES_RAMPA + FASES_HOLD + FASES_UNICAS)}
    r = pd.DataFrame(filas)
    return r.sort_values(['sensor', 'fase', 'experimento'], key=lambda s: s.map(orden) if s.name == 'fase' else s)


# %% [markdown]
# ## 4. Gráficas

# %%
def fig_panorama(datos, sensor):
    """Señal completa de cada experimento: cruda (tenue) y filtrada, con fallidos marcados."""
    fig, axs = plt.subplots(len(datos), 1, figsize=(13, 2.6 * len(datos)), sharex=True, squeeze=False)
    for ax, (nombre, (df, meta)) in zip(axs[:, 0], datos.items()):
        ax.plot(df['t_s'], df[sensor], color='#b8b7b0', lw=0.6, label='cruda')
        ax.plot(df['t_s'], df[sensor + '_f'], color=COLORES[0], lw=0.9, label='filtrada')
        for c in meta['failed_list']:
            g = df[(df.cycle == c) & (df.phase == 'ext_ramp')]
            if len(g):
                ax.axvspan(g.t_s.iloc[0], g.t_s.iloc[-1], color=COLOR_FALLIDO, alpha=0.25, lw=0)
        ax.set_title(f'{nombre} — {sensor}' + (f"  (ciclos fallidos: {meta['failed_list']})" if meta['failed_list'] else ''),
                     loc='left', fontsize=10)
        ax.set_ylabel('T (°C)')
    axs[0, 0].legend(loc='lower right', ncol=2, frameon=False)
    axs[-1, 0].set_xlabel('Tiempo (s)')
    fig.tight_layout()
    return fig


def fig_filtro(df, nombre, sensor, ciclo=10):
    """Zoom a un ciclo: señal cruda vs filtrada, para revisar que el filtro no deforme las rampas."""
    g = df[df.cycle == ciclo]
    fig, ax = plt.subplots(figsize=(13, 3.6))
    ax.plot(g.t_s, g[sensor], color='#b8b7b0', lw=1.2, label='cruda')
    ax.plot(g.t_s, g[sensor + '_f'], color=COLORES[0], lw=1.6, label='filtrada')
    for _, gg in g.groupby('seg'):
        ax.axvline(gg.t_s.iloc[0], color='#52514e', lw=0.5, ls=':')
        ax.text(gg.t_s.iloc[0] + 0.3, ax.get_ylim()[1], gg.phase.iloc[0], rotation=90, va='top', fontsize=7, color='#52514e')
    ax.set_title(f'{nombre} — {sensor}: ciclo {ciclo}, cruda vs filtrada', loc='left', fontsize=10)
    ax.set_xlabel('Tiempo (s)'); ax.set_ylabel('T (°C)'); ax.legend(frameon=False, loc='lower left')
    fig.tight_layout()
    return fig


def fig_superpuestas(df, meta, metr, nombre, sensor):
    """Todos los ciclos alineados en t = 0 por fase, con curva promedio."""
    fases = FASES_RAMPA + FASES_HOLD
    fig, axs = plt.subplots(2, 3, figsize=(14, 7.5))
    for ax, fase in zip(axs.ravel(), fases):
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
            tmax = np.median([c[0][-1] for c in curvas])
            grid = np.linspace(0, tmax, 200)
            ym = np.mean([np.interp(grid, t, y) for t, y in curvas], axis=0)
            ax.plot(grid, ym, color='#0b0b0b', lw=2, label='promedio')
        d = mf[~mf.excluido]
        if fase in FASES_RAMPA:
            txt = (f"10–90 %: {d.rate_10_90.mean():+.2f} ± {d.rate_10_90.std():.2f} °C/s\nΔT/Δt fase: {d.rate_avg.mean():+.2f} ± {d.rate_avg.std():.2f} °C/s"
                   f"\nretardo = {d.retardo_s.mean():.2f} ± {d.retardo_s.std():.2f} s")
        else:
            sp = setpoint(fase, meta)
            ax.axhline(sp, color='#52514e', ls='--', lw=1, label=f'setpoint {sp:.0f} °C')
            txt = f"T = {d.T_media.mean():.2f} ± {d.T_media.std():.2f} °C\ndur = {d.duracion_s.mean():.2f} ± {d.duracion_s.std():.2f} s"
        ax.text(0.98, 0.04 if fase != 'cooling' else 0.96, txt, transform=ax.transAxes, ha='right',
                va='bottom' if fase != 'cooling' else 'top', fontsize=8,
                bbox=dict(fc='white', ec='none', alpha=0.85))
        ax.set_title(NOMBRE_FASE[fase] + f'  (n = {len(d)})', loc='left', fontsize=10)
        ax.set_xlabel('t desde inicio de fase (s)'); ax.set_ylabel('T (°C)')
    if (metr.fallido).any():
        axs[0, 0].plot([], [], color=COLOR_FALLIDO, label='ciclo fallido')
    axs[0, 0].legend(frameon=False, fontsize=8, loc='upper left')
    axs[1, 0].legend(frameon=False, fontsize=8, loc='upper left')
    fig.suptitle(f'{nombre} — {sensor}: ciclos superpuestos', x=0.01, ha='left', fontsize=12, fontweight='bold')
    fig.tight_layout()
    return fig


def fig_ciclo_a_ciclo(metr, nombre, sensor):
    """Cada métrica graficada contra el número de ciclo (vista 'separada', para ver repetibilidad)."""
    paneles = [(f, TASA_PRINCIPAL, 'rate (°C/s)') for f in FASES_RAMPA] + \
              [(f, 'T_media', 'T media (°C)') for f in FASES_HOLD] + \
              [(f, 'duracion_s', 'duración (s)') for f in FASES_RAMPA + FASES_HOLD]
    fig, axs = plt.subplots(4, 3, figsize=(14, 11), sharex=True)
    for ax, (fase, var, et) in zip(axs.ravel(), paneles):
        d = metr[metr.fase == fase].sort_values('ciclo')
        ok = d[~d.excluido]
        if var == 'T_media':
            ax.fill_between(d.ciclo, d.T_min, d.T_max, color=COLORES[0], alpha=0.15, lw=0, label='mín–máx')
        mu, sd = ok[var].mean(), ok[var].std()
        ax.axhspan(mu - sd, mu + sd, color='#52514e', alpha=0.1, lw=0)
        ax.axhline(mu, color='#52514e', lw=1, ls='--', label=f'media ± DE: {mu:.2f} ± {sd:.2f}')
        ax.plot(ok.ciclo, ok[var], 'o-', color=COLORES[0], ms=3.5, lw=1)
        a = ok[ok.atipico]
        if len(a):
            ax.plot(a.ciclo, a[var], 'o', mfc='none', mec='#0b0b0b', ms=8, mew=1.2, label='atípico')
        f = d[d.fallido]
        if len(f):
            ax.plot(f.ciclo, f[var], 'X', color=COLOR_FALLIDO, ms=7, label='fallido (excluido)' if EXCLUIR_CICLOS_FALLIDOS else 'fallido')
        ax.set_title(f'{NOMBRE_FASE[fase]} — {et}', loc='left', fontsize=9)
        ax.legend(frameon=False, fontsize=7, loc='best')
    for ax in axs[-1]:
        ax.set_xlabel('ciclo')
    fig.suptitle(f'{nombre} — {sensor}: métricas ciclo tras ciclo', x=0.01, ha='left', fontsize=12, fontweight='bold')
    fig.tight_layout()
    return fig


def fig_comparacion(res, sensor):
    """Barras media ± DE por experimento para las métricas clave."""
    r = res[res.sensor == sensor]
    exps = sorted(r.experimento.unique())
    paneles = [(f, TASA_PRINCIPAL, 'rate (°C/s)') for f in FASES_RAMPA] + \
              [(f, 'T_media', 'T media (°C)') for f in FASES_HOLD] + \
              [(f, 'duracion_s', 'duración (s)') for f in FASES_HOLD] + \
              [(f, 'duracion_s', 'duración (s)') for f in FASES_RAMPA] + \
              [(f, 'T_media', 'T media (°C)') for f in FASES_UNICAS]
    fig, axs = plt.subplots(3, 5, figsize=(17, 9.5))
    axs = axs.ravel()
    for ax, (fase, var, et) in zip(axs, paneles):
        d = r[r.fase == fase].set_index('experimento').reindex(exps)
        mu, sd = d[f'{var}_media'].values, d.get(f'{var}_de', pd.Series(np.nan, index=d.index)).values
        x = np.arange(len(exps))
        ax.bar(x, mu, yerr=np.nan_to_num(sd), color=COLORES[:len(exps)], width=0.6, capsize=4,
               edgecolor='white', linewidth=2, error_kw=dict(lw=1, ecolor='#52514e'))
        lo = np.nanmin(mu - np.nan_to_num(sd)); hi = np.nanmax(mu + np.nan_to_num(sd))
        if not var.startswith('rate'):
            pad = (hi - lo) * 0.6 + 0.5
            ax.set_ylim(lo - pad, hi + pad * 0.6)
        elif hi > 0:
            ax.set_ylim(0, hi * 1.3)
        else:
            ax.set_ylim(lo * 1.35, 0)
        for xi, m_, s_ in zip(x, mu, sd):
            ax.annotate(f'{m_:.2f}' + (f'\n±{s_:.2f}' if np.isfinite(s_) else ''), (xi, m_ + (np.nan_to_num(s_) if m_ >= 0 else -np.nan_to_num(s_))),
                        ha='center', va='bottom' if m_ >= 0 else 'top', fontsize=7, xytext=(0, 2 if m_ >= 0 else -2), textcoords='offset points')
        ax.set_xticks(x, [e.split('_', 1)[-1].replace('_', '\n') for e in exps], fontsize=7)
        ax.set_title(f'{NOMBRE_FASE[fase]}\n{et}', loc='left', fontsize=9)
    for ax in axs[len(paneles):]:
        ax.axis('off')
    # tabla con duraciones de denat hold y final extension (1 valor por experimento)
    ax = axs[-1]
    filas = []
    for fase in FASES_UNICAS:
        d = r[r.fase == fase].set_index('experimento').reindex(exps)['duracion_s_media']
        filas.append([{'denat_hold': 'Denat', 'final_ext': 'Final ext'}[fase]] + [f'{v:.2f}' for v in d] + [f'{d.mean():.2f} ± {d.std():.2f}'])
    tb = ax.table(cellText=filas, colLabels=['dur. (s)'] + [e.split('_')[-1] for e in exps] + ['media ± DE'],
                  loc='center', cellLoc='center')
    tb.auto_set_font_size(False); tb.set_fontsize(7); tb.scale(1, 1.6)
    for c in tb.get_celld().values():
        c.set_edgecolor('#d9d8d3')
    fig.suptitle(f'Comparación entre experimentos — {sensor} (media ± DE ciclo tras ciclo)', x=0.01, ha='left',
                 fontsize=12, fontweight='bold')
    fig.tight_layout()
    return fig


def fig_sensores(res):
    """Termopar vs IR para las métricas clave (paneles separados, mismo eje dentro de cada panel)."""
    exps = sorted(res.experimento.unique())
    paneles = [(f, TASA_PRINCIPAL, 'rate (°C/s)') for f in FASES_RAMPA] + [(f, 'T_media', 'T media (°C)') for f in FASES_HOLD]
    fig, axs = plt.subplots(2, 3, figsize=(14, 7))
    for ax, (fase, var, et) in zip(axs.ravel(), paneles):
        x = np.arange(len(exps)); w = 0.36
        for i, sen in enumerate(['Thermocouple', 'IR Object']):
            d = res[(res.fase == fase) & (res.sensor == sen)].set_index('experimento').reindex(exps)
            ax.bar(x + (i - 0.5) * w, d[f'{var}_media'], w, yerr=d[f'{var}_de'], color=COLORES[i], capsize=3,
                   edgecolor='white', linewidth=2, error_kw=dict(lw=1, ecolor='#52514e'), label=sen)
        if var == 'T_media':
            vals = res[(res.fase == fase)][f'{var}_media']
            ax.set_ylim(vals.min() - 4, vals.max() + 3)
        ax.set_xticks(x, [e.split('_', 1)[-1].replace('_', '\n') for e in exps], fontsize=7)
        ax.set_title(f'{NOMBRE_FASE[fase]} — {et}', loc='left', fontsize=9)
    axs[0, 0].legend(frameon=False, fontsize=8)
    fig.suptitle('Termopar vs sensor IR', x=0.01, ha='left', fontsize=12, fontweight='bold')
    fig.tight_layout()
    return fig


# %% [markdown]
# ## 5. Ejecutar

# %%
def guardar(fig, ruta):
    fig.savefig(ruta, bbox_inches='tight')
    plt.close(fig)


def correr():
    os.makedirs(CARPETA_SALIDA, exist_ok=True)
    rutas = sorted(glob.glob(os.path.join(CARPETA_DATOS, '*.xlsx')) + glob.glob(os.path.join(CARPETA_DATOS, '*.xls'))
                   + glob.glob(os.path.join(CARPETA_DATOS, '*.csv')))
    print(f'{len(rutas)} archivos encontrados')
    sensores = ['Thermocouple', 'IR Object'] if SENSOR == 'ambos' else [SENSOR]
    datos, todas = {}, []
    for r in rutas:
        df, meta, nombre = cargar_experimento(r)
        df = preparar(df, meta)
        datos[nombre] = (df, meta)
        print(f"  {nombre}: {len(df)} puntos, {df.t_s.iloc[-1]:.0f} s, ciclos fallidos: {meta['failed_list'] or 'ninguno'}")
        for s in sensores:
            todas.append(metricas_experimento(df, meta, nombre, s))
    metricas = pd.concat(todas, ignore_index=True)
    res = resumen(metricas)

    for s in sensores:
        tag = 'TC' if s == 'Thermocouple' else 'IR'
        guardar(fig_panorama(datos, s), f'{CARPETA_SALIDA}/01_panorama_{tag}.png')
        for nombre, (df, meta) in datos.items():
            guardar(fig_filtro(df, nombre, s), f'{CARPETA_SALIDA}/00_filtro_{nombre}_{tag}.png')
            m = metricas[(metricas.experimento == nombre) & (metricas.sensor == s)]
            guardar(fig_superpuestas(df, meta, m, nombre, s), f'{CARPETA_SALIDA}/02_superpuestas_{nombre}_{tag}.png')
            guardar(fig_ciclo_a_ciclo(m, nombre, s), f'{CARPETA_SALIDA}/03_ciclo_a_ciclo_{nombre}_{tag}.png')
        guardar(fig_comparacion(res, s), f'{CARPETA_SALIDA}/04_comparacion_{tag}.png')
    if len(sensores) == 2:
        guardar(fig_sensores(res), f'{CARPETA_SALIDA}/05_TC_vs_IR.png')

    plt.close('all')
    with pd.ExcelWriter(f'{CARPETA_SALIDA}/resumen_pcr.xlsx') as xw:
        res.round(4).to_excel(xw, sheet_name='resumen', index=False)
        metricas.round(4).to_excel(xw, sheet_name='por_ciclo', index=False)
        pd.DataFrame([{k: v for k, v in m.items() if k != 'failed_list'} | {'experimento': n}
                      for n, (_, m) in datos.items()]).to_excel(xw, sheet_name='metadata', index=False)
    print('Listo →', CARPETA_SALIDA)
    return datos, metricas, res


if __name__ == '__main__':
    import sys
    if len(sys.argv) > 2:
        CARPETA_DATOS, CARPETA_SALIDA = sys.argv[1], sys.argv[2]
        import matplotlib; matplotlib.use('Agg')
    datos, metricas, res = correr()
