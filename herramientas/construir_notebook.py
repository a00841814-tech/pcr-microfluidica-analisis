import json, re, sys
src_py, out = sys.argv[1], sys.argv[2]
s = open(src_py).read().split("\nif __name__ == '__main__':")[0]
cells = []
for blk in re.split(r'\n?# %%', s):
    blk = blk.strip('\n')
    if not blk.strip(): continue
    if blk.startswith(' [markdown]'):
        body = blk[len(' [markdown]'):].strip('\n')
        cells.append(('markdown', '\n'.join(l[2:] if l.startswith('# ') else l.lstrip('#') for l in body.split('\n'))))
    else:
        cells.append(('code', blk))
carga = '''# Cargar archivos: (A) USAR_DRIVE = True → usa la carpeta de Drive en CARPETA_DATOS
#                  (B) USAR_DRIVE = False → te pide subirlos desde tu computadora
# Ejemplo Drive: CARPETA_DATOS = '/content/drive/MyDrive/Datos Microfluídicos'
import os
USAR_DRIVE = True

try:
    from google.colab import files, drive
    if USAR_DRIVE:
        drive.mount('/content/drive')
        print('Archivos en la carpeta:', os.listdir(CARPETA_DATOS))
    else:
        os.makedirs(CARPETA_DATOS, exist_ok=True)
        for n, contenido in files.upload().items():
            with open(os.path.join(CARPETA_DATOS, n), 'wb') as f:
                f.write(contenido)
except ImportError:
    print('No estás en Colab: se usan los archivos que ya estén en', CARPETA_DATOS)'''
cells.insert(2, ('code', carga))
cells += [('code', 'datos, metricas, res = correr()\nres.round(3)'),
          ('markdown', '## 7. Ver la selección de gráficas y descargar resultados'),
          ('code', "from IPython.display import Image, display\nfor p in sorted(glob.glob(f'{CARPETA_SALIDA}/seleccion/*.png')):\n    print(os.path.basename(p)); display(Image(p))"),
          ('code', "# Tabla por ciclo de un experimento/sensor/fase concreto\nmetricas[(metricas.sensor == 'Thermocouple') & (metricas.fase == 'ext_hold')].head(10)"),
          ('code', "# Descargar todo en un .zip\nimport shutil\nshutil.make_archive('/content/resultados_pcr', 'zip', CARPETA_SALIDA)\ntry:\n    from google.colab import files; files.download('/content/resultados_pcr.zip')\nexcept ImportError:\n    pass")]
cfg = cells[1][1].replace("CARPETA_DATOS = '/content/datos'          # carpeta con los .xlsx / .csv",
                          "CARPETA_DATOS = '/content/drive/MyDrive/Datos Microfluídicos'   # carpeta con los .xlsx / .csv")
cells[1] = ('code', cfg)
nb = {'nbformat': 4, 'nbformat_minor': 5,
      'metadata': {'kernelspec': {'name': 'python3', 'display_name': 'Python 3'}, 'language_info': {'name': 'python'}, 'colab': {'provenance': []}},
      'cells': [{'cell_type': t, 'metadata': {}, 'source': c.splitlines(True), **({'outputs': [], 'execution_count': None} if t == 'code' else {})} for t, c in cells]}
json.dump(nb, open(out, 'w'), ensure_ascii=False, indent=1)
print(len(cells), 'celdas')
