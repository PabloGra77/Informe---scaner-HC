# RenombrePanacea — Renombre y Reporte de PDF

Aplicacion de escritorio (Windows) para procesar lotes de PDF de historia
clinica/escaneo: extrae datos por OCR-regex, renombra los archivos por
`numero_documento` + `fecha_atencion`, y genera un informe Excel.

Incluye interfaz grafica, modo CLI y un boton **Buscar actualizaciones** que
descarga e instala la ultima version publicada en GitHub Releases.

- Repositorio: https://github.com/PabloGra77/Informe---scaner-HC
- Version actual: ver `version.py` / `CHANGELOG.md`

## Caracteristicas

- Extrae: `codigo_actividad`, `programa`, `apellidos`, `nombres`,
  `tipo_identificacion`, `numero_documento`, `fecha_atencion`.
- Renombra los PDF en MAYUSCULA usando `NUMERODOCUMENTO_DDMMYYYY.pdf`.
- Encarpeta archivos por numero de documento desde un Excel.
- Procesamiento paralelo (multiprocessing).
- Informe Excel con division automatica si supera el millon de filas.
- Auto-actualizacion desde GitHub Releases (solo en la version `.exe`).

---

## Instalacion para usuarios finales (.exe)

1. Ve a la pagina de [Releases](https://github.com/PabloGra77/Informe---scaner-HC/releases/latest).
2. Descarga `RenombrePanacea.exe`.
3. Doble click. No requiere instalar Python.
4. Para futuras versiones usa el boton **Buscar actualizaciones** en la app.

## Instalacion para desarrolladores (codigo fuente)

Requiere Python 3.10 o superior en Windows.

```powershell
# Opcion automatica
.\INSTALAR_PROYECTO.bat
.\INICIAR_INTERFAZ.bat
```

```powershell
# Opcion manual
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## Uso por linea de comandos

```powershell
python procesar_pdfs.py --entrada "EJEMPLO PDF" --salida "INFORME_PDFS.xlsx"
```

Opciones utiles:

- `--sin-renombrar` solo informe.
- `--recursivo` busca en subcarpetas.
- `--workers N` procesos en paralelo.
- `--max-filas-excel N` (default 1.000.000).
- `--progreso-cada N` cada cuantos PDF se imprime el progreso.

### Encarpetar desde Excel

```powershell
python procesar_pdfs.py --accion encarpetar-desde-excel ^
  --excel-entrada "INFORME_MASIVO.xlsx" ^
  --excel-salida-encarpetado "INFORME_ENCARPETADO.xlsx"
```

### Informe HC NUEVOS (formato nuevo)

Extrae para cada PDF: `numero_ide`, `paciente`, `ingreso`, `tipo_servicio`
(EVOLUCION PSICOLOGIA, NOTA DE ENFERMERIA, etc.), `fecha_atencion`,
`diagnostico` (tipo / clase / codigo / descripcion), `nota` y `ruta_archivo`.

```powershell
python procesar_pdfs.py --accion informe-nuevos ^
  --entrada "PDF NUEVOS" --salida "INFORME_NUEVOS.xlsx" --recursivo
```

En la interfaz, eleige el modo **"Informe HC NUEVOS"** en el desplegable.

---

## Compilar el `.exe` localmente

```powershell
.\COMPILAR_EXE.bat
```

El binario queda en `dist\RenombrePanacea.exe`.

## Publicar una nueva version (mantenedor)

1. Edita `version.py` y sube `APP_VERSION` (ej. `1.1.0`).
2. Actualiza `CHANGELOG.md`.
3. Commit + push a `main`.
4. Crea y publica un tag con el mismo numero, prefijo `v`:

   ```powershell
   git tag v1.1.0
   git push origin v1.1.0
   ```

5. GitHub Actions (`.github/workflows/release.yml`) compila el `.exe` y lo
   publica como asset del release.
6. Los usuarios veran la actualizacion al pulsar **Buscar actualizaciones**.

> El comparador de versiones del updater espera tags tipo `vX.Y.Z`.
> El asset debe ser un archivo `.exe` (lo es por defecto del workflow).

---

## Estructura del proyecto

```
app.py                     # Punto de entrada (.exe): GUI o --cli
procesar_pdfs.py           # Logica de procesamiento (CLI)
interfaz_procesador.py     # Interfaz Tkinter
updater.py                 # Cliente de GitHub Releases
version.py                 # Nombre, version y repo
RenombrePanacea.spec       # Spec de PyInstaller
COMPILAR_EXE.bat           # Build local del .exe
INSTALAR_PROYECTO.bat      # Crea .venv e instala dependencias
INICIAR_INTERFAZ.bat       # Lanza la GUI desde codigo fuente
.github/workflows/release.yml  # CI: build + release automatico
requirements.txt
CHANGELOG.md
LICENSE
```

## Subir el proyecto a GitHub por primera vez

```powershell
git init
git add .
git commit -m "chore: primer commit"
git branch -M main
git remote add origin https://github.com/PabloGra77/Informe---scaner-HC.git
git push -u origin main
```

Luego publica la primera release siguiendo la seccion *Publicar una nueva version*.

## Licencia

MIT — ver [LICENSE](LICENSE).
