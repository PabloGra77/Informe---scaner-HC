# Changelog

Todas las versiones notables de este proyecto se documentan aqui.
El formato sigue [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/)
y este proyecto usa [SemVer](https://semver.org/lang/es/).

## [1.2.2] - 2026-05-08
### Added
- Modo "Informe HC NUEVOS": dos columnas nuevas en el Excel de salida,
  `novedades` y `fecha_novedad`, extraidas de la seccion "12. NOVEDADES"
  del formato de historia clinica (con fallback si el titulo no esta numerado).

## [1.2.1] - 2026-05-08
### Fixed
- La interfaz ahora muestra el progreso en vivo (TOTAL / PROCESADOS / FALTAN
  y log) durante toda la ejecucion. Antes el subproceso retenia la salida en
  buffer y los contadores se quedaban en 0 hasta el final.
- Se desactiva el buffering del subproceso (`python -u`, `PYTHONUNBUFFERED=1`,
  `bufsize=1`) y los `print` del procesador hacen flush inmediato.
- Se emite progreso tambien en el primer y ultimo PDF, y en cada archivo
  cuando el lote es pequeno (<= 50 PDFs).

## [1.2.0] - 2026-05-07
### Changed
- Rediseno completo de la interfaz: tema moderno, header azul, tarjetas
  blancas con borde, botones con color (primary / danger), panel de
  estadisticas tipo cards, log con estilo consola y coloreo de mensajes,
  selector inteligente segun el modo y hint dinamico.

## [1.1.0] - 2026-05-07
### Added
- Nuevo modo "Informe HC NUEVOS" (`--accion informe-nuevos`):
  extrae `numero_ide`, `paciente`, `ingreso`, `tipo_servicio`,
  `fecha_atencion`, `diagnostico` (tipo / clase / codigo / descripcion),
  `nota` y `ruta_archivo` para el formato nuevo de historia clinica.
- Selector de modo en la interfaz.

## [1.0.0] - 2026-05-07
### Added
- Procesamiento masivo de PDF con extraccion de campos por regex.
- Renombrado por `numero_documento` y `fecha_atencion` en MAYUSCULA.
- Encarpetado desde Excel.
- Interfaz grafica (Tkinter) con progreso, log y controles.
- Sistema de actualizaciones desde GitHub Releases (boton "Buscar actualizaciones").
- Build de `.exe` con PyInstaller y workflow de release automatico.
