# Changelog

Todas las versiones notables de este proyecto se documentan aqui.
El formato sigue [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/)
y este proyecto usa [SemVer](https://semver.org/lang/es/).

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
