"""Sistema simple de actualizaciones contra GitHub Releases.

- Consulta /releases/latest del repositorio configurado en version.py
- Compara la version actual con la del ultimo release (tag tipo v1.2.3)
- Si hay version nueva descarga el .exe del asset y lo aplica reemplazando
  el ejecutable actual mediante un script .bat de relevo.

Uso desde la interfaz:

    from updater import check_for_updates, apply_update_in_background

    info = check_for_updates()
    if info.has_update:
        apply_update_in_background(info)
"""
from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import threading
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

from version import APP_VERSION, GITHUB_OWNER, GITHUB_REPO


GITHUB_API_LATEST = (
    f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
)
USER_AGENT = "RenombrePanacea-Updater"


@dataclass
class UpdateInfo:
    has_update: bool
    current_version: str
    latest_version: str
    download_url: Optional[str]
    release_url: Optional[str]
    release_notes: str = ""
    error: Optional[str] = None


def _parse_version(value: str) -> tuple:
    value = (value or "").strip().lstrip("vV")
    parts = re.split(r"[.\-+]", value)
    out = []
    for part in parts:
        if part.isdigit():
            out.append(int(part))
        else:
            # Solo se considera la parte numerica para comparar.
            break
    return tuple(out) or (0,)


def _is_newer(latest: str, current: str) -> bool:
    return _parse_version(latest) > _parse_version(current)


def check_for_updates(timeout: float = 10.0) -> UpdateInfo:
    """Consulta GitHub y retorna info sobre la version disponible."""
    try:
        req = urllib.request.Request(
            GITHUB_API_LATEST,
            headers={"User-Agent": USER_AGENT, "Accept": "application/vnd.github+json"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:  # nosec - URL fija
            data = json.loads(resp.read().decode("utf-8", errors="replace"))
    except Exception as exc:
        return UpdateInfo(
            has_update=False,
            current_version=APP_VERSION,
            latest_version="",
            download_url=None,
            release_url=None,
            error=f"No se pudo consultar GitHub: {exc}",
        )

    tag = str(data.get("tag_name") or data.get("name") or "").strip()
    notes = str(data.get("body") or "")
    html_url = data.get("html_url")

    download_url: Optional[str] = None
    for asset in data.get("assets", []) or []:
        name = str(asset.get("name") or "").lower()
        if name.endswith(".exe"):
            download_url = asset.get("browser_download_url")
            break

    has_update = bool(tag) and _is_newer(tag, APP_VERSION)

    return UpdateInfo(
        has_update=has_update,
        current_version=APP_VERSION,
        latest_version=tag.lstrip("vV"),
        download_url=download_url,
        release_url=html_url,
        release_notes=notes,
    )


def _download_file(url: str, dest: Path, on_progress: Optional[Callable[[int, int], None]] = None) -> None:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=30) as resp:  # nosec - URL viene de release
        total = int(resp.headers.get("Content-Length") or 0)
        downloaded = 0
        chunk = 64 * 1024
        with open(dest, "wb") as fh:
            while True:
                buf = resp.read(chunk)
                if not buf:
                    break
                fh.write(buf)
                downloaded += len(buf)
                if on_progress:
                    on_progress(downloaded, total)


def _running_as_frozen_exe() -> bool:
    return getattr(sys, "frozen", False) and sys.executable.lower().endswith(".exe")


def _write_replace_script(new_exe: Path, current_exe: Path) -> Path:
    """Crea un .bat que espera a que se cierre el .exe, lo reemplaza y lo relanza."""
    script = Path(tempfile.gettempdir()) / "renombre_panacea_update.bat"
    content = f"""@echo off
setlocal
set "TARGET={current_exe}"
set "SOURCE={new_exe}"

rem Esperar a que el ejecutable actual se cierre
:wait
tasklist /FI "IMAGENAME eq {current_exe.name}" | find /I "{current_exe.name}" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait
)

rem Reemplazar
copy /Y "%SOURCE%" "%TARGET%" >nul
if errorlevel 1 (
    echo No se pudo reemplazar el ejecutable.
    pause
    exit /b 1
)

del /Q "%SOURCE%" >nul 2>nul

rem Relanzar
start "" "%TARGET%"
exit /b 0
"""
    script.write_text(content, encoding="utf-8")
    return script


def apply_update_in_background(
    info: UpdateInfo,
    on_progress: Optional[Callable[[int, int], None]] = None,
    on_done: Optional[Callable[[bool, str], None]] = None,
) -> None:
    """Descarga el .exe nuevo y lanza el reemplazo en un hilo aparte."""

    def _worker() -> None:
        try:
            if not info.download_url:
                raise RuntimeError("El release no tiene un .exe adjunto")

            if not _running_as_frozen_exe():
                raise RuntimeError(
                    "La actualizacion automatica solo funciona en la version .exe. "
                    "Descarga manualmente desde: " + (info.release_url or "GitHub")
                )

            current_exe = Path(sys.executable).resolve()
            new_exe = Path(tempfile.gettempdir()) / f"{current_exe.stem}_new.exe"

            _download_file(info.download_url, new_exe, on_progress=on_progress)

            script = _write_replace_script(new_exe, current_exe)

            # Lanza el .bat sin ventana y termina la app actual.
            os.startfile(str(script))  # noqa: S606 - intencional
            if on_done:
                on_done(True, "Actualizacion descargada. La aplicacion se reiniciara.")

            # Pequeno margen y salir
            os._exit(0)
        except Exception as exc:
            if on_done:
                on_done(False, str(exc))

    threading.Thread(target=_worker, daemon=True).start()
