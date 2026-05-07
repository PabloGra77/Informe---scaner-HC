"""Punto de entrada unico de la aplicacion empaquetada (.exe).

- Sin argumentos: abre la interfaz grafica.
- Con --cli + argumentos: ejecuta el procesador en modo linea de comandos.

Esto permite que el .exe se reinvoque a si mismo para correr el procesador
sin necesidad de incluir Python por separado.
"""
from __future__ import annotations

import multiprocessing
import sys


def main() -> None:
    # Necesario para que ProcessPoolExecutor funcione en el .exe.
    multiprocessing.freeze_support()

    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        # Reescribir argv para que argparse del procesador funcione igual.
        sys.argv = [sys.argv[0], *sys.argv[2:]]
        from procesar_pdfs import main as cli_main
        cli_main()
        return

    from interfaz_procesador import main as gui_main
    gui_main()


if __name__ == "__main__":
    main()
