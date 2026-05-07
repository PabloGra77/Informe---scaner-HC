import os
import queue
import re
import subprocess
import sys
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from version import APP_NAME, APP_VERSION
from updater import check_for_updates, apply_update_in_background


def _app_dir() -> Path:
    """Carpeta donde vive el script o el .exe empaquetado."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _run_processor_args(args: list) -> list:
    """Construye el comando para llamar al procesador.

    - En modo .exe (frozen) se reinvoca el mismo ejecutable con --cli
      para reutilizar el runtime y la logica empaquetada.
    - En modo script se llama a procesar_pdfs.py con el python actual.
    """
    if getattr(sys, "frozen", False):
        return [sys.executable, "--cli", *args]
    return [sys.executable, str(_app_dir() / "procesar_pdfs.py"), *args]


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}  -  Renombre y Reporte de PDF")
        self.geometry("960x680")

        self.output_queue: queue.Queue = queue.Queue()
        self.process = None
        self.start_time = 0.0
        self.total_detectado = 0
        self.procesados = 0

        self._build_ui()
        self.after(150, self._poll_queue)

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # Barra superior con version y boton de actualizaciones
        top = ttk.Frame(main)
        top.grid(row=0, column=0, columnspan=2, sticky="we", pady=(0, 8))
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text=f"{APP_NAME}  v{APP_VERSION}", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Button(top, text="Buscar actualizaciones", command=self._on_check_updates).grid(row=0, column=1, sticky="e")

        ttk.Label(main, text="Carpeta de PDF:").grid(row=1, column=0, sticky="w", pady=4)
        self.carpeta_var = tk.StringVar(value=str(_app_dir() / "EJEMPLO PDF"))
        ttk.Entry(main, textvariable=self.carpeta_var, width=90).grid(row=2, column=0, sticky="we", padx=(0, 8))
        ttk.Button(main, text="Seleccionar", command=self._select_folder).grid(row=2, column=1, sticky="we")

        ttk.Label(main, text="Archivo de salida (.xlsx):").grid(row=3, column=0, sticky="w", pady=(10, 4))
        self.salida_var = tk.StringVar(value=str(_app_dir() / "INFORME_MASIVO.xlsx"))
        ttk.Entry(main, textvariable=self.salida_var, width=90).grid(row=4, column=0, sticky="we", padx=(0, 8))
        ttk.Button(main, text="Elegir", command=self._select_output).grid(row=4, column=1, sticky="we")

        # Selector de modo
        modo_frame = ttk.Frame(main)
        modo_frame.grid(row=5, column=0, columnspan=2, sticky="we", pady=(8, 0))
        ttk.Label(modo_frame, text="Modo:").pack(side=tk.LEFT, padx=(0, 6))
        self.modo_var = tk.StringVar(value="Clasico (renombrar + informe)")
        modo_cb = ttk.Combobox(
            modo_frame,
            textvariable=self.modo_var,
            state="readonly",
            width=45,
            values=[
                "Clasico (renombrar + informe)",
                "Encarpetar desde Excel",
                "Informe HC NUEVOS (Numero Ide, Paciente, Nota, Diagnostico...)",
            ],
        )
        modo_cb.pack(side=tk.LEFT)

        options = ttk.Frame(main)
        options.grid(row=6, column=0, columnspan=2, sticky="we", pady=10)
        options.columnconfigure(7, weight=1)

        self.recursivo_var = tk.BooleanVar(value=False)
        self.renombrar_var = tk.BooleanVar(value=True)

        ttk.Checkbutton(options, text="Escaneo recursivo", variable=self.recursivo_var).grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Checkbutton(options, text="Renombrar PDF", variable=self.renombrar_var).grid(row=0, column=1, sticky="w", padx=(0, 12))

        ttk.Label(options, text="Workers:").grid(row=0, column=2, sticky="e")
        self.workers_var = tk.StringVar(value="4")
        ttk.Entry(options, textvariable=self.workers_var, width=6).grid(row=0, column=3, sticky="w", padx=(4, 12))

        ttk.Label(options, text="Progreso cada N:").grid(row=0, column=4, sticky="e")
        self.progreso_cada_var = tk.StringVar(value="100")
        ttk.Entry(options, textvariable=self.progreso_cada_var, width=8).grid(row=0, column=5, sticky="w", padx=(4, 12))

        ttk.Label(options, text="Max filas Excel:").grid(row=0, column=6, sticky="e")
        self.max_filas_var = tk.StringVar(value="1000000")
        ttk.Entry(options, textvariable=self.max_filas_var, width=10).grid(row=0, column=7, sticky="w", padx=(4, 0))

        controls = ttk.Frame(main)
        controls.grid(row=7, column=0, columnspan=2, sticky="we", pady=(2, 8))
        self.btn_iniciar = ttk.Button(controls, text="Iniciar", command=self._start)
        self.btn_iniciar.pack(side=tk.LEFT)
        self.btn_detener = ttk.Button(controls, text="Detener", command=self._stop, state=tk.DISABLED)
        self.btn_detener.pack(side=tk.LEFT, padx=(8, 0))

        self.progress = ttk.Progressbar(main, orient=tk.HORIZONTAL, mode="determinate")
        self.progress.grid(row=8, column=0, columnspan=2, sticky="we", pady=6)

        stats = ttk.Frame(main)
        stats.grid(row=9, column=0, columnspan=2, sticky="we")
        for i in range(5):
            stats.columnconfigure(i, weight=1)

        self.estado_var = tk.StringVar(value="Estado: Listo")
        self.total_var = tk.StringVar(value="Total PDF: 0")
        self.procesados_var = tk.StringVar(value="Procesados: 0")
        self.faltan_var = tk.StringVar(value="Faltan: 0")
        self.tiempo_var = tk.StringVar(value="Tiempo: 00:00:00")

        ttk.Label(stats, textvariable=self.estado_var).grid(row=0, column=0, sticky="w")
        ttk.Label(stats, textvariable=self.total_var).grid(row=0, column=1, sticky="w")
        ttk.Label(stats, textvariable=self.procesados_var).grid(row=0, column=2, sticky="w")
        ttk.Label(stats, textvariable=self.faltan_var).grid(row=0, column=3, sticky="w")
        ttk.Label(stats, textvariable=self.tiempo_var).grid(row=0, column=4, sticky="w")

        self.log = tk.Text(main, height=20, wrap="word")
        self.log.grid(row=10, column=0, columnspan=2, sticky="nsew", pady=(8, 0))

        scrollbar = ttk.Scrollbar(main, orient="vertical", command=self.log.yview)
        scrollbar.grid(row=10, column=2, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)

        main.columnconfigure(0, weight=1)
        main.rowconfigure(10, weight=1)

    # ----- Actualizaciones -----
    def _on_check_updates(self) -> None:
        self._append_log("Buscando actualizaciones en GitHub...")

        def _worker() -> None:
            info = check_for_updates()
            self.after(0, lambda: self._handle_update_info(info))

        threading.Thread(target=_worker, daemon=True).start()

    def _handle_update_info(self, info) -> None:
        if info.error:
            messagebox.showwarning("Actualizaciones", info.error)
            self._append_log(f"Actualizaciones: {info.error}")
            return

        if not info.has_update:
            messagebox.showinfo(
                "Actualizaciones",
                f"Estas usando la ultima version ({info.current_version}).",
            )
            self._append_log(f"Sin actualizaciones (actual: {info.current_version}).")
            return

        msg = (
            f"Hay una nueva version disponible.\n\n"
            f"Actual: {info.current_version}\n"
            f"Nueva:  {info.latest_version}\n\n"
            f"{(info.release_notes or '').strip()[:600]}\n\n"
            f"Deseas descargar e instalar ahora?"
        )
        if not messagebox.askyesno("Nueva actualizacion", msg):
            return

        if not getattr(sys, "frozen", False):
            messagebox.showinfo(
                "Actualizaciones",
                "La instalacion automatica solo funciona en la version .exe.\n"
                f"Descarga manualmente desde:\n{info.release_url}",
            )
            return

        self._append_log("Descargando actualizacion...")

        def _progress(d: int, t: int) -> None:
            def _ui() -> None:
                txt = f"Descargando: {d/1024/1024:.1f} MB"
                if t:
                    txt += f" / {t/1024/1024:.1f} MB"
                self.estado_var.set(txt)
            self.after(0, _ui)

        def _done(ok: bool, m: str) -> None:
            def _ui() -> None:
                if ok:
                    messagebox.showinfo("Actualizaciones", m)
                else:
                    messagebox.showerror("Actualizaciones", m)
            self.after(0, _ui)

        apply_update_in_background(info, on_progress=_progress, on_done=_done)

    # ----- Selectores -----
    def _select_folder(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.carpeta_var.set(folder)

    def _select_output(self) -> None:
        output = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="INFORME_MASIVO.xlsx",
        )
        if output:
            self.salida_var.set(output)

    def _append_log(self, text: str) -> None:
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)

    def _start(self) -> None:
        carpeta = self.carpeta_var.get().strip().strip('"')
        salida = self.salida_var.get().strip().strip('"')

        if not carpeta:
            messagebox.showerror("Error", "Debes seleccionar una carpeta o archivo de entrada")
            return

        if not Path(carpeta).exists():
            messagebox.showerror("Error", "La ruta indicada no existe")
            return

        self.total_detectado = 0
        self.procesados = 0
        self.progress["value"] = 0
        self.progress["maximum"] = 100
        self.estado_var.set("Estado: Ejecutando")
        self.total_var.set("Total PDF: 0")
        self.procesados_var.set("Procesados: 0")
        self.faltan_var.set("Faltan: 0")
        self.log.delete("1.0", tk.END)

        cmd_args = [
            "--workers", self.workers_var.get().strip() or "4",
            "--progreso-cada", self.progreso_cada_var.get().strip() or "100",
            "--max-filas-excel", self.max_filas_var.get().strip() or "1000000",
        ]

        modo = self.modo_var.get()
        if modo.startswith("Encarpetar"):
            # Modo encarpetar desde Excel: la "carpeta" se reinterpreta como Excel de entrada
            cmd_args = [
                "--accion", "encarpetar-desde-excel",
                "--excel-entrada", carpeta,
                "--excel-salida-encarpetado", salida,
            ]
        elif modo.startswith("Informe HC NUEVOS"):
            cmd_args = [
                "--accion", "informe-nuevos",
                "--entrada", carpeta,
                "--salida", salida,
                *cmd_args,
            ]
            if self.recursivo_var.get():
                cmd_args.append("--recursivo")
        else:
            cmd_args = [
                "--entrada", carpeta,
                "--salida", salida,
                *cmd_args,
            ]
            if self.recursivo_var.get():
                cmd_args.append("--recursivo")
            if not self.renombrar_var.get():
                cmd_args.append("--sin-renombrar")

        cmd = _run_processor_args(cmd_args)

        self.start_time = time.time()
        self.btn_iniciar.configure(state=tk.DISABLED)
        self.btn_detener.configure(state=tk.NORMAL)

        thread = threading.Thread(target=self._run_process, args=(cmd,), daemon=True)
        thread.start()

    def _run_process(self, cmd: list) -> None:
        try:
            creationflags = 0
            if os.name == "nt":
                creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

            self.process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=creationflags,
            )

            assert self.process.stdout is not None
            for line in self.process.stdout:
                self.output_queue.put(("line", line.rstrip()))

            code = self.process.wait()
            self.output_queue.put(("done", code))
        except Exception as exc:
            self.output_queue.put(("line", f"ERROR: {exc}"))
            self.output_queue.put(("done", 1))

    def _stop(self) -> None:
        if self.process and self.process.poll() is None:
            self.process.terminate()
            self._append_log("Proceso detenido por usuario")

    def _update_time(self) -> None:
        if self.start_time:
            elapsed = int(time.time() - self.start_time)
            h = elapsed // 3600
            m = (elapsed % 3600) // 60
            s = elapsed % 60
            self.tiempo_var.set(f"Tiempo: {h:02d}:{m:02d}:{s:02d}")

    def _handle_line(self, line: str) -> None:
        self._append_log(line)

        m_total = re.search(r"Total detectado:\s*(\d+)", line)
        if m_total:
            self.total_detectado = int(m_total.group(1))
            self.total_var.set(f"Total PDF: {self.total_detectado}")
            self.progress["maximum"] = max(self.total_detectado, 1)
            return

        m_prog = re.search(r"Progreso:\s*(\d+)\s+de\s+(\d+)\s+\|\s+Faltan:\s*(\d+)", line)
        if m_prog:
            self.procesados = int(m_prog.group(1))
            total = int(m_prog.group(2))
            faltan = int(m_prog.group(3))
            self.total_detectado = max(self.total_detectado, total)
            self.procesados_var.set(f"Procesados: {self.procesados}")
            self.faltan_var.set(f"Faltan: {faltan}")
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = min(self.procesados, total)
            return

    def _handle_done(self, code: int) -> None:
        self.btn_iniciar.configure(state=tk.NORMAL)
        self.btn_detener.configure(state=tk.DISABLED)

        if self.total_detectado > 0 and self.procesados <= self.total_detectado:
            faltan = max(self.total_detectado - self.procesados, 0)
            self.faltan_var.set(f"Faltan: {faltan}")

        if code == 0:
            self.estado_var.set("Estado: Finalizado")
            messagebox.showinfo("Completado", "Proceso finalizado correctamente")
        else:
            self.estado_var.set("Estado: Error")
            messagebox.showerror("Error", "El proceso termino con error. Revisa el log.")

        self.process = None

    def _poll_queue(self) -> None:
        self._update_time()
        while True:
            try:
                kind, payload = self.output_queue.get_nowait()
            except queue.Empty:
                break

            if kind == "line":
                self._handle_line(payload)
            elif kind == "done":
                self._handle_done(payload)

        self.after(150, self._poll_queue)


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
