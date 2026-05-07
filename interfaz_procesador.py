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
    # Paleta de colores
    COLOR_BG = "#f4f6fa"
    COLOR_HEADER = "#1e3a5f"
    COLOR_HEADER_FG = "#ffffff"
    COLOR_ACCENT = "#2a7de1"
    COLOR_ACCENT_HOVER = "#1f63b8"
    COLOR_SUCCESS = "#2ecc71"
    COLOR_DANGER = "#e74c3c"
    COLOR_WARN = "#f39c12"
    COLOR_CARD = "#ffffff"
    COLOR_BORDER = "#d6dbe4"
    COLOR_TEXT = "#1f2a3c"
    COLOR_MUTED = "#6b7384"
    COLOR_LOG_BG = "#1e1e1e"
    COLOR_LOG_FG = "#dcdcdc"

    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("1040x740")
        self.minsize(900, 640)
        self.configure(bg=self.COLOR_BG)

        self.output_queue: queue.Queue = queue.Queue()
        self.process = None
        self.start_time = 0.0
        self.total_detectado = 0
        self.procesados = 0

        self._setup_styles()
        self._build_ui()
        self.after(150, self._poll_queue)

    def _setup_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=("Segoe UI", 10), background=self.COLOR_BG, foreground=self.COLOR_TEXT)
        style.configure("TFrame", background=self.COLOR_BG)
        style.configure("Card.TFrame", background=self.COLOR_CARD, relief="flat")
        style.configure("Header.TFrame", background=self.COLOR_HEADER)
        style.configure("Stat.TFrame", background=self.COLOR_CARD, relief="flat", borderwidth=1)

        style.configure("TLabel", background=self.COLOR_BG, foreground=self.COLOR_TEXT)
        style.configure("Card.TLabel", background=self.COLOR_CARD, foreground=self.COLOR_TEXT)
        style.configure("Header.TLabel", background=self.COLOR_HEADER, foreground=self.COLOR_HEADER_FG)
        style.configure("HeaderTitle.TLabel", background=self.COLOR_HEADER, foreground=self.COLOR_HEADER_FG, font=("Segoe UI Semibold", 16))
        style.configure("HeaderSub.TLabel", background=self.COLOR_HEADER, foreground="#bcd2ee", font=("Segoe UI", 9))
        style.configure("Section.TLabel", background=self.COLOR_BG, foreground=self.COLOR_MUTED, font=("Segoe UI Semibold", 9))
        style.configure("StatTitle.TLabel", background=self.COLOR_CARD, foreground=self.COLOR_MUTED, font=("Segoe UI", 9))
        style.configure("StatValue.TLabel", background=self.COLOR_CARD, foreground=self.COLOR_TEXT, font=("Segoe UI Semibold", 14))
        style.configure("FieldLabel.TLabel", background=self.COLOR_BG, foreground=self.COLOR_TEXT, font=("Segoe UI Semibold", 9))

        style.configure(
            "TEntry",
            fieldbackground=self.COLOR_CARD,
            foreground=self.COLOR_TEXT,
            bordercolor=self.COLOR_BORDER,
            lightcolor=self.COLOR_BORDER,
            darkcolor=self.COLOR_BORDER,
            padding=6,
        )
        style.configure(
            "TCombobox",
            fieldbackground=self.COLOR_CARD,
            background=self.COLOR_CARD,
            foreground=self.COLOR_TEXT,
            padding=4,
        )

        style.configure("TCheckbutton", background=self.COLOR_BG, foreground=self.COLOR_TEXT)

        # Botones
        style.configure("TButton", padding=(12, 6), background="#e9ecf2", foreground=self.COLOR_TEXT, borderwidth=0)
        style.map("TButton", background=[("active", "#dde2eb")])

        style.configure("Primary.TButton", padding=(16, 8), background=self.COLOR_ACCENT, foreground="#ffffff", borderwidth=0, font=("Segoe UI Semibold", 10))
        style.map("Primary.TButton",
                  background=[("active", self.COLOR_ACCENT_HOVER), ("disabled", "#a8bddc")],
                  foreground=[("disabled", "#eef3fb")])

        style.configure("Danger.TButton", padding=(16, 8), background=self.COLOR_DANGER, foreground="#ffffff", borderwidth=0, font=("Segoe UI Semibold", 10))
        style.map("Danger.TButton",
                  background=[("active", "#c0392b"), ("disabled", "#e9b0a9")],
                  foreground=[("disabled", "#fbe9e7")])

        style.configure("Ghost.TButton", padding=(10, 6), background=self.COLOR_HEADER, foreground=self.COLOR_HEADER_FG, borderwidth=1)
        style.map("Ghost.TButton", background=[("active", "#274d7c")])

        style.configure("Horizontal.TProgressbar",
                        troughcolor="#e2e6ee",
                        background=self.COLOR_ACCENT,
                        bordercolor="#e2e6ee",
                        lightcolor=self.COLOR_ACCENT,
                        darkcolor=self.COLOR_ACCENT,
                        thickness=14)

    def _card(self, parent: tk.Widget):
        outer = tk.Frame(parent, bg=self.COLOR_BORDER)
        inner = ttk.Frame(outer, style="Card.TFrame", padding=14)
        inner.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        outer.inner = inner  # type: ignore[attr-defined]
        return outer

    def _build_ui(self) -> None:
        # ===== Header =====
        header = ttk.Frame(self, style="Header.TFrame", padding=(20, 14))
        header.pack(fill=tk.X)
        header.columnconfigure(0, weight=1)

        title_box = ttk.Frame(header, style="Header.TFrame")
        title_box.grid(row=0, column=0, sticky="w")
        ttk.Label(title_box, text=f"📄  {APP_NAME}", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(title_box, text=f"Renombre y Reporte de PDF  •  v{APP_VERSION}", style="HeaderSub.TLabel").pack(anchor="w")

        ttk.Button(header, text="🔄  Buscar actualizaciones", style="Ghost.TButton",
                   command=self._on_check_updates).grid(row=0, column=1, sticky="e")

        # ===== Body =====
        body = ttk.Frame(self, padding=20)
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)

        # --- Card de configuracion ---
        cfg_card = self._card(body)
        cfg_card.grid(row=0, column=0, sticky="we")
        cfg = cfg_card.inner  # type: ignore[attr-defined]
        cfg.columnconfigure(1, weight=1)

        ttk.Label(cfg, text="CONFIGURACION", style="Section.TLabel", background=self.COLOR_CARD).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        # Modo
        ttk.Label(cfg, text="Modo de procesamiento", style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        self.modo_var = tk.StringVar(value="Clasico (renombrar + informe)")
        self.modo_cb = ttk.Combobox(
            cfg,
            textvariable=self.modo_var,
            state="readonly",
            values=[
                "Clasico (renombrar + informe)",
                "Encarpetar desde Excel",
                "Informe HC NUEVOS (Numero Ide, Paciente, Nota, Diagnostico)",
            ],
        )
        self.modo_cb.grid(row=1, column=1, columnspan=2, sticky="we", pady=4)
        self.modo_cb.bind("<<ComboboxSelected>>", lambda e: self._update_mode_hint())

        # Entrada
        ttk.Label(cfg, text="Carpeta / Excel de entrada", style="Card.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(10, 4))
        self.carpeta_var = tk.StringVar(value=str(_app_dir() / "EJEMPLO PDF"))
        ttk.Entry(cfg, textvariable=self.carpeta_var).grid(row=2, column=1, sticky="we", pady=(10, 4))
        ttk.Button(cfg, text="📁  Seleccionar", command=self._select_input).grid(row=2, column=2, sticky="we", padx=(8, 0), pady=(10, 4))

        # Salida
        ttk.Label(cfg, text="Archivo de salida (.xlsx)", style="Card.TLabel").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=4)
        self.salida_var = tk.StringVar(value=str(_app_dir() / "INFORME_MASIVO.xlsx"))
        ttk.Entry(cfg, textvariable=self.salida_var).grid(row=3, column=1, sticky="we", pady=4)
        ttk.Button(cfg, text="💾  Elegir", command=self._select_output).grid(row=3, column=2, sticky="we", padx=(8, 0), pady=4)

        # Hint dinamico
        self.hint_var = tk.StringVar(value="")
        ttk.Label(cfg, textvariable=self.hint_var, style="Card.TLabel", foreground=self.COLOR_MUTED, wraplength=900, justify="left").grid(
            row=4, column=0, columnspan=3, sticky="we", pady=(6, 4))

        # Opciones avanzadas
        sep = ttk.Separator(cfg, orient="horizontal")
        sep.grid(row=5, column=0, columnspan=3, sticky="we", pady=(10, 10))

        opts = ttk.Frame(cfg, style="Card.TFrame")
        opts.grid(row=6, column=0, columnspan=3, sticky="we")
        for i in range(8):
            opts.columnconfigure(i, weight=0)

        self.recursivo_var = tk.BooleanVar(value=False)
        self.renombrar_var = tk.BooleanVar(value=True)

        chk_style = "TCheckbutton"
        ttk.Checkbutton(opts, text="Escaneo recursivo", variable=self.recursivo_var, style=chk_style).grid(row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Checkbutton(opts, text="Renombrar PDF", variable=self.renombrar_var, style=chk_style).grid(row=0, column=1, sticky="w", padx=(0, 16))

        ttk.Label(opts, text="Workers", style="Card.TLabel").grid(row=0, column=2, padx=(0, 4))
        self.workers_var = tk.StringVar(value="4")
        ttk.Entry(opts, textvariable=self.workers_var, width=5).grid(row=0, column=3, padx=(0, 16))

        ttk.Label(opts, text="Progreso cada N", style="Card.TLabel").grid(row=0, column=4, padx=(0, 4))
        self.progreso_cada_var = tk.StringVar(value="100")
        ttk.Entry(opts, textvariable=self.progreso_cada_var, width=7).grid(row=0, column=5, padx=(0, 16))

        ttk.Label(opts, text="Max filas Excel", style="Card.TLabel").grid(row=0, column=6, padx=(0, 4))
        self.max_filas_var = tk.StringVar(value="1000000")
        ttk.Entry(opts, textvariable=self.max_filas_var, width=10).grid(row=0, column=7)

        # --- Controles + Stats ---
        ctrl_row = ttk.Frame(body)
        ctrl_row.grid(row=1, column=0, sticky="we", pady=(14, 8))
        ctrl_row.columnconfigure(1, weight=1)

        btns = ttk.Frame(ctrl_row)
        btns.grid(row=0, column=0, sticky="w")
        self.btn_iniciar = ttk.Button(btns, text="▶  Iniciar", style="Primary.TButton", command=self._start)
        self.btn_iniciar.pack(side=tk.LEFT)
        self.btn_detener = ttk.Button(btns, text="■  Detener", style="Danger.TButton", command=self._stop, state=tk.DISABLED)
        self.btn_detener.pack(side=tk.LEFT, padx=(8, 0))

        # Stats cards
        stats = ttk.Frame(ctrl_row)
        stats.grid(row=0, column=1, sticky="e")
        self.estado_var = tk.StringVar(value="Listo")
        self.total_var = tk.StringVar(value="0")
        self.procesados_var = tk.StringVar(value="0")
        self.faltan_var = tk.StringVar(value="0")
        self.tiempo_var = tk.StringVar(value="00:00:00")

        for col, (titulo, var) in enumerate([
            ("Estado", self.estado_var),
            ("Total", self.total_var),
            ("Procesados", self.procesados_var),
            ("Faltan", self.faltan_var),
            ("Tiempo", self.tiempo_var),
        ]):
            card = self._card(stats)
            card.grid(row=0, column=col, padx=(8 if col else 0, 0))
            inner = card.inner  # type: ignore[attr-defined]
            inner.configure(padding=(14, 8))
            ttk.Label(inner, text=titulo.upper(), style="StatTitle.TLabel").pack(anchor="w")
            ttk.Label(inner, textvariable=var, style="StatValue.TLabel").pack(anchor="w")

        # Progress bar
        self.progress = ttk.Progressbar(body, orient=tk.HORIZONTAL, mode="determinate", style="Horizontal.TProgressbar")
        self.progress.grid(row=2, column=0, sticky="we", pady=(4, 12))

        # --- Log ---
        log_card = self._card(body)
        log_card.grid(row=3, column=0, sticky="nsew")
        body.rowconfigure(3, weight=1)
        log_inner = log_card.inner  # type: ignore[attr-defined]
        log_inner.columnconfigure(0, weight=1)
        log_inner.rowconfigure(1, weight=1)
        ttk.Label(log_inner, text="REGISTRO DE ACTIVIDAD", style="Section.TLabel", background=self.COLOR_CARD).grid(row=0, column=0, sticky="w", pady=(0, 8))

        log_wrap = tk.Frame(log_inner, bg=self.COLOR_LOG_BG)
        log_wrap.grid(row=1, column=0, sticky="nsew")
        log_wrap.columnconfigure(0, weight=1)
        log_wrap.rowconfigure(0, weight=1)

        self.log = tk.Text(
            log_wrap,
            wrap="word",
            bg=self.COLOR_LOG_BG,
            fg=self.COLOR_LOG_FG,
            insertbackground=self.COLOR_LOG_FG,
            font=("Consolas", 10),
            relief="flat",
            padx=10,
            pady=8,
            borderwidth=0,
            highlightthickness=0,
        )
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_wrap, orient="vertical", command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)

        self.log.tag_configure("ok", foreground="#7ee08a")
        self.log.tag_configure("err", foreground="#ff8a80")
        self.log.tag_configure("info", foreground="#82b1ff")

        # ===== Footer =====
        footer = ttk.Frame(self, padding=(20, 8))
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Label(footer, text="© IPS Goleman  •  github.com/PabloGra77/Informe---scaner-HC",
                  foreground=self.COLOR_MUTED).pack(side=tk.LEFT)

        self._update_mode_hint()

    def _update_mode_hint(self) -> None:
        modo = self.modo_var.get()
        if modo.startswith("Encarpetar"):
            self.hint_var.set("ℹ  La 'entrada' debe ser un archivo Excel con columnas 'ruta_archivo' y 'numero de documento'.")
            self.renombrar_var.set(False)
        elif modo.startswith("Informe HC NUEVOS"):
            self.hint_var.set("ℹ  Modo para el formato nuevo de HC. Extrae paciente, ingreso, tipo servicio, diagnostico y nota.")
            self.renombrar_var.set(False)
        else:
            self.hint_var.set("ℹ  Modo clasico. Renombra los PDF por numero de documento y genera informe Excel.")
            self.renombrar_var.set(True)

    # ----- Actualizaciones -----
    def _on_check_updates(self) -> None:
        self._append_log("Buscando actualizaciones en GitHub...", "info")

        def _worker() -> None:
            info = check_for_updates()
            self.after(0, lambda: self._handle_update_info(info))

        threading.Thread(target=_worker, daemon=True).start()

    def _handle_update_info(self, info) -> None:
        if info.error:
            messagebox.showwarning("Actualizaciones", info.error)
            self._append_log(f"Actualizaciones: {info.error}", "err")
            return

        if not info.has_update:
            messagebox.showinfo(
                "Actualizaciones",
                f"Estas usando la ultima version ({info.current_version}).",
            )
            self._append_log(f"Sin actualizaciones (actual: {info.current_version}).", "ok")
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

        self._append_log("Descargando actualizacion...", "info")

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
    def _select_input(self) -> None:
        """Selector inteligente: archivo Excel si modo Encarpetar, carpeta en otros casos."""
        modo = self.modo_var.get()
        if modo.startswith("Encarpetar"):
            path = filedialog.askopenfilename(
                title="Selecciona el Excel de entrada",
                filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
            )
        else:
            path = filedialog.askdirectory(title="Selecciona la carpeta con PDF")
        if path:
            self.carpeta_var.set(path)

    def _select_output(self) -> None:
        output = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="INFORME_MASIVO.xlsx",
        )
        if output:
            self.salida_var.set(output)

    def _append_log(self, text: str, tag: str = "") -> None:
        if tag:
            self.log.insert(tk.END, text + "\n", tag)
        else:
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
        self.estado_var.set("Ejecutando")
        self.total_var.set("0")
        self.procesados_var.set("0")
        self.faltan_var.set("0")
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
            self._append_log("Proceso detenido por usuario", "err")

    def _update_time(self) -> None:
        if self.start_time:
            elapsed = int(time.time() - self.start_time)
            h = elapsed // 3600
            m = (elapsed % 3600) // 60
            s = elapsed % 60
            self.tiempo_var.set(f"{h:02d}:{m:02d}:{s:02d}")

    def _handle_line(self, line: str) -> None:
        tag = ""
        low = line.lower()
        if "error" in low or "err:" in low:
            tag = "err"
        elif "finalizado" in low or "informe generado" in low or line.startswith("OK"):
            tag = "ok"
        elif "progreso" in low or "total detectado" in low:
            tag = "info"
        self._append_log(line, tag)

        m_total = re.search(r"Total detectado:\s*(\d+)", line)
        if m_total:
            self.total_detectado = int(m_total.group(1))
            self.total_var.set(str(self.total_detectado))
            self.progress["maximum"] = max(self.total_detectado, 1)
            return

        m_prog = re.search(r"Progreso:\s*(\d+)\s+de\s+(\d+)\s+\|\s+Faltan:\s*(\d+)", line)
        if m_prog:
            self.procesados = int(m_prog.group(1))
            total = int(m_prog.group(2))
            faltan = int(m_prog.group(3))
            self.total_detectado = max(self.total_detectado, total)
            self.procesados_var.set(str(self.procesados))
            self.faltan_var.set(str(faltan))
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = min(self.procesados, total)
            return

    def _handle_done(self, code: int) -> None:
        self.btn_iniciar.configure(state=tk.NORMAL)
        self.btn_detener.configure(state=tk.DISABLED)

        if self.total_detectado > 0 and self.procesados <= self.total_detectado:
            faltan = max(self.total_detectado - self.procesados, 0)
            self.faltan_var.set(str(faltan))

        if code == 0:
            self.estado_var.set("Finalizado")
            messagebox.showinfo("Completado", "Proceso finalizado correctamente")
        else:
            self.estado_var.set("Error")
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
