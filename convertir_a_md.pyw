"""
convertir_a_md.pyw — App de escritorio para convertir PDF y archivos de
Office (.docx, .xlsx, .pptx y otros) a Markdown (.md).

Mejoras:
  1) Arrastrar y soltar archivos sobre la ventana (si tkinterdnd2 está
     instalado; si no, los botones siguen funcionando).
  2) Acepta archivos como argumentos al ejecutarse (para el menú
     contextual de Windows "Convertir a Markdown").
  3) Barra de progreso real con contador "N de M".
  4) Botón "Abrir carpeta de salida" tras convertir.
  6) Nunca sobrescribe: si ya existe, añade " (1)", " (2)", …
  7) Vista previa del .md generado al hacer clic en un archivo de la lista.
  8) Post-procesado opcional para limpiar PDFs (une líneas partidas, quita
     números de página sueltos, colapsa saltos en exceso).
  9) Detecta PDFs escaneados sin texto extraído (sugiere OCR).
 10) Front-matter YAML con metadatos (título, origen, fecha, páginas).

Requisitos:
    pip install "markitdown[all]"
Opcionales (solo si los quieres aprovechar):
    pip install tkinterdnd2     # para arrastrar y soltar
    pip install pypdf           # para contar páginas en el front-matter
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime
from pathlib import Path

__version__ = "1.2.0"

# URL a un JSON con el formato:
#   { "version": "1.3.0",
#     "url":     "https://...página de descarga o release",
#     "notes":   "Texto corto con novedades (opcional)" }
# Déjalo en "" para desactivar la comprobación de actualizaciones.
UPDATE_URL = "https://raw.githubusercontent.com/Jamper-sys/pdf-office-to-md/main/version.json"

# --- Auto-sync del script de accesos directos al repo git (opcional) --------
# Al abrir la app, si crear_accesos_directos.ps1 (que vive junto a esta app)
# difiere de la copia versionada en el repo, se copia y se hace commit/push.
# Es silencioso, corre en segundo plano y NUNCA interrumpe ni rompe la app.
# Solo actúa en el PC donde existe el repo; en cualquier otro, no hace nada.
# Pon AUTO_SYNC_SHORTCUT_TO_REPO = False para desactivarlo.
AUTO_SYNC_SHORTCUT_TO_REPO = True
REPO_DIR = Path(r"C:\Users\javie\Repos\pdf-office-to-md")
SHORTCUT_SCRIPT_NAME = "crear_accesos_directos.ps1"
from tkinter import (
    BooleanVar, StringVar, IntVar, END, Listbox, SINGLE,
    filedialog, messagebox, ttk, scrolledtext,
)

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
    BaseTk = TkinterDnD.Tk
except ImportError:
    from tkinter import Tk as BaseTk
    DND_AVAILABLE = False
    DND_FILES = None


# ----------------------------------------------------------------- constantes

SUPPORTED_EXTS = {
    ".pdf",
    ".docx", ".doc",
    ".xlsx", ".xls", ".csv",
    ".pptx", ".ppt",
    ".html", ".htm",
    ".txt", ".rtf",
    ".epub",
    ".png", ".jpg", ".jpeg",
    ".mp3", ".wav", ".m4a",
    ".json", ".xml",
    ".zip",
}

FILE_TYPES = [
    ("Todos los soportados",
     "*.pdf *.docx *.doc *.xlsx *.xls *.csv *.pptx *.ppt "
     "*.html *.htm *.txt *.rtf *.epub *.png *.jpg *.jpeg "
     "*.mp3 *.wav *.m4a *.json *.xml *.zip"),
    ("PDF", "*.pdf"),
    ("Word", "*.docx *.doc"),
    ("Excel", "*.xlsx *.xls *.csv"),
    ("PowerPoint", "*.pptx *.ppt"),
    ("Todos los archivos", "*.*"),
]

CONFIG_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "convertir_a_md"
CONFIG_FILE = CONFIG_DIR / "config.json"


# -------------------------------------------------------------------- utilidades

def default_downloads_dir() -> Path:
    return Path.home() / "Downloads"


def load_config() -> dict:
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(data: dict) -> None:
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False),
                               encoding="utf-8")
    except Exception:
        pass


def unique_path(p: Path) -> Path:
    """Devuelve una ruta que no choque con un archivo existente."""
    if not p.exists():
        return p
    stem, suf, parent = p.stem, p.suffix, p.parent
    i = 1
    while True:
        cand = parent / f"{stem} ({i}){suf}"
        if not cand.exists():
            return cand
        i += 1


def get_pdf_page_count(path: Path) -> int | None:
    try:
        from pypdf import PdfReader
        return len(PdfReader(str(path)).pages)
    except Exception:
        return None


def build_frontmatter(src: Path, extra: dict | None = None) -> str:
    lines = ["---"]
    title = src.stem.replace('"', "'")
    lines.append(f'title: "{title}"')
    lines.append(f'source: "{src.name}"')
    lines.append(f'source_path: "{str(src).replace(chr(92), "/")}"')
    lines.append(f'converted_at: "{datetime.now().isoformat(timespec="seconds")}"')
    lines.append('converter: "MarkItDown"')
    if extra:
        for k, v in extra.items():
            if isinstance(v, str):
                lines.append(f'{k}: "{v}"')
            else:
                lines.append(f'{k}: {v}')
    lines.append("---")
    lines.append("")
    return "\n".join(lines)


def postprocess_markdown(text: str) -> str:
    """Limpia ruido típico de PDFs convertidos."""
    if not text:
        return text

    # 1) Quitar líneas que son solo número de página (con o sin guiones).
    text = re.sub(r'(?m)^\s*[-–—]?\s*\d{1,4}\s*[-–—]?\s*$\n?', '', text)

    # 2) Colapsar 3+ líneas vacías a 2.
    text = re.sub(r'\n{3,}', '\n\n', text)

    # 3) Unir líneas partidas a la mitad: línea acaba en letra/coma y
    #    la siguiente arranca en minúscula → era un único párrafo.
    lines = text.split("\n")
    out: list[str] = []
    i = 0
    skip_prefixes = ("#", "-", "*", ">", "|", "```", "    ")
    while i < len(lines):
        cur = lines[i]
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        if (cur and nxt
                and not cur.startswith(skip_prefixes)
                and not nxt.startswith(skip_prefixes)
                and re.search(r'[A-Za-záéíóúñü,]\s*$', cur)
                and re.match(r'^[a-záéíóúñü]', nxt)):
            out.append(cur.rstrip() + " " + nxt.lstrip())
            i += 2
        else:
            out.append(cur)
            i += 1

    return "\n".join(out).strip() + "\n"


def _version_tuple(v: str) -> tuple:
    try:
        return tuple(int(x) for x in re.findall(r'\d+', v))
    except Exception:
        return (0,)


def is_newer_version(remote: str, local: str) -> bool:
    return _version_tuple(remote) > _version_tuple(local)


def sync_shortcut_script_to_repo() -> None:
    """Sincroniza crear_accesos_directos.ps1 (junto a la app) hacia el repo git.

    Pensada para ejecutarse en segundo plano al abrir la app. Es totalmente
    silenciosa y a prueba de fallos: si el repo no existe, si no hay git, si
    no hay red o si algo sale mal, simplemente no hace nada. Solo crea un
    commit cuando git detecta un cambio real en ese archivo (sin commits de
    ruido).
    """
    if not AUTO_SYNC_SHORTCUT_TO_REPO:
        return
    try:
        app_dir = Path(__file__).resolve().parent
        local = app_dir / SHORTCUT_SCRIPT_NAME
        if not local.exists():
            return
        if not (REPO_DIR / ".git").exists():
            return  # no estamos en el PC del repo: no hacer nada

        repo_copy = REPO_DIR / "registry" / SHORTCUT_SCRIPT_NAME
        repo_copy.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(local, repo_copy)  # actualiza el working tree del repo

        rel = "registry/" + SHORTCUT_SCRIPT_NAME
        no_window = 0x08000000  # CREATE_NO_WINDOW: git no abre consola

        def git(*args):
            return subprocess.run(
                ["git", *args], cwd=str(REPO_DIR),
                capture_output=True, text=True, creationflags=no_window,
                timeout=30,
            )

        # Si git no ve cambios en ese archivo, no hay nada que hacer.
        status = git("status", "--porcelain", rel)
        if not status.stdout.strip():
            return

        git("commit", rel, "-m",
            "Auto-sync crear_accesos_directos.ps1 desde la app")
        git("push")
    except Exception:
        pass


def looks_like_scanned_pdf(src: Path, md_text: str) -> bool:
    """Heurística: PDF con muy poco texto → probablemente escaneado."""
    if src.suffix.lower() != ".pdf":
        return False
    # Limpia front-matter y espacios
    body = re.sub(r'(?s)^---.*?---\s*', '', md_text or '').strip()
    return len(body) < 100


# ----------------------------------------------------------------- aplicación

class ConvertirApp:
    def __init__(self, root, archivos_iniciales: list[Path] | None = None) -> None:
        self.root = root
        root.title("Convertir a Markdown")
        root.geometry("820x540")
        root.minsize(680, 420)

        cfg = load_config()

        self.archivos: list[Path] = []
        self.resultados: dict[str, Path] = {}  # src -> ruta .md generado
        self.cancelar = threading.Event()

        self.salida = StringVar(value=cfg.get("salida", str(default_downloads_dir())))
        self.recursivo = BooleanVar(value=cfg.get("recursivo", True))
        self.junto_al_original = BooleanVar(value=cfg.get("junto_al_original", False))
        self.postprocesar = BooleanVar(value=cfg.get("postprocesar", True))
        self.frontmatter = BooleanVar(value=cfg.get("frontmatter", True))
        self.aviso_ocr = BooleanVar(value=cfg.get("aviso_ocr", True))

        self.conversion_en_curso = False

        self._construir_menu()
        self._construir_ui()
        self._toggle_salida()

        if DND_AVAILABLE:
            self._activar_dnd()

        # Tareas de fondo (red + git): se lanzan DESPUÉS de mostrar la ventana
        # para que el arranque sea inmediato y no compitan con el primer pintado.
        self.root.after(1500, self._lanzar_tareas_fondo)

        # Asegura que la ventana aparezca al frente (no detrás del Explorador
        # cuando se abre desde el menú contextual).
        self.root.after(0, self._traer_al_frente)

        if archivos_iniciales:
            for p in archivos_iniciales:
                if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
                    self.archivos.append(p)
            self._refrescar_lista()

        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ----------------------------------------------------- arranque/ventana

    def _traer_al_frente(self) -> None:
        """Levanta la ventana al frente al abrir (incluido desde el menú
        contextual, donde si no podría quedar tras el Explorador)."""
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(500, lambda: self.root.attributes("-topmost", False))
            self.root.focus_force()
        except Exception:
            pass

    def _lanzar_tareas_fondo(self) -> None:
        """Tareas que no deben retrasar el arranque: chequeo de
        actualizaciones (red) y auto-sync del script de accesos (git).
        Se ejecutan en hilos demonio, ya con la ventana visible."""
        if UPDATE_URL:
            threading.Thread(target=self._check_actualizaciones,
                             args=(True,), daemon=True).start()
        threading.Thread(target=sync_shortcut_script_to_repo,
                         daemon=True).start()

    # --------------------------------------------------------------- menú

    def _construir_menu(self) -> None:
        from tkinter import Menu
        barra = Menu(self.root)
        ayuda = Menu(barra, tearoff=0)
        ayuda.add_command(label="Buscar actualizaciones…",
                          command=lambda: threading.Thread(
                              target=self._check_actualizaciones,
                              args=(False,), daemon=True).start())
        ayuda.add_separator()
        ayuda.add_command(label="Acerca de…", command=self._mostrar_acerca_de)
        barra.add_cascade(label="Ayuda", menu=ayuda)
        self.root.config(menu=barra)

    def _mostrar_acerca_de(self) -> None:
        messagebox.showinfo(
            "Acerca de",
            f"Convertir a Markdown\nVersión {__version__}\n\n"
            "PDF y archivos de Office a Markdown vía MarkItDown.",
        )

    def _check_actualizaciones(self, silencioso: bool) -> None:
        if not UPDATE_URL:
            if not silencioso:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Actualizaciones",
                    "No hay URL de actualización configurada en el script "
                    "(constante UPDATE_URL).",
                ))
            return
        try:
            import urllib.request
            with urllib.request.urlopen(UPDATE_URL, timeout=6) as r:
                data = json.loads(r.read().decode("utf-8"))
        except Exception as e:
            if not silencioso:
                self.root.after(0, lambda: messagebox.showerror(
                    "Actualizaciones",
                    f"No se pudo comprobar la actualización:\n{e}",
                ))
            return

        remota = str(data.get("version", "")).strip()
        url_descarga = data.get("url", "")
        notas = data.get("notes", "")

        if remota and is_newer_version(remota, __version__):
            def _notify():
                if messagebox.askyesno(
                    "Nueva versión disponible",
                    f"Versión instalada: {__version__}\n"
                    f"Versión disponible: {remota}\n\n"
                    + (f"Novedades:\n{notas}\n\n" if notas else "")
                    + "¿Abrir la página de descarga?",
                ):
                    if url_descarga:
                        webbrowser.open(url_descarga)
            self.root.after(0, _notify)
        elif not silencioso:
            self.root.after(0, lambda: messagebox.showinfo(
                "Actualizaciones",
                f"Estás en la última versión ({__version__}).",
            ))

    # ----------------------------------------------------------------- UI

    def _construir_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill="both", expand=True)

        # --- Encabezado / botones de carga -----------------------------
        cabecera = ttk.Frame(outer)
        cabecera.pack(fill="x")
        ttk.Label(
            cabecera,
            text="Arrastra archivos aquí, o úsalos botones para añadirlos."
                 if DND_AVAILABLE else
                 "Añade archivos con los botones (instala 'tkinterdnd2' para arrastrar y soltar).",
            font=("Segoe UI", 10, "bold"),
        ).pack(side="left")

        botones = ttk.Frame(outer)
        botones.pack(fill="x", pady=(8, 6))
        ttk.Button(botones, text="Elegir archivos…",
                   command=self.elegir_archivos).pack(side="left")
        ttk.Button(botones, text="Elegir carpeta…",
                   command=self.elegir_carpeta).pack(side="left", padx=6)
        ttk.Button(botones, text="Quitar selección",
                   command=self.quitar_seleccion).pack(side="left")
        ttk.Button(botones, text="Limpiar lista",
                   command=self.limpiar).pack(side="left", padx=6)

        # --- Opciones --------------------------------------------------
        opciones = ttk.LabelFrame(outer, text="Opciones", padding=8)
        opciones.pack(fill="x", pady=(4, 6))

        fila1 = ttk.Frame(opciones)
        fila1.pack(fill="x")
        ttk.Checkbutton(fila1, text="Incluir subcarpetas (recursivo)",
                        variable=self.recursivo).pack(side="left")
        ttk.Checkbutton(fila1, text="Guardar junto al archivo original",
                        variable=self.junto_al_original,
                        command=self._toggle_salida).pack(side="left", padx=12)

        fila2 = ttk.Frame(opciones)
        fila2.pack(fill="x", pady=(4, 0))
        ttk.Checkbutton(fila2, text="Limpiar Markdown de PDF (post-procesado)",
                        variable=self.postprocesar).pack(side="left")
        ttk.Checkbutton(fila2, text="Añadir front-matter YAML",
                        variable=self.frontmatter).pack(side="left", padx=12)
        ttk.Checkbutton(fila2, text="Avisar si un PDF parece escaneado (OCR)",
                        variable=self.aviso_ocr).pack(side="left")

        fila3 = ttk.Frame(opciones)
        fila3.pack(fill="x", pady=(6, 0))
        ttk.Label(fila3, text="Carpeta de salida:").pack(side="left")
        self.entry_salida = ttk.Entry(fila3, textvariable=self.salida)
        self.entry_salida.pack(side="left", fill="x", expand=True, padx=6)
        self.btn_salida = ttk.Button(fila3, text="…", width=3,
                                     command=self.elegir_salida)
        self.btn_salida.pack(side="left")

        # --- Fila de acciones (anclada abajo, SIEMPRE visible) --------
        accion = ttk.Frame(outer)
        accion.pack(side="bottom", fill="x", pady=(6, 0))

        progreso = ttk.Frame(outer)
        progreso.pack(side="bottom", fill="x")

        # --- Panel dividido: lista | vista previa ----------------------
        paned = ttk.PanedWindow(outer, orient="horizontal")
        paned.pack(fill="both", expand=True, pady=(4, 6))

        izq = ttk.LabelFrame(paned, text="Archivos en cola", padding=4)
        paned.add(izq, weight=1)

        self.listbox = Listbox(izq, selectmode=SINGLE, activestyle="dotbox")
        self.listbox.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(izq, orient="vertical",
                               command=self.listbox.yview)
        scroll.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scroll.set)
        self.listbox.bind("<<ListboxSelect>>", self._on_seleccion)

        der = ttk.LabelFrame(paned, text="Vista previa del .md", padding=4)
        paned.add(der, weight=2)
        self.preview = scrolledtext.ScrolledText(
            der, wrap="word", font=("Consolas", 10),
        )
        self.preview.pack(fill="both", expand=True)
        self.preview.insert(
            END,
            "(Selecciona un archivo de la lista para previsualizar su Markdown. "
            "Si todavía no lo has convertido, pulsa 'Convertir' primero.)",
        )
        self.preview.configure(state="disabled")

        # --- Progreso + acciones (widgets dentro de los frames ya creados arriba)
        self.progress = ttk.Progressbar(progreso, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True)
        self.lbl_progreso = ttk.Label(progreso, text="", width=12,
                                      anchor="e")
        self.lbl_progreso.pack(side="left", padx=(8, 0))

        self.estado = ttk.Label(accion, text="Listo.")
        self.estado.pack(side="left")

        self.btn_abrir_salida = ttk.Button(
            accion, text="Abrir carpeta de salida",
            command=self.abrir_salida, state="disabled",
        )
        self.btn_abrir_salida.pack(side="right", padx=(6, 0))

        self.btn_cancelar = ttk.Button(
            accion, text="Cancelar", command=self.cancelar_conversion,
            state="disabled",
        )
        self.btn_cancelar.pack(side="right", padx=(6, 0))

        self.btn_convertir = ttk.Button(
            accion, text="Convertir a Markdown", command=self.convertir,
        )
        self.btn_convertir.pack(side="right")

    def _activar_dnd(self) -> None:
        widgets = [self.root, self.listbox, self.preview]
        for w in widgets:
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    def _toggle_salida(self) -> None:
        state = "disabled" if self.junto_al_original.get() else "normal"
        self.entry_salida.configure(state=state)
        self.btn_salida.configure(state=state)

    # ----------------------------------------------------------- eventos

    def _on_drop(self, event) -> None:
        # event.data viene como '{ruta con espacios} otra_ruta {…}'
        raw = event.data
        rutas = self.root.tk.splitlist(raw)
        agregados = 0
        for r in rutas:
            p = Path(r)
            if p.is_dir():
                it = p.rglob("*") if self.recursivo.get() else p.iterdir()
                for child in it:
                    if child.is_file() and child.suffix.lower() in SUPPORTED_EXTS:
                        self.archivos.append(child)
                        agregados += 1
            elif p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
                self.archivos.append(p)
                agregados += 1
        if agregados:
            self._refrescar_lista()
            self.estado.config(text=f"Añadidos {agregados} archivo(s) por arrastrar.")

    def _on_seleccion(self, _event) -> None:
        sel = self.listbox.curselection()
        if not sel:
            return
        src = self.archivos[sel[0]]
        md_path = self.resultados.get(str(src))
        self.preview.configure(state="normal")
        self.preview.delete("1.0", END)
        if md_path and md_path.exists():
            try:
                self.preview.insert(END, md_path.read_text(encoding="utf-8"))
            except Exception as e:
                self.preview.insert(END, f"(No se pudo leer la vista previa: {e})")
        else:
            self.preview.insert(
                END,
                f"(Aún no convertido)\n\nOrigen: {src}\n\n"
                "Pulsa 'Convertir a Markdown' para generar el .md.",
            )
        self.preview.configure(state="disabled")

    # --------------------------------------------------------- selección

    def elegir_archivos(self) -> None:
        rutas = filedialog.askopenfilenames(
            title="Selecciona archivos para convertir",
            filetypes=FILE_TYPES,
        )
        if rutas:
            for r in rutas:
                self.archivos.append(Path(r))
            self._refrescar_lista()

    def elegir_carpeta(self) -> None:
        carpeta = filedialog.askdirectory(title="Selecciona una carpeta")
        if not carpeta:
            return
        base = Path(carpeta)
        it = base.rglob("*") if self.recursivo.get() else base.iterdir()
        agregados = 0
        for p in it:
            if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
                self.archivos.append(p)
                agregados += 1
        if agregados == 0:
            messagebox.showinfo(
                "Sin resultados",
                "No se encontraron archivos soportados en esa carpeta.",
            )
        self._refrescar_lista()

    def elegir_salida(self) -> None:
        carpeta = filedialog.askdirectory(
            title="Carpeta donde guardar los .md",
            initialdir=self.salida.get() or str(default_downloads_dir()),
        )
        if carpeta:
            self.salida.set(carpeta)

    def quitar_seleccion(self) -> None:
        sel = self.listbox.curselection()
        if not sel:
            return
        del self.archivos[sel[0]]
        self._refrescar_lista()

    def limpiar(self) -> None:
        self.archivos.clear()
        self.resultados.clear()
        self._refrescar_lista()
        self.estado.config(text="Listo.")
        self.progress.config(value=0, maximum=100)
        self.lbl_progreso.config(text="")
        self.btn_abrir_salida.config(state="disabled")

    def _refrescar_lista(self) -> None:
        self.listbox.delete(0, END)
        for p in self.archivos:
            marca = "✓ " if str(p) in self.resultados else "   "
            self.listbox.insert(END, f"{marca}{p}")
        self.estado.config(text=f"{len(self.archivos)} archivo(s) en cola.")

    # ----------------------------------------------------- abrir salida

    def abrir_salida(self) -> None:
        ruta = self._ultima_salida if hasattr(self, "_ultima_salida") else None
        if not ruta:
            ruta = Path(self.salida.get()) if self.salida.get() \
                else default_downloads_dir()
        try:
            os.startfile(str(ruta))  # Windows
        except AttributeError:
            messagebox.showinfo("Carpeta de salida", str(ruta))

    # --------------------------------------------------------- conversión

    def convertir(self) -> None:
        if not self.archivos:
            messagebox.showwarning(
                "Nada que convertir",
                "Añade archivos o una carpeta antes de convertir.",
            )
            return
        try:
            import markitdown  # noqa: F401
        except ImportError:
            messagebox.showerror(
                "Falta MarkItDown",
                'Esta app necesita la librería MarkItDown.\n\n'
                'Abre PowerShell y ejecuta:\n\n'
                '    pip install "markitdown[all]"',
            )
            return

        self.cancelar.clear()
        self.conversion_en_curso = True
        self.btn_convertir.config(state="disabled")
        self.btn_cancelar.config(state="normal")
        self.btn_abrir_salida.config(state="disabled")
        self.progress.config(value=0, maximum=len(self.archivos))
        self.lbl_progreso.config(text=f"0 de {len(self.archivos)}")
        self.estado.config(text="Convirtiendo…")
        threading.Thread(target=self._worker, daemon=True).start()

    def cancelar_conversion(self) -> None:
        self.cancelar.set()
        self.estado.config(text="Cancelando…")

    def _worker(self) -> None:
        from markitdown import MarkItDown
        md = MarkItDown()

        if self.junto_al_original.get():
            out_dir_base = None
        else:
            out_dir_base = Path(self.salida.get()).expanduser() \
                if self.salida.get() else default_downloads_dir()
            out_dir_base.mkdir(parents=True, exist_ok=True)
        self._ultima_salida = out_dir_base or self.archivos[0].parent

        ok = 0
        fail = 0
        ocr_avisos: list[str] = []

        for idx, src in enumerate(list(self.archivos), start=1):
            if self.cancelar.is_set():
                break

            try:
                target_dir = src.parent if self.junto_al_original.get() else out_dir_base
                target_dir.mkdir(parents=True, exist_ok=True)
                target = unique_path(target_dir / (src.stem + ".md"))

                result = md.convert(str(src))
                text = result.text_content or ""

                if self.postprocesar.get():
                    text = postprocess_markdown(text)

                if self.frontmatter.get():
                    extra = {}
                    if src.suffix.lower() == ".pdf":
                        n = get_pdf_page_count(src)
                        if n is not None:
                            extra["pages"] = n
                    text = build_frontmatter(src, extra) + text

                target.write_text(text, encoding="utf-8")
                self.resultados[str(src)] = target

                if self.aviso_ocr.get() and looks_like_scanned_pdf(src, text):
                    ocr_avisos.append(src.name)

                ok += 1
            except Exception as e:
                fail += 1
                self.root.after(0, lambda n=src.name, m=str(e):
                                self.estado.config(text=f"Error en {n}: {m}"))

            self.root.after(0, self._actualizar_progreso, idx, len(self.archivos))

        cancelado = self.cancelar.is_set()
        self.root.after(0, self._fin_conversion, ok, fail, cancelado, ocr_avisos)

    def _actualizar_progreso(self, hecho: int, total: int) -> None:
        self.progress.config(value=hecho)
        self.lbl_progreso.config(text=f"{hecho} de {total}")
        self._refrescar_lista()

    def _fin_conversion(self, ok: int, fail: int, cancelado: bool,
                        ocr_avisos: list[str]) -> None:
        self.conversion_en_curso = False
        self.btn_convertir.config(state="normal")
        self.btn_cancelar.config(state="disabled")
        if ok > 0:
            self.btn_abrir_salida.config(state="normal")

        if cancelado:
            resumen = f"Cancelado. Convertidos: {ok}. Fallidos: {fail}."
        else:
            resumen = f"Terminado. Convertidos: {ok}. Fallidos: {fail}."
        self.estado.config(text=resumen)

        if ocr_avisos:
            messagebox.showwarning(
                "Posibles PDFs escaneados",
                "Estos PDFs apenas tienen texto extraíble — probablemente "
                "son escaneos sin capa de texto. Considera pasarles OCR antes:\n\n"
                + "\n".join(f"• {n}" for n in ocr_avisos)
                + "\n\nUna opción es la herramienta 'ocrmypdf':\n"
                "    pip install ocrmypdf\n"
                "    ocrmypdf entrada.pdf salida.pdf"
            )
        elif ok > 0 and not cancelado:
            messagebox.showinfo(
                "Conversión terminada",
                f"Se convirtieron {ok} archivo(s).\n"
                f"Destino: {self._ultima_salida}",
            )

    # -------------------------------------------------------------- cierre

    def _on_close(self) -> None:
        if self.conversion_en_curso:
            if not messagebox.askyesno(
                "Conversión en curso",
                "Hay una conversión en marcha. ¿Cancelar y salir?",
            ):
                return
            self.cancelar.set()
            # Espera breve a que el hilo termine el archivo actual
            self.estado.config(text="Cerrando…")
            self.root.after(500, self._on_close)
            return
        save_config({
            "salida": self.salida.get(),
            "recursivo": self.recursivo.get(),
            "junto_al_original": self.junto_al_original.get(),
            "postprocesar": self.postprocesar.get(),
            "frontmatter": self.frontmatter.get(),
            "aviso_ocr": self.aviso_ocr.get(),
        })
        self.root.destroy()


# -------------------------------------------------------------------- main

def _log_launch() -> None:
    """Deja constancia de cada arranque (y sus argumentos) en
    %APPDATA%\\convertir_a_md\\launch.log. Sirve para confirmar si el menú
    contextual / 'Enviar a' realmente lanza la app. Silencioso."""
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_DIR / "launch.log", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat(timespec='seconds')}  "
                    f"argv={sys.argv[1:]}\n")
    except Exception:
        pass


def main() -> None:
    _log_launch()
    archivos = []
    for a in sys.argv[1:]:
        try:
            p = Path(a)
            if p.exists():
                archivos.append(p)
        except Exception:
            pass

    root = BaseTk()
    try:
        ttk.Style().theme_use("vista")
    except Exception:
        pass
    ConvertirApp(root, archivos_iniciales=archivos)
    root.mainloop()


if __name__ == "__main__":
    main()
