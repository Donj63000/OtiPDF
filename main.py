#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mini‑application de conversion multi‑format ➜ PDF
Auteur : Valentin GIDON (OTI) — licence MIT

Formats pris en charge par défaut
---------------------------------
• Images : .png .jpg .jpeg .bmp .gif .tif .tiff
• Textes : .txt .md .csv .log
• Word   : .docx
• PDF    : déjà PDF — option de fusion

Dépendances externes  (≥ versions stables compatibles Python 3.8 → 3.13)
-----------------------------------------------------------------------
pip install pillow reportlab python-docx pypdf2 ttkbootstrap
"""

# ====== Imports standards ===================================================
import threading
import queue
from pathlib import Path
from datetime import datetime
import textwrap
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ====== Imports optionnels (gestion thème) ==================================
try:
    import ttkbootstrap as tb                 # thème sombre/bleu
    from ttkbootstrap.constants import *
    _THEME_OK = True
except ImportError:
    _THEME_OK = False

# ====== Imports dépendances fonctionnelles ==================================
# → les bloc try/except permettent d’afficher un message clair si l’une manque
try:
    from PIL import Image
except ImportError as e:  # pragma: no cover
    messagebox.showerror("Module manquant",
                         "Le module Pillow n’est pas installé !\n"
                         "Exécute :  pip install pillow")
    raise e

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
except ImportError as e:  # pragma: no cover
    messagebox.showerror("Module manquant",
                         "Le module reportlab n’est pas installé !\n"
                         "Exécute :  pip install reportlab")
    raise e

try:
    import docx
except ImportError:            # pragma: no cover
    docx = None                # Le .docx sera désactivé proprement

try:
    from PyPDF2 import PdfMerger
except ImportError as e:       # pragma: no cover
    messagebox.showerror("Module manquant",
                         "Le module PyPDF2 n’est pas installé !\n"
                         "Exécute :  pip install pypdf2")
    raise e

# ====== Constantes de formats ===============================================
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}
TEXT_EXTS  = {".txt", ".md", ".csv", ".log"}
DOCX_EXTS  = {".docx"} if docx else set()

# ----------------------------------------------------------------------------
#                         Convertisseurs de fichiers
# ----------------------------------------------------------------------------
def convert_image(path: Path, out_dir: Path) -> Path:
    """Image → PDF (1 page)"""
    with Image.open(path) as img:
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        pdf_path = out_dir / f"{path.stem}.pdf"
        img.save(pdf_path, "PDF", resolution=300.0)
    return pdf_path


def _write_text(c: canvas.Canvas, text: str, page_w: float, page_h: float,
                margin: int = 72, leading: int = 14) -> None:
    """Écrit `text` dans le canvas avec wrap et sauts de page."""
    x = margin
    y = page_h - margin
    for line in text.splitlines():
        for seg in textwrap.wrap(line, width=95) or [" "]:
            if y < margin:
                c.showPage()
                y = page_h - margin
            c.drawString(x, y, seg)
            y -= leading


def convert_text(path: Path, out_dir: Path) -> Path:
    """Fichier texte (UTF‑8) → PDF A4."""
    pdf_path = out_dir / f"{path.stem}.pdf"
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read().replace("\t", "    ")

    c = canvas.Canvas(str(pdf_path), pagesize=A4)
    _write_text(c, content, *A4)
    c.save()
    return pdf_path


def convert_docx(path: Path, out_dir: Path) -> Path:
    """DOCX → PDF par extraction du texte brut."""
    if not docx:
        raise RuntimeError("python-docx non disponible !")
    document = docx.Document(str(path))
    full_text = "\n".join(p.text for p in document.paragraphs)
    tmp_txt = out_dir / f"_{path.stem}.txt"
    tmp_txt.write_text(full_text, encoding="utf-8")
    pdf_path = convert_text(tmp_txt, out_dir)
    tmp_txt.unlink(missing_ok=True)
    return pdf_path


# Dictionnaire dynamique extension ➜ handler
HANDLERS = {ext: convert_image for ext in IMAGE_EXTS}
HANDLERS.update({ext: convert_text  for ext in TEXT_EXTS})
HANDLERS.update({ext: convert_docx  for ext in DOCX_EXTS})


# ----------------------------------------------------------------------------
#                               GUI
# ----------------------------------------------------------------------------
class PDFConverterApp:
    def __init__(self) -> None:
        # -------- Fenêtre / thème ----------
        if _THEME_OK:
            self.root = tb.Window(title="Universal → PDF Converter",
                                  themename="superhero",                   # sombre/bleu
                                  size=(780, 480), minsize=(640, 400))
            self.ttk = tb
        else:
            self.root = tk.Tk()
            self.root.title("Universal → PDF Converter")
            self.root.geometry("780x480")
            self.root.minsize(640, 400)
            self.ttk = ttk

        # -------- Données ----------
        self.files: list[str] = []
        self.out_dir = Path.cwd()
        self._q: queue.Queue = queue.Queue()

        # -------- Interface ----------
        self._build_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ GUI
    def _build_widgets(self) -> None:
        ttkb = self.ttk

        # --- Barre de boutons haut
        top = ttkb.Frame(self.root)
        top.pack(fill="x", padx=10, pady=5)

        ttkb.Button(top, text="Ajouter fichiers", command=self._add_files).pack(side="left", padx=5)
        ttkb.Button(top, text="Supprimer sélection", command=self._remove_selected).pack(side="left", padx=5)
        ttkb.Button(top, text="Vider liste", command=self._clear_list).pack(side="left", padx=5)
        ttkb.Button(top, text="Dossier de sortie…", command=self._choose_out_dir).pack(side="left", padx=5)

        self.merge_var = tk.BooleanVar(value=False)
        ttkb.Checkbutton(top, text="Fusionner en un seul PDF", variable=self.merge_var).pack(side="left", padx=15)

        # --- Liste des fichiers
        lst_frame = ttkb.Frame(self.root)
        lst_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.lb = ttkb.Treeview(lst_frame, columns=("path",), show="headings", selectmode="extended")
        self.lb.heading("path", text="Fichiers à convertir")
        self.lb.column("path", anchor="w")
        vsb_kwargs = {"bootstyle": "info-round"} if _THEME_OK else {}
        vsb = ttkb.Scrollbar(lst_frame, orient="vertical", command=self.lb.yview, **vsb_kwargs)
        self.lb.configure(yscrollcommand=vsb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # --- Barre de progression
        self.progress = ttkb.Progressbar(self.root, orient="horizontal",
                                         mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=10, pady=5)

        # --- Bouton démarrer
        btn_opts = {"bootstyle": "success"} if _THEME_OK else {}
        ttkb.Button(self.root, text="Démarrer la conversion",
                    command=self._start_conversion, **btn_opts).pack(pady=8)

        # --- Statut
        self.status = tk.StringVar(value="Prêt.")
        ttkb.Label(self.root, textvariable=self.status, anchor="w").pack(fill="x", padx=10, pady=(0, 5))

    # ---------------------------------------------------------------- Callbacks
    def _add_files(self) -> None:
        types = [("Fichiers pris en charge",
                  "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.txt *.md *.csv *.log *.docx *.pdf"),
                 ("Images", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff"),
                 ("Textes", "*.txt *.md *.csv *.log"),
                 ("Word (DOCX)", "*.docx"),
                 ("PDF", "*.pdf"),
                 ("Tous", "*.*")]
        paths = filedialog.askopenfilenames(title="Sélectionnez vos fichiers", filetypes=types)
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.lb.insert("", "end", values=(p,))
        self.status.set(f"{len(self.files)} fichier(s) en attente.")

    def _remove_selected(self) -> None:
        for iid in self.lb.selection():
            path = self.lb.item(iid, "values")[0]
            self.files.remove(path)
            self.lb.delete(iid)
        self.status.set(f"{len(self.files)} fichier(s) restants.")

    def _clear_list(self) -> None:
        self.lb.delete(*self.lb.get_children())
        self.files.clear()
        self.status.set("Liste vidée.")

    def _choose_out_dir(self) -> None:
        new = filedialog.askdirectory(title="Choisissez le dossier de sortie", mustexist=True)
        if new:
            self.out_dir = Path(new)
            self.status.set(f"Dossier de sortie : {self.out_dir}")

    # ---------------------------------------------------------- Conversion
    def _start_conversion(self) -> None:
        if not self.files:
            messagebox.showwarning("Aucun fichier", "Ajoutez au moins un fichier avant de lancer la conversion.")
            return
        self.progress.configure(value=0, maximum=len(self.files))
        self.status.set("Conversion en cours…")
        threading.Thread(target=self._worker, daemon=True).start()
        self.root.after(150, self._poll_queue)

    def _worker(self) -> None:
        try:
            produced: list[Path] = []
            for idx, file_path in enumerate(self.files, start=1):
                p = Path(file_path)
                ext = p.suffix.lower()
                if ext in HANDLERS:
                    produced.append(HANDLERS[ext](p, self.out_dir))
                elif ext == ".pdf":
                    dest = self.out_dir / p.name
                    if dest.resolve() != p.resolve():
                        dest.write_bytes(p.read_bytes())
                    produced.append(dest)
                else:
                    self._q.put(("warn", f"Extension non prise en charge : {p.name}"))
                self._q.put(("progress", idx))

            # Fusion
            if self.merge_var.get() and produced:
                merged = self.out_dir / f"merged_{datetime.now():%Y%m%d_%H%M%S}.pdf"
                with PdfMerger() as m:
                    for pdf in produced:
                        m.append(str(pdf))
                    m.write(str(merged))
                self._q.put(("info", f"Fusion terminée : {merged.name}"))

            self._q.put(("done", "Conversion terminée."))
        except Exception as exc:                # pragma: no cover
            self._q.put(("error", str(exc)))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, msg = self._q.get_nowait()
                if kind == "progress":
                    self.progress["value"] = msg
                    self.status.set(f"Traitement {msg}/{len(self.files)}…")
                elif kind == "warn":
                    messagebox.showwarning("Attention", msg)
                elif kind == "info":
                    messagebox.showinfo("Information", msg)
                elif kind == "error":
                    messagebox.showerror("Erreur", msg)
                    self.status.set("Erreur rencontrée !")
                elif kind == "done":
                    self.progress["value"] = self.progress["maximum"]
                    self.status.set(msg)
        except queue.Empty:
            pass
        finally:
            if self.progress["value"] < self.progress["maximum"]:
                self.root.after(150, self._poll_queue)

    # -------------------------------------------------------------- Divers
    def _on_close(self) -> None:
        if messagebox.askokcancel("Quitter", "Voulez‑vous vraiment quitter ?"):
            self.root.destroy()

    # -------------------------------------------------------------- Run
    def run(self) -> None:
        self.root.mainloop()


# ======== Point d’entrée =====================================================
if __name__ == "__main__":
    PDFConverterApp().run()
