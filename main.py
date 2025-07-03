#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mini-application de conversion multi‑format ➜ PDF
Auteur : (c) 2025, libre de droits – vous pouvez l’utiliser/vendre/modifier.

Formats gérés par défaut :
  • Images  : .png, .jpg, .jpeg, .bmp, .gif, .tif, .tiff
  • Texte   : .txt, .md, .csv, .log…
  • DOCX    : .docx
  • PDF     : (déjà PDF) – peut être fusionné
"""

# ------------------ Imports standard ------------------
import os
import sys
import threading
import queue
from pathlib import Path
from datetime import datetime

# ------------------ Imports externes ------------------
try:
    import ttkbootstrap as tb                 # Thème sombre/bleu
    from ttkbootstrap.constants import *
    TTKBOOTSTRAP_AVAILABLE = True
except ImportError:
    import tkinter as tk                      # Fallback sans thème
    from tkinter import ttk
    TTKBOOTSTRAP_AVAILABLE = False

from tkinter import filedialog, messagebox

from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import utils

import docx
from PyPDF2 import PdfMerger

# ------------------ Conversion helpers ------------------
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}
TEXT_EXTS  = {".txt", ".md", ".csv", ".log"}
DOCX_EXTS  = {".docx"}

def convert_image(path: Path, out_dir: Path) -> Path:
    """Convertit une image en PDF et renvoie le chemin du PDF créé."""
    img = Image.open(path)
    if img.mode in ("RGBA", "P"):             # convertit tout en RVB
        img = img.convert("RGB")
    pdf_path = out_dir / f"{path.stem}.pdf"
    img.save(pdf_path, "PDF", resolution=300.0)
    return pdf_path

def _write_lines_to_canvas(c, lines, page_w, page_h, margin=72, leading=14):
    """Écrit du texte ligne par ligne avec retour page automatique."""
    x = margin
    y = page_h - margin
    for line in lines:
        if y < margin:
            c.showPage()
            y = page_h - margin
        c.drawString(x, y, line)
        y -= leading

def convert_text(path: Path, out_dir: Path) -> Path:
    """Convertit un fichier texte (UTF‑8) en PDF (A4)."""
    pdf_path = out_dir / f"{path.stem}.pdf"
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read().replace("\t", "    ")

    c = canvas.Canvas(str(pdf_path), pagesize=A4)
    page_w, page_h = A4
    # Découpe le texte sur 95 car./ligne env. (simple, peut être raffiné)
    import textwrap
    wrapped = []
    for paragraph in content.splitlines():
        wrapped.extend(textwrap.wrap(paragraph, width=95) or [" "])
    _write_lines_to_canvas(c, wrapped, page_w, page_h)
    c.save()
    return pdf_path

def convert_docx(path: Path, out_dir: Path) -> Path:
    """Convertit un DOCX en PDF (texte brut)."""
    document = docx.Document(str(path))
    full_text = []
    for para in document.paragraphs:
        full_text.append(para.text)
    tmp_txt = out_dir / f"_{path.stem}.txt"
    tmp_txt.write_text("\n".join(full_text), encoding="utf-8")
    pdf_path = convert_text(tmp_txt, out_dir)
    tmp_txt.unlink(missing_ok=True)
    return pdf_path

# Mapping extension ➜ fonction
handlers = {}
for ext in IMAGE_EXTS:
    handlers[ext] = convert_image
for ext in TEXT_EXTS:
    handlers[ext] = convert_text
for ext in DOCX_EXTS:
    handlers[ext] = convert_docx

# ------------------ GUI principale ------------------
class PDFConverterApp:
    def __init__(self):
        # ----- Fenêtre / thème -----
        if TTKBOOTSTRAP_AVAILABLE:
            self.root = tb.Window(title="Universal‑to‑PDF Converter",
                                  themename="superhero",
                                  size=(780, 480),
                                  minsize=(640, 400))
            ttkb = tb
        else:
            self.root = tk.Tk()
            self.root.title("Universal‑to‑PDF Converter")
            ttkb = ttk
            self.root.geometry("780x480")
            self.root.minsize(640, 400)

        self.files = []            # liste interne des chemins sélectionnés
        self.out_dir = Path.cwd()  # dossier de sortie par défaut
        self.queue = queue.Queue() # pour progres bar

        # ----- Widgets -----
        self._build_widgets(ttkb)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_widgets(self, ttkb):
        # Boutons haut
        btn_frame = ttkb.Frame(self.root)
        btn_frame.pack(fill="x", padx=10, pady=5)

        ttkb.Button(btn_frame, text="Ajouter fichiers", command=self.add_files).pack(side="left", padx=5)
        ttkb.Button(btn_frame, text="Supprimer sélection", command=self.remove_selected).pack(side="left", padx=5)
        ttkb.Button(btn_frame, text="Vider liste", command=self.clear_list).pack(side="left", padx=5)
        ttkb.Button(btn_frame, text="Dossier de sortie…", command=self.choose_out_dir).pack(side="left", padx=5)

        # Checkbox fusion
        self.merge_var = ttkb.BooleanVar(value=False)
        ttkb.Checkbutton(btn_frame, text="Fusionner en un seul PDF", variable=self.merge_var).pack(side="left", padx=15)

        # Listbox des fichiers
        list_frame = ttkb.Frame(self.root)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        if TTKBOOTSTRAP_AVAILABLE:
            vsb_style = {"bootstyle": "info-round"}
        else:
            vsb_style = {}
        self.lb = ttkb.Treeview(list_frame, columns=("path",), show="headings", selectmode="extended")
        self.lb.heading("path", text="Fichiers à convertir")
        self.lb.column("path", anchor="w")
        vsb = ttkb.Scrollbar(list_frame, orient="vertical", command=self.lb.yview, **vsb_style)
        self.lb.configure(yscrollcommand=vsb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Barre de progression
        self.progress = ttkb.Progressbar(self.root, orient="horizontal", mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=10, pady=5)

        # Convert button
        ttkb.Button(self.root, text="Démarrer la conversion", bootstyle="success" if TTKBOOTSTRAP_AVAILABLE else None,
                    command=self.start_conversion).pack(pady=8)

        # Statut
        self.status_var = ttkb.StringVar(value="Prêt.")
        ttkb.Label(self.root, textvariable=self.status_var, anchor="w").pack(fill="x", padx=10, pady=(0,5))

    # ---------- Callbacks ----------
    def add_files(self):
        types = [("Fichiers pris en charge", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.txt *.md *.csv *.log *.docx *.pdf"),
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
        self.status_var.set(f"{len(self.files)} fichier(s) en file d’attente.")

    def remove_selected(self):
        for iid in self.lb.selection():
            path = self.lb.item(iid, "values")[0]
            self.files.remove(path)
            self.lb.delete(iid)
        self.status_var.set(f"{len(self.files)} fichier(s) restants.")

    def clear_list(self):
        self.lb.delete(*self.lb.get_children())
        self.files.clear()
        self.status_var.set("Liste vidée.")

    def choose_out_dir(self):
        new_dir = filedialog.askdirectory(title="Choisissez le dossier de sortie", mustexist=True)
        if new_dir:
            self.out_dir = Path(new_dir)
            self.status_var.set(f"Dossier de sortie : {self.out_dir}")

    # ---------- Conversion threading ----------
    def start_conversion(self):
        if not self.files:
            messagebox.showwarning("Aucun fichier", "Ajoutez au moins un fichier avant de lancer la conversion.")
            return
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.files)
        self.status_var.set("Conversion en cours…")
        t = threading.Thread(target=self._convert_worker, daemon=True)
        t.start()
        self.root.after(100, self._poll_queue)

    def _convert_worker(self):
        try:
            produced_pdfs = []
            for idx, file_path in enumerate(self.files, start=1):
                path = Path(file_path)
                ext = path.suffix.lower()
                if ext in handlers:
                    pdf_path = handlers[ext](path, self.out_dir)
                    produced_pdfs.append(pdf_path)
                elif ext == ".pdf":
                    # On copie simplement ou ajoute pour fusion
                    dest = self.out_dir / path.name
                    if dest.resolve() != path.resolve():
                        dest.write_bytes(path.read_bytes())
                    produced_pdfs.append(dest)
                else:
                    self.queue.put(("warn", f"Extension non prise en charge : {path.name}"))
                self.queue.put(("progress", idx))

            # Fusion ?
            if self.merge_var.get() and produced_pdfs:
                merged_name = f"merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                merged_path = self.out_dir / merged_name
                merger = PdfMerger()
                for pdf in produced_pdfs:
                    merger.append(str(pdf))
                merger.write(str(merged_path))
                merger.close()
                self.queue.put(("info", f"Fusion terminée : {merged_path.name}"))

            self.queue.put(("done", "Conversion terminée."))
        except Exception as e:
            self.queue.put(("error", str(e)))

    def _poll_queue(self):
        try:
            while True:
                kind, msg = self.queue.get_nowait()
                if kind == "progress":
                    self.progress["value"] = msg
                    self.status_var.set(f"Traitement {msg}/{len(self.files)}…")
                elif kind == "warn":
                    messagebox.showwarning("Attention", msg)
                elif kind == "info":
                    messagebox.showinfo("Information", msg)
                elif kind == "error":
                    messagebox.showerror("Erreur", msg)
                    self.status_var.set("Erreur rencontrée !")
                elif kind == "done":
                    self.progress["value"] = self.progress["maximum"]
                    self.status_var.set(msg)
        except queue.Empty:
            pass
        finally:
            # On continue l’écoute tant que la progression n’est pas complète
            if self.progress["value"] < self.progress["maximum"]:
                self.root.after(100, self._poll_queue)

    def on_close(self):
        if messagebox.askokcancel("Quitter", "Voulez‑vous vraiment quitter ?"):
            self.root.destroy()

    # ---------- Run ----------
    def run(self):
        self.root.mainloop()


# ---------------------- Lancement direct ----------------------
if __name__ == "__main__":
    app = PDFConverterApp()
    app.run()
