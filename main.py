#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cody PDF 1.0 — Universal‑to‑PDF Converter
Auteur : Valentin GIDON (OTI) – MIT
"""

# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
# Imports standard
# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
import sys, subprocess, importlib, shutil, textwrap, threading, queue, os
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

print(f"► Cody PDF 1.0 — Python : {Path(sys.executable)}", flush=True)

# pip‑name → import‑name
_REQ = {
    "Pillow":       "PIL",         "img2pdf":      "img2pdf",
    "reportlab":    "reportlab",   "PyPDF2":       "PyPDF2",
    "docx2pdf":     "docx2pdf",    "pdfkit":       "pdfkit",
    "markdown":     "markdown",    "ttkbootstrap": "ttkbootstrap",
}

def _ensure(pip_name, mod_name):
    try:
        return importlib.import_module(mod_name)
    except ImportError:
        print(f"• {mod_name} absent — pip install {pip_name}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
        return importlib.import_module(mod_name)

# Dépendances de base
PIL         = _ensure("Pillow",   "PIL")
img2pdf_mod = _ensure("img2pdf",  "img2pdf")
reportlab   = _ensure("reportlab","reportlab")
PyPDF2_mod  = _ensure("PyPDF2",   "PyPDF2")

# Dépendances optionnelles
try:    docx2pdf_mod = _ensure("docx2pdf", "docx2pdf")
except Exception: docx2pdf_mod = None

try:
    pdfkit_mod = _ensure("pdfkit", "pdfkit")
    md         = _ensure("markdown","markdown")
except Exception: pdfkit_mod = None

# Thème bootstrap
try:
    import ttkbootstrap as tb; from ttkbootstrap.constants import *
    _THEME_OK = True
except Exception:
    _THEME_OK = False


# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
# Imports dépendants
# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from PyPDF2 import PdfMerger


# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
# Helpers
# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
def _exe_in_path(bin_name: str) -> bool: return shutil.which(bin_name) is not None

def _unique_path(p: Path) -> Path:
    """Renvoie un chemin libre : foo.pdf → foo (1).pdf si déjà présent."""
    if not p.exists(): return p
    stem, suf, i = p.stem, p.suffix, 1
    while True:
        test = p.with_name(f"{stem} ({i}){suf}")
        if not test.exists(): return test
        i += 1


# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
# Convertisseurs
# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
def convert_image(src: Path, dest_dir: Path) -> Path:
    dest = _unique_path(dest_dir / f"{src.stem}.pdf")
    with open(dest, "wb") as out, open(src, "rb") as img:
        out.write(img2pdf_mod.convert(img))
    return dest

def convert_text(src: Path, dest_dir: Path) -> Path:
    dest = _unique_path(dest_dir / f"{src.stem}.pdf")
    txt = src.read_text(encoding="utf-8", errors="ignore").replace("\t", "    ")
    c   = canvas.Canvas(str(dest), pagesize=A4)
    x,y,margin,lead = 72, A4[1]-72, 72, 14
    for line in txt.splitlines():
        for seg in textwrap.wrap(line, 95) or [" "]:
            if y < margin: c.showPage(); y = A4[1]-margin
            c.drawString(x, y, seg); y -= lead
    c.save(); return dest

def convert_office(src: Path, dest_dir: Path) -> Path:
    dest = _unique_path(dest_dir / f"{src.stem}.pdf")
    if src.suffix.lower() == ".docx" and docx2pdf_mod:
        try: docx2pdf_mod.convert(str(src), str(dest)); return dest
        except Exception: pass
    if _exe_in_path("soffice"):
        subprocess.check_call(["soffice","--headless","--convert-to","pdf",str(src),"--outdir",str(dest_dir)],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        final = dest_dir / f"{src.stem}.pdf"
        if final != dest: final.rename(dest)
        return dest
    raise RuntimeError("LibreOffice/Word non disponible.")

def convert_odf(src: Path, dest_dir: Path) -> Path:
    if not _exe_in_path("soffice"): raise RuntimeError("LibreOffice requis pour ODF.")
    dest = _unique_path(dest_dir / f"{src.stem}.pdf")
    subprocess.check_call(["soffice","--headless","--convert-to","pdf",str(src),"--outdir",str(dest_dir)],
                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    final = dest_dir / f"{src.stem}.pdf"
    if final != dest: final.rename(dest)
    return dest

def convert_html(src: Path, dest_dir: Path) -> Path:
    if not pdfkit_mod or not _exe_in_path("wkhtmltopdf"):
        raise RuntimeError("wkhtmltopdf manquant")
    dest = _unique_path(dest_dir / f"{src.stem}.pdf")
    pdfkit_mod.from_file(str(src), str(dest)); return dest

def convert_md(src: Path, dest_dir: Path) -> Path:
    if not pdfkit_mod or not _exe_in_path("wkhtmltopdf"):
        raise RuntimeError("wkhtmltopdf manquant")
    html = md.markdown(src.read_text(encoding="utf-8", errors="ignore"))
    tmp  = dest_dir / f"_{src.stem}.html";  tmp.write_text(html, encoding="utf-8")
    dest = convert_html(tmp, dest_dir);     tmp.unlink(missing_ok=True)
    return dest


# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
# Registre extensions → fonction
# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
HANDLERS = {}
for e in ".png .jpg .jpeg .bmp .gif .tif .tiff".split(): HANDLERS[e]=convert_image
for e in ".txt .log .csv".split():                       HANDLERS[e]=convert_text
for e in ".docx .pptx .xlsx .rtf".split():               HANDLERS[e]=convert_office
for e in ".odt .odp .ods".split():                       HANDLERS[e]=convert_odf
for e in ".html .htm .mht".split():                      HANDLERS[e]=convert_html
HANDLERS[".md"]=convert_md


# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
# GUI
# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
class CodyPDF:
    def __init__(self):
        # ---- Fenêtre / thème ----------------------------------------------
        if _THEME_OK:
            self.root = tb.Window(title="Cody PDF 1.0 — Universal → PDF",
                                  themename="darkly", size=(850, 540),
                                  minsize=(720, 440))
            ttkb = tb
            ttkb.Style().configure(".", foreground="#e8f0ff")
        else:
            self.root = tk.Tk(); self.root.title("Cody PDF 1.0 — Universal → PDF")
            self.root.geometry("850x540"); self.root.minsize(720, 440)
            ttkb = ttk
        self.ttk, self.style = ttkb, ttkb.Style()

        self.files: list[str] = []
        self.out_dir: Path|None = None
        self.q = queue.Queue()

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------------------------------------------------------- UI
    def _build_ui(self):
        ttkb = self.ttk
        top = ttkb.Frame(self.root); top.pack(fill="x", padx=10, pady=6)
        ttkb.Button(top, text="Ajouter fichiers", command=self._add).pack(side="left", padx=5)
        ttkb.Button(top, text="Supprimer",        command=self._rm ).pack(side="left", padx=5)
        ttkb.Button(top, text="Vider",            command=self._clear).pack(side="left", padx=5)

        self.same_dir = tk.BooleanVar(self.root, False)
        ttkb.Checkbutton(top, text="Même dossier que l’original", variable=self.same_dir)\
            .pack(side="left", padx=15)

        self.merge = tk.BooleanVar(self.root, False)
        ttkb.Checkbutton(top, text="Fusionner en un seul PDF", variable=self.merge)\
            .pack(side="left", padx=15)

        lst = ttkb.Frame(self.root); lst.pack(fill="both", expand=True, padx=10, pady=5)
        self.lb = ttkb.Treeview(lst, columns=("path",), show="headings", selectmode="extended")
        self.lb.heading("path", text="Fichiers à convertir"); self.lb.column("path", anchor="w")
        vsb = ttkb.Scrollbar(lst, orient="vertical", command=self.lb.yview)
        self.lb.configure(yscrollcommand=vsb.set)
        self.lb.pack(side="left", fill="both", expand=True); vsb.pack(side="right", fill="y")

        self.progress = ttkb.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=6)

        ttkb.Button(self.root, text="Démarrer", command=self._start,
                    bootstyle="success" if _THEME_OK else None).pack(pady=6)

        self.status = tk.StringVar(value="Prêt.")
        lbl = ttkb.Label(self.root, textvariable=self.status, anchor="w", cursor="hand2")
        lbl.pack(fill="x", padx=10, pady=(0,6))
        lbl.bind("<Button-1>", lambda e: self._open_out_dir())

    # -------------------------------------------------------------- callbacks
    def _add(self):
        exts = {e.lstrip('.') for e in HANDLERS}|{"pdf"}
        paths = filedialog.askopenfilenames(
            title="Choisissez des fichiers",
            filetypes=[("Formats gérés"," ".join(f"*.{e}" for e in sorted(exts))),
                       ("Tous", "*.*")])
        for p in paths:
            if p not in self.files:
                self.files.append(p); self.lb.insert("", "end", values=(p,))
        self.status.set(f"{len(self.files)} fichier(s) en attente.")

    def _rm(self):
        for iid in self.lb.selection():
            p=self.lb.item(iid,"values")[0]; self.files.remove(p); self.lb.delete(iid)
        self.status.set(f"{len(self.files)} fichier(s) restant(s).")

    def _clear(self):
        self.lb.delete(*self.lb.get_children()); self.files.clear()
        self.status.set("Liste vidée.")

    # ----------------------------------------------------------- conversion
    def _start(self):
        if not self.files:
            messagebox.showwarning("Aucun fichier", "Ajoutez des fichiers avant de démarrer."); return

        # — Sélection du dossier cible (si nécessaire) —
        if not self.same_dir.get():
            target = filedialog.askdirectory(
                title="Choisissez le dossier de destination",
                initialdir=str(self.out_dir or Path.home()))
            if not target:
                self.status.set("Opération annulée."); return
            self.out_dir = Path(target)
            self.status.set(f"Dossier choisi : {self.out_dir}")
        else:
            self.out_dir = None  # signaler qu’on travaille par‑fichier

        self.progress.configure(maximum=len(self.files), value=0)
        self.status.set("Conversion…")
        threading.Thread(target=self._worker, daemon=True).start()
        self.root.after(200, self._poll)

    def _worker(self):
        produced=[]
        for idx, fp in enumerate(self.files, 1):
            src = Path(fp); ext = src.suffix.lower()
            dest_dir = src.parent if self.same_dir.get() else self.out_dir
            try:
                if ext in HANDLERS:
                    produced.append(HANDLERS[ext](src, dest_dir))
                elif ext==".pdf":
                    dest=_unique_path(dest_dir/src.name); dest.write_bytes(src.read_bytes())
                    produced.append(dest)
                else:
                    self.q.put(("warn", f"Format non pris : {src.name}"))
            except Exception as e:
                self.q.put(("error", f"{src.name} : {e}"))
            self.q.put(("progress", idx))

        if self.merge.get() and produced:
            merge_dir = produced[0].parent
            merged=_unique_path(merge_dir/f"merged_{datetime.now():%Y%m%d_%H%M%S}.pdf")
            with PdfMerger() as m:
                for pdf in produced: m.append(str(pdf))
                m.write(str(merged))
            produced.append(merged)
            self.q.put(("info", f"Fusion : {merged.name}"))

        final_dir = produced[0].parent if produced else (self.out_dir or Path.cwd())
        self.q.put(("done", f"{len(produced)} PDF créés dans {final_dir}"))

    def _poll(self):
        try:
            while True:
                kind,msg=self.q.get_nowait()
                if kind=="progress":
                    self.progress["value"]=msg; self.status.set(f"{msg}/{len(self.files)} traités…")
                elif kind=="warn":  messagebox.showwarning("Attention", msg)
                elif kind=="info":  messagebox.showinfo("Info", msg)
                elif kind=="error": messagebox.showerror("Erreur", msg)
                elif kind=="done":
                    self.progress["value"]=self.progress["maximum"]
                    self.status.set(f"{msg} — cliquer pour ouvrir")
        except queue.Empty:
            pass
        finally:
            if self.progress["value"]<self.progress["maximum"]:
                self.root.after(200, self._poll)

    # -------------------------------------------------------------- utilitaires
    def _open_out_dir(self):
        dir_to_open = None
        if self.same_dir.get():
            if self.files: dir_to_open = Path(self.files[0]).parent
        else:
            dir_to_open = self.out_dir
        if dir_to_open and dir_to_open.exists():
            if sys.platform.startswith("win"):   os.startfile(dir_to_open)
            elif sys.platform.startswith("darwin"): subprocess.call(["open", dir_to_open])
            else: subprocess.call(["xdg-open", dir_to_open])
        else:
            messagebox.showinfo("Info", "Aucun dossier de sortie défini.")

    def _on_close(self):
        if messagebox.askokcancel("Quitter", "Fermer Cody PDF ?"): self.root.destroy()

    def run(self): self.root.mainloop()


# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
# Entrée
# :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
if __name__ == "__main__":
    try:
        CodyPDF().run()
    except subprocess.CalledProcessError as cpe:
        messagebox.showerror("Conversion externe échouée",
                             f"Commande terminée avec le code {cpe.returncode}\n{cpe.cmd}")
