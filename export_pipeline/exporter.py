#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exportiert alle *.scatt‑Dateien in einem gewählten Überordner mittels samples.vbs
in gleichnamige *.scatt.txt‑Dateien.
• nutzt 32‑Bit‑cscript.exe (SysWoW64) für die SCATT‑COM‑Komponente
• beschädigte Dateien werden still übersprungen
• Fortschrittsfenster mit grünem Balken & Prozentanzeige
Legen Sie export_scatt.py und samples.vbs in denselben Ordner.
"""

import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# ---------------------------------------------------------------------------
#  Pfade zu Hilfsprogrammen
# ---------------------------------------------------------------------------
THIS_DIR    = Path(sys.argv[0]).resolve().parent
SAMPLES_VBS = THIS_DIR / "samples.vbs"
CSCRIPT32   = Path(os.environ["windir"]) / "SysWoW64" / "cscript.exe"   # 32‑Bit!


# ---------------------------------------------------------------------------
#  Hilfsfunktionen
# ---------------------------------------------------------------------------
def collect_pending(root: Path):
    """Alle .scatt‑Dateien ohne vorhandene .scatt.txt sammeln."""
    return [p for p in root.rglob("*.scatt")
            if not p.with_suffix(p.suffix + ".txt").exists()]


def export_file(scatt: Path):
    """
    Führt den VBS‑Export durch.
    Gibt Tuple (done=True, skipped_bool) zurück.
    Fehler (korrupt) → skipped = True, ansonsten False.
    """
    try:
        subprocess.run(
            [str(CSCRIPT32), "//NoLogo", str(SAMPLES_VBS), scatt.name],
            cwd=scatt.parent,                 # Arbeitsverzeichnis = Ordner der Datei
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        return True, False
    except subprocess.CalledProcessError:
        # Datei fehlerhaft → überspringen
        return True, True


# ---------------------------------------------------------------------------
#  Hauptprogramm
# ---------------------------------------------------------------------------
def main():
    # Ordner wählen
    root_tk = tk.Tk()
    root_tk.withdraw()
    sel = filedialog.askdirectory(title="Überordner mit .scatt‑Dateien auswählen")
    if not sel:
        return
    root_dir = Path(sel)

    # Grundchecks
    if not SAMPLES_VBS.exists():
        messagebox.showerror("Fehler",
                             f"'{SAMPLES_VBS.name}' nicht neben diesem Skript gefunden.")
        return
    if not CSCRIPT32.exists():
        messagebox.showerror("Fehler",
                             f"32‑Bit‑cscript.exe nicht gefunden:\n{CSCRIPT32}")
        return

    pending = collect_pending(root_dir)
    if not pending:
        messagebox.showinfo("Nichts zu tun",
                            "Alle .scatt‑Dateien in diesem Ordner wurden bereits exportiert.")
        return

    # ---------- Fortschrittsfenster (verschönert) ----------
    win = tk.Toplevel()
    win.title("Export läuft …")
    win.resizable(False, False)
    win.configure(padx=24, pady=20)

    ttk.Style().theme_use("clam")  # Theme, um Farbe zu setzen
    style = ttk.Style()
    style.configure(
        "Green.Horizontal.TProgressbar",
        troughcolor="#d9d9d9",
        bordercolor="#d9d9d9",
        background="#4caf50",
        lightcolor="#4caf50",
        darkcolor="#4caf50",
        thickness=18
    )

    pb = ttk.Progressbar(
        win, style="Green.Horizontal.TProgressbar",
        mode="determinate",
        maximum=len(pending),
        length=500
    )
    pb.pack(fill="x")

    percent_lbl = tk.Label(win, text="0 %", font=("Segoe UI", 10, "bold"))
    percent_lbl.pack(anchor="e", pady=(2, 14))

    status = tk.Label(win, anchor="w", justify="left", wraplength=500)
    status.pack(fill="x")

    # ---------- Export‑Loop ----------
    skipped = 0
    for done, scatt in enumerate(pending, 1):
        # Dateipfad anzeigen (relativ)
        status.config(text=os.path.relpath(scatt, root_dir))
        status.update_idletasks()

        # Export durchführen
        _, was_skipped = export_file(scatt)
        if was_skipped:
            skipped += 1

        # Fortschritt aktualisieren
        pb["value"] = done
        percent_lbl.config(text=f"{done * 100 // len(pending)} %")
        win.update_idletasks()

    # ---------- Ergebnis ----------
    win.destroy()
    exported = len(pending) - skipped
    messagebox.showinfo(
        "Fertig",
        f"Vorgang abgeschlossen.\n"
        f"Exportiert:   {exported}\n"
        f"Übersprungen: {skipped}"
    )


if __name__ == "__main__":
    main()
