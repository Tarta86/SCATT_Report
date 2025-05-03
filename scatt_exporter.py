# -*- coding: utf-8 -*-
"""
scatt_exporter.py
Erstellt eine Text‑Datei im selben Format wie das ursprüngliche VBS‑Export‑Skript.
"""
import sys
import pathlib
import win32com.client as win32

# ---------------------------------------------------------------------------
# Hilfs‑Funktionen
# ---------------------------------------------------------------------------

def format_header(session):
    """Erzeugt die ersten vier Kopfzeilen."""
    gun      = session.Gun.Name           # z. B. "10m Air Rifle (AR10)"
    dt       = session.DateTime           # COM‑Date -> python datetime
    athlete  = session.Shooter.Name
    shots    = session.Shots.Count
    return ["{}"
            .format(gun),
            dt.strftime("%d.%m.%Y %H:%M:%S"),
            athlete,
            f"Final  {shots:2d}  shots",
            ""]                           # Leerzeile

def format_shot(shot, number):
    """
    Konvertiert einen einzelnen Shot‑Trace in das gewünschte Text‑Layout.
    `shot.Trace` liefert eine Liste von Tripeln (t, x, y)
    """
    lines = [f"Shot #{number}"]
    for t, x, y in shot.Trace:
        # t, x, y liegen im SCATT‑Koordinaten­system (t in ms, x/y in mm)
        lines.append(f"{t:0.3f} x={x:+.2f} y={y:+.2f}")
    lines.append("")                      # Leerzeile nach jedem Schuss
    return lines

def main(path_scatt: pathlib.Path, path_out: pathlib.Path = None):
    if path_out is None:
        path_out = path_scatt.with_suffix(".txt")

    # ---------------- COM Initialisierung ----------------
    scatt = win32.Dispatch("ScattPro.Application")
    session = scatt.OpenSession(str(path_scatt))   # lädt die .scatt‑Datei

    # ---------------- Export erstellen -------------------
    textlines = []
    textlines += format_header(session)

    for idx, shot in enumerate(session.Shots, start=1):
        textlines += format_shot(shot, idx)

    # ---------------- Speichern --------------------------
    path_out.write_text("\n".join(textlines), encoding="utf‑8")
    print(f"Export fertig: {path_out}")

    # Wichtig: COM‑Objekt sauber schließen
    session.Close()
    scatt.Quit()

# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Aufruf: python scatt_exporter.py <datei.scatt> [<ausgabe.txt>]")
        sys.exit(1)

    main(pathlib.Path(sys.argv[1]),
         pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else None)
