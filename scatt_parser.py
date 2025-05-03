# scatt_parser.py – clean & parameterised (2025‑05‑03)
"""Pure‑Python parser for SCATT *.scatt* files (OLE Compound Binary format).

Core function
-------------
parse_scatt(file_obj, sample_rate=120.0) -> (shots, meta)
    *shots* – list[pandas.DataFrame] with columns t_s, x_mm, y_mm
    *meta*  – dict(shooter, n_shots)

Sample‑rate can be 60 Hz, 120 Hz, etc. The caller must pick the correct rate.
"""
from __future__ import annotations

import io, struct
from typing import List, Dict, Tuple

import numpy as np
import olefile  # type: ignore
import pandas as pd

__all__ = ["parse_scatt", "parse_scatt_long"]

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def _ole_date_to_unix_seconds(ole_date: float) -> float:
    return (ole_date - 25569.0) * 86400.0


def _decode_utf16_maybe(data: bytes) -> str:
    if len(data) % 2:
        data += b"\x00"
    try:
        return data.decode("utf-16-le", errors="strict").rstrip("\x00")
    except UnicodeDecodeError:
        return data.decode("latin-1", errors="replace")

# -----------------------------------------------------------------------------
# Core parser
# -----------------------------------------------------------------------------

def parse_scatt(file_obj: io.BufferedIOBase, sample_rate: float = 120.0
                ) -> Tuple[List[pd.DataFrame], Dict]:
    with olefile.OleFileIO(file_obj) as ole:
        raw = ole.openstream("Contents").read()

    n_shots, hdr_len = struct.unpack_from("<II", raw, 0)
    pos = hdr_len

    terminator = raw.find(b"\x00\x00", pos)
    name_end = terminator if terminator != -1 else pos + 40
    shooter = _decode_utf16_maybe(raw[pos:name_end])
    pos = name_end + 2 if terminator != -1 else name_end

    shots: List[pd.DataFrame] = []
    scale_mm = 0.01

    for _ in range(n_shots):
        # ---- per‑shot header ------------------------------------------------
        if pos + 16 > len(raw):
            break  # truncated file
        # read: 4 B n_samples, 4 B dummy, 8 B FILETIME (ignored)
        n_samples, _dummy, _filetime = struct.unpack_from("<IIQ", raw, pos)
        pos += 16

        # ensure we don't read past EOF
        if pos + 2 * n_samples > len(raw):
            n_samples = (len(raw) - pos) // 2

        # ---- trace bytes ----------------------------------------------------
        block = np.frombuffer(raw, np.int8, count=2 * n_samples, offset=pos)
        pos += 2 * n_samples

        # layout autodetect (see comment above)
        xA, yA = block[:n_samples], block[n_samples:]
        if np.count_nonzero(yA) < n_samples * 0.05 or np.count_nonzero(xA) < n_samples * 0.05:
            x_deltas, y_deltas = block[0::2], block[1::2]
        else:
            x_deltas, y_deltas = xA, yA

        xs = np.cumsum(x_deltas, dtype=np.float32) * scale_mm
        ys = np.cumsum(y_deltas, dtype=np.float32) * scale_mm
        ts = np.arange(n_samples, dtype=np.float32) / sample_rate  # relative time (0 = shot)

        shots.append(pd.DataFrame({"t_s": ts, "x_mm": xs, "y_mm": ys}))

    return shots, {"shooter": shooter, "n_shots": len(shots)}

# -----------------------------------------------------------------------------
# Convenience
# -----------------------------------------------------------------------------

def parse_scatt_long(file_obj: io.BufferedIOBase, sample_rate: float = 120.0
                     ) -> pd.DataFrame:
    shots, _ = parse_scatt(file_obj, sample_rate)
    return pd.concat(
        (df.assign(shot=i + 1) for i, df in enumerate(shots)), ignore_index=True
    )
