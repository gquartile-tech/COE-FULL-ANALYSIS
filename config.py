# config.py — Google CoE shared constants
# This file is the shared contract used by all 5 pillars.
from __future__ import annotations

from dataclasses import dataclass

# ── Status constants (locked) ─────────────────────────────────────────────────
STATUS_OK      = "OK"
STATUS_FLAG    = "FLAG"
STATUS_PARTIAL = "PARTIAL"


# ── ControlResult ─────────────────────────────────────────────────────────────
@dataclass(frozen=True)
class ControlResult:
    status: str
    what:   str = ""   # What We Saw — observed data, always includes real numbers
    why:    str = ""   # Why It Matters — one action-oriented sentence
    wysd:   str = ""   # What You Should Do — remediation step (optional, Framework uses this)
