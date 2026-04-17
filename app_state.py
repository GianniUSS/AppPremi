"""
Stato globale dell'applicazione — periodo di lavoro condiviso tra tutte le schermate.
Persiste su disco in un file JSON nella cartella utente.
"""
import datetime
import json
import os
from pathlib import Path
from typing import Tuple

_STATE_FILE = Path.home() / ".appeden_state.json"

_MESI = [
    "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre",
]

# Periodo corrente in memoria
_today = datetime.date.today()
_periodo_anno: int = _today.year
_periodo_mese: int = _today.month   # 1-12


def _load() -> None:
    """Carica il periodo dal file di stato se esiste."""
    global _periodo_anno, _periodo_mese
    try:
        if _STATE_FILE.exists():
            data = json.loads(_STATE_FILE.read_text(encoding="utf-8"))
            a = int(data.get("anno", _periodo_anno))
            m = int(data.get("mese", _periodo_mese))
            if 2000 <= a <= 2100 and 1 <= m <= 12:
                _periodo_anno = a
                _periodo_mese = m
    except Exception:
        pass


def _save() -> None:
    """Salva il periodo corrente su disco."""
    try:
        _STATE_FILE.write_text(
            json.dumps({"anno": _periodo_anno, "mese": _periodo_mese}),
            encoding="utf-8",
        )
    except Exception:
        pass


def get_periodo() -> Tuple[int, int]:
    """Restituisce (anno, mese) del periodo globale corrente."""
    return _periodo_anno, _periodo_mese


def get_anno() -> int:
    return _periodo_anno


def get_mese() -> int:
    """Mese 1-12."""
    return _periodo_mese


def get_mese_label() -> str:
    """Nome del mese (es. 'Febbraio')."""
    return _MESI[_periodo_mese - 1]


def set_periodo(anno: int, mese: int) -> None:
    """Imposta il periodo globale e lo persiste su disco."""
    global _periodo_anno, _periodo_mese
    _periodo_anno = anno
    _periodo_mese = mese
    _save()


# Carica all'import
_load()
