"""Utility per la gestione degli aggiornamenti dell'applicazione."""

from __future__ import annotations

import hashlib
import json
import os
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Callable, Optional, Tuple
from urllib.error import URLError
from urllib.request import urlopen

from version import APP_VERSION, UPDATE_MANIFEST_URL

# Dimensione blocco download (64 KiB)
_CHUNK_SIZE = 64 * 1024
# Cartella locale dove salvare gli aggiornamenti scaricati
_DOWNLOAD_DIR = Path.home() / "Downloads" / "AppEdenAggiornamenti"


@dataclass
class UpdateInfo:
    """Informazioni estratte dal manifesto remoto."""

    version: str
    release_notes: str
    download_url: str
    sha256: str
    file_name: str


def _parse_version(value: str) -> Tuple[int, ...]:
    """Converte una stringa versione in tupla di interi."""
    parts = []
    for item in value.split('.'):
        try:
            parts.append(int(item))
        except ValueError:
            # Ignora eventuali suffissi non numerici
            digits = ''.join(ch for ch in item if ch.isdigit())
            parts.append(int(digits) if digits else 0)
    return tuple(parts)


def is_newer(remote: str, local: str) -> bool:
    """Ritorna True se la versione remota è più nuova della locale."""
    return _parse_version(remote) > _parse_version(local)


def fetch_manifest(url: str = UPDATE_MANIFEST_URL) -> Tuple[Optional[UpdateInfo], Optional[str]]:
    """Scarica e interpreta il manifesto aggiornamenti.

    Ritorna una coppia (info, errore). Se avviene un errore info è None.
    """

    try:
        with urlopen(url, timeout=10) as response:
            content = response.read().decode('utf-8')
    except URLError as exc:
        return None, f"Impossibile recuperare il manifesto: {exc.reason}"
    except Exception as exc:  # pragma: no cover - protezione generale
        return None, f"Errore durante il download del manifesto: {exc}"

    try:
        payload = json.loads(content)
    except json.JSONDecodeError as exc:
        return None, f"Manifesto non valido: {exc}"

    latest_version = str(payload.get('latest_version', '')).strip()
    if not latest_version:
        return None, "Manifesto privo del campo 'latest_version'"

    download_url = str(payload.get('download_url', '')).strip()
    if not download_url:
        return None, "Manifesto privo del campo 'download_url'"

    release_notes = payload.get('release_notes', [])
    if isinstance(release_notes, list):
        release_notes = '\n'.join(f"• {note}" for note in release_notes)
    else:
        release_notes = str(release_notes)

    sha256 = str(payload.get('sha256', '')).strip()
    file_name = str(payload.get('file_name', '')).strip() or Path(download_url).name

    return UpdateInfo(
        version=latest_version,
        release_notes=release_notes,
        download_url=download_url,
        sha256=sha256,
        file_name=file_name,
    ), None


def _download_file(url: str, destination: Path, progress_cb: Callable[[int, Optional[int]], None]) -> None:
    """Scarica un file aggiornando la progress bar tramite callback."""
    with urlopen(url) as response:
        total = response.headers.get('Content-Length')
        total_bytes = int(total) if total is not None else None

        downloaded = 0
        with destination.open('wb') as fout:
            while True:
                chunk = response.read(_CHUNK_SIZE)
                if not chunk:
                    break
                fout.write(chunk)
                downloaded += len(chunk)
                progress_cb(downloaded, total_bytes)


def _compute_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open('rb') as fobj:
        for block in iter(lambda: fobj.read(_CHUNK_SIZE), b''):
            digest.update(block)
    return digest.hexdigest()


class UpdateDialog(tk.Toplevel):
    """Finestra modale per la gestione degli aggiornamenti."""

    def __init__(self, master: tk.Tk):
        super().__init__(master)
        self.title("Aggiornamenti disponibili")
        self.resizable(False, False)
        self.grab_set()  # Comportamento modale

        self._manifest: Optional[UpdateInfo] = None
        self._download_path: Optional[Path] = None

        # UI elements
        self.status_var = tk.StringVar(value="Ricerca aggiornamenti in corso...")
        self.progress_var = tk.DoubleVar(value=0.0)

        self._build_ui()

        # Avvia controllo in thread separato
        threading.Thread(target=self._check_for_updates, daemon=True).start()

    # ------------------------------------------------------------------ UI setup
    def _build_ui(self) -> None:
        padding = {"padx": 20, "pady": 10}

        tk.Label(self, textvariable=self.status_var, font=("Segoe UI", 11, "bold")).pack(**padding)

        self.notes_text = tk.Text(self, width=70, height=8, wrap="word", state="disabled")
        self.notes_text.pack(padx=20, pady=(0, 10))

        self.progress = ttk.Progressbar(self, orient="horizontal", length=320, mode="determinate")
        self.progress.pack(padx=20, pady=(0, 10))

        self.buttons_frame = tk.Frame(self)
        self.buttons_frame.pack(pady=(0, 20))

        self.install_button = ttk.Button(self.buttons_frame, text="Scarica e installa", command=self._on_install_clicked, state="disabled")
        self.install_button.grid(row=0, column=0, padx=5)

        self.open_folder_button = ttk.Button(self.buttons_frame, text="Apri cartella download", command=self._open_download_folder, state="disabled")
        self.open_folder_button.grid(row=0, column=1, padx=5)

        ttk.Button(self.buttons_frame, text="Chiudi", command=self.destroy).grid(row=0, column=2, padx=5)

    # ------------------------------------------------------------------ actions
    def _check_for_updates(self) -> None:
        info, error = fetch_manifest()
        if error:
            self._set_status(error)
            return

        if not info:
            self._set_status("Manifesto aggiornamenti non disponibile")
            return

        self._manifest = info

        if is_newer(info.version, APP_VERSION):
            message = f"Versione disponibile: {info.version}\nVersione installata: {APP_VERSION}\n\nNote di rilascio:\n{info.release_notes or 'N/A'}"
            self._set_status(message)
            self._enable_install()
        else:
            self._set_status(f"Sei già all'ultima versione ({APP_VERSION}).")

    def _set_status(self, message: str) -> None:
        def _update():
            self.status_var.set(message)
            self.notes_text.config(state="normal")
            self.notes_text.delete("1.0", tk.END)
            self.notes_text.insert(tk.END, message)
            self.notes_text.config(state="disabled")
        self.after(0, _update)

    def _enable_install(self) -> None:
        def _update():
            self.install_button.config(state="normal")
        self.after(0, _update)

    def _on_install_clicked(self) -> None:
        if not self._manifest:
            return

        self.install_button.config(state="disabled")
        self.progress_var.set(0.0)
        self.progress.config(mode="determinate", value=0)
        self._set_status("Download dell'aggiornamento in corso...")

        threading.Thread(target=self._download_update, daemon=True).start()

    def _download_update(self) -> None:
        assert self._manifest is not None
        info = self._manifest

        try:
            _DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
            destination = _DOWNLOAD_DIR / info.file_name

            def progress_cb(done: int, total: Optional[int]) -> None:
                if total:
                    percent = (done / total) * 100
                    self.after(0, lambda: self.progress.config(value=percent, maximum=100))
                else:
                    # Se la dimensione è sconosciuta usa barra indeterminata
                    self.after(0, lambda: self.progress.config(mode="indeterminate"))

            _download_file(info.download_url, destination, progress_cb)

            if info.sha256:
                computed = _compute_sha256(destination)
                if computed.lower() != info.sha256.lower():
                    raise ValueError("Checksum SHA256 non corrispondente")

        except Exception as exc:
            self._set_status(f"Errore durante il download: {exc}")
            self.after(0, lambda: self.install_button.config(state="normal"))
            return

        self._download_path = destination
        self.after(0, self._on_download_complete)

    def _on_download_complete(self) -> None:
        if not self._download_path:
            return

        self.progress.config(value=100, maximum=100, mode="determinate")

        message = (
            f"Aggiornamento scaricato correttamente in:\n{self._download_path}\n\n"
            "Chiudi l'applicazione corrente e avvia il nuovo eseguibile per completare l'aggiornamento."
        )
        self._set_status(message)
        self.open_folder_button.config(state="normal")

    def _open_download_folder(self) -> None:
        if not self._download_path:
            _DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
            target = _DOWNLOAD_DIR
        else:
            target = self._download_path.parent

        try:
            os.startfile(target)  # Disponibile su Windows
        except OSError:
            messagebox.showinfo("Cartella", f"Apri manualmente la cartella: {target}")


def open_update_dialog(master: tk.Tk) -> None:
    """Apre la finestra di gestione degli aggiornamenti."""
    UpdateDialog(master)

