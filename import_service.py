"""
Servizio principale per l'importazione dei dati Excel nel database.
"""
from pathlib import Path
from typing import Optional
import datetime
import time
import tkinter as tk
from tkinter import messagebox, IntVar
from tkinter import ttk

from database import ensure_table_and_indexes, insert_batch_data, update_penalita_picking
from parsers import (
    parse_preparatori,
    parse_carrelisti,
    parse_ricevitori,
    parse_doppia_spunta,
    DoppiaSpuntaResult,
)
from utils import prepare_dataframe_for_db


class ImportService:
    """Servizio che gestisce la logica di importazione dei dati."""

    def import_excel(
        self,
        file_path: str,
        tipo: str,
        data_rif: Optional[datetime.date],
        progress_var: IntVar,
        status_label: ttk.Label,
        root: tk.Tk,
    ) -> None:
        """
        Esegue l'import del file Excel nel DB, scegliendo il parser per tipo attività.
        
        Args:
            file_path: Percorso del file Excel
            tipo: Tipo di attività (Preparatori/Carrellisti/Ricevitori)
            data_rif: Data di riferimento (opzionale, richiesta per Carrellisti)
            progress_var: Variabile per la progress bar
            status_label: Label per messaggi di stato
            root: Finestra principale per gli aggiornamenti UI
        """
        start_time = time.time()
        try:
            self._reset_ui(progress_var, status_label, root)
            ensure_table_and_indexes()

            # Validazione file
            if not self._validate_file(file_path):
                messagebox.showwarning("Attenzione", "Seleziona un file Excel valido.")
                return

            # Parsing file
            elapsed = time.time() - start_time
            self._update_status(f"Lettura file... ({elapsed:.1f}s)", 10, progress_var, status_label, root)
            parse_output = self._parse_file(file_path, tipo, data_rif)
            penalita_picking_df = None

            if isinstance(parse_output, DoppiaSpuntaResult):
                penalita_picking_df = parse_output.penalita_picking
                df_grouped = parse_output.records
            else:
                df_grouped = parse_output

            if df_grouped.empty:
                self._show_no_data_warning(status_label)
                return

            # Preparazione dati
            values = prepare_dataframe_for_db(df_grouped)

            # Inserimento in database
            elapsed = time.time() - start_time
            self._update_status(f"Scrittura su database... ({elapsed:.1f}s)", 60, progress_var, status_label, root)
            records_count = insert_batch_data(values)

            # Aggiornamento penalità per attività PICKING
            if penalita_picking_df is not None and not penalita_picking_df.empty:
                elapsed = time.time() - start_time
                self._update_status(
                    f"Aggiornamento penalità PICKING... ({elapsed:.1f}s)",
                    80,
                    progress_var,
                    status_label,
                    root,
                )
                update_values = [
                    (
                        row["data"],
                        str(row["codice_preparatore"]),
                        int(row["penalita"]),
                    )
                    for _, row in penalita_picking_df.iterrows()
                ]
                updated = update_penalita_picking(update_values)
                if updated < len(update_values):
                    print(
                        f"⚠️ Penalità PICKING aggiornate parzialmente: {updated}/{len(update_values)}"
                    )
                else:
                    print(
                        f"✓ Penalità PICKING aggiornate: {updated}/{len(update_values)}"
                    )

            # Completamento
            elapsed_total = time.time() - start_time
            self._show_success(records_count, elapsed_total, progress_var, status_label, root)

        except Exception as e:
            self._show_error(str(e), status_label)

    def _reset_ui(self, progress_var: IntVar, status_label: ttk.Label, root: tk.Tk):
        """Resetta l'interfaccia utente all'inizio dell'importazione."""
        status_label.config(text="")
        progress_var.set(0)
        root.update_idletasks()

    def _validate_file(self, file_path: str) -> bool:
        """Valida che il file esista e sia accessibile."""
        file_path = (file_path or "").strip()
        return bool(file_path and Path(file_path).exists())

    def _parse_file(self, file_path: str, tipo: str, data_rif: Optional[datetime.date]):
        """
        Seleziona e applica il parser appropriato basato sul tipo di attività.
        
        Returns:
            DataFrame pandas con i dati parsati oppure DoppiaSpuntaResult
        """
        if "Preparatori" in tipo:
            return parse_preparatori(file_path)
        elif "Carrellisti" in tipo:
            # I Carrellisti ora hanno la colonna Data nel file, quindi data_rif è opzionale
            # Se non fornita, verrà usata quella dal file
            if not data_rif:
                data_rif = datetime.date.today()  # Fallback, ma il parser userà la data dal file
            return parse_carrelisti(file_path, data_rif)
        elif "Doppia" in tipo:
            return parse_doppia_spunta(file_path)
        else:
            return parse_ricevitori(file_path)

    def _update_status(
        self, 
        message: str, 
        progress: int, 
        progress_var: IntVar, 
        status_label: ttk.Label, 
        root: tk.Tk
    ):
        """Aggiorna lo stato dell'interfaccia."""
        status_label.config(text=message)
        progress_var.set(progress)
        root.update_idletasks()

    def _show_no_data_warning(self, status_label: ttk.Label):
        """Mostra avviso per dati non trovati."""
        status_label.config(text="Nessun dato importabile.")
        messagebox.showwarning("Attenzione", "Nessun dato valido trovato nel file selezionato.")

    def _show_success(
        self, 
        records_count: int, 
        elapsed_seconds: float,
        progress_var: IntVar, 
        status_label: ttk.Label, 
        root: tk.Tk
    ):
        """Mostra messaggio di successo."""
        progress_var.set(100)
        root.update_idletasks()
        minutes = int(elapsed_seconds // 60)
        seconds = int(elapsed_seconds % 60)
        time_str = f"{minutes}m {seconds}s" if minutes > 0 else f"{seconds}s"
        status_label.config(text=f"Importazione completata ✅ ({time_str})")
        messagebox.showinfo("Successo", f"Importazione completata in {time_str}\n\nRecord importati: {records_count}")

    def _show_error(self, error_message: str, status_label: ttk.Label):
        """Mostra messaggio di errore."""
        status_label.config(text="Errore ❌")
        messagebox.showerror("Errore durante l'importazione", error_message)