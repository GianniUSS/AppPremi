from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Dict, Optional

import tkinter as tk
from tkinter import ttk, messagebox

from ui_components import create_button

from database import (
    delete_fascia_premio,
    fetch_fasce_premi,
    insert_fascia_premio,
    update_fascia_premio,
)


DEFAULT_UNITS: Dict[str, tuple[str, str]] = {
    "PICKING": ("COLLI/h", "€/COL"),
    "CARRELLISTI": ("Mov/h", "€/Plt"),
    "RICEVITORI": ("Plt/h", "€/gg"),
    "DOPPIA_SPUNTA": ("Colli/h", "€/collo"),
}


@dataclass
class FasciaPremio:
    id: Optional[int]
    tipo_attivita: str
    valore_riferimento: Decimal
    valore_premio: Decimal
    unita_riferimento: str
    unita_premio: str
    note: Optional[str]

    @staticmethod
    def from_row(row: Dict[str, object]) -> "FasciaPremio":
        row_id = row.get("id")
        parsed_id: Optional[int]
        if row_id is None:
            parsed_id = None
        else:
            try:
                parsed_id = int(row_id)  # type: ignore[arg-type]
            except (TypeError, ValueError):
                parsed_id = None
        return FasciaPremio(
            id=parsed_id,
            tipo_attivita=str(row.get("tipo_attivita", "")),
            valore_riferimento=Decimal(str(row.get("valore_riferimento", "0"))),
            valore_premio=Decimal(str(row.get("valore_premio", "0"))),
            unita_riferimento=str(row.get("unita_riferimento", "")),
            unita_premio=str(row.get("unita_premio", "")),
            note=str(row.get("note")) if row.get("note") is not None else None,
        )


class FascePremiView(ttk.Frame):
    """Gestione fasce premi per attività."""

    def __init__(self, parent: tk.Widget):
        super().__init__(parent)
        self.pack(fill=tk.BOTH, expand=True)

        self.tipo_var = tk.StringVar(value="TUTTI")

        self._build_header()
        self._build_tree()
        self._build_footer()
        self._load_data()

    def _build_header(self) -> None:
        header = ttk.Frame(self)
        header.pack(fill=tk.X, padx=16, pady=(16, 8))

        ttk.Label(header, text="Tipo attività:").pack(side=tk.LEFT)

        tipi = ["TUTTI", *DEFAULT_UNITS.keys()]
        tipo_combo = ttk.Combobox(
            header,
            textvariable=self.tipo_var,
            values=tipi,
            state="readonly",
            width=18,
        )
        tipo_combo.pack(side=tk.LEFT, padx=(8, 16))
        tipo_combo.bind("<<ComboboxSelected>>", lambda _event: self._load_data())

        create_button(
            header,
            text="🔄 Ricarica",
            command=self._load_data,
            variant="primary",
            width=12,
        ).pack(side=tk.LEFT)

    def _build_tree(self) -> None:
        columns = ("tipo", "valore", "unita_valore", "premio", "unita_premio", "note")

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=8)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            height=12,
        )

        self.tree.heading("tipo", text="Tipo")
        self.tree.heading("valore", text="Valore riferimento")
        self.tree.heading("unita_valore", text="Unità riferimento")
        self.tree.heading("premio", text="Premio")
        self.tree.heading("unita_premio", text="Unità premio")
        self.tree.heading("note", text="Note")

        self.tree.column("tipo", width=140, anchor=tk.CENTER)
        self.tree.column("valore", width=150, anchor=tk.CENTER)
        self.tree.column("unita_valore", width=140, anchor=tk.CENTER)
        self.tree.column("premio", width=120, anchor=tk.CENTER)
        self.tree.column("unita_premio", width=120, anchor=tk.CENTER)
        self.tree.column("note", width=200, anchor=tk.W)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_footer(self) -> None:
        footer = ttk.Frame(self)
        footer.pack(fill=tk.X, padx=16, pady=(8, 16))

        create_button(
            footer,
            text="➕ Aggiungi",
            command=self._on_add,
            variant="primary",
            width=12,
        ).pack(side=tk.LEFT)
        create_button(
            footer,
            text="✏️ Modifica",
            command=self._on_edit,
            variant="primary",
            width=12,
        ).pack(side=tk.LEFT, padx=8)
        create_button(
            footer,
            text="🗑️ Elimina",
            command=self._on_delete,
            variant="danger",
            width=12,
        ).pack(side=tk.LEFT)

        ttk.Label(
            footer,
            text="Gestisci fasce di premio per ogni attività. Le modifiche sono salvate immediatamente.",
        ).pack(side=tk.RIGHT)

    def _load_data(self) -> None:
        tipo = self.tipo_var.get()
        rows = fetch_fasce_premi(None if tipo == "TUTTI" else tipo)

        self.tree.delete(*self.tree.get_children())

        for row in rows:
            fascia = FasciaPremio.from_row(row)
            self.tree.insert(
                "",
                tk.END,
                iid=str(fascia.id or ""),
                values=(
                    fascia.tipo_attivita,
                    f"{fascia.valore_riferimento:.2f}",
                    fascia.unita_riferimento,
                    f"{fascia.valore_premio:.5f}",
                    fascia.unita_premio,
                    fascia.note or "",
                ),
            )

    def _get_selected_fascia(self) -> Optional[FasciaPremio]:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Selezione mancante", "Seleziona una fascia per continuare.")
            return None
        item_id = selection[0]
        values = self.tree.item(item_id, "values")
        if not values:
            return None
        return FasciaPremio(
            id=int(item_id),
            tipo_attivita=str(values[0]),
            valore_riferimento=Decimal(str(values[1])),
            valore_premio=Decimal(str(values[3])),
            unita_riferimento=str(values[2]),
            unita_premio=str(values[4]),
            note=str(values[5]) if values[5] else None,
        )

    def _on_add(self) -> None:
        dialog = FasciaDialog(self, title="Aggiungi fascia")
        self.wait_window(dialog)
        if dialog.result:
            result = dialog.result
            insert_fascia_premio(
                result.tipo_attivita,
                float(result.valore_riferimento),
                float(result.valore_premio),
                result.unita_riferimento,
                result.unita_premio,
                result.note,
            )
            self._load_data()

    def _on_edit(self) -> None:
        fascia = self._get_selected_fascia()
        if not fascia:
            return
        dialog = FasciaDialog(self, title="Modifica fascia", fascia=fascia)
        self.wait_window(dialog)
        if dialog.result:
            if fascia.id is None:
                messagebox.showerror(
                    "Aggiornamento non riuscito",
                    "Impossibile aggiornare la fascia selezionata.",
                )
                return
            result = dialog.result
            update_fascia_premio(
                fascia.id,
                result.tipo_attivita,
                float(result.valore_riferimento),
                float(result.valore_premio),
                result.unita_riferimento,
                result.unita_premio,
                result.note,
            )
            self._load_data()

    def _on_delete(self) -> None:
        fascia = self._get_selected_fascia()
        if not fascia:
            return
        if messagebox.askyesno("Conferma eliminazione", "Vuoi eliminare la fascia selezionata?"):
            if fascia.id is None:
                messagebox.showerror(
                    "Eliminazione non riuscita",
                    "Impossibile eliminare la fascia selezionata.",
                )
                return
            delete_fascia_premio(fascia.id)
            self._load_data()


class FasciaDialog(tk.Toplevel):
    """Finestra di dialogo per aggiunta/modifica fascia premio."""

    def __init__(
        self,
        parent: FascePremiView,
        title: str,
        fascia: Optional[FasciaPremio] = None,
    ):
        super().__init__(parent)
        self.title(title)
        toplevel = parent.winfo_toplevel()
        self.transient(toplevel)
        self.grab_set()

        self._editing_existing = fascia is not None
        self.result: Optional[FasciaPremio] = None
        self._build_ui(fascia)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.resizable(False, False)
        self.wait_visibility()
        self.focus()

    def _build_ui(self, fascia: Optional[FasciaPremio]) -> None:
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

        self.tipo_var = tk.StringVar(value=fascia.tipo_attivita if fascia else "PICKING")
        self.valore_var = tk.StringVar(
            value=f"{fascia.valore_riferimento}" if fascia else ""
        )
        self.unita_valore_var = tk.StringVar(
            value=fascia.unita_riferimento if fascia else DEFAULT_UNITS["PICKING"][0]
        )
        self.premio_var = tk.StringVar(value=f"{fascia.valore_premio}" if fascia else "")
        self.unita_premio_var = tk.StringVar(
            value=fascia.unita_premio if fascia else DEFAULT_UNITS["PICKING"][1]
        )
        note_value = fascia.note if fascia else ""
        self.note_var = tk.StringVar(value=note_value or "")

        ttk.Label(frame, text="Tipo attività").grid(row=0, column=0, sticky=tk.W)
        tipo_combo = ttk.Combobox(
            frame,
            textvariable=self.tipo_var,
            values=list(DEFAULT_UNITS.keys()),
            state="readonly",
            width=22,
        )
        tipo_combo.grid(row=0, column=1, sticky=tk.W, pady=4)
        tipo_combo.bind("<<ComboboxSelected>>", self._on_tipo_change)

        ttk.Label(frame, text="Valore riferimento").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.valore_var, width=24).grid(
            row=1,
            column=1,
            sticky=tk.W,
            pady=4,
        )

        ttk.Label(frame, text="Unità riferimento").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.unita_valore_var, width=24).grid(
            row=2,
            column=1,
            sticky=tk.W,
            pady=4,
        )

        ttk.Label(frame, text="Premio").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.premio_var, width=24).grid(
            row=3,
            column=1,
            sticky=tk.W,
            pady=4,
        )

        ttk.Label(frame, text="Unità premio").grid(row=4, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.unita_premio_var, width=24).grid(
            row=4,
            column=1,
            sticky=tk.W,
            pady=4,
        )

        ttk.Label(frame, text="Note").grid(row=5, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.note_var, width=24).grid(
            row=5,
            column=1,
            sticky=tk.W,
            pady=4,
        )

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=(12, 0))

        create_button(
            button_frame,
            text="Annulla",
            command=self.destroy,
            variant="secondary",
            width=12,
        ).pack(side=tk.RIGHT)
        create_button(
            button_frame,
            text="Salva",
            command=self._on_save,
            variant="primary",
            width=12,
        ).pack(side=tk.RIGHT, padx=8)

        frame.columnconfigure(1, weight=1)

        # Aggiorna unità di default per le nuove fasce
        if not self._editing_existing:
            self._apply_default_units(self.tipo_var.get(), force=True)

    def _on_tipo_change(self, _event: object) -> None:
        self._apply_default_units(self.tipo_var.get(), force=True)

    def _apply_default_units(self, tipo: str, force: bool = False) -> None:
        if tipo in DEFAULT_UNITS:
            unita_val, unita_premio = DEFAULT_UNITS[tipo]
            if force or not self._editing_existing:
                self.unita_valore_var.set(unita_val)
                self.unita_premio_var.set(unita_premio)

    def _on_save(self) -> None:
        try:
            valore_riferimento = Decimal(self.valore_var.get().replace(",", "."))
            valore_premio = Decimal(self.premio_var.get().replace(",", "."))
        except InvalidOperation:
            messagebox.showerror("Valori non validi", "Inserisci numeri validi per valore e premio.")
            return

        tipo = self.tipo_var.get()
        if not tipo:
            messagebox.showerror("Tipo mancante", "Seleziona il tipo di attività.")
            return

        unita_val = self.unita_valore_var.get() or DEFAULT_UNITS.get(tipo, ("", ""))[0]
        unita_premio = self.unita_premio_var.get() or DEFAULT_UNITS.get(tipo, ("", ""))[1]

        self.result = FasciaPremio(
            id=None,
            tipo_attivita=tipo,
            valore_riferimento=valore_riferimento,
            valore_premio=valore_premio,
            unita_riferimento=unita_val,
            unita_premio=unita_premio,
            note=self.note_var.get() or None,
        )
        self.destroy()
