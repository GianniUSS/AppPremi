from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import List, Optional

import tkinter as tk
from tkinter import ttk, messagebox

from ui_components import create_button

from database import (
    delete_peso_movimento,
    fetch_pesi_movimenti,
    insert_peso_movimento,
    update_peso_movimento,
)


ATTIVITA_SUPPORTATE: List[str] = [
    "CARRELLISTI",
    "PICKING",
    "RICEVITORI",
    "DOPPIA_SPUNTA",
]


@dataclass
class PesoMovimento:
    id: Optional[int]
    tipo_attivita: str
    tipo: str
    peso: Decimal
    note: Optional[str]

    @staticmethod
    def from_row(row: dict) -> "PesoMovimento":
        row_id = row.get("id")
        parsed_id: Optional[int]
        if row_id is None:
            parsed_id = None
        else:
            try:
                parsed_id = int(row_id)  # type: ignore[arg-type]
            except (TypeError, ValueError):
                parsed_id = None
        return PesoMovimento(
            id=parsed_id,
            tipo_attivita=str(row.get("tipo_attivita", "")),
            tipo=str(row.get("tipo", "")),
            peso=Decimal(str(row.get("peso", "0"))),
            note=str(row.get("note")) if row.get("note") is not None else None,
        )


class PesoMovimentiView(ttk.Frame):
    """Gestione del peso dei movimenti per i carrellisti."""

    def __init__(self, parent: tk.Widget):
        super().__init__(parent)
        self.pack(fill=tk.BOTH, expand=True)

        self.attivita_var = tk.StringVar(value="CARRELLISTI")

        self._build_header()
        self._build_tree()
        self._build_footer()
        self._load_data()

    def _build_header(self) -> None:
        header = ttk.Frame(self)
        header.pack(fill=tk.X, padx=16, pady=(16, 8))

        ttk.Label(header, text="Tipo attività:").pack(side=tk.LEFT)

        valori_combo = ["TUTTE", *ATTIVITA_SUPPORTATE]
        combo = ttk.Combobox(
            header,
            textvariable=self.attivita_var,
            values=valori_combo,
            state="readonly",
            width=18,
        )
        combo.pack(side=tk.LEFT, padx=(8, 16))
        combo.bind("<<ComboboxSelected>>", lambda _evt: self._load_data())

        create_button(
            header,
            text="🔄 Ricarica",
            command=self._load_data,
            variant="secondary",
            width=12,
        ).pack(side=tk.LEFT)

    def _build_tree(self) -> None:
        columns = ("attivita", "tipo", "peso", "note")

        frame = ttk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))

        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=12)
        self.tree.heading("attivita", text="Attività")
        self.tree.heading("tipo", text="Tipo")
        self.tree.heading("peso", text="Peso")
        self.tree.heading("note", text="Note")

        self.tree.column("attivita", width=160, anchor=tk.CENTER)
        self.tree.column("tipo", width=140, anchor=tk.CENTER)
        self.tree.column("peso", width=140, anchor=tk.CENTER)
        self.tree.column("note", width=240, anchor=tk.W)

        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_footer(self) -> None:
        footer = ttk.Frame(self)
        footer.pack(fill=tk.X, padx=16, pady=(0, 16))

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
            variant="secondary",
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
            text="Gestisci il peso per tipo movimento. Valori usati nel calcolo premi.",
        ).pack(side=tk.RIGHT)

    def _load_data(self) -> None:
        selected = self.attivita_var.get()
        filtro = None if selected in ("", "TUTTE") else selected
        rows = fetch_pesi_movimenti(filtro)
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            peso = PesoMovimento.from_row(row)
            self.tree.insert(
                "",
                tk.END,
                iid=str(peso.id or ""),
                values=(
                    peso.tipo_attivita,
                    peso.tipo,
                    f"{peso.peso:.3f}",
                    peso.note or "",
                ),
            )

    def _get_selected(self) -> Optional[PesoMovimento]:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Selezione mancante", "Seleziona una riga per continuare.")
            return None
        item_id = selection[0]
        values = self.tree.item(item_id, "values")
        if not values:
            return None
        return PesoMovimento(
            id=int(item_id),
            tipo_attivita=str(values[0]),
            tipo=str(values[0]),
            peso=Decimal(str(values[1])),
            note=str(values[2]) if values[2] else None,
        )

    def _on_add(self) -> None:
        dialog = PesoDialog(self, title="Aggiungi peso", attivita_default=self.attivita_var.get())
        self.wait_window(dialog)
        if dialog.result:
            result = dialog.result
            insert_peso_movimento(
                result.tipo_attivita,
                result.tipo,
                float(result.peso),
                result.note,
            )
            self._load_data()

    def _on_edit(self) -> None:
        current = self._get_selected()
        if not current:
            return
        dialog = PesoDialog(self, title="Modifica peso", peso=current)
        self.wait_window(dialog)
        if dialog.result:
            if current.id is None:
                messagebox.showerror("Aggiornamento non riuscito", "Record non valido.")
                return
            result = dialog.result
            update_peso_movimento(
                current.id,
                result.tipo_attivita,
                result.tipo,
                float(result.peso),
                result.note,
            )
            self._load_data()

    def _on_delete(self) -> None:
        current = self._get_selected()
        if not current:
            return
        if messagebox.askyesno("Conferma eliminazione", "Eliminare il peso selezionato?"):
            if current.id is None:
                messagebox.showerror("Eliminazione non riuscita", "Record non valido.")
                return
            delete_peso_movimento(current.id)
            self._load_data()


class PesoDialog(tk.Toplevel):
    """Dialogo per aggiungere o modificare i pesi movimento."""

    def __init__(
        self,
        parent: PesoMovimentiView,
        title: str,
        peso: Optional[PesoMovimento] = None,
        attivita_default: str = "",
    ):
        super().__init__(parent)
        self.title(title)
        toplevel = parent.winfo_toplevel()
        self.transient(toplevel)
        self.grab_set()

        self.result: Optional[PesoMovimento] = None
        self._build_ui(peso, attivita_default)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.resizable(False, False)
        self.wait_visibility()
        self.focus()

    def _build_ui(self, peso: Optional[PesoMovimento], attivita_default: str) -> None:
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

        default_attivita = attivita_default if attivita_default not in ("", "TUTTE") else "CARRELLISTI"
        self.attivita_var = tk.StringVar(value=peso.tipo_attivita if peso else default_attivita)
        self.tipo_var = tk.StringVar(value=peso.tipo if peso else "")
        self.peso_var = tk.StringVar(value=f"{peso.peso}" if peso else "")
        self.note_var = tk.StringVar(value=(peso.note or "") if peso else "")

        ttk.Label(frame, text="Attività").grid(row=0, column=0, sticky=tk.W)
        attivita_combo = ttk.Combobox(
            frame,
            textvariable=self.attivita_var,
            values=ATTIVITA_SUPPORTATE,
            state="readonly",
            width=20,
        )
        attivita_combo.grid(row=0, column=1, sticky=tk.W, pady=4)

        ttk.Label(frame, text="Tipo").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.tipo_var, width=18).grid(
            row=1,
            column=1,
            sticky=tk.W,
            pady=4,
        )

        ttk.Label(frame, text="Peso").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.peso_var, width=18).grid(
            row=2,
            column=1,
            sticky=tk.W,
            pady=4,
        )

        ttk.Label(frame, text="Note").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.note_var, width=24).grid(
            row=3,
            column=1,
            sticky=tk.W,
            pady=4,
        )

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(12, 0))

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

    def _on_save(self) -> None:
        attivita = (self.attivita_var.get() or "").strip().upper()
        if not attivita:
            messagebox.showerror("Attività mancante", "Seleziona il tipo di attività.")
            return
        tipo = (self.tipo_var.get() or "").strip().upper()
        if not tipo:
            messagebox.showerror("Tipo mancante", "Inserisci il tipo di movimento.")
            return
        try:
            peso_value = Decimal(self.peso_var.get().replace(",", "."))
        except InvalidOperation:
            messagebox.showerror("Peso non valido", "Inserisci un numero valido per il peso.")
            return

        self.result = PesoMovimento(
            id=None,
            tipo_attivita=attivita,
            tipo=tipo,
            peso=peso_value,
            note=self.note_var.get().strip() or None,
        )
        self.destroy()
