from __future__ import annotations

import datetime
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Iterable, Optional

import tkinter as tk
from tkinter import ttk, messagebox

from ui_components import create_button

from database import (
    delete_malus_bonus,
    fetch_malus_bonus,
    get_malus_bonus,
    upsert_malus_bonus,
)


ITALIAN_MONTHS = [
    "",
    "Gennaio",
    "Febbraio",
    "Marzo",
    "Aprile",
    "Maggio",
    "Giugno",
    "Luglio",
    "Agosto",
    "Settembre",
    "Ottobre",
    "Novembre",
    "Dicembre",
]


AVAILABLE_BONUS_ACTIVITIES: dict[str, str] = {
    "PICKING": "Preparatori (PICKING)",
    "CARRELLISTI": "Carrellisti (CARRELLISTI)",
    "RICEVITORI": "Ricevitori",
    "DOPPIA_SPUNTA": "Doppia Spunta",
}


@dataclass
class MalusBonusRecord:
    id: Optional[int]
    anno: int
    mese: int
    importo_rotture: Decimal
    importo_differenze: Decimal
    soglia_rotture: Decimal
    soglia_differenze: Decimal
    attivita_bonus: tuple[str, ...]
    note: Optional[str]

    @property
    def totale(self) -> Decimal:
        return self.importo_rotture + self.importo_differenze

    @property
    def bonus_attivo(self) -> bool:
        return (
            self.importo_rotture <= self.soglia_rotture
            and self.importo_differenze <= self.soglia_differenze
            and bool(self.attivita_bonus)
        )

    @staticmethod
    def from_row(row: dict) -> "MalusBonusRecord":
        row_id = row.get("id")
        parsed_id = int(row_id) if row_id is not None else None
        raw_attivita = row.get("attivita_bonus")
        if raw_attivita is None:
            attivita: tuple[str, ...] = tuple(AVAILABLE_BONUS_ACTIVITIES.keys())
        else:
            cleaned = str(raw_attivita).strip()
            if not cleaned or cleaned.upper() == "NONE":
                attivita = tuple()
            else:
                attivita = tuple(
                    part.strip().upper()
                    for part in cleaned.split(",")
                    if part and part.strip()
                )
        fallback_soglia = Decimal(str(row.get("soglia_bonus", "2500")))
        soglia_rotture = row.get("soglia_rotture")
        soglia_differenze = row.get("soglia_differenze")
        return MalusBonusRecord(
            id=parsed_id,
            anno=int(row.get("anno", 0)),
            mese=int(row.get("mese", 0)),
            importo_rotture=Decimal(str(row.get("importo_rotture", "0"))),
            importo_differenze=Decimal(str(row.get("importo_differenze", "0"))),
            soglia_rotture=Decimal(str(soglia_rotture if soglia_rotture is not None else fallback_soglia)),
            soglia_differenze=Decimal(
                str(soglia_differenze if soglia_differenze is not None else fallback_soglia)
            ),
            attivita_bonus=attivita,
            note=str(row.get("note")) if row.get("note") is not None else None,
        )


class MalusBonusView(ttk.Frame):
    """Gestione mensile di malus/bonus (rotture e differenze inventariali)."""

    def __init__(self, parent: tk.Widget):
        super().__init__(parent)
        self.pack(fill=tk.BOTH, expand=True)

        today = datetime.date.today()
        self.anno_var = tk.StringVar(value=str(today.year))
        self.mese_var = tk.IntVar(value=today.month)

        self.rotture_var = tk.StringVar()
        self.differenze_var = tk.StringVar()
        self.soglia_rotture_var = tk.StringVar(value="2500")
        self.soglia_differenze_var = tk.StringVar(value="2500")
        self.note_var = tk.StringVar()
        self.activity_vars: dict[str, tk.BooleanVar] = {
            code: tk.BooleanVar(value=True)
            for code in AVAILABLE_BONUS_ACTIVITIES
        }

        self._build_header()
        self._build_form()
        self._build_footer()
        self._build_history()

        self._load_record()
        self._refresh_history()

    def _build_header(self) -> None:
        header = ttk.Frame(self)
        header.pack(fill=tk.X, padx=16, pady=(16, 8))

        ttk.Label(header, text="Anno").pack(side=tk.LEFT)
        years = [str(y) for y in range(datetime.date.today().year - 5, datetime.date.today().year + 2)]
        year_combo = ttk.Combobox(
            header,
            values=years,
            textvariable=self.anno_var,
            width=6,
            state="readonly",
        )
        year_combo.pack(side=tk.LEFT, padx=(4, 16))
        year_combo.bind("<<ComboboxSelected>>", lambda _evt: self._on_period_change())

        ttk.Label(header, text="Mese").pack(side=tk.LEFT)
        month_names = ITALIAN_MONTHS[1:]
        month_combo = ttk.Combobox(
            header,
            values=month_names,
            width=12,
            state="readonly",
        )
        month_combo.pack(side=tk.LEFT, padx=(4, 16))
        month_combo.bind("<<ComboboxSelected>>", lambda evt: self._on_month_combo(evt, month_names))
        month_combo.current(self.mese_var.get() - 1)

        create_button(
            header,
            text="🔄 Carica",
            command=self._load_record,
            variant="secondary",
            width=12,
        ).pack(side=tk.LEFT)

    def _build_form(self) -> None:
        form = ttk.Frame(self)
        form.pack(fill=tk.X, padx=16, pady=8)

        self._create_labeled_entry(form, "Importo rotture", self.rotture_var, row=0)
        self._create_labeled_entry(form, "Importo differenze inventariali", self.differenze_var, row=1)
        self._create_labeled_entry(form, "Soglia rotture", self.soglia_rotture_var, row=2)
        self._create_labeled_entry(form, "Soglia differenze", self.soglia_differenze_var, row=3)

        ttk.Label(form, text="Note opzionali").grid(row=4, column=0, sticky=tk.W, pady=4)
        ttk.Entry(form, textvariable=self.note_var, width=40).grid(row=4, column=1, sticky=tk.W, pady=4)

        ttk.Label(form, text="Attività interessate").grid(row=5, column=0, sticky=tk.W, pady=(8, 4))
        activities_frame = ttk.Frame(form)
        activities_frame.grid(row=5, column=1, sticky=tk.W, pady=(8, 4))
        for idx, (code, label) in enumerate(AVAILABLE_BONUS_ACTIVITIES.items()):
            chk = ttk.Checkbutton(
                activities_frame,
                text=label,
                variable=self.activity_vars[code],
                command=self._update_status,
            )
            chk.grid(row=idx, column=0, sticky=tk.W, pady=2)

        self.status_label = ttk.Label(form, font=("Segoe UI", 10, "bold"))
        self.status_label.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(12, 0))

        for var in (
            self.rotture_var,
            self.differenze_var,
            self.soglia_rotture_var,
            self.soglia_differenze_var,
        ):
            var.trace_add("write", lambda *_args: self._update_status())

    def _build_footer(self) -> None:
        footer = ttk.Frame(self)
        footer.pack(fill=tk.X, padx=16, pady=8)

        create_button(
            footer,
            text="💾 Salva",
            command=self._on_save,
            variant="primary",
            width=12,
        ).pack(side=tk.LEFT)
        create_button(
            footer,
            text="↺ Reset",
            command=self._on_reset,
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

    def _build_history(self) -> None:
        history_frame = ttk.LabelFrame(self, text="Storico valorizzazioni")
        history_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(8, 16))

        columns = (
            "anno",
            "mese",
            "rotture",
            "differenze",
            "soglia_rotture",
            "soglia_differenze",
            "totale",
            "bonus",
            "attivita",
            "note",
        )
        self.tree = ttk.Treeview(history_frame, columns=columns, show="headings", height=10)

        headings = {
            "anno": "Anno",
            "mese": "Mese",
            "rotture": "Rotture",
            "differenze": "Differenze",
            "soglia_rotture": "Soglia rotture",
            "soglia_differenze": "Soglia differenze",
            "totale": "Totale",
            "bonus": "Bonus",
            "attivita": "Attività bonus",
            "note": "Note",
        }
        widths = {
            "anno": 80,
            "mese": 120,
            "rotture": 120,
            "differenze": 140,
            "soglia_rotture": 140,
            "soglia_differenze": 160,
            "totale": 120,
            "bonus": 120,
            "attivita": 180,
            "note": 200,
        }
        for key in columns:
            self.tree.heading(key, text=headings[key])
            anchor = tk.W if key in {"note", "attivita"} else tk.CENTER
            self.tree.column(key, width=widths[key], anchor=anchor)

        vsb = ttk.Scrollbar(history_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_labeled_entry(self, parent: ttk.Frame, label: str, variable: tk.StringVar, row: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=4)
        ttk.Entry(parent, textvariable=variable, width=20).grid(row=row, column=1, sticky=tk.W, pady=4)

    def _set_activity_selection(self, activities: Optional[Iterable[str]] = None) -> None:
        if activities is None:
            selected = set(AVAILABLE_BONUS_ACTIVITIES.keys())
        else:
            selected = {code.strip().upper() for code in activities}
        for code, var in self.activity_vars.items():
            var.set(code in selected)

    def _get_selected_activities(self) -> list[str]:
        return [code for code, var in self.activity_vars.items() if var.get()]

    def _format_activity_labels(self, activities: Optional[Iterable[str]] = None) -> str:
        selected = [code.upper() for code in (activities or self._get_selected_activities())]
        if not selected:
            return ""

        selected_set = set(selected)
        labels = [
            AVAILABLE_BONUS_ACTIVITIES.get(code, code)
            for code in AVAILABLE_BONUS_ACTIVITIES
            if code in selected_set
        ]
        extras = [code for code in selected if code not in AVAILABLE_BONUS_ACTIVITIES]
        if extras:
            labels.extend(extras)
        return ", ".join(labels)

    def _on_month_combo(self, event: tk.Event, month_names: list[str]) -> None:
        index = event.widget.current()  # type: ignore[attr-defined]
        self.mese_var.set(index + 1)
        self._on_period_change()

    def _on_period_change(self) -> None:
        self._load_record()
        self._refresh_history()

    def _load_record(self) -> None:
        anno = int(self.anno_var.get())
        mese = self.mese_var.get()
        row = get_malus_bonus(anno, mese)
        if row:
            record = MalusBonusRecord.from_row(row)
            self.rotture_var.set(f"{record.importo_rotture:.2f}")
            self.differenze_var.set(f"{record.importo_differenze:.2f}")
            self.soglia_rotture_var.set(f"{record.soglia_rotture:.2f}")
            self.soglia_differenze_var.set(f"{record.soglia_differenze:.2f}")
            self.note_var.set(record.note or "")
            self._set_activity_selection(record.attivita_bonus)
        else:
            self.rotture_var.set("0")
            self.differenze_var.set("0")
            self.soglia_rotture_var.set("2500")
            self.soglia_differenze_var.set("2500")
            self.note_var.set("")
            self._set_activity_selection(None)
        self._update_status()

    def _refresh_history(self) -> None:
        self.tree.delete(*self.tree.get_children())
        records = fetch_malus_bonus(int(self.anno_var.get()))
        for row in records:
            record = MalusBonusRecord.from_row(row)
            month_name = ITALIAN_MONTHS[record.mese]
            bonus = "+15%" if record.bonus_attivo else "-"
            activity_label = self._format_activity_labels(record.attivita_bonus)
            self.tree.insert(
                "",
                tk.END,
                values=(
                    record.anno,
                    month_name,
                    f"{record.importo_rotture:.2f}",
                    f"{record.importo_differenze:.2f}",
                    f"{record.soglia_rotture:.2f}",
                    f"{record.soglia_differenze:.2f}",
                    f"{record.totale:.2f}",
                    bonus,
                    activity_label or "-",
                    record.note or "",
                ),
            )

    def _update_status(self) -> None:
        try:
            rotture = Decimal(self.rotture_var.get().replace(",", "."))
            differenze = Decimal(self.differenze_var.get().replace(",", "."))
            soglia_rotture = Decimal(self.soglia_rotture_var.get().replace(",", "."))
            soglia_differenze = Decimal(self.soglia_differenze_var.get().replace(",", "."))
        except InvalidOperation:
            self.status_label.config(text="Valori non validi", foreground="#d35400")
            return

        totale = rotture + differenze
        rotture_ok = rotture <= soglia_rotture
        differenze_ok = differenze <= soglia_differenze
        selected_codes = self._get_selected_activities()
        activities_line = self._format_activity_labels(selected_codes)

        if rotture_ok and differenze_ok:
            if selected_codes:
                self.status_label.config(
                    text=(
                        f"Bonus attivo: rotture {rotture:.2f} ≤ soglia rotture {soglia_rotture:.2f} "
                        f"e differenze {differenze:.2f} ≤ soglia differenze {soglia_differenze:.2f} → +15%"
                        f"\nAttività: {activities_line}"
                    ),
                    foreground="#2ecc71",
                )
            else:
                self.status_label.config(
                    text=(
                        "Soglia rispettata ma nessuna attività selezionata per il bonus."
                        "\nSeleziona almeno un'attività per applicare il +15%."
                    ),
                    foreground="#f39c12",
                )
        else:
            reasons = []
            if not rotture_ok:
                reasons.append(
                    f"rotture {rotture:.2f} > soglia rotture {soglia_rotture:.2f}"
                )
            if not differenze_ok:
                reasons.append(
                    f"differenze {differenze:.2f} > soglia differenze {soglia_differenze:.2f}"
                )
            details = ", ".join(reasons) if reasons else f"totale {totale:.2f}"
            self.status_label.config(
                text=(
                    f"Nessun bonus: {details}"
                    f"\nAttività: {activities_line or 'nessuna selezionata'}"
                ),
                foreground="#c0392b",
            )

    def _read_decimal(self, value: str) -> Decimal:
        return Decimal(value.replace(",", "."))

    def _on_save(self) -> None:
        try:
            rotture = self._read_decimal(self.rotture_var.get())
            differenze = self._read_decimal(self.differenze_var.get())
            soglia_rotture = self._read_decimal(self.soglia_rotture_var.get())
            soglia_differenze = self._read_decimal(self.soglia_differenze_var.get())
        except (InvalidOperation, AttributeError):
            messagebox.showerror("Valori non validi", "Usa numeri validi (puoi usare la virgola).")
            return

        anno = int(self.anno_var.get())
        mese = self.mese_var.get()
        note = self.note_var.get().strip() or None
        selected_codes = self._get_selected_activities()
        attivita_value = ",".join(selected_codes)

        upsert_malus_bonus(
            anno,
            mese,
            float(rotture),
            float(differenze),
            float(soglia_rotture),
            float(soglia_differenze),
            attivita_value,
            note,
        )
        messagebox.showinfo("Salvato", "Valori malus/bonus salvati correttamente.")
        self._update_status()
        self._refresh_history()

    def _on_reset(self) -> None:
        self._load_record()

    def _on_delete(self) -> None:
        anno = int(self.anno_var.get())
        mese = self.mese_var.get()
        row = get_malus_bonus(anno, mese)
        if not row:
            messagebox.showinfo("Nessun dato", "Non esiste un record da eliminare per il periodo selezionato.")
            return
        record_id = row.get("id")
        if record_id is None:
            messagebox.showerror("Errore", "Record non valido.")
            return
        if messagebox.askyesno("Confermi eliminazione?", "Eliminare i valori salvati per questo mese?"):
            delete_malus_bonus(int(record_id))
            self._load_record()
            self._refresh_history()

    def update(self) -> None:  # type: ignore[override]
        super().update()
        self._update_status()

def on_field_change(view: MalusBonusView) -> None:
    view._update_status()
