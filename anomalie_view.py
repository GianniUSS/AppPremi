import datetime
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import DateEntry

import pandas as pd

from config import COLORS, FONTS
from database import (
    delete_anomalia,
    ensure_table_and_indexes,
    fetch_anomalie,
    fetch_report_templates,
    execute_custom_query,
    update_anomalia_stato,
)
from ui_components import create_button


EXPORTS_DIR = Path(__file__).resolve().parent / "exports"

ANOMALIA_TIPO_CHOICES: List[str] = [
    "CODICE_NON_ABBINATO",
    "ORE_SENZA_PRODUZIONE",
    "PRODUZIONE_SENZA_ORE",
    "DIFFERENZA_60_120",
    "DIFFERENZA_>120",
]

ANOMALIA_STATO_CHOICES: List[str] = [
    "Tutti",
    "APERTA",
    "VERIFICATA",
    "RISOLTA"
]

MONTH_CHOICES: List[Tuple[str, Optional[int]]] = [
    ("Tutti", None),
    ("Gennaio", 1),
    ("Febbraio", 2),
    ("Marzo", 3),
    ("Aprile", 4),
    ("Maggio", 5),
    ("Giugno", 6),
    ("Luglio", 7),
    ("Agosto", 8),
    ("Settembre", 9),
    ("Ottobre", 10),
    ("Novembre", 11),
    ("Dicembre", 12),
]

MONTH_LABEL_TO_VALUE: Dict[str, Optional[int]] = {
    label: value for label, value in MONTH_CHOICES
}


class AnomalieView:
    """Gestione delle anomalie con filtri periodo o intervallo date."""
    def __init__(self, parent: Optional[tk.Misc] = None, use_toplevel: bool = True) -> None:
        self.parent = parent
        self.use_toplevel = use_toplevel

        if self.use_toplevel:
            self.root = tk.Toplevel(parent) if parent is not None else tk.Tk()
            self.root.title("Gestione Anomalie")
            self.root.geometry("1150x620")
            self.root.configure(bg=COLORS["background"])
        else:
            if parent is None:
                raise ValueError("Per l'uso embedded è richiesto un parent.")
            self.root = parent

        self.dialog_parent = self.root if self.use_toplevel else self.root.winfo_toplevel()

        try:
            ensure_table_and_indexes()
        except Exception as exc:  # pragma: no cover - avvisa l'utente
            messagebox.showerror(
                "Errore",
                f"Impossibile inizializzare la tabella anomalie:\n{exc}",
                parent=self.dialog_parent,
            )
            if self.use_toplevel:
                self.root.destroy()
            raise

        today = datetime.date.today()
        self.today = today

        self.search_var = tk.StringVar()
        self.codice_var = tk.StringVar()
        self.tipo_var = tk.StringVar(value="Tutte")
        self.stato_var = tk.StringVar(value=ANOMALIA_STATO_CHOICES[0])
        self.dal_var = tk.StringVar()
        self.al_var = tk.StringVar()
        self.anno_var = tk.StringVar(value=str(today.year))
        default_month_label = next(
            (label for label, value in MONTH_CHOICES if value == today.month),
            MONTH_CHOICES[0][0],
        )
        self.mese_var = tk.StringVar(value=default_month_label)
        self.use_date_range_var = tk.BooleanVar(value=False)
        self.report_var = tk.StringVar()
        self.report_options: Dict[str, Dict[str, Any]] = {}

        self._setup_ui()
        self._load_report_templates()
        self._load_anomalie()

    def _setup_ui(self) -> None:
        style = ttk.Style(self.root)

        # Barra di ricerca in alto
        search_frame = tk.Frame(self.root, bg=COLORS["background"])
        search_frame.pack(fill="x", padx=16, pady=(12, 8))

        tk.Label(
            search_frame,
            text="🔍 Ricerca:",
            font=FONTS["big"],
            bg=COLORS["background"],
            foreground=COLORS["text_dark"],
        ).pack(side="left", padx=(0, 8))

        search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=FONTS["big"],
            width=40,
        )
        search_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        search_entry.bind("<Return>", lambda e: self._load_anomalie())

        create_button(
            search_frame,
            text="Cerca",
            command=self._load_anomalie,
            variant="primary",
            width=12,
        ).pack(side="left", padx=(0, 8))

        create_button(
            search_frame,
            text="Pulisci Filtri",
            command=self._reset_filters,
            variant="secondary",
            width=14,
        ).pack(side="left")

        # Filtro - layout più grande e omogeneo
        filter_frame = tk.Frame(self.root, bg=COLORS["background"], bd=2, relief="groove")
        filter_frame.pack(fill="x", padx=16, pady=(0, 0), ipady=12)

        # Colonne: i controlli si espandono, le etichette restano compatte
        control_columns = {1, 3, 5, 7, 9, 12, 14}
        for col in range(15):
            weight = 1 if col in control_columns else 0
            filter_frame.grid_columnconfigure(col, weight=weight)

        # Tipo - Listbox per selezione multipla
        tk.Label(filter_frame, text="Tipo:", font=FONTS["label"], bg=COLORS["background"], anchor="w").grid(row=0, column=0, sticky="nw", padx=(0, 6), pady=10)
        tipo_frame = tk.Frame(filter_frame, bg=COLORS["background"])
        tipo_frame.grid(row=0, column=1, sticky="ew", padx=(0, 16), pady=10)
        
        tipo_scrollbar = tk.Scrollbar(tipo_frame, orient="vertical")
        self.tipo_listbox = tk.Listbox(
            tipo_frame,
            selectmode="multiple",
            height=4,
            font=FONTS["input"],
            yscrollcommand=tipo_scrollbar.set,
            exportselection=False
        )
        tipo_scrollbar.config(command=self.tipo_listbox.yview)
        self.tipo_listbox.pack(side="left", fill="both", expand=True)
        tipo_scrollbar.pack(side="right", fill="y")
        
        # Popola la listbox
        for tipo in ANOMALIA_TIPO_CHOICES:
            self.tipo_listbox.insert(tk.END, tipo)

        # Stato
        tk.Label(filter_frame, text="Stato:", font=FONTS["label"], bg=COLORS["background"], anchor="w").grid(row=0, column=2, sticky="w", padx=(0, 6), pady=10)
        self.stato_var = tk.StringVar(value="Tutti")
        stato_menu = ttk.Combobox(filter_frame, textvariable=self.stato_var, values=["Tutti", "APERTA", "VERIFICATA", "RISOLTA"], width=16, state="readonly", font=FONTS["input"])
        stato_menu.grid(row=0, column=3, sticky="ew", padx=(0, 16), pady=10)

        # Codice
        tk.Label(filter_frame, text="Codice:", font=FONTS["label"], bg=COLORS["background"], anchor="w").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=(0, 10))
        codice_entry = ttk.Entry(filter_frame, textvariable=self.codice_var, width=18, font=FONTS["input"])
        codice_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=(0, 16), pady=(0, 10))

        # Attività
        tk.Label(filter_frame, text="Attività:", font=FONTS["label"], bg=COLORS["background"], anchor="w").grid(row=0, column=4, sticky="w", padx=(0, 6), pady=10)
        self.attivita_var = tk.StringVar(value="Tutte")
        attivita_menu = ttk.Combobox(filter_frame, textvariable=self.attivita_var, values=["Tutte", "PICKING", "CARRELLISTI", "RICEVITORI", "DOPPIA_SPUNTA"], width=16, state="readonly", font=FONTS["input"])
        attivita_menu.grid(row=0, column=5, sticky="ew", padx=(0, 16), pady=10)
        attivita_menu.bind("<<ComboboxSelected>>", self._on_attivita_changed)

        # Anno
        tk.Label(filter_frame, text="Anno:", font=FONTS["label"], bg=COLORS["background"], anchor="w").grid(row=0, column=6, sticky="w", padx=(0, 6), pady=10)
        current_year = datetime.date.today().year
        anni = [str(y) for y in range(current_year - 5, current_year + 2)]
        self.anno_var = tk.StringVar(value=str(current_year))
        anno_menu = ttk.Combobox(filter_frame, textvariable=self.anno_var, values=anni, width=10, state="readonly", font=FONTS["input"])
        anno_menu.grid(row=0, column=7, sticky="ew", padx=(0, 16), pady=10)

        # Mese
        tk.Label(filter_frame, text="Mese:", font=FONTS["label"], bg=COLORS["background"], anchor="w").grid(row=0, column=8, sticky="w", padx=(0, 6), pady=10)
        mesi = [label for label, _ in MONTH_CHOICES]
        self.mese_var = tk.StringVar(value=mesi[0])
        mese_menu = ttk.Combobox(filter_frame, textvariable=self.mese_var, values=mesi, width=14, state="readonly", font=FONTS["input"])
        mese_menu.grid(row=0, column=9, sticky="ew", padx=(0, 16), pady=10)

        # Intervallo date
        date_range_cb = tk.Checkbutton(
            filter_frame,
            text="Usa intervallo date",
            variable=self.use_date_range_var,
            font=FONTS["label"],
            bg=COLORS["background"],
            activebackground=COLORS["background"],
            anchor="w",
        )
        date_range_cb.grid(row=0, column=10, sticky="w", padx=(0, 16), pady=10)
        # Assicura che il filtro intervallo date sia disattivato di default
        self.use_date_range_var.set(False)
        date_range_cb.deselect()

        # Dal
        tk.Label(filter_frame, text="Dal:", font=FONTS["label"], bg=COLORS["background"], anchor="w").grid(row=0, column=11, sticky="w", padx=(0, 6), pady=10)
        self.dal_var = tk.StringVar()
        dal_entry = DateEntry(filter_frame, textvariable=self.dal_var, font=FONTS["input"], date_pattern="dd/mm/yyyy", width=12)
        dal_entry.grid(row=0, column=12, sticky="ew", padx=(0, 16), pady=10)

        # Al
        tk.Label(filter_frame, text="Al:", font=FONTS["label"], bg=COLORS["background"], anchor="w").grid(row=0, column=13, sticky="w", padx=(0, 6), pady=10)
        self.al_var = tk.StringVar()
        al_entry = DateEntry(filter_frame, textvariable=self.al_var, font=FONTS["input"], date_pattern="dd/mm/yyyy", width=12)
        al_entry.grid(row=0, column=14, sticky="ew", padx=(0, 16), pady=10)

        # Pulsanti filtri - più grandi e omogenei
        button_row = tk.Frame(self.root, bg=COLORS["background"])
        button_row.pack(fill="x", padx=16, pady=(0, 12))

        create_button(
            button_row,
            text="🔍 Applica filtri",
            command=self._load_anomalie,
            variant="primary",
            width=18,
        ).pack(side="left", padx=8, pady=2)

        create_button(
            button_row,
            text="↺ Reset",
            command=self._reset_filters,
            variant="primary",
            width=14,
        ).pack(side="left", padx=8, pady=2)

        # Sezione export report
        export_frame = tk.Frame(self.root, bg=COLORS["background"], bd=2, relief="groove")
        export_frame.pack(fill="x", padx=16, pady=(0, 12))
        export_frame.grid_columnconfigure(1, weight=1)

        tk.Label(
            export_frame,
            text="Report preset:",
            font=FONTS["label"],
            bg=COLORS["background"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(8, 6), pady=10)

        self.report_combo = ttk.Combobox(
            export_frame,
            textvariable=self.report_var,
            state="readonly",
            font=FONTS["input"],
            width=45,
        )
        self.report_combo.grid(row=0, column=1, sticky="ew", padx=(0, 16), pady=10)

        self.export_button = create_button(
            export_frame,
            text="📄 Esporta Excel",
            command=self._on_export_report,
            variant="primary",
            width=18,
        )
        self.export_button.grid(row=0, column=2, sticky="e", padx=(0, 8), pady=10)

        # Tabella anomalie
        table_frame = tk.Frame(
            self.root,
            bg=COLORS["white"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        table_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        columns = (
            "id",
            "tipo",
            "anno",
            "mese",
            "codice",
            "nome",
            "attivita",
            "ore",
            "dettagli",
            "stato",
        )
        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
        )
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        headers = {
            "id": "ID",
            "tipo": "Tipo",
            "anno": "Anno",
            "mese": "Mese",
            "codice": "Codice",
            "nome": "Nome",
            "attivita": "Attività",
            "ore": "Ore TIM",
            "dettagli": "Dettagli",
            "stato": "Stato",
        }
        for column, title in headers.items():
            self.tree.heading(column, text=title)

        self.tree.column("id", width=60, anchor="center")
        self.tree.column("tipo", width=170, anchor="w")
        self.tree.column("anno", width=70, anchor="center")
        self.tree.column("mese", width=90, anchor="center")
        self.tree.column("codice", width=90, anchor="center")
        self.tree.column("nome", width=180, anchor="w")
        self.tree.column("attivita", width=110, anchor="center")
        self.tree.column("ore", width=80, anchor="center")
        self.tree.column("dettagli", width=320, anchor="w")
        self.tree.column("stato", width=110, anchor="center")

        self.tree.tag_configure("APERTA", background="#FFE5E0")
        self.tree.tag_configure("VERIFICATA", background="#FFF6DA")
        self.tree.tag_configure("RISOLTA", background="#E4FFDF")

        # Footer con azioni
        footer_frame = tk.Frame(
            self.root,
            bg=COLORS["white"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        footer_frame.pack(fill="x", padx=16)

        actions_frame = tk.Frame(footer_frame, bg=COLORS["white"])
        actions_frame.pack(side="left", padx=12, pady=10)

        create_button(
            actions_frame,
            text="✓ Verificata",
            command=lambda: self._change_stato("VERIFICATA"),
            variant="primary",
        ).pack(side="left", padx=4)

        create_button(
            actions_frame,
            text="✓ Risolta",
            command=lambda: self._change_stato("RISOLTA"),
            variant="primary",
        ).pack(side="left", padx=4)

        create_button(
            actions_frame,
            text="⟲ Riapri",
            command=lambda: self._change_stato("APERTA"),
            variant="primary",
            width=12,
        ).pack(side="left", padx=4)

        create_button(
            actions_frame,
            text="🗑️ Elimina",
            command=self._delete_anomalia,
            variant="danger",
        ).pack(side="left", padx=4)

        self.stats_label = tk.Label(
            footer_frame,
            text="",
            font=FONTS.get("subtitle", ("Segoe UI", 10)),
            bg=COLORS["white"],
            fg=COLORS["text_light"],
        )
        self.stats_label.pack(side="right", padx=12)

    def _reset_filters(self) -> None:
        self.search_var.set("")
        self.tipo_listbox.selection_clear(0, tk.END)
        self.stato_var.set(ANOMALIA_STATO_CHOICES[0])
        self.attivita_var.set("Tutte")
        self.use_date_range_var.set(False)
        self.dal_var.set("")
        self.al_var.set("")
        self.codice_var.set("")
        self.anno_var.set(str(self.today.year))
        default_month_label = next(
            (label for label, value in MONTH_CHOICES if value == self.today.month),
            MONTH_CHOICES[0][0],
        )
        self.mese_var.set(default_month_label)
        self._update_filter_states()
        self._load_anomalie()
        self._load_report_templates()

    def _update_filter_states(self) -> None:
        # Metodo rimosso: i nuovi filtri non richiedono questa logica
        pass

    def _on_filter_mode_toggle(self) -> None:
        # Metodo rimosso: i nuovi filtri non richiedono questa logica  
        pass

    def _on_period_filters_changed(self, _event: Optional[tk.Event]) -> None:
        # Metodo rimosso: i nuovi filtri non richiedono questa logica
        pass

    def _on_attivita_changed(self, _event: Optional[tk.Event] = None) -> None:
        self._load_report_templates()

    def _gather_filters(self) -> Dict[str, Any]:
        search_text = self.search_var.get().strip()
        selected_indices = self.tipo_listbox.curselection()
        tipi_selezionati = [self.tipo_listbox.get(i) for i in selected_indices]

        stato_raw = (self.stato_var.get() or "").strip()
        stato = None if stato_raw in {"", "Tutti"} else stato_raw

        attivita_raw = (self.attivita_var.get() or "").strip()
        attivita = None if attivita_raw in {"", "Tutte"} else attivita_raw

        codice = (self.codice_var.get() or "").strip().upper() or None

        anno_str = (self.anno_var.get() or "").strip()
        anno: Optional[int] = None
        if anno_str:
            try:
                anno = int(anno_str)
            except ValueError as exc:
                raise ValueError("Inserisci un anno valido (es. 2025).") from exc

        mese_label = (self.mese_var.get() or "").strip()
        if mese_label and mese_label not in MONTH_LABEL_TO_VALUE:
            raise ValueError(f"Mese selezionato non valido: {mese_label}")
        mese_value = MONTH_LABEL_TO_VALUE.get(mese_label)

        use_date_range = bool(self.use_date_range_var.get())
        data_da: Optional[datetime.date] = None
        data_a: Optional[datetime.date] = None

        if use_date_range:
            data_da_str = (self.dal_var.get() or "").strip()
            data_a_str = (self.al_var.get() or "").strip()
            try:
                if data_da_str:
                    data_da = datetime.datetime.strptime(data_da_str, "%d/%m/%Y").date()
                if data_a_str:
                    data_a = datetime.datetime.strptime(data_a_str, "%d/%m/%Y").date()
            except ValueError as exc:
                raise ValueError("Inserisci le date nel formato GG/MM/AAAA.") from exc
        else:
            if anno is not None:
                if mese_value is None:
                    data_da = datetime.date(anno, 1, 1)
                    data_a = datetime.date(anno, 12, 31)
                else:
                    data_da = datetime.date(anno, mese_value, 1)
                    if mese_value == 12:
                        data_a = datetime.date(anno, 12, 31)
                    else:
                        next_month = mese_value + 1
                        next_year = anno if next_month <= 12 else anno + 1
                        next_month = 1 if next_month > 12 else next_month
                        data_a = datetime.date(next_year, next_month, 1) - datetime.timedelta(days=1)

        tipi_upper = [t.upper() for t in tipi_selezionati]

        return {
            "search_text": search_text,
            "tipi": tipi_upper,
            "tipi_original": tipi_selezionati,
            "stato": stato,
            "attivita": attivita,
            "codice": codice,
            "anno": anno,
            "mese": mese_value,
            "mese_label": mese_label,
            "use_date_range": use_date_range,
            "data_da": data_da,
            "data_a": data_a,
        }

    def _load_anomalie(self) -> None:
        try:
            filters = self._gather_filters()
        except ValueError as exc:
            messagebox.showwarning(
                "Filtri non validi",
                str(exc),
                parent=self.dialog_parent,
            )
            return

        tipi_selezionati = filters["tipi"] or None
        stato = filters["stato"]
        attivita = filters["attivita"]
        data_da = filters["data_da"]
        data_a = filters["data_a"]
        codice = filters["codice"]

        try:
            anomalie = fetch_anomalie(
                tipo_anomalia=tipi_selezionati,
                stato=stato,
                data_da=data_da,
                data_a=data_a,
                tipo_attivita=attivita,
                codice_preparatore=codice,
            )
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile recuperare le anomalie:\n{exc}",
                parent=self.dialog_parent,
            )
            return

        search_text = filters["search_text"]

        if search_text:
            search_lower = search_text.lower()
            anomalie = [
                a for a in anomalie
                if search_lower in str(a.get("codice_preparatore", "")).lower()
                or search_lower in str(a.get("nome_preparatore", "")).lower()
            ]

        for item in self.tree.get_children():
            self.tree.delete(item)

        for anomalia in anomalie:
            ore_tim = anomalia.get("ore_tim")
            ore_str = f"{float(ore_tim):.2f}" if ore_tim is not None else ""
            
            # Usa anno e mese dal database se disponibili, altrimenti dalla data_rilevamento
            anno = anomalia.get("anno")
            mese = anomalia.get("mese")
            
            if not anno or not mese:
                # Fallback: estrae da data_rilevamento se anno/mese non presenti
                data_ril = anomalia.get("data_rilevamento")
                if data_ril:
                    anno = data_ril.year
                    mese = data_ril.month
                else:
                    anno = ""
                    mese = ""
            
            mese_label = MONTH_CHOICES[mese][0] if mese and 0 < mese < len(MONTH_CHOICES) else ""

            self.tree.insert(
                "",
                "end",
                values=(
                    anomalia.get("id", ""),
                    anomalia.get("tipo_anomalia", ""),
                    anno,
                    mese_label,
                    anomalia.get("codice_preparatore", ""),
                    anomalia.get("nome_preparatore", ""),
                    anomalia.get("tipo_attivita", ""),
                    ore_str,
                    anomalia.get("dettagli", ""),
                    anomalia.get("stato", ""),
                ),
                tags=(anomalia.get("stato", ""),),
            )

        total = len(anomalie)
        aperte = sum(1 for row in anomalie if row.get("stato") == "APERTA")
        verificate = sum(1 for row in anomalie if row.get("stato") == "VERIFICATA")
        risolte = sum(1 for row in anomalie if row.get("stato") == "RISOLTA")

        if total == 0:
            self.stats_label.config(text="Nessuna anomalia trovata")
        else:
            self.stats_label.config(
                text=(
                    f"Totale: {total} | "
                    f"Aperte: {aperte} | "
                    f"Verificate: {verificate} | "
                    f"Risolte: {risolte}"
                )
            )

    def _load_report_templates(self) -> None:
        attivita = (self.attivita_var.get() or "").strip()
        attivita_filter = None if attivita in {"", "Tutte"} else attivita

        try:
            templates = fetch_report_templates(attivi_solo=True, attivita=attivita_filter)
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile recuperare i report configurati:\n{exc}",
                parent=self.dialog_parent,
            )
            templates = []
        self.report_options = {tpl["nome"]: tpl for tpl in templates}
        values = list(self.report_options.keys())
        self.report_combo["values"] = values

        if values:
            if not self.report_var.get() or self.report_var.get() not in values:
                self.report_combo.current(0)
            self.export_button.configure(state="normal")
        else:
            self.report_var.set("")
            self.export_button.configure(state="disabled")

    def _build_placeholder_values(self, filters: Dict[str, Any]) -> Dict[str, Any]:
        oggi = datetime.date.today()
        now = datetime.datetime.now()
        mese_label = filters.get("mese_label", "") or ""
        anno = filters.get("anno")
        mese = filters.get("mese")
        data_da = filters.get("data_da")
        data_a = filters.get("data_a")
        codice = filters.get("codice")
        tipi: List[str] = filters.get("tipi", [])
        tipi_original: List[str] = filters.get("tipi_original", [])
        search_text: str = filters.get("search_text", "") or ""

        if anno is not None and mese:
            periodo_label = f"{mese_label} {anno}".strip()
            periodo_key = f"{anno:04d}-{mese:02d}"
        elif anno is not None:
            periodo_label = str(anno)
            periodo_key = str(anno)
        else:
            periodo_label = mese_label
            periodo_key = mese_label

        values: Dict[str, Any] = {
            "anno": anno,
            "mese": mese,
            "mese_nome": mese_label,
            "mese_label": mese_label,
            "mese_testo": mese_label,
            "stato": filters.get("stato"),
            "attivita": filters.get("attivita"),
            "codice": codice,
            "codice_lower": codice.lower() if isinstance(codice, str) else None,
            "utente": codice,
            "tipi": tipi,
            "tipi_list": tipi,
            "tipi_csv": ",".join(tipi),
            "tipi_count": len(tipi),
            "ha_tipi": bool(tipi),
            "tipi_original": tipi_original,
            "tipi_original_csv": ",".join(tipi_original),
            "ricerca": search_text or None,
            "ricerca_like": f"%{search_text}%" if search_text else None,
            "search": search_text or None,
            "search_like": f"%{search_text}%" if search_text else None,
            "data_da": data_da,
            "data_da_iso": data_da.isoformat() if isinstance(data_da, datetime.date) else None,
            "data_a": data_a,
            "data_a_iso": data_a.isoformat() if isinstance(data_a, datetime.date) else None,
            "use_date_range": filters.get("use_date_range", False),
            "periodo_label": periodo_label,
            "periodo_key": periodo_key,
            "oggi": oggi,
            "oggi_iso": oggi.isoformat(),
            "ora_corrente": now,
            "anno_corrente": oggi.year,
            "mese_corrente": oggi.month,
            "mese_corrente_nome": MONTH_CHOICES[oggi.month][0],
        }

        return values

    def _render_sql_template(
        self,
        template: str,
        placeholder_values: Dict[str, Any],
    ) -> Tuple[str, List[Any]]:
        pattern = re.compile(r"@([A-Za-z_][A-Za-z0-9_]*)")
        params: List[Any] = []

        def _replace(match: re.Match[str]) -> str:
            name = match.group(1)
            if name not in placeholder_values:
                raise KeyError(name)
            value = placeholder_values[name]

            if isinstance(value, (list, tuple)):
                if not value:
                    raise ValueError(
                        f"Nessun valore disponibile per la lista '@{name}'. Seleziona almeno un elemento oppure usa un placeholder alternativo."
                    )
                params.extend(list(value))
                return ", ".join(["%s"] * len(value))

            params.append(value)
            return "%s"

        query = pattern.sub(_replace, template)
        return query, params

    def _get_export_directory(self) -> Path:
        export_dir = EXPORTS_DIR / "anomalie"
        export_dir.mkdir(parents=True, exist_ok=True)
        return export_dir

    @staticmethod
    def _slugify_report_name(name: str) -> str:
        slug = re.sub(r"[^A-Za-z0-9]+", "_", name.strip().lower())
        slug = re.sub(r"_+", "_", slug).strip("_")
        return slug or "report"

    def _on_export_report(self) -> None:
        report_name = self.report_var.get().strip()
        if not report_name:
            messagebox.showwarning(
                "Selezione mancante",
                "Seleziona un report prima di procedere con l'export.",
                parent=self.dialog_parent,
            )
            return

        template = self.report_options.get(report_name)
        if not template:
            messagebox.showerror(
                "Errore",
                "Definizione del report non trovata. Aggiorna l'elenco dei report.",
                parent=self.dialog_parent,
            )
            return
        sql_template = template.get("sql_template")
        if not sql_template:
            messagebox.showerror(
                "Errore",
                "La definizione del report non contiene alcuna query SQL.",
                parent=self.dialog_parent,
            )
            return

        try:
            filters = self._gather_filters()
        except ValueError as exc:
            messagebox.showwarning(
                "Filtri non validi",
                str(exc),
                parent=self.dialog_parent,
            )
            return

        placeholder_values = self._build_placeholder_values(filters)

        try:
            query, params = self._render_sql_template(sql_template, placeholder_values)
        except KeyError as missing:
            messagebox.showerror(
                "Placeholder mancante",
                (
                    f"La query richiede il placeholder '@{missing.args[0]}',"
                    " ma non è stato possibile determinarne il valore dai filtri correnti."
                ),
                parent=self.dialog_parent,
            )
            return
        except ValueError as exc:
            messagebox.showerror(
                "Placeholder non valido",
                str(exc),
                parent=self.dialog_parent,
            )
            return

        try:
            rows, columns = execute_custom_query(query, params)
        except Exception as exc:
            messagebox.showerror(
                "Errore SQL",
                f"Errore durante l'esecuzione della query:\n{exc}",
                parent=self.dialog_parent,
            )
            return

        df = pd.DataFrame(rows)
        if columns:
            if df.empty:
                df = pd.DataFrame(columns=columns)
            else:
                df = df.loc[:, columns]

        export_dir = self._get_export_directory()
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        slug = self._slugify_report_name(report_name)
        filename = f"{slug}_{timestamp}.xlsx"
        file_path = export_dir / filename

        try:
            df.to_excel(file_path, index=False)
        except Exception as exc:
            messagebox.showerror(
                "Errore scrittura file",
                f"Impossibile salvare il file Excel:\n{exc}",
                parent=self.dialog_parent,
            )
            return

        descrizione = template.get("descrizione") or ""
        rows_count = len(df)
        msg_lines = [
            f"Report '{report_name}' esportato con successo.",
            f"Righe esportate: {rows_count}",
            f"File creato in: {file_path}",
        ]
        if descrizione:
            msg_lines.insert(1, f"Descrizione: {descrizione}")

        messagebox.showinfo(
            "Export completato",
            "\n".join(msg_lines),
            parent=self.dialog_parent,
        )

    def _change_stato(self, nuovo_stato: str) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning(
                "Nessuna selezione",
                "Seleziona un'anomalia dalla tabella.",
                parent=self.dialog_parent,
            )
            return

        anomalia_id = self.tree.item(selected[0])["values"][0]
        try:
            update_anomalia_stato(int(anomalia_id), nuovo_stato)
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile aggiornare l'anomalia:\n{exc}",
                parent=self.dialog_parent,
            )
            return

        messagebox.showinfo(
            "Aggiornamento completato",
            f"Anomalia #{anomalia_id} aggiornata a '{nuovo_stato}'.",
            parent=self.dialog_parent,
        )
        self._load_anomalie()

    def _delete_anomalia(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning(
                "Nessuna selezione",
                "Seleziona un'anomalia dalla tabella.",
                parent=self.dialog_parent,
            )
            return

        anomalia_id = self.tree.item(selected[0])["values"][0]
        if not messagebox.askyesno(
            "Conferma eliminazione",
            f"Vuoi eliminare l'anomalia #{anomalia_id}?",
            parent=self.dialog_parent,
        ):
            return

        try:
            delete_anomalia(int(anomalia_id))
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile eliminare l'anomalia:\n{exc}",
                parent=self.dialog_parent,
            )
            return

        messagebox.showinfo(
            "Eliminazione completata",
            f"Anomalia #{anomalia_id} eliminata.",
            parent=self.dialog_parent,
        )
        self._load_anomalie()

    def show(self) -> None:
        if self.use_toplevel and self.parent is None:
            self.root.mainloop()


def main() -> None:
    view = AnomalieView()
    view.show()


if __name__ == "__main__":
    main()
