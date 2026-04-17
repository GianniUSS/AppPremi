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
    fetch_anomalia_by_id,
    fetch_attivita_tim,
    fetch_codici_gestionali_by_utente,
    fetch_tipi_attivita_tim,
    fetch_report_templates,
    execute_custom_query,
    insert_codice_gestionale_tim,
    resolve_anomalie_codice_non_abbinato,
    search_utenti_tim,
    split_attivita_tim,
    ripristina_attivita_tim,
    fetch_pause_by_attivita_ids,
    update_anomalia_stato,
    update_attivita_tim,
    update_codice_gestionale_tim,
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


_TIPO_ANOMALIA_LABEL: Dict[str, str] = {
    "PRODUZIONE_SENZA_ORE": "Produzione senza ore TIM",
    "ORE_SENZA_PRODUZIONE": "Ore TIM senza produzione",
    "DIFFERENZA_60_120": "Differenza ore 60-120 min",
    "CODICE_NON_ABBINATO": "Codice non abbinato",
}


class AnomalieView:
    """Gestione delle anomalie con filtri periodo o intervallo date."""

    # Stato filtri persistente tra navigazioni
    _saved_state: Dict[str, Any] = {}

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

        s = AnomalieView._saved_state
        self.search_var = tk.StringVar(value=s.get("search", ""))
        self.codice_var = tk.StringVar(value=s.get("codice", ""))
        self.tipo_var = tk.StringVar(value="Tutte")
        self.stato_var = tk.StringVar(value=s.get("stato", ANOMALIA_STATO_CHOICES[0]))
        self.dal_var = tk.StringVar(value=s.get("dal", ""))
        self.al_var = tk.StringVar(value=s.get("al", ""))
        self.anno_var = tk.StringVar(value=s.get("anno", str(today.year)))
        default_month_label = next(
            (label for label, value in MONTH_CHOICES if value == today.month),
            MONTH_CHOICES[0][0],
        )
        self.mese_var = tk.StringVar(value=s.get("mese", default_month_label))
        self.use_date_range_var = tk.BooleanVar(value=s.get("use_range", False))
        self._saved_tipo_indices: list = s.get("tipo_indices", [])
        self.report_var = tk.StringVar()
        self.report_options: Dict[str, Dict[str, Any]] = {}
        self._sort_reverse: Dict[str, bool] = {}

        self._setup_ui()
        # Ripristina selezione listbox tipo dopo la creazione dell'UI
        if self._saved_tipo_indices:
            self.tipo_listbox.selection_clear(0, tk.END)
            for idx in self._saved_tipo_indices:
                try:
                    self.tipo_listbox.selection_set(idx)
                except Exception:
                    pass
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
        stato_menu = ttk.Combobox(filter_frame, textvariable=self.stato_var, values=["Tutti", "APERTA", "VERIFICATA"], width=16, state="readonly", font=FONTS["input"])
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
            selectmode="extended",
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
            self.tree.heading(column, text=title, command=lambda c=column: self._sort_by_column(c))

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

        self.tree.bind("<Double-1>", self._open_attivita_form)

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
            text="Da verificare",
            command=lambda: self._change_stato("VERIFICATA"),
            variant="primary",
        ).pack(side="left", padx=4)

        create_button(
            actions_frame,
            text="🗑️ Elimina",
            command=self._delete_anomalia,
            variant="danger",
        ).pack(side="left", padx=4)

        create_button(
            actions_frame,
            text="🔄 Aggiorna",
            command=self._load_anomalie,
            variant="secondary",
        ).pack(side="left", padx=4)

        self.stats_label = tk.Label(
            footer_frame,
            text="",
            font=FONTS.get("subtitle", ("Segoe UI", 10)),
            bg=COLORS["white"],
            fg=COLORS["text_light"],
        )
        self.stats_label.pack(side="right", padx=12)

    def _open_attivita_form(self, event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return

        item = self.tree.item(selection[0])
        values = item.get("values") or []
        if not values:
            return

        tipo_anomalia = values[1] if len(values) > 1 else ""
        if str(tipo_anomalia).strip() == "CODICE_NON_ABBINATO":
            self._open_codici_gestionali_form(values)
            return

        try:
            anomalia_id = int(values[0])
        except (TypeError, ValueError):
            return

        anomalia = fetch_anomalia_by_id(anomalia_id)
        if not anomalia:
            messagebox.showwarning(
                "Dettagli non disponibili",
                "Impossibile recuperare i dettagli dell'anomalia selezionata.",
                parent=self.dialog_parent,
            )
            return

        codice = anomalia.get("codice_preparatore") or ""
        nome = anomalia.get("nome_preparatore") or ""
        data_ril = anomalia.get("data_rilevamento")
        if not data_ril:
            messagebox.showwarning(
                "Dettagli non disponibili",
                "Data di riferimento mancante per l'anomalia selezionata.",
                parent=self.dialog_parent,
            )
            return

        try:
            attivita_rows = fetch_attivita_tim(codice, data_ril)
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile recuperare le attività da TIM:\n{exc}",
                parent=self.dialog_parent,
            )
            return

        dialog = tk.Toplevel(self.dialog_parent)
        dialog.title("Attività TIM")
        dlg_w, dlg_h = 1050, 460
        scr_w = dialog.winfo_screenwidth()
        scr_h = dialog.winfo_screenheight()
        dialog.geometry(f"{dlg_w}x{dlg_h}+{(scr_w - dlg_w) // 2}+{(scr_h - dlg_h) // 2}")
        dialog.configure(bg=COLORS["background"])

        header = tk.Frame(dialog, bg=COLORS["background"])
        header.pack(fill="x", padx=16, pady=(12, 6))

        data_str = data_ril.strftime("%d/%m/%Y") if hasattr(data_ril, "strftime") else str(data_ril)
        titolo = f"{codice} - {nome} | {data_str}" if nome else f"{codice} | {data_str}"

        tk.Label(
            header,
            text=titolo,
            font=FONTS["big"],
            bg=COLORS["background"],
            fg=COLORS.get("text_dark", "#222"),
        ).pack(anchor="w")

        try:
            tipi_attivita = fetch_tipi_attivita_tim()
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile recuperare i tipi attività da TIM:\n{exc}",
                parent=dialog,
            )
            tipi_attivita = []

        tipi_by_desc = {str(t.get("descrizione") or "").strip(): t.get("id") for t in tipi_attivita}
        tipi_by_id = {t.get("id"): str(t.get("descrizione") or "").strip() for t in tipi_attivita}

        editor_frame = tk.Frame(dialog, bg=COLORS["background"])
        editor_frame.pack(fill="x", padx=16, pady=(0, 8))

        tk.Label(
            editor_frame,
            text="Tipo attività:",
            font=FONTS["label"],
            bg=COLORS["background"],
            fg=COLORS.get("text_dark", "#222"),
        ).pack(side="left", padx=(0, 8))

        tipo_var = tk.StringVar()
        tipo_combo = ttk.Combobox(
            editor_frame,
            textvariable=tipo_var,
            values=list(tipi_by_desc.keys()),
            state="readonly" if tipi_by_desc else "disabled",
            font=FONTS.get("input", FONTS.get("label")),
            width=30,
        )
        tipo_combo.pack(side="left", padx=(0, 8))

        table_frame = tk.Frame(
            dialog,
            bg=COLORS["white"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        table_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        columns = ("attivita_id", "tipo_id", "attivita", "inizio", "fine", "ore", "pausa_min", "pagata", "pranzo")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        tree.heading("attivita_id", text="ID")
        tree.heading("tipo_id", text="Tipo ID")
        tree.heading("attivita", text="Attività")
        tree.heading("inizio", text="Inizio")
        tree.heading("fine", text="Fine")
        tree.heading("ore", text="Ore")
        tree.heading("pausa_min", text="Pausa (min)")
        tree.heading("pagata", text="Pausa Pagata")
        tree.heading("pranzo", text="Pausa Pranzo")

        tree.column("attivita_id", width=0, stretch=False)
        tree.column("tipo_id", width=0, stretch=False)
        tree.column("attivita", width=180, anchor="w")
        tree.column("inizio", width=70, anchor="center")
        tree.column("fine", width=70, anchor="center")
        tree.column("ore", width=60, anchor="center")
        tree.column("pausa_min", width=80, anchor="center")
        tree.column("pagata", width=150, anchor="center")
        tree.column("pranzo", width=150, anchor="center")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def refresh_attivita_rows() -> None:
            tree.delete(*tree.get_children())
            if not attivita_rows:
                tree.insert("", "end", values=("", "", "Nessuna attività trovata", "", "", "", "", "", ""))
                return

            # Recupera le pause dalla tabella pausa
            att_ids = [int(r["attivita_id"]) for r in attivita_rows if r.get("attivita_id")]
            try:
                pause_map = fetch_pause_by_attivita_ids(att_ids) if att_ids else {}
            except Exception:
                pause_map = {}

            for row in attivita_rows:
                attivita_id = row.get("attivita_id")
                tipo_id = row.get("tipo_attivita_id")
                attivita = row.get("attivita") or ""
                inizio = row.get("data_inizio")
                fine = row.get("data_fine")
                minuti = row.get("durata_minuti") or 0
                pausa_min = row.get("pausa")

                inizio_str = inizio.strftime("%H:%M") if hasattr(inizio, "strftime") else str(inizio or "")
                fine_str = fine.strftime("%H:%M") if hasattr(fine, "strftime") else str(fine or "")
                ore_str = f"{float(minuti) / 60:.2f}" if minuti is not None else ""
                pausa_min_str = str(pausa_min) if pausa_min is not None else ""

                # Orari pausa dalla tabella pausa, separati per tipo
                pause_list = pause_map.get(int(attivita_id), []) if attivita_id else []
                pagata_parts = []
                pranzo_parts = []
                for p in pause_list:
                    pi = p.get("data_inizio")
                    pf = p.get("data_fine")
                    pi_str = pi.strftime("%H:%M") if hasattr(pi, "strftime") else str(pi or "")
                    pf_str = pf.strftime("%H:%M") if hasattr(pf, "strftime") else str(pf or "")
                    orario = f"{pi_str}-{pf_str}"
                    if p.get("pagata"):
                        pagata_parts.append(orario)
                    else:
                        pranzo_parts.append(orario)
                pagata_str = " / ".join(pagata_parts)
                pranzo_str = " / ".join(pranzo_parts)

                tree.insert(
                    "",
                    "end",
                    values=(attivita_id, tipo_id, attivita, inizio_str, fine_str, ore_str, pausa_min_str, pagata_str, pranzo_str),
                )

        refresh_attivita_rows()

        def on_tree_select(_event=None) -> None:
            selection = tree.selection()
            if not selection:
                return
            item = tree.item(selection[0])
            values = item.get("values") or []
            if len(values) < 2:
                return
            tipo_id = values[1]
            tipo_desc = tipi_by_id.get(tipo_id)
            if tipo_desc:
                tipo_var.set(tipo_desc)

        tree.bind("<<TreeviewSelect>>", on_tree_select)

        def on_update_attivita() -> None:
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(
                    "Selezione mancante",
                    "Seleziona una riga da aggiornare.",
                    parent=dialog,
                )
                return

            new_desc = tipo_var.get().strip()
            new_tipo_id = tipi_by_desc.get(new_desc)
            if not new_tipo_id:
                messagebox.showwarning(
                    "Tipo non valido",
                    "Seleziona un tipo attività valido.",
                    parent=dialog,
                )
                return

            item = tree.item(selection[0])
            values = item.get("values") or []
            if len(values) < 1:
                return
            attivita_id = values[0]
            if not attivita_id:
                return

            try:
                update_attivita_tim(int(attivita_id), int(new_tipo_id))
                nonlocal_rows = fetch_attivita_tim(codice, data_ril)
                attivita_rows.clear()
                attivita_rows.extend(nonlocal_rows)
                refresh_attivita_rows()
                on_tree_select()
            except Exception as exc:
                messagebox.showerror(
                    "Errore",
                    f"Impossibile aggiornare l'attività:\n{exc}",
                    parent=dialog,
                )

        def on_split_attivita() -> None:
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(
                    "Selezione mancante",
                    "Seleziona una riga da spezzare.",
                    parent=dialog,
                )
                return

            item = tree.item(selection[0])
            values = item.get("values") or []
            if len(values) < 6:
                return

            sel_attivita_id = values[0]
            if not sel_attivita_id:
                return

            row_data = None
            for r in attivita_rows:
                if r.get("attivita_id") == int(sel_attivita_id):
                    row_data = r
                    break
            if not row_data:
                messagebox.showerror(
                    "Errore",
                    "Impossibile trovare i dati dell'attività selezionata.",
                    parent=dialog,
                )
                return

            orig_inizio = row_data["data_inizio"]
            orig_fine = row_data["data_fine"]
            orig_attivita_desc = row_data.get("attivita") or ""
            orig_pausa = row_data.get("pausa")

            split_dlg = tk.Toplevel(dialog)
            split_dlg.title("Spezza Attività")
            sp_w, sp_h = 460, 280
            sp_sw = split_dlg.winfo_screenwidth()
            sp_sh = split_dlg.winfo_screenheight()
            split_dlg.geometry(f"{sp_w}x{sp_h}+{(sp_sw - sp_w) // 2}+{(sp_sh - sp_h) // 2}")
            split_dlg.configure(bg=COLORS["background"])
            split_dlg.transient(dialog)
            split_dlg.grab_set()

            info_frame = tk.Frame(split_dlg, bg=COLORS["background"])
            info_frame.pack(fill="x", padx=16, pady=(12, 8))

            inizio_display = orig_inizio.strftime("%H:%M") if hasattr(orig_inizio, "strftime") else str(orig_inizio)
            fine_display = orig_fine.strftime("%H:%M") if hasattr(orig_fine, "strftime") else str(orig_fine)
            pausa_display = f" | Pausa: {orig_pausa} min" if orig_pausa else ""

            tk.Label(
                info_frame,
                text=f"Attività: {orig_attivita_desc}",
                font=FONTS["big"],
                bg=COLORS["background"],
                fg=COLORS.get("text_dark", "#222"),
            ).pack(anchor="w")
            tk.Label(
                info_frame,
                text=f"Orario: {inizio_display} - {fine_display}{pausa_display}",
                font=FONTS.get("subtitle", ("Segoe UI", 10)),
                bg=COLORS["background"],
                fg=COLORS.get("text_light", "#666"),
            ).pack(anchor="w", pady=(2, 0))

            form_frame = tk.Frame(split_dlg, bg=COLORS["background"])
            form_frame.pack(fill="x", padx=16, pady=(4, 8))

            tk.Label(
                form_frame, text="Spezza alle (HH:MM):",
                font=FONTS["label"], bg=COLORS["background"],
            ).grid(row=0, column=0, sticky="w", pady=4)

            part1_fine_var = tk.StringVar()
            ttk.Entry(
                form_frame, textvariable=part1_fine_var, width=8,
                font=FONTS.get("input", FONTS.get("label")),
            ).grid(row=0, column=1, sticky="w", padx=(8, 0), pady=4)

            tk.Label(
                form_frame, text="Attività spezzata:",
                font=FONTS["label"], bg=COLORS["background"],
            ).grid(row=1, column=0, sticky="w", pady=4)

            split_tipo_var = tk.StringVar(value=orig_attivita_desc)
            split_tipo_combo = ttk.Combobox(
                form_frame,
                textvariable=split_tipo_var,
                values=list(tipi_by_desc.keys()),
                state="readonly" if tipi_by_desc else "disabled",
                font=FONTS.get("input", FONTS.get("label")),
                width=28,
            )
            split_tipo_combo.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=4)

            tk.Label(
                form_frame,
                text="Le pause verranno riassegnate automaticamente.",
                font=FONTS.get("subtitle", ("Segoe UI", 9)),
                bg=COLORS["background"],
                fg=COLORS.get("text_light", "#888"),
            ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

            btn_frame = tk.Frame(split_dlg, bg=COLORS["background"])
            btn_frame.pack(fill="x", padx=16, pady=(8, 12))

            def on_confirm_split():
                p1_fine_str = part1_fine_var.get().strip()

                if not p1_fine_str:
                    messagebox.showwarning(
                        "Campo obbligatorio",
                        "Inserisci l'orario in cui spezzare l'attività.",
                        parent=split_dlg,
                    )
                    return

                time_pattern = re.compile(r"^(\d{1,2}):(\d{2})$")
                m1 = time_pattern.match(p1_fine_str)
                if not m1:
                    messagebox.showwarning(
                        "Formato non valido",
                        "Usa il formato HH:MM per l'orario.",
                        parent=split_dlg,
                    )
                    return

                ref_date = orig_inizio.date() if hasattr(orig_inizio, "date") else data_ril
                try:
                    p1_fine_dt = datetime.datetime.combine(
                        ref_date,
                        datetime.time(int(m1.group(1)), int(m1.group(2))),
                    )
                except ValueError as ve:
                    messagebox.showwarning(
                        "Orario non valido", str(ve), parent=split_dlg,
                    )
                    return

                # Tipo attività per la parte spezzata (Parte 1)
                selected_tipo_desc = split_tipo_var.get().strip()
                p1_tipo_id = tipi_by_desc.get(selected_tipo_desc)
                tipo_p1_label = selected_tipo_desc or orig_attivita_desc

                confirm = messagebox.askyesno(
                    "Conferma Spezzatura",
                    f"Parte 1: {inizio_display} - {p1_fine_str}  ({tipo_p1_label})\n"
                    f"Parte 2: {p1_fine_str} - {fine_display}  ({orig_attivita_desc})\n\n"
                    f"Le pause verranno riassegnate automaticamente.\n"
                    f"Procedere?",
                    parent=split_dlg,
                )
                if not confirm:
                    return

                try:
                    split_attivita_tim(
                        original_id=int(sel_attivita_id),
                        part1_fine=p1_fine_dt,
                        part1_tipo_attivita_id=p1_tipo_id,
                    )
                except ValueError as ve:
                    messagebox.showwarning(
                        "Validazione fallita", str(ve), parent=split_dlg,
                    )
                    return
                except Exception as exc:
                    messagebox.showerror(
                        "Errore",
                        f"Impossibile spezzare l'attività:\n{exc}",
                        parent=split_dlg,
                    )
                    return

                try:
                    refreshed = fetch_attivita_tim(codice, data_ril)
                    attivita_rows.clear()
                    attivita_rows.extend(refreshed)
                    refresh_attivita_rows()
                except Exception:
                    pass

                split_dlg.destroy()
                messagebox.showinfo(
                    "Operazione completata",
                    "Attività spezzata con successo.",
                    parent=dialog,
                )

            create_button(
                btn_frame, text="Conferma", command=on_confirm_split,
                variant="primary", width=12,
            ).pack(side="left", padx=(0, 8))

            create_button(
                btn_frame, text="Annulla", command=split_dlg.destroy,
                variant="secondary", width=12,
            ).pack(side="left")

        create_button(
            editor_frame,
            text="Aggiorna",
            command=on_update_attivita,
            variant="primary",
            width=12,
        ).pack(side="left")

        create_button(
            editor_frame,
            text="Spezza",
            command=on_split_attivita,
            variant="secondary",
            width=12,
        ).pack(side="left", padx=(8, 0))

        def on_ripristina_attivita() -> None:
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(
                    "Selezione mancante",
                    "Seleziona una riga da ripristinare.",
                    parent=dialog,
                )
                return

            item = tree.item(selection[0])
            values = item.get("values") or []
            if len(values) < 6:
                return

            sel_attivita_id = values[0]
            if not sel_attivita_id:
                return

            confirm = messagebox.askyesno(
                "Conferma Ripristino",
                "Riunire le parti spezzate dell'attività selezionata?\n\n"
                "Il record derivato verrà eliminato e l'originale\n"
                "ripristinato con gli orari completi.",
                parent=dialog,
            )
            if not confirm:
                return

            try:
                ripristina_attivita_tim(int(sel_attivita_id))
            except ValueError as ve:
                messagebox.showwarning(
                    "Ripristino non possibile",
                    str(ve),
                    parent=dialog,
                )
                return
            except Exception as exc:
                messagebox.showerror(
                    "Errore",
                    f"Impossibile ripristinare l'attività:\n{exc}",
                    parent=dialog,
                )
                return

            try:
                refreshed = fetch_attivita_tim(codice, data_ril)
                attivita_rows.clear()
                attivita_rows.extend(refreshed)
                refresh_attivita_rows()
            except Exception:
                pass

            messagebox.showinfo(
                "Operazione completata",
                "Attività ripristinata con successo.",
                parent=dialog,
            )

        create_button(
            editor_frame,
            text="Ripristina",
            command=on_ripristina_attivita,
            variant="secondary",
            width=12,
        ).pack(side="left", padx=(8, 0))

        footer = tk.Frame(dialog, bg=COLORS["background"])
        footer.pack(fill="x", padx=16, pady=(0, 12))

        create_button(
            footer,
            text="Chiudi",
            command=dialog.destroy,
            variant="secondary",
            width=12,
        ).pack(side="right")

    def _open_codici_gestionali_form(self, row_values: List[Any]) -> None:
        try:
            anomalia_id = int(row_values[0]) if row_values else None
        except (TypeError, ValueError):
            anomalia_id = None
        dialog = tk.Toplevel(self.dialog_parent)
        dialog.title("Codici Gestionali")
        cg_w, cg_h = 820, 520
        cg_sw = dialog.winfo_screenwidth()
        cg_sh = dialog.winfo_screenheight()
        dialog.geometry(f"{cg_w}x{cg_h}+{(cg_sw - cg_w) // 2}+{(cg_sh - cg_h) // 2}")
        dialog.configure(bg=COLORS["background"])

        header = tk.Frame(dialog, bg=COLORS["background"])
        header.pack(fill="x", padx=16, pady=(12, 6))

        codice = row_values[4] if len(row_values) > 4 else ""
        nome = row_values[5] if len(row_values) > 5 else ""

        titolo = "Elenco Codici Gestionale"
        if codice or nome:
            titolo = f"{titolo}: {codice} {nome}".strip()

        tk.Label(
            header,
            text=titolo,
            font=FONTS["big"],
            bg=COLORS["background"],
            fg=COLORS.get("text_dark", "#222"),
        ).pack(anchor="w")

        search_frame = tk.Frame(dialog, bg=COLORS["background"])
        search_frame.pack(fill="x", padx=16, pady=(0, 8))

        tk.Label(
            search_frame,
            text="Nominativo:",
            font=FONTS["label"],
            bg=COLORS["background"],
            fg=COLORS.get("text_dark", "#222"),
        ).pack(side="left", padx=(0, 8))

        search_var = tk.StringVar(value=str(nome or "").strip())
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=32, font=FONTS["input"])
        search_entry.pack(side="left", padx=(0, 8))

        users: List[Dict[str, Any]] = []
        user_display_map: Dict[str, int] = {}
        current_user_id: List[Optional[int]] = [None]

        table_frame = tk.Frame(
            dialog,
            bg=COLORS["white"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        table_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        columns = ("record_id", "tipo_id", "descrizione", "attivita", "codice", "dal", "al", "reparto")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        tree.heading("record_id", text="ID")
        tree.heading("tipo_id", text="Tipo ID")
        tree.heading("descrizione", text="Descrizione")
        tree.heading("attivita", text="Attività")
        tree.heading("codice", text="Codice")
        tree.heading("dal", text="Dal")
        tree.heading("al", text="Al")
        tree.heading("reparto", text="Reparto")

        tree.column("record_id", width=0, stretch=False)
        tree.column("tipo_id", width=0, stretch=False)
        tree.column("descrizione", width=160, anchor="w")
        tree.column("attivita", width=140, anchor="center")
        tree.column("codice", width=120, anchor="center")
        tree.column("dal", width=120, anchor="center")
        tree.column("al", width=120, anchor="center")
        tree.column("reparto", width=120, anchor="center")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def render_codici(rows: List[Dict[str, Any]]) -> None:
            tree.delete(*tree.get_children())
            if not rows:
                tree.insert("", "end", values=("", "", "Nessun codice trovato", "", "", "", "", ""))
                return
            for row in rows:
                record_id = row.get("record_id")
                tipo_id = row.get("tipo_attivita_id")
                descrizione = row.get("descrizione") or ""
                attivita = row.get("attivita") or ""
                cod = row.get("codice") or ""
                dal = row.get("valido_dal")
                al = row.get("valido_al")
                reparto = row.get("reparto") or ""

                dal_str = dal.strftime("%d/%m/%Y") if hasattr(dal, "strftime") else str(dal or "")
                al_str = al.strftime("%d/%m/%Y") if hasattr(al, "strftime") else str(al or "")

                tree.insert(
                    "",
                    "end",
                    values=(record_id, tipo_id, descrizione, attivita, cod, dal_str, al_str, reparto),
                )

        def on_tree_select(_event=None) -> None:
            selection = tree.selection()
            if not selection:
                return
            item = tree.item(selection[0])
            values = item.get("values") or []
            if len(values) < 8:
                return
            record_id = values[0]
            tipo_id = values[1]
            descr = values[2]
            att = values[3]
            cod = values[4]
            dal = values[5]
            al = values[6]
            rep = values[7]

            selected_record_id[0] = int(record_id) if record_id else None
            descr_var.set(str(descr or ""))
            codice_var.set(str(cod or ""))
            reparto_var.set(str(rep or ""))
            dal_var.set(str(dal or ""))
            al_var.set(str(al or ""))
            no_end_var.set(not bool(al))
            toggle_end_date()

            tipo_desc = tipo_by_id.get(tipo_id) or str(att or "")
            if tipo_desc:
                attivita_var.set(tipo_desc)

        tree.bind("<<TreeviewSelect>>", on_tree_select)

        def parse_date(value: str) -> Optional[datetime.date]:
            text = (value or "").strip()
            if not text:
                return None
            try:
                return datetime.datetime.strptime(text, "%d/%m/%Y").date()
            except ValueError:
                try:
                    return datetime.datetime.strptime(text, "%Y-%m-%d").date()
                except ValueError:
                    return None

        def collect_form_values() -> Optional[Dict[str, Any]]:
            tipo_desc = attivita_var.get().strip()
            tipo_id = tipo_by_desc.get(tipo_desc)
            if not tipo_id:
                messagebox.showwarning(
                    "Dati mancanti",
                    "Seleziona un'attività valida.",
                    parent=dialog,
                )
                return None

            codice_val = codice_var.get().strip()
            if not codice_val:
                messagebox.showwarning(
                    "Dati mancanti",
                    "Inserisci il codice.",
                    parent=dialog,
                )
                return None

            dal_date = parse_date(dal_var.get())
            if not dal_date:
                messagebox.showwarning(
                    "Dati mancanti",
                    "Inserisci una data di inizio valida.",
                    parent=dialog,
                )
                return None

            al_date = parse_date(al_var.get())

            return {
                "tipo_id": int(tipo_id),
                "codice": codice_val,
                "dal": dal_date,
                "al": al_date,
                "descrizione": descr_var.get().strip() or None,
                "reparto": reparto_var.get().strip() or None,
            }

        def on_update() -> None:
            if not selected_record_id[0]:
                messagebox.showwarning(
                    "Selezione mancante",
                    "Seleziona un record da modificare.",
                    parent=dialog,
                )
                return
            payload = collect_form_values()
            if not payload:
                return
            try:
                update_codice_gestionale_tim(
                    record_id=selected_record_id[0],
                    tipo_attivita_id=payload["tipo_id"],
                    codice=payload["codice"],
                    valido_dal=payload["dal"],
                    valido_al=payload["al"],
                    descrizione=payload["descrizione"],
                    reparto=payload["reparto"],
                )
            except Exception as exc:
                messagebox.showerror(
                    "Errore",
                    f"Impossibile aggiornare il codice gestionale:\n{exc}",
                    parent=dialog,
                )
                return
            if current_user_id[0]:
                load_codici_for_user(current_user_id[0])
            try:
                resolve_anomalie_codice_non_abbinato(str(codice), note="Codice gestionale assegnato")
                self._load_anomalie()
            except Exception:
                pass

        def on_insert() -> None:
            if not current_user_id[0]:
                messagebox.showwarning(
                    "Utente mancante",
                    "Seleziona un utente.",
                    parent=dialog,
                )
                return
            payload = collect_form_values()
            if not payload:
                return
            try:
                insert_codice_gestionale_tim(
                    utente_id=current_user_id[0],
                    tipo_attivita_id=payload["tipo_id"],
                    codice=payload["codice"],
                    valido_dal=payload["dal"],
                    valido_al=payload["al"],
                    descrizione=payload["descrizione"],
                    reparto=payload["reparto"],
                )
            except Exception as exc:
                messagebox.showerror(
                    "Errore",
                    f"Impossibile inserire il codice gestionale:\n{exc}",
                    parent=dialog,
                )
                return
            if current_user_id[0]:
                load_codici_for_user(current_user_id[0])
            clear_form()
            try:
                resolve_anomalie_codice_non_abbinato(str(codice), note="Codice gestionale assegnato")
                self._load_anomalie()
            except Exception:
                pass

        user_frame = tk.Frame(dialog, bg=COLORS["background"])
        user_frame.pack(fill="x", padx=16, pady=(0, 8))

        tk.Label(
            user_frame,
            text="Utente:",
            font=FONTS["label"],
            bg=COLORS["background"],
            fg=COLORS.get("text_dark", "#222"),
        ).pack(side="left", padx=(0, 8))

        user_var = tk.StringVar()
        user_combo = ttk.Combobox(
            user_frame,
            textvariable=user_var,
            state="readonly",
            font=FONTS["input"],
            width=40,
        )
        user_combo.pack(side="left", padx=(0, 8))

        edit_frame = tk.Frame(dialog, bg=COLORS["background"])
        edit_frame.pack(fill="x", padx=16, pady=(0, 8))

        tk.Label(
            edit_frame,
            text="Descrizione:",
            font=FONTS["label"],
            bg=COLORS["background"],
        ).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)

        descr_var = tk.StringVar()
        descr_entry = ttk.Entry(edit_frame, textvariable=descr_var, width=24, font=FONTS["input"])
        descr_entry.grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=4)

        tk.Label(
            edit_frame,
            text="Attività:",
            font=FONTS["label"],
            bg=COLORS["background"],
        ).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=4)

        try:
            tipi_attivita = fetch_tipi_attivita_tim()
        except Exception:
            tipi_attivita = []

        tipo_by_desc = {str(t.get("descrizione") or "").strip(): t.get("id") for t in tipi_attivita}
        tipo_by_id = {t.get("id"): str(t.get("descrizione") or "").strip() for t in tipi_attivita}

        attivita_var = tk.StringVar()
        attivita_combo = ttk.Combobox(
            edit_frame,
            textvariable=attivita_var,
            values=list(tipo_by_desc.keys()),
            state="normal" if tipo_by_desc else "disabled",
            font=FONTS["input"],
            width=42,
        )
        attivita_combo.grid(row=0, column=3, sticky="ew", padx=(0, 12), pady=4)

        attivita_values = list(tipo_by_desc.keys())

        def _filter_attivita(_event=None) -> None:
            text = attivita_var.get().strip().lower()
            if not text:
                attivita_combo["values"] = attivita_values
                return
            filtered = [v for v in attivita_values if v.lower().startswith(text)]
            attivita_combo["values"] = filtered

        attivita_combo.bind("<KeyRelease>", _filter_attivita)

        tk.Label(
            edit_frame,
            text="Codice:",
            font=FONTS["label"],
            bg=COLORS["background"],
        ).grid(row=1, column=0, sticky="w", padx=(0, 6), pady=4)

        codice_var = tk.StringVar()
        codice_entry = ttk.Entry(edit_frame, textvariable=codice_var, width=16, font=FONTS["input"])
        codice_entry.grid(row=1, column=1, sticky="ew", padx=(0, 12), pady=4)

        tk.Label(
            edit_frame,
            text="Dal:",
            font=FONTS["label"],
            bg=COLORS["background"],
        ).grid(row=1, column=2, sticky="w", padx=(0, 6), pady=4)

        dal_var = tk.StringVar()
        dal_entry = DateEntry(edit_frame, textvariable=dal_var, font=FONTS["input"], date_pattern="dd/mm/yyyy", width=12)
        dal_entry.grid(row=1, column=3, sticky="w", padx=(0, 12), pady=4)

        tk.Label(
            edit_frame,
            text="Al:",
            font=FONTS["label"],
            bg=COLORS["background"],
        ).grid(row=1, column=4, sticky="w", padx=(0, 6), pady=4)

        al_var = tk.StringVar()
        al_entry = DateEntry(edit_frame, textvariable=al_var, font=FONTS["input"], date_pattern="dd/mm/yyyy", width=12)
        al_entry.grid(row=1, column=5, sticky="w", padx=(0, 12), pady=4)

        no_end_var = tk.BooleanVar(value=False)

        def toggle_end_date() -> None:
            if no_end_var.get():
                al_var.set("")
                al_entry.configure(state="disabled")
            else:
                al_entry.configure(state="normal")

        tk.Checkbutton(
            edit_frame,
            text="Senza fine",
            variable=no_end_var,
            command=toggle_end_date,
            bg=COLORS["background"],
        ).grid(row=1, column=6, sticky="w", padx=(0, 12), pady=4)

        tk.Label(
            edit_frame,
            text="Reparto:",
            font=FONTS["label"],
            bg=COLORS["background"],
        ).grid(row=1, column=7, sticky="w", padx=(0, 6), pady=4)

        reparto_var = tk.StringVar()
        reparto_entry = ttk.Entry(edit_frame, textvariable=reparto_var, width=14, font=FONTS["input"])
        reparto_entry.grid(row=1, column=8, sticky="w", padx=(0, 12), pady=4)

        edit_frame.grid_columnconfigure(1, weight=1)
        edit_frame.grid_columnconfigure(3, weight=1)

        selected_record_id: List[Optional[int]] = [None]

        def clear_form() -> None:
            selected_record_id[0] = None
            descr_var.set("")
            attivita_var.set("")
            codice_var.set("")
            dal_var.set("")
            al_var.set("")
            no_end_var.set(False)
            toggle_end_date()
            reparto_var.set("")

        def load_codici_for_user(utente_id: int) -> None:
            try:
                codici = fetch_codici_gestionali_by_utente(int(utente_id))
            except Exception as exc:
                messagebox.showerror(
                    "Errore",
                    f"Impossibile recuperare i codici gestionali:\n{exc}",
                    parent=dialog,
                )
                return
            current_user_id[0] = int(utente_id)
            render_codici(codici)

        def on_user_selected(_event=None) -> None:
            label = user_var.get().strip()
            utente_id = user_display_map.get(label)
            if utente_id:
                load_codici_for_user(utente_id)

        user_combo.bind("<<ComboboxSelected>>", on_user_selected)

        def on_search() -> None:
            nome_val = search_var.get().strip()
            if not nome_val:
                messagebox.showwarning(
                    "Ricerca",
                    "Inserisci un nominativo.",
                    parent=dialog,
                )
                return
            try:
                nonlocal_users = search_utenti_tim(nome_val)
            except Exception as exc:
                messagebox.showerror(
                    "Errore",
                    f"Impossibile cercare il nominativo su TIM:\n{exc}",
                    parent=dialog,
                )
                return

            users.clear()
            users.extend(nonlocal_users)
            user_display_map.clear()

            if not users:
                render_codici([])
                return

            display_values: List[str] = []
            for utente in users:
                cognome = str(utente.get("cognome") or "").strip()
                nome_u = str(utente.get("nome") or "").strip()
                matricola = str(utente.get("matricola") or "").strip()
                badge = str(utente.get("badge") or "").strip()
                label = f"{cognome} {nome_u}".strip()
                extra_parts: List[str] = []
                if matricola:
                    extra_parts.append(f"Matr: {matricola}")
                if badge:
                    extra_parts.append(f"Badge: {badge}")
                if extra_parts:
                    label = f"{label} ({' | '.join(extra_parts)})"
                display_values.append(label)
                user_display_map[label] = int(utente.get("id"))

            user_combo["values"] = display_values
            user_var.set(display_values[0])
            load_codici_for_user(user_display_map[display_values[0]])

        create_button(
            search_frame,
            text="Cerca",
            command=on_search,
            variant="primary",
            width=12,
        ).pack(side="left")

        if search_var.get().strip():
            on_search()

        actions_frame = tk.Frame(dialog, bg=COLORS["background"])
        actions_frame.pack(fill="x", padx=16, pady=(0, 8))

        create_button(
            actions_frame,
            text="Modifica",
            command=on_update,
            variant="primary",
            width=12,
        ).pack(side="left", padx=6)

        create_button(
            actions_frame,
            text="Nuovo",
            command=on_insert,
            variant="secondary",
            width=12,
        ).pack(side="left", padx=6)

        create_button(
            actions_frame,
            text="Pulisci",
            command=clear_form,
            variant="secondary",
            width=12,
        ).pack(side="left", padx=6)

        footer = tk.Frame(dialog, bg=COLORS["background"])
        footer.pack(fill="x", padx=16, pady=(0, 12))

        create_button(
            footer,
            text="Chiudi",
            command=dialog.destroy,
            variant="secondary",
            width=12,
        ).pack(side="right")

    def _sort_by_column(self, col: str) -> None:
        """Ordina la tabella client-side in base alla colonna selezionata."""
        reverse = self._sort_reverse.get(col, False)
        items = list(self.tree.get_children(""))

        def sort_key(item_id: str):
            value = self.tree.set(item_id, col)
            return self._coerce_sort_value(value)

        items.sort(key=sort_key, reverse=reverse)

        for index, item_id in enumerate(items):
            self.tree.move(item_id, "", index)

        self._sort_reverse[col] = not reverse

    @staticmethod
    def _coerce_sort_value(value: Any) -> Any:
        if value is None:
            return (1, "")

        if isinstance(value, (int, float, datetime.date, datetime.datetime)):
            return (0, value)

        text = str(value).strip()
        if not text:
            return (1, "")

        # Date parsing
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return (0, datetime.datetime.strptime(text, fmt))
            except ValueError:
                pass

        # Number parsing (supporta formati it/en)
        normalized = text.replace(" ", "")
        if re.match(r"^-?\d{1,3}(\.\d{3})*(,\d+)?$", normalized):
            normalized = normalized.replace(".", "").replace(",", ".")
        elif re.match(r"^-?\d+([\.,]\d+)?$", normalized):
            normalized = normalized.replace(",", ".")

        try:
            return (0, float(normalized))
        except ValueError:
            return (0, text.lower())

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

        # Salva stato filtri per la prossima apertura della schermata
        AnomalieView._saved_state = {
            "search": self.search_var.get(),
            "codice": self.codice_var.get(),
            "stato": self.stato_var.get(),
            "dal": self.dal_var.get(),
            "al": self.al_var.get(),
            "anno": self.anno_var.get(),
            "mese": self.mese_var.get(),
            "use_range": self.use_date_range_var.get(),
            "tipo_indices": list(self.tipo_listbox.curselection()),
        }

        for item in self.tree.get_children():
            self.tree.delete(item)

        for anomalia in anomalie:
            ore_tim = anomalia.get("ore_tim")
            # Converti minuti → ore per la visualizzazione
            ore_str = f"{float(ore_tim)/60:.2f}" if ore_tim is not None else ""

            # Costruisce dettagli leggibili con numero colli
            totale_colli = anomalia.get("totale_colli") or 0
            dettagli_db = (anomalia.get("dettagli") or "").strip()
            parti = []
            if totale_colli:
                parti.append(f"Colli: {int(totale_colli)}")
            if dettagli_db:
                parti.append(dettagli_db)
            dettagli_display = " | ".join(parti) if parti else ""

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
                    dettagli_display,
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
        else:
            self.report_var.set("")
        self.export_button.configure(state="normal")

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

        anomalia_ids: List[int] = []
        for item_id in selected:
            values = self.tree.item(item_id).get("values") or []
            if values:
                try:
                    anomalia_ids.append(int(values[0]))
                except (TypeError, ValueError):
                    continue

        if not anomalia_ids:
            messagebox.showwarning(
                "Nessuna selezione",
                "Seleziona un'anomalia dalla tabella.",
                parent=self.dialog_parent,
            )
            return

        if len(anomalia_ids) == 1:
            confirm_text = f"Vuoi eliminare l'anomalia #{anomalia_ids[0]}?"
        else:
            confirm_text = f"Vuoi eliminare {len(anomalia_ids)} anomalie selezionate?"

        if not messagebox.askyesno(
            "Conferma eliminazione",
            confirm_text,
            parent=self.dialog_parent,
        ):
            return

        try:
            for anomalia_id in anomalia_ids:
                delete_anomalia(anomalia_id)
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile eliminare le anomalie:\n{exc}",
                parent=self.dialog_parent,
            )
            return

        if len(anomalia_ids) == 1:
            success_text = f"Anomalia #{anomalia_ids[0]} eliminata."
        else:
            success_text = f"Eliminate {len(anomalia_ids)} anomalie."

        messagebox.showinfo(
            "Eliminazione completata",
            success_text,
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
