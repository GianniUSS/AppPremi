"""
Finestra di visualizzazione dati in stile Excel.
"""
import datetime
import calendar
import threading
import time
import tkinter as tk
from contextlib import closing
from tkinter import messagebox, ttk
from typing import Any, Dict, List, Optional, cast

import mysql.connector

from config import COLORS, FONTS, MYSQL_CONFIG, MYSQL_CONFIG_MAIN, TABLE_NAME
from database import load_nuove_aperture, save_nuove_aperture
from ui_components import create_button


MONTH_CHOICES: List[tuple[str, Optional[int]]] = [
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


class DataViewer:
    """Visualizzatore dati in stile Excel con funzionalità di ricerca e filtro."""

    def __init__(self, parent: Optional[tk.Misc] = None):
        self.parent_window = parent  # Salva riferimento alla finestra di importazione
        self.is_standalone = parent is None

        self.window: tk.Misc
        if self.is_standalone:
            self.window = tk.Tk()
            self.window.title("Visualizzatore Dati Produzione")
            self.window.geometry("1200x700")
            self.window.configure(bg=COLORS["background"])
        else:
            if parent is None:
                raise ValueError("Per l'uso embedded è richiesto un parent.")
            # Quando è embedded, usa il parent direttamente come container
            self.window = parent
            # Non chiama pack qui perché il parent è già nel layout

        today = datetime.date.today()
        self.search_var = tk.StringVar()
        self.tipo_attivita_var = tk.StringVar(value="Tutti")
        self.data_da_var = tk.StringVar()
        self.data_a_var = tk.StringVar()
        self.use_date_filter_var = tk.BooleanVar(value=False)
        self.anno_var = tk.StringVar(value=str(today.year))
        self.mese_var = tk.StringVar(value=MONTH_CHOICES[today.month][0])
        self._sync_in_progress = False
        self._last_filters = None
        self._stats_text_before_sync = ""

        self._setup_ui()
        self._update_filter_states()
        self._load_data(self._collect_filters())

    def _setup_ui(self) -> None:
        """Configura l'interfaccia utente della finestra."""
        style = ttk.Style()
        style.configure("Main.TFrame", background=COLORS["background"])
        style.configure("Card.TFrame", background=COLORS["white"])
        # In modalità embedded, rimuovi padding per massimizzare lo spazio
        if self.is_standalone:
            main_frame = ttk.Frame(self.window, style="Main.TFrame")
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        else:
            main_frame = ttk.Frame(self.window, style="Main.TFrame")
            main_frame.pack(fill="both", expand=True)

        header_frame = ttk.Frame(main_frame, style="Card.TFrame")
        header_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(
            header_frame,
            text="📊 Visualizzatore Dati Produzione",
            font=FONTS["title"],
            background=COLORS["white"],
            foreground=COLORS["text_dark"],
        ).pack(pady=15)

        filter_frame = ttk.Frame(main_frame, style="Card.TFrame")
        filter_frame.pack(fill="x", pady=(0, 10), padx=5)

        search_row = ttk.Frame(filter_frame, style="Card.TFrame")
        search_row.pack(fill="x", padx=10, pady=(10, 5))

        ttk.Label(
            search_row,
            text="🔍 Ricerca:",
            font=FONTS["big"],
            background=COLORS["white"],
        ).pack(side="left", padx=(0, 10))

        search_entry = ttk.Entry(
            search_row,
            textvariable=self.search_var,
            font=FONTS["big"],
            width=40,
        )
        search_entry.pack(side="left", padx=(0, 10))

        create_button(
            search_row,
            text="Cerca",
            command=self._apply_filters,
            variant="primary",
            width=12,
        ).pack(side="left", padx=5)

        create_button(
            search_row,
            text="Pulisci Filtri",
            command=self._clear_filters,
            variant="primary",
            width=14,
        ).pack(side="left", padx=5)

        create_button(
            search_row,
            text="🔄 Ricarica",
            command=lambda: self._load_data(),
            variant="primary",
            width=14,
        ).pack(side="left", padx=5)

        filter_row = ttk.Frame(filter_frame, style="Card.TFrame")
        filter_row.pack(fill="x", padx=10, pady=(5, 10))

        ttk.Label(
            filter_row,
            text="Tipo Attività:",
            font=FONTS["big"],
            background=COLORS["white"],
        ).pack(side="left", padx=(0, 5))

        tipo_combo = ttk.Combobox(
            filter_row,
            textvariable=self.tipo_attivita_var,
            values=["Tutti", "PICKING", "CARRELLISTI", "RICEVITORI", "DOPPIA_SPUNTA"],
            state="readonly",
            font=FONTS["big"],
            width=20,
        )
        tipo_combo.pack(side="left", padx=(0, 15))
        tipo_combo.bind("<<ComboboxSelected>>", self._on_tipo_attivita_changed)

        current_year = datetime.date.today().year
        year_options = [str(current_year - i) for i in range(0, 6)]

        ttk.Label(
            filter_row,
            text="Anno:",
            font=FONTS["big"],
            background=COLORS["white"],
        ).pack(side="left", padx=(0, 5))

        self.anno_combo = ttk.Combobox(
            filter_row,
            textvariable=self.anno_var,
            values=year_options,
            state="readonly",
            font=FONTS["big"],
            width=6,
        )
        self.anno_combo.pack(side="left", padx=(0, 15))
        self.anno_combo.bind("<<ComboboxSelected>>", self._on_period_filters_changed)

        ttk.Label(
            filter_row,
            text="Mese:",
            font=FONTS["big"],
            background=COLORS["white"],
        ).pack(side="left", padx=(0, 5))

        month_labels = [label for label, _ in MONTH_CHOICES]
        self.mese_combo = ttk.Combobox(
            filter_row,
            textvariable=self.mese_var,
            values=month_labels,
            state="readonly",
            font=FONTS["big"],
            width=14,
        )
        self.mese_combo.pack(side="left", padx=(0, 15))
        self.mese_combo.bind("<<ComboboxSelected>>", self._on_period_filters_changed)

        self.date_filter_check = ttk.Checkbutton(
            filter_row,
            text="Filtra per intervallo date",
            variable=self.use_date_filter_var,
            command=self._on_filter_mode_toggle,
        )
        self.date_filter_check.pack(side="left", padx=(0, 15))
        # Garantisce che il filtro per intervallo date sia disattivato all'apertura
        self.use_date_filter_var.set(False)
        self.date_filter_check.state(["!selected"])


        from tkcalendar import DateEntry
        ttk.Label(
            filter_row,
            text="Data da:",
            font=FONTS["big"],
            background=COLORS["white"],
        ).pack(side="left", padx=(0, 5))
        self.date_da_entry = DateEntry(
            filter_row,
            textvariable=self.data_da_var,
            font=FONTS["big"],
            width=12,
            date_pattern="yyyy-mm-dd",
            background=COLORS["white"],
            foreground=COLORS["text_dark"],
        )
        self.date_da_entry.pack(side="left", padx=(0, 15))

        ttk.Label(
            filter_row,
            text="Data a:",
            font=FONTS["big"],
            background=COLORS["white"],
        ).pack(side="left", padx=(0, 5))
        self.date_a_entry = DateEntry(
            filter_row,
            textvariable=self.data_a_var,
            font=FONTS["big"],
            width=12,
            date_pattern="yyyy-mm-dd",
            background=COLORS["white"],
            foreground=COLORS["text_dark"],
        )
        self.date_a_entry.pack(side="left", padx=(0, 15))

        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(table_frame, orient="vertical")
        hsb = ttk.Scrollbar(table_frame, orient="horizontal")

        columns = (
            "ID",
            "Data",
            "Codice",
            "Nome",
            "Colli",
            "Penalità",
            "Tipo Attività",
            "Tipo",
            "Ore TIM",
            "Ore Gestionale",
        )
        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            height=20,
        )

        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        widths = [60, 100, 120, 200, 80, 80, 140, 150, 100, 120]
        for col, width in zip(columns, widths):
            self.tree.heading(col, text=col, command=lambda c=col: self._sort_by_column(c))
            anchor = "w" if col == "Nome" else "center"
            self.tree.column(col, width=width, anchor=anchor)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree.tag_configure("oddrow", background="#F8F9FA")
        self.tree.tag_configure("evenrow", background=COLORS["white"])

        # Footer con padding aumentato per visibilità
        footer_frame = ttk.Frame(main_frame, style="Card.TFrame", relief="solid", borderwidth=2)
        footer_frame.pack(fill="x", pady=(10, 10), padx=5)

        # Stats a sinistra
        self.stats_label = ttk.Label(
            footer_frame,
            text="",
            font=FONTS["big"],
            background=COLORS["white"],
            foreground=COLORS["text_dark"],
        )
        self.stats_label.pack(side="left", padx=10, pady=10)

        # Pulsante sincronizzazione a destra
        self.sync_button = create_button(
            footer_frame,
            text="🔄 Sincronizza con TIM",
            command=self._sync_with_tim,
            variant="primary",
            width=20,
        )
        self.sync_button.configure(state=tk.NORMAL)
        self.sync_button.pack(side="right", padx=10, pady=10)

        # Pulsante nuove aperture (visibile solo per Doppia Spunta)
        self.nuove_aperture_button = create_button(
            footer_frame,
            text="🏪 Nuove Aperture",
            command=self._gestisci_nuove_aperture,
            variant="primary",
            width=18,
        )
        # Inizialmente nascosto (pack_forget viene chiamato automaticamente se non si fa pack)

        self.progress_window = None
        self.sync_progress = None
        self.sync_progress_label = None

        self.tree.bind("<Double-1>", self._show_details)

    def _collect_filters(self) -> Dict[str, Any]:
        filters: Dict[str, Any] = {
            "search": self.search_var.get().strip(),
            "tipo_attivita": self.tipo_attivita_var.get(),
            "use_date_filter": self.use_date_filter_var.get(),
        }

        if filters["use_date_filter"]:
            filters["data_da"] = self.data_da_var.get().strip()
            filters["data_a"] = self.data_a_var.get().strip()
        else:
            anno_value = self.anno_var.get().strip()
            filters["anno"] = anno_value or None
            month_value = MONTH_LABEL_TO_VALUE.get(self.mese_var.get())
            filters["mese"] = month_value

        return filters

    def _update_filter_states(self) -> None:
        use_date = self.use_date_filter_var.get()

        if use_date:
            self.anno_combo.configure(state="disabled")
            self.mese_combo.configure(state="disabled")
            self.date_da_entry.configure(state="normal")
            self.date_a_entry.configure(state="normal")
        else:
            self.anno_combo.configure(state="readonly")
            self.mese_combo.configure(state="readonly")
            self.date_da_entry.configure(state="disabled")
            self.date_a_entry.configure(state="disabled")

    def _on_filter_mode_toggle(self) -> None:
        self._update_filter_states()
        self._load_data(self._collect_filters())

    def _on_period_filters_changed(self, _event=None) -> None:
        if not self.use_date_filter_var.get():
            self._load_data(self._collect_filters())

    def _load_data(self, filters: dict | None = None) -> None:
        """Carica i dati dal database locale applicando eventuali filtri."""
        filters = filters or self._collect_filters()
        self._last_filters = filters.copy()
        try:
            for item in self.tree.get_children():
                self.tree.delete(item)

            query = (
                "SELECT id, data, codice_preparatore, nome_preparatore, totale_colli, "
                "penalita, tipo_attivita, tipo, ore_tim, ore_gestionale FROM "
                f"{TABLE_NAME}"
            )
            conditions: List[str] = []
            params: List[str] = []

            search_text = filters.get("search")
            if search_text:
                like = f"%{search_text}%"
                conditions.append(
                    "(codice_preparatore LIKE %s OR nome_preparatore LIKE %s OR tipo_attivita LIKE %s)"
                )
                params.extend([like, like, like])

            tipo_attivita = filters.get("tipo_attivita")
            if tipo_attivita and tipo_attivita != "Tutti":
                conditions.append("tipo_attivita = %s")
                params.append(tipo_attivita)

            if filters.get("use_date_filter"):
                data_da = filters.get("data_da")
                if data_da:
                    conditions.append("data >= %s")
                    params.append(data_da)

                data_a = filters.get("data_a")
                if data_a:
                    conditions.append("data <= %s")
                    params.append(data_a)
            else:
                anno = filters.get("anno")
                if anno:
                    conditions.append("YEAR(data) = %s")
                    params.append(str(anno))
                mese = filters.get("mese")
                if mese:
                    conditions.append("MONTH(data) = %s")
                    params.append(str(mese))

            if conditions:
                query += " WHERE " + " AND ".join(conditions)

            query += " ORDER BY data DESC, id DESC LIMIT 1000"

            with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
                with closing(conn.cursor(dictionary=True)) as cur:
                    cur.execute(query, params)
                    rows = cast(List[Dict[str, Any]], cur.fetchall())

            total_colli = 0
            total_penalita = 0
            total_ore = 0.0
            total_ore_gestionale = 0.0

            for idx, row in enumerate(rows):
                values = (
                    row.get("id"),
                    row.get("data"),
                    row.get("codice_preparatore"),
                    row.get("nome_preparatore"),
                    row.get("totale_colli"),
                    row.get("penalita"),
                    row.get("tipo_attivita"),
                    row.get("tipo"),
                    row.get("ore_tim", 0),
                    row.get("ore_gestionale", 0),
                )
                tag = "evenrow" if idx % 2 == 0 else "oddrow"
                self.tree.insert("", "end", values=values, tags=(tag,))

                colli = row.get("totale_colli") or 0
                penalita = row.get("penalita") or 0
                ore_tim = row.get("ore_tim") or 0
                ore_gestionale = row.get("ore_gestionale") or 0
                try:
                    total_colli += int(float(colli))
                except (TypeError, ValueError):
                    pass
                try:
                    total_penalita += int(float(penalita))
                except (TypeError, ValueError):
                    pass
                try:
                    total_ore += float(ore_tim)
                except (TypeError, ValueError):
                    pass
                try:
                    total_ore_gestionale += float(ore_gestionale)
                except (TypeError, ValueError):
                    pass

            self.stats_label.config(
                text=(
                    f"📈 Record trovati: {len(rows)} | Totale Ore TIM: {total_ore:.2f} | "
                    f"Totale Ore Gestionale: {total_ore_gestionale:.2f} | "
                    f"Totale Colli: {total_colli:,} | Totale Penalità: {total_penalita:,}"
                )
            )
        except Exception as exc:
            messagebox.showerror("Errore", f"Errore nel caricamento dati:\n{exc}")

    def _apply_filters(self) -> None:
        """Applica i filtri scelti dall'utente."""
        filters = self._collect_filters()

        if filters["use_date_filter"]:
            for key in ("data_da", "data_a"):
                value = filters.get(key)
                if value:
                    try:
                        datetime.datetime.strptime(value, "%Y-%m-%d")
                    except ValueError:
                        messagebox.showwarning(
                            "Attenzione",
                            f"Formato data non valido per '{key}'. Usa: YYYY-MM-DD",
                        )
                        return
        else:
            anno = filters.get("anno")
            if anno and not str(anno).isdigit():
                messagebox.showwarning(
                    "Attenzione",
                    "Inserisci un anno numerico (es. 2025).",
                )
                return

        self._load_data(filters)

    def _clear_filters(self) -> None:
        """Ripristina i filtri di ricerca ai valori di default."""
        today = datetime.date.today()
        self.search_var.set("")
        self.tipo_attivita_var.set("Tutti")
        self.data_da_var.set("")
        self.data_a_var.set("")
        self.use_date_filter_var.set(False)
        self.anno_var.set(str(today.year))
        self.mese_var.set(MONTH_CHOICES[today.month][0])
        self._update_filter_states()
        self._load_data(self._collect_filters())
        self._update_nuove_aperture_button()

    def _on_tipo_attivita_changed(self, event=None) -> None:
        """Mostra/nasconde il pulsante Nuove Aperture in base al tipo attività selezionato."""
        self._update_nuove_aperture_button()
    
    def _update_nuove_aperture_button(self) -> None:
        """Aggiorna la visibilità del pulsante Nuove Aperture."""
        tipo = self.tipo_attivita_var.get()
        if tipo == "DOPPIA_SPUNTA":
            # Mostra il pulsante tra stats e sync (side=right, ma prima del sync)
            self.nuove_aperture_button.pack(side="right", padx=10, pady=10, before=self.sync_button)
        else:
            self.nuove_aperture_button.pack_forget()

    def _sort_by_column(self, col: str) -> None:
        """Ordina rapidamente la tabella ricaricando i dati (stub)."""
        # TODO: implementare un ordinamento client-side se necessario.
        pass

    def _show_details(self, event) -> None:
        """Mostra i dettagli del record selezionato."""
        selection = self.tree.selection()
        if not selection:
            return

        item = self.tree.item(selection[0])
        values = item["values"]

        details = (
            "Dettagli Record:\n\n"
            f"ID: {values[0]}\n"
            f"Data: {values[1]}\n"
            f"Codice Preparatore: {values[2]}\n"
            f"Nome Preparatore: {values[3]}\n"
            f"Totale Colli: {values[4]}\n"
            f"Penalità: {values[5]}\n"
            f"Tipo Attività: {values[6]}\n"
            f"Tipo: {values[7]}"
        )
        messagebox.showinfo("Dettagli Record", details)

    def _gestisci_nuove_aperture(self) -> None:
        """Apre una finestra per selezionare le nuove aperture dalla colonna tipo."""
        # Recupera i filtri applicati (legge direttamente dai widget)
        conditions = ["tipo_attivita = 'DOPPIA_SPUNTA'", "tipo IS NOT NULL", "tipo != ''"]
        params = []
        
        # Usa l'intervallo date solo se il filtro dedicato è attivo
        if self.use_date_filter_var.get():
            data_da = self.data_da_var.get().strip() if self.data_da_var.get() else None
            data_a = self.data_a_var.get().strip() if self.data_a_var.get() else None
        else:
            anno_str = self.anno_var.get().strip()
            mese_label = self.mese_var.get().strip()
            mese_value = MONTH_LABEL_TO_VALUE.get(mese_label)

            if not (anno_str.isdigit() and mese_value):
                messagebox.showwarning(
                    "Attenzione",
                    "Per gestire le nuove aperture seleziona un mese specifico oppure abilita il filtro per intervallo date.",
                    parent=self.window,
                )
                return

            anno_int = int(anno_str)
            first_day = datetime.date(anno_int, mese_value, 1)
            last_day = datetime.date(anno_int, mese_value, calendar.monthrange(anno_int, mese_value)[1])
            data_da = first_day.isoformat()
            data_a = last_day.isoformat()

        if not data_da or not data_a:
            messagebox.showwarning(
                "Attenzione",
                "Specifica un intervallo valido di date per salvare le nuove aperture.",
                parent=self.window,
            )
            return
        
        # Applica i filtri di data (campo 'data' nella tabella)
        if data_da:
            conditions.append("data >= %s")
            params.append(data_da)
        if data_a:
            conditions.append("data <= %s")
            params.append(data_a)
        
        query = f"""
            SELECT DISTINCT tipo 
            FROM dati_produzione 
            WHERE {' AND '.join(conditions)}
            ORDER BY tipo
        """
        
        print(f"\n{'='*60}")
        print(f"QUERY NEGOZI DOPPIA SPUNTA:")
        print(f"Query: {query}")
        print(f"Parametri: {params}")
        print(f"Date filtro: da={data_da}, a={data_a}")
        print(f"{'='*60}\n")
        
        try:
            with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
                with closing(conn.cursor()) as cursor:
                    cursor.execute(query, params)
                    negozi = [row[0] for row in cursor.fetchall()]
                    print(f"✓ Trovati {len(negozi)} negozi: {negozi[:10]}{'...' if len(negozi) > 10 else ''}")
        except mysql.connector.Error as e:
            print(f"❌ Errore query: {e}")
            messagebox.showerror("Errore", f"Errore nel recupero dei negozi:\n{e}", parent=self.window)
            return
        
        if not negozi:
            msg = "Nessun negozio trovato per la Doppia Spunta"
            if data_da or data_a:
                msg += f"\nnel periodo selezionato ({data_da} - {data_a})."
            else:
                msg += ".\n\nAssicurati di aver importato dati di tipo 'Doppia Spunta'."
            messagebox.showinfo("Nessun dato", msg, parent=self.window)
            return
        
        # Carica le nuove aperture salvate per questo periodo
        nuove_aperture_salvate = set(load_nuove_aperture(data_da, data_a))
        
        # Nascondi la finestra di importazione se esiste (solo in modalità standalone)
        if self.parent_window and hasattr(self.parent_window, 'withdraw'):
            self.parent_window.withdraw()
        
        # Porta in primo piano la finestra del visualizzatore (solo se è standalone)
        if self.is_standalone:
            self.window.lift()
            self.window.focus_force()
        
        # Trova la finestra root (per il popup)
        if self.is_standalone:
            root_window = self.window
        else:
            # Trova il root principale navigando verso l'alto
            root_window = self.window.winfo_toplevel()
        
        # Crea finestra popup
        popup = tk.Toplevel(root_window)
        popup.title("Gestione Nuove Aperture")
        popup.geometry("700x600")
        popup.transient(root_window)
        popup.grab_set()
        
        # Centra la finestra
        popup.update_idletasks()
        x = root_window.winfo_x() + (root_window.winfo_width() // 2) - (700 // 2)
        y = root_window.winfo_y() + (root_window.winfo_height() // 2) - (600 // 2)
        popup.geometry(f"700x600+{x}+{y}")
        
        # Frame principale
        main_frame = ttk.Frame(popup, padding=20)
        main_frame.pack(fill="both", expand=True)
        
        # Titolo
        ttk.Label(
            main_frame,
            text="Seleziona le Nuove Aperture",
            font=FONTS["title"]
        ).pack(pady=(0, 10))
        
        ttk.Label(
            main_frame,
            text="Seleziona i negozi che sono nuove aperture per la Doppia Spunta:",
            font=FONTS["subtitle"]
        ).pack(pady=(0, 20))
        
        # Frame con scrollbar per la lista di checkbox
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill="both", expand=True)
        
        canvas = tk.Canvas(list_frame, bg=COLORS["white"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas_frame = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Funzione per lo scroll con la rotella del mouse
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        # Bind della rotella del mouse
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Quando la finestra viene chiusa, rimuovi il binding e ripristina la finestra di importazione
        def on_close():
            canvas.unbind_all("<MouseWheel>")
            popup.destroy()
            # Ripristina la finestra di importazione se era nascosta
            if self.parent_window:
                self.parent_window.deiconify()
        
        popup.protocol("WM_DELETE_WINDOW", on_close)
        
        # Posiziona canvas e scrollbar PRIMA di creare i checkbox
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Aggiorna la larghezza del canvas quando viene ridimensionato
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_frame, width=event.width)
        
        canvas.bind("<Configure>", _on_canvas_configure)
        
        # Dizionario per memorizzare i checkbox
        checkbox_vars = {}
        
        # Crea checkbox per ogni negozio
        for negozio in negozi:
            var = tk.BooleanVar(value=negozio in nuove_aperture_salvate)
            checkbox_vars[negozio] = var
            
            cb = ttk.Checkbutton(
                scrollable_frame,
                text=str(negozio),
                variable=var
            )
            cb.pack(anchor="w", pady=2, padx=10)
        
        # Aggiorna il canvas per calcolare la regione scrollabile
        scrollable_frame.update_idletasks()
        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        
        # Frame pulsanti
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        
        def save_and_close():
            # Salva le nuove aperture selezionate nel database
            selected = [negozio for negozio, var in checkbox_vars.items() if var.get()]
            try:
                count = save_nuove_aperture(data_da, data_a, selected)
                messagebox.showinfo("Salvato", f"Salvate {count} nuove aperture per il periodo {data_da} - {data_a}.", parent=popup)
            except Exception as e:
                messagebox.showerror("Errore", f"Errore nel salvataggio:\n{e}", parent=popup)
            popup.destroy()
            # Ripristina la finestra di importazione se era nascosta
            if self.parent_window:
                self.parent_window.deiconify()
        
        create_button(
            button_frame,
            text="✓ Salva",
            command=save_and_close,
            variant="primary",
            width=10,
        ).pack(side="right", padx=5)

        create_button(
            button_frame,
            text="✗ Annulla",
            command=on_close,
            variant="secondary",
            width=10,
        ).pack(side="right")

    def _load_nuove_aperture(self) -> set:
        """Carica la lista delle nuove aperture da file."""
        import json
        try:
            with open("nuove_aperture.json", "r", encoding="utf-8") as f:
                return set(json.load(f))
        except FileNotFoundError:
            return set()
        except Exception:
            return set()
    
    def _save_nuove_aperture(self, nuove_aperture: list) -> None:
        """Salva la lista delle nuove aperture su file."""
        import json
        try:
            with open("nuove_aperture.json", "w", encoding="utf-8") as f:
                json.dump(nuove_aperture, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Errore", f"Errore nel salvataggio:\n{e}")

    def _sync_with_tim(self) -> None:
        """Avvia la sincronizzazione nel thread secondario con feedback visivo."""
        # Log immediato per debug
        try:
            with open("sync_debug.txt", "w", encoding="utf-8") as f:
                f.write("_sync_with_tim chiamato!\n")
        except:
            pass
        
        if self._sync_in_progress:
            return

        self._sync_in_progress = True
        self._stats_text_before_sync = self.stats_label.cget("text")
        self.sync_button.config(state="disabled")
        
        # Trova la finestra root (per il popup)
        if self.is_standalone:
            root_window = self.window
        else:
            root_window = self.window.winfo_toplevel()
        
        # Crea finestra popup per progress bar
        self.progress_window = tk.Toplevel(root_window)
        self.progress_window.title("Sincronizzazione")
        self.progress_window.geometry("500x120")
        self.progress_window.resizable(False, False)
        self.progress_window.transient(root_window)
        self.progress_window.grab_set()
        
        # Centra la finestra
        self.progress_window.update_idletasks()
        x = root_window.winfo_x() + (root_window.winfo_width() // 2) - (500 // 2)
        y = root_window.winfo_y() + (root_window.winfo_height() // 2) - (120 // 2)
        self.progress_window.geometry(f"500x120+{x}+{y}")
        
        frame = ttk.Frame(self.progress_window, padding=20)
        frame.pack(fill="both", expand=True)
        
        ttk.Label(
            frame,
            text="⏳ Sincronizzazione in corso...",
            font=FONTS["big"]
        ).pack(pady=(0, 10))
        
        self.sync_progress_label = ttk.Label(
            frame,
            text="0%",
            font=FONTS["big"]
        )
        self.sync_progress_label.pack(pady=(0, 5))
        
        self.sync_progress = ttk.Progressbar(
            frame,
            mode="determinate",
            length=450,
            maximum=100,
        )
        self.sync_progress.pack(fill="x")
        self.sync_progress['value'] = 0
        
        self.stats_label.config(text="⏳ Sincronizzazione in corso...")

        threading.Thread(target=self._sync_background, daemon=True).start()

    def _sync_background(self) -> None:
        """Esegue la sincronizzazione fuori dal thread GUI."""
        self._sync_start_time = time.time()
        result = self._perform_sync_with_progress()
        result['elapsed_time'] = time.time() - self._sync_start_time
        self.window.after(0, lambda: self._on_sync_complete(result))
    def _update_progress(self, percent: float) -> None:
        def update():
            if self.sync_progress and self.sync_progress_label:
                self.sync_progress['value'] = percent
                elapsed = int(time.time() - self._sync_start_time)
                self.sync_progress_label.config(text=f"{percent}% - {elapsed}s trascorsi")
        self.window.after(0, update)

    def _perform_sync_with_progress(self) -> Dict[str, Any]:
        """Logica di sincronizzazione con avanzamento progressivo."""
        
        # Step 1: Leggi i codici e tipi da dati_produzione CON I FILTRI APPLICATI
        local_records = []
        try:
            with closing(mysql.connector.connect(**MYSQL_CONFIG)) as local_conn:
                with closing(local_conn.cursor(dictionary=True)) as local_cursor:
                    # Costruisci query con gli stessi filtri della visualizzazione
                    query = """
                        SELECT DISTINCT codice_preparatore, tipo_attivita, data, nome_preparatore
                        FROM dati_produzione
                        WHERE codice_preparatore IS NOT NULL
                          AND tipo_attivita IS NOT NULL
                          AND data IS NOT NULL
                    """
                    conditions: List[str] = []
                    params: List[Any] = []
                    
                    # Applica gli stessi filtri della visualizzazione
                    if self._last_filters:
                        if self._last_filters.get("data_da"):
                            conditions.append("data >= %s")
                            params.append(self._last_filters["data_da"])
                        if self._last_filters.get("data_a"):
                            conditions.append("data <= %s")
                            params.append(self._last_filters["data_a"])
                        if self._last_filters.get("codice"):
                            conditions.append("LOWER(codice_preparatore) LIKE LOWER(%s)")
                            params.append(f"%{self._last_filters['codice']}%")
                        if self._last_filters.get("nome"):
                            conditions.append("LOWER(nome_preparatore) LIKE LOWER(%s)")
                            params.append(f"%{self._last_filters['nome']}%")
                        tipo_filter = self._last_filters.get("tipo_attivita")
                        if tipo_filter and tipo_filter != "Tutti":
                            conditions.append("tipo_attivita = %s")
                            params.append(tipo_filter)
                        search_term = (self._last_filters.get("search") or "").strip()
                        if search_term:
                            conditions.append("(LOWER(codice_preparatore) LIKE LOWER(%s) OR LOWER(nome_preparatore) LIKE LOWER(%s))")
                            params.extend([f"%{search_term}%", f"%{search_term}%"])
                    
                    if conditions:
                        query += " AND " + " AND ".join(conditions)
                    
                    local_cursor.execute(query, params)
                    local_records = cast(List[Dict[str, Any]], local_cursor.fetchall())
        except mysql.connector.Error as err:
            return {
                "success": False,
                "message": f"Errore durante la lettura del database locale:\n{err}",
            }
        
        if not local_records:
            return {
                "success": True,
                "no_data": True,
                "updated": 0,
                "non_mapped": [],
                "non_mapped_details": [],
                "mapping_warning": False,
            }

        aggiornati = 0
        non_trovati_dettaglio: List[Dict[str, Any]] = []  # Codici non trovati in TIM
        anomalie_ore_count = 0
        anomalie_x_xx_count = 0  # Inizializza contatore anomalie X/XX
        anomalie_senza_ore_count = 0  # Inizializza contatore anomalie PRODUZIONE_SENZA_ORE

        local_nome_map: Dict[tuple[str, str, str], str] = {}

        # Import per anomalie
        from database import insert_anomalia
        today = datetime.date.today()

        # Step 3: OTTIMIZZAZIONE - Recupera TUTTE le durate da TIM in una query
        try:
            with closing(mysql.connector.connect(**MYSQL_CONFIG)) as app_conn:
                with closing(app_conn.cursor(dictionary=True)) as app_cursor:
                    # Mappa tipo_attivita locale -> tipo_attivita_id TIM
                    tipo_mapping = {
                        "PICKING": "PICKING",
                        "CARRELLISTI": "CARRELLISTI",
                        "RICEVITORI": "MAG. RICEVIMENTO",
                        "DOPPIA_SPUNTA": "DOPPIA SPUNTA"
                    }
                    
                    with closing(mysql.connector.connect(**MYSQL_CONFIG_MAIN)) as tim_conn:
                        with closing(tim_conn.cursor(dictionary=True)) as tim_cursor:
                            
                            print("🚀 Recupero TUTTE le durate da TIM in una query...")
                            
                            # Costruisci la lista di (codice, tipo_tim, data) distinti
                            codici_date: set[tuple[str, str, str]] = set()
                            for rec in local_records:
                                codice = rec.get("codice_preparatore")
                                tipo_locale = rec.get("tipo_attivita")
                                data_prod = rec.get("data")
                                nome_locale = rec.get("nome_preparatore")
                                
                                if not codice or not tipo_locale or not data_prod:
                                    continue
                                    
                                # Converti la data
                                if isinstance(data_prod, datetime.datetime):
                                    data_rif = data_prod.date()
                                elif isinstance(data_prod, str):
                                    data_rif = datetime.datetime.strptime(data_prod, "%Y-%m-%d").date()
                                else:
                                    data_rif = data_prod
                                
                                tipo_tim = tipo_mapping.get(tipo_locale)
                                if tipo_tim:
                                    key = (codice.upper(), tipo_tim, str(data_rif))
                                    codici_date.add(key)
                                    if nome_locale:
                                        nome_str = str(nome_locale).strip()
                                        if nome_str:
                                            local_nome_map.setdefault(key, nome_str)
                            
                            # Query batch per tutte le durate (ottimizzato per data)
                            durate_map: Dict[tuple[str, str, str], Dict[str, Any]] = {}

                            codici_unici = sorted({cod.upper() for cod, _, _ in codici_date})
                            tipi_unici = sorted({tipo for _, tipo, _ in codici_date})
                            date_uniche = sorted({data for _, _, data in codici_date})

                            if codici_unici and tipi_unici and date_uniche:
                                codici_lower = [c.lower() for c in codici_unici]
                                code_placeholders = ", ".join(["%s"] * len(codici_lower))
                                tipo_placeholders = ", ".join(["%s"] * len(tipi_unici))

                                durate_query = f"""
                                    SELECT LOWER(cg.codice) AS codice,
                                           ta.descrizione AS tipo,
                                           %s AS data_riferimento,
                                           u.nome,
                                           u.cognome,
                                           COALESCE(SUM(a.durata), 0) AS durata_totale
                                    FROM codicegestionale cg
                                    JOIN utente u ON cg.utente_id = u.id
                                    JOIN tipoattivita ta ON cg.tipo_attivita_id = ta.id
                                    LEFT JOIN attivita a ON a.utente_id = u.id
                                        AND a.tipo_attivita_id = ta.id
                                        AND a.data_riferimento = %s
                                    WHERE LOWER(cg.codice) IN ({code_placeholders})
                                      AND ta.descrizione IN ({tipo_placeholders})
                                      AND %s BETWEEN cg.valido_dal AND COALESCE(cg.valido_al, '9999-12-31')
                                    GROUP BY codice, tipo, u.nome, u.cognome
                                """

                                for data_str in date_uniche:
                                    params: List[Any] = [data_str, data_str]
                                    params.extend(codici_lower)
                                    params.extend(tipi_unici)
                                    params.append(data_str)

                                    tim_cursor.execute(durate_query, params)
                                    rows = cast(List[Dict[str, Any]], tim_cursor.fetchall())

                                    for row in rows:
                                        codice_res = str(row.get("codice") or "").upper()
                                        tipo_res = str(row.get("tipo") or "")
                                        key = (codice_res, tipo_res, data_str)
                                        durata_totale = row.get("durata_totale") or 0
                                        durate_map[key] = {
                                            "nome": str(row.get("nome") or "").strip(),
                                            "cognome": str(row.get("cognome") or "").strip(),
                                            "durata": float(durata_totale),
                                        }

                                # Identifica codici mancanti (anomalia tipo 1)
                                missing_keys = codici_date - set(durate_map.keys())
                                for codice, tipo_tim, data_str in missing_keys:
                                    nome_locale = local_nome_map.get((codice, tipo_tim, data_str))
                                    nome_pulito = nome_locale.strip() if nome_locale else None
                                    dettaglio = {
                                        "codice": codice,
                                        "tipo": tipo_tim,
                                        "data": data_str,
                                        "motivo": "Codice non trovato in TIM",
                                    }
                                    if nome_pulito:
                                        dettaglio["nome"] = nome_pulito
                                    non_trovati_dettaglio.append(dettaglio)
                                    
                                    # Converti data_str in datetime.date per data_rilevamento
                                    try:
                                        data_anomalia = datetime.datetime.strptime(data_str, "%Y-%m-%d").date()
                                    except ValueError:
                                        data_anomalia = today
                                    
                                    insert_anomalia(
                                        tipo_anomalia="CODICE_NON_ABBINATO",
                                        data_rilevamento=data_anomalia,
                                        codice_preparatore=codice,
                                        nome_preparatore=nome_pulito,
                                        tipo_attivita=tipo_tim,
                                        ore_tim=None,
                                        dettagli=f"Data: {data_str} - Codice non trovato in TIM",
                                        note=None,
                                    )

                            print(f"✅ Recuperate durate per {len(durate_map)} combinazioni codice/tipo/data")
                            
                            # Recupera TUTTI i colli in una query
                            print("🚀 Recupero TUTTI i colli locali...")
                            
                            app_cursor.execute("""
                                SELECT codice_preparatore, tipo_attivita, data, tipo, totale_colli
                                FROM dati_produzione
                            """)
                            all_records = cast(List[Dict[str, Any]], app_cursor.fetchall())
                            
                            # Raggruppa per (codice, tipo_attivita, data)
                            colli_map = {}  # (codice, tipo, data) -> [(tipo_negozio, colli)]
                            for row in all_records:
                                codice = str(row.get("codice_preparatore") or "")
                                tipo_att = str(row.get("tipo_attivita") or "")
                                data = str(row.get("data") or "")
                                tipo_negozio = row.get("tipo")
                                colli = row.get("totale_colli") or 0
                                
                                tipo_tim = tipo_mapping.get(tipo_att)
                                if codice and tipo_tim and data:
                                    key = (codice.upper(), tipo_tim, data)
                                    if key not in colli_map:
                                        colli_map[key] = []
                                    colli_map[key].append((codice, tipo_att, tipo_negozio, colli))
                            
                            print(f"✅ Recuperati colli per {len(colli_map)} combinazioni")
                            
                            # Anomalia tipo 2: Ore TIM senza produzione per attività a premi
                            print("🔍 Controllo anomalie ore senza produzione...")
                            attivita_premi = ["PICKING", "CARRELLISTI", "MAG. RICEVIMENTO", "DOPPIA SPUNTA"]
                            
                            for key, tim_data in durate_map.items():
                                codice_upper, tipo_tim, data_str = key
                                
                                # Controlla solo attività a premi
                                if tipo_tim not in attivita_premi:
                                    continue
                                
                                # Se ci sono ore in TIM ma nessun dato di produzione locale
                                ore_tim = float(tim_data['durata'])
                                if ore_tim > 0 and key not in colli_map:
                                    # Anomalia: ore registrate in TIM ma nessuna produzione locale
                                    nome = tim_data['nome'].strip()
                                    cognome = tim_data['cognome'].strip()
                                    # Formato: "COGNOME NOME"
                                    if cognome and nome:
                                        nominativo = f"{cognome.upper()} {nome.upper()}"
                                    elif cognome:
                                        nominativo = cognome.upper()
                                    elif nome:
                                        nominativo = nome.upper()
                                    else:
                                        nominativo = ""
                                    
                                    ore_tim_decimal = ore_tim / 60.0  # Converti minuti in ore
                                    
                                    # Converti data_str in datetime.date per data_rilevamento
                                    try:
                                        data_anomalia = datetime.datetime.strptime(data_str, "%Y-%m-%d").date()
                                    except ValueError:
                                        data_anomalia = today
                                    
                                    insert_anomalia(
                                        tipo_anomalia="ORE_SENZA_PRODUZIONE",
                                        data_rilevamento=data_anomalia,
                                        codice_preparatore=codice_upper,
                                        nome_preparatore=nominativo or None,
                                        tipo_attivita=tipo_tim,
                                        ore_tim=ore_tim_decimal,
                                        dettagli=f"Data: {data_str} - {ore_tim} minuti ({ore_tim_decimal:.2f} ore) in TIM ma nessuna produzione locale",
                                        note=None
                                    )
                                    anomalie_ore_count += 1
                                    
                                    print(f"  ⚠️ Anomalia: {codice_upper} ({tipo_tim}) - {ore_tim} min in TIM, 0 colli locali (data: {data_str})")
                            
                            # Ora calcola e prepara gli update
                            updates_batch = []
                            
                            for idx, (key, records_list) in enumerate(colli_map.items()):
                                codice_upper, tipo_tim, data_str = key
                                
                                # Trova durata in TIM
                                tim_data = durate_map.get(key)
                                if not tim_data:
                                    continue
                                
                                nome = tim_data['nome'].strip()
                                cognome = tim_data['cognome'].strip()
                                # Formato: "COGNOME NOME"
                                if cognome and nome:
                                    nominativo = f"{cognome.upper()} {nome.upper()}"
                                elif cognome:
                                    nominativo = cognome.upper()
                                elif nome:
                                    nominativo = nome.upper()
                                else:
                                    nominativo = ""
                                durata_totale = float(tim_data['durata'])
                                
                                # Calcola totale colli
                                totale_colli_globale = sum(float(r[3] or 0) for r in records_list)
                                
                                # Proporziona per ogni negozio
                                for codice_orig, tipo_att, tipo_negozio, colli in records_list:
                                    if totale_colli_globale > 0:
                                        ore_prop = round((durata_totale * float(colli or 0)) / totale_colli_globale / 60.0, 2)
                                    else:
                                        ore_prop = 0.0
                                    
                                    updates_batch.append((nominativo, ore_prop, codice_orig, tipo_att, data_str, tipo_negozio))
                                
                                if idx % 100 == 0:
                                    percent = int((idx + 1) / len(colli_map) * 100)
                                    self._update_progress(percent)
                            
                            # Esegui TUTTI gli update in batch
                            print(f"🚀 Eseguo {len(updates_batch)} update...")
                            
                            update_query = """
                                UPDATE dati_produzione
                                SET nome_preparatore = %s,
                                    ore_tim = %s
                                WHERE LOWER(codice_preparatore) = LOWER(%s)
                                  AND tipo_attivita = %s
                                  AND data = %s
                                  AND tipo = %s
                            """
                            
                            app_cursor.executemany(update_query, updates_batch)
                            aggiornati = app_cursor.rowcount
                            
                            # GENERA ANOMALIE X/XX per TUTTE le attività dopo il sync
                            print(f"\n🔍 Controllo anomalie X/XX per tutte le attività...")
                            
                            # Query per raggruppare per data+codice+tipo_attivita e sommare ore
                            # ESCLUDE i record con ore_tim = 0 (che generano PRODUZIONE_SENZA_ORE)
                            check_query = """
                                SELECT data, 
                                       codice_preparatore, 
                                       tipo_attivita,
                                       SUM(CAST(ore_tim AS DECIMAL(10,2))) as ore_tim_totali, 
                                       SUM(CAST(ore_gestionale AS DECIMAL(10,2))) as ore_gestionale_totali,
                                       GROUP_CONCAT(DISTINCT tipo ORDER BY tipo SEPARATOR ', ') as tipi
                                FROM dati_produzione
                                WHERE ore_tim IS NOT NULL
                                  AND ore_tim > 0
                                  AND ore_gestionale IS NOT NULL
                                GROUP BY data, codice_preparatore, tipo_attivita
                                HAVING ABS((ore_gestionale_totali - ore_tim_totali) * 60) >= 60
                            """
                            
                            app_cursor.execute(check_query)
                            records_con_diff = app_cursor.fetchall()
                            
                            print(f"  Trovati {len(records_con_diff)} giorni con differenza >= 60 min")
                            
                            for row in records_con_diff:
                                # Row è un dizionario con dati aggregati
                                data = row['data']
                                codice = row['codice_preparatore']
                                tipo_attivita = row['tipo_attivita']
                                tipi = row['tipi']  # Lista dei tipi (ST, AP, CM)
                                ore_tim_totali = row['ore_tim_totali']
                                ore_gestionale_totali = row['ore_gestionale_totali']
                                
                                # Recupera nome e cognome da durate_map (da TIM)
                                data_str = data.strftime('%Y-%m-%d') if hasattr(data, 'strftime') else str(data)
                                durata_info = durate_map.get((codice.upper(), tipo_attivita, data_str))
                                
                                nome_formattato = None
                                if durata_info:
                                    nome_tim = durata_info.get("nome", "").strip()
                                    cognome_tim = durata_info.get("cognome", "").strip()
                                    if cognome_tim and nome_tim:
                                        nome_formattato = f"{cognome_tim.upper()} {nome_tim.upper()}"
                                    elif cognome_tim:
                                        nome_formattato = cognome_tim.upper()
                                    elif nome_tim:
                                        nome_formattato = nome_tim.upper()
                                
                                # Converti Decimal in float
                                ore_tim_val = float(ore_tim_totali) if ore_tim_totali is not None else 0.0
                                ore_gestionale_val = float(ore_gestionale_totali) if ore_gestionale_totali is not None else 0.0
                                differenza_minuti = (ore_gestionale_val - ore_tim_val) * 60
                                differenza_assoluta = abs(differenza_minuti)
                                
                                if differenza_assoluta >= 120:
                                    # Anomalia XX - differenza > 120 min
                                    from database import insert_anomalia
                                    insert_anomalia(
                                        tipo_anomalia='DIFFERENZA_>120',
                                        data_rilevamento=data,
                                        codice_preparatore=codice,
                                        nome_preparatore=nome_formattato,
                                        tipo_attivita=tipo_attivita,
                                        ore_tim=ore_tim_val,
                                        dettagli=f"Data: {data_str} - Ore TIM: {ore_tim_val:.2f}h, Ore Gestionale: {ore_gestionale_val:.2f}h - Differenza: {differenza_minuti:+.0f} min - Tipi: {tipi}"
                                    )
                                    anomalie_x_xx_count += 1
                                    print(f"  ⚠️ XX: {codice} ({nome_formattato}) {data_str} - {differenza_minuti:+.0f} min (TIM: {ore_tim_val:.2f}h, Gest: {ore_gestionale_val:.2f}h)")
                                elif differenza_assoluta >= 60:
                                    # Anomalia X - differenza 60-120 min
                                    from database import insert_anomalia
                                    insert_anomalia(
                                        tipo_anomalia='DIFFERENZA_60_120',
                                        data_rilevamento=data,
                                        codice_preparatore=codice,
                                        nome_preparatore=nome_formattato,
                                        tipo_attivita=tipo_attivita,
                                        ore_tim=ore_tim_val,
                                        dettagli=f"Data: {data_str} - Ore TIM: {ore_tim_val:.2f}h, Ore Gestionale: {ore_gestionale_val:.2f}h - Differenza: {differenza_minuti:+.0f} min - Tipi: {tipi}"
                                    )
                                    anomalie_x_xx_count += 1
                                    print(f"  ⚠️ X: {codice} ({nome_formattato}) {data_str} - {differenza_minuti:+.0f} min (TIM: {ore_tim_val:.2f}h, Gest: {ore_gestionale_val:.2f}h)")
                            
                            print(f"✅ Anomalie X/XX generate: {anomalie_x_xx_count}")
                            
                            # GENERA ANOMALIE PRODUZIONE_SENZA_ORE (0 ore TIM ma con ore gestionale)
                            print(f"\n🔍 Controllo anomalie PRODUZIONE_SENZA_ORE...")
                            
                            produzione_senza_ore_query = """
                                SELECT data, 
                                       codice_preparatore, 
                                       tipo_attivita,
                                       SUM(CAST(ore_gestionale AS DECIMAL(10,2))) as ore_gestionale_totali,
                                       GROUP_CONCAT(DISTINCT tipo ORDER BY tipo SEPARATOR ', ') as tipi
                                FROM dati_produzione
                                WHERE (ore_tim IS NULL OR ore_tim = 0)
                                  AND ore_gestionale > 0
                                GROUP BY data, codice_preparatore, tipo_attivita
                            """
                            
                            app_cursor.execute(produzione_senza_ore_query)
                            records_senza_tim = app_cursor.fetchall()
                            
                            print(f"  Trovati {len(records_senza_tim)} giorni con produzione senza ore TIM")
                            
                            anomalie_senza_ore_count = 0
                            for row in records_senza_tim:
                                data = row['data']
                                codice = row['codice_preparatore']
                                tipo_attivita = row['tipo_attivita']
                                tipi = row['tipi']
                                ore_gestionale_totali = row['ore_gestionale_totali']
                                
                                # Recupera nome da durate_map (da TIM)
                                data_str = data.strftime('%Y-%m-%d') if hasattr(data, 'strftime') else str(data)
                                durata_info = durate_map.get((codice.upper(), tipo_attivita, data_str))
                                
                                # SALTA se il codice non è in TIM (già generato CODICE_NON_ABBINATO)
                                if not durata_info:
                                    print(f"  ⏭️ SKIP: {codice} {data_str} - già gestito da CODICE_NON_ABBINATO")
                                    continue
                                
                                nome_formattato = None
                                nome_tim = durata_info.get("nome", "").strip()
                                cognome_tim = durata_info.get("cognome", "").strip()
                                if cognome_tim and nome_tim:
                                    nome_formattato = f"{cognome_tim.upper()} {nome_tim.upper()}"
                                elif cognome_tim:
                                    nome_formattato = cognome_tim.upper()
                                elif nome_tim:
                                    nome_formattato = nome_tim.upper()
                                
                                # Converti Decimal in float
                                ore_gestionale_val = float(ore_gestionale_totali) if ore_gestionale_totali is not None else 0.0
                                
                                from database import insert_anomalia
                                insert_anomalia(
                                    tipo_anomalia='PRODUZIONE_SENZA_ORE',
                                    data_rilevamento=data,
                                    codice_preparatore=codice,
                                    nome_preparatore=nome_formattato,
                                    tipo_attivita=tipo_attivita,
                                    ore_tim=0.0,
                                    dettagli=f"Data: {data_str} - Ore TIM: 0.00h, Ore Gestionale: {ore_gestionale_val:.2f}h - Tipi: {tipi}"
                                )
                                anomalie_senza_ore_count += 1
                                print(f"  ⚠️ PRODUZIONE_SENZA_ORE: {codice} ({nome_formattato}) {data_str} - Gest: {ore_gestionale_val:.2f}h - Tipi: {tipi}")
                            
                            print(f"✅ Anomalie PRODUZIONE_SENZA_ORE generate: {anomalie_senza_ore_count}")

                    app_conn.commit()
                    print(f"✅ Aggiornati {aggiornati} record!")
                    
        except mysql.connector.Error as err:
            return {
                "success": False,
                "message": f"Errore durante l'aggiornamento del database locale:\n{err}",
            }

        # Totale anomalie generate (X/XX + PRODUZIONE_SENZA_ORE, non i non trovati)
        totale_anomalie = anomalie_x_xx_count + anomalie_senza_ore_count
        
        # Log fine sincronizzazione
        log_final = f"\n{'='*60}\nFINE SINCRONIZZAZIONE\nRecord aggiornati: {aggiornati}\nRecord non trovati: {len(non_trovati_dettaglio)}\nAnomalie X/XX: {anomalie_x_xx_count}\nAnomalie PRODUZIONE_SENZA_ORE: {anomalie_senza_ore_count}\nTotale anomalie generate: {totale_anomalie}\n{'='*60}\n"
        print(log_final)
        
        # Scrivi su file
        try:
            with open("sync_log.txt", "a", encoding="utf-8") as f:
                f.write(log_final)
                if non_trovati_dettaglio:
                    f.write("\n=== RIEPILOGO CODICI NON TROVATI ===\n")
                    for item in non_trovati_dettaglio:
                        f.write(f"  Codice: {item['codice']}, Tipo: {item['tipo']}, Motivo: {item['motivo']}\n")
        except Exception as e:
            print(f"Errore scrittura log: {e}")

        return {
            "success": True,
            "no_data": False,
            "updated": aggiornati,
            "non_trovati_details": non_trovati_dettaglio,
            "anomalie_count": totale_anomalie,
        }

    def _on_sync_complete(self, result: Dict[str, Any]) -> None:
        """Ripristina l'interfaccia e mostra l'esito al termine della sincronizzazione."""
        if self.progress_window:
            self.progress_window.destroy()
            self.progress_window = None
        
        self.sync_progress = None
        self.sync_progress_label = None
        self.sync_button.config(state="normal")
        self._sync_in_progress = False

        if result.get("success"):
            filters_to_apply = self._last_filters.copy() if isinstance(self._last_filters, dict) else None
            reloaded = False
            try:
                self._load_data(filters_to_apply)
                reloaded = True
            except Exception:
                # eventuali errori sono già gestiti da _load_data
                pass
            if not reloaded:
                self.stats_label.config(text=self._stats_text_before_sync)

            # Calcola tempo trascorso
            elapsed_seconds = result.get('elapsed_time', 0)
            minutes = int(elapsed_seconds // 60)
            seconds = int(elapsed_seconds % 60)
            time_str = f"{minutes}m {seconds}s" if minutes > 0 else f"{seconds}s"

            if result.get("no_data"):
                message = "Nessun dato da sincronizzare dal database locale."
            else:
                message = f"Sincronizzazione completata in {time_str}\n\nRecord aggiornati: {result.get('updated', 0)}"

                anomalie_count = int(result.get("anomalie_count", 0) or 0)
                if anomalie_count:
                    message += "\n\nSono presenti anomalie registrate durante la sincronizzazione.\nApri la sezione Anomalie per gestirle."

                # Log dettagliato dei codici non trovati in TIM
                non_trovati = result.get("non_trovati_details") or []
                if non_trovati:
                    print("\n=== CODICI NON TROVATI IN TIM ===")
                    for item in non_trovati:
                        print(f"  Codice: {item['codice']}, Tipo: {item['tipo']}, Motivo: {item['motivo']}")
                    print(f"Totale non trovati: {len(non_trovati)}")
                    print("==================================\n")

            self.window.lift()
            self.window.focus_force()
            messagebox.showinfo("Sincronizzazione completata", message, parent=self.window)
        else:
            self.stats_label.config(text=self._stats_text_before_sync)
            self.window.lift()
            self.window.focus_force()
            messagebox.showerror(
                "Errore di sincronizzazione",
                result.get("message", "Errore sconosciuto durante la sincronizzazione."),
                parent=self.window,
            )

    def _gestisci_anomalie(self) -> None:
        """Apre la finestra di gestione anomalie."""
        try:
            # Import dinamico per evitare dipendenza circolare
            from anomalie_view import AnomalieView
            AnomalieView(parent=self.window)
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"ERRORE apertura anomalie:\n{error_details}")
            messagebox.showerror(
                "Errore",
                f"Impossibile aprire la gestione anomalie:\n{str(e)}\n\nVedi console per dettagli completi.",
                parent=self.window
            )

    def show(self) -> None:
        """Avvia il loop principale della finestra (solo modalità standalone)."""
        if self.is_standalone:
            self.window.mainloop()


# Alias per compatibilità con main_menu.py
DataViewerApp = DataViewer


def main() -> None:
    """Permette l'esecuzione standalone del visualizzatore."""
    viewer = DataViewer()
    viewer.show()


if __name__ == "__main__":
    main()
                            