"""Vista per consultazione dati TIM (attività dal gestionale)."""

import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict, Any

from config import COLORS, FONTS
from ui_components import create_button
from database import fetch_attivita_tim_range


MONTH_CHOICES = [
    ("Gennaio", 1), ("Febbraio", 2), ("Marzo", 3), ("Aprile", 4),
    ("Maggio", 5), ("Giugno", 6), ("Luglio", 7), ("Agosto", 8),
    ("Settembre", 9), ("Ottobre", 10), ("Novembre", 11), ("Dicembre", 12),
]


class TimView(tk.Frame):
    """Visualizzatore dati attività TIM con filtri data e ricerca."""

    def __init__(self, parent: tk.Widget):
        super().__init__(parent, bg=COLORS["background"])
        self.pack(fill="both", expand=True)

        self._sort_reverse: Dict[str, bool] = {}

        self._build_ui()
        self._carica_dati()

    def _build_ui(self) -> None:
        # Title
        tk.Label(
            self,
            text="\u2699 Dati TIM",
            font=FONTS["title"],
            bg=COLORS["background"],
            fg=COLORS["text_dark"],
        ).pack(pady=(12, 6), anchor="w", padx=16)

        # Filters
        filter_frame = tk.Frame(self, bg=COLORS["background"])
        filter_frame.pack(fill="x", padx=16, pady=(0, 8))

        tk.Label(
            filter_frame, text="Data Da:", font=FONTS["label"], bg=COLORS["background"],
        ).grid(row=0, column=0, sticky="w", padx=(0, 6), pady=8)

        self.data_da_var = tk.StringVar()
        ttk.Entry(
            filter_frame, textvariable=self.data_da_var, width=12, font=FONTS["input"],
        ).grid(row=0, column=1, sticky="ew", padx=(0, 16), pady=8)

        tk.Label(
            filter_frame, text="Data A:", font=FONTS["label"], bg=COLORS["background"],
        ).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=8)

        self.data_a_var = tk.StringVar()
        ttk.Entry(
            filter_frame, textvariable=self.data_a_var, width=12, font=FONTS["input"],
        ).grid(row=0, column=3, sticky="ew", padx=(0, 16), pady=8)

        tk.Label(
            filter_frame, text="\U0001f50d Ricerca:", font=FONTS["label"], bg=COLORS["background"],
        ).grid(row=0, column=4, sticky="w", padx=(0, 6), pady=8)

        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(
            filter_frame, textvariable=self.search_var, width=24, font=FONTS["input"],
        )
        search_entry.grid(row=0, column=5, sticky="ew", padx=(0, 16), pady=8)
        search_entry.bind("<Return>", lambda e: self._carica_dati())

        create_button(
            filter_frame, text="Cerca", command=self._carica_dati, variant="primary", width=12,
        ).grid(row=0, column=6, padx=(6, 6), pady=8)

        create_button(
            filter_frame, text="Pulisci", command=self._pulisci_filtri, variant="secondary", width=12,
        ).grid(row=0, column=7, padx=(6, 6), pady=8)

        filter_frame.grid_columnconfigure(5, weight=1)

        # Set default dates: current month
        today = datetime.date.today()
        first_day = today.replace(day=1)
        self.data_da_var.set(first_day.strftime("%d/%m/%Y"))
        self.data_a_var.set(today.strftime("%d/%m/%Y"))

        # Table
        table_frame = tk.Frame(
            self,
            bg=COLORS["white"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        table_frame.pack(fill="both", expand=True, padx=16, pady=(0, 0))

        columns = (
            "codice", "nome", "data", "attivita",
            "inizio", "fine", "ore", "pausa",
        )
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        headings = {
            "codice": "Codice",
            "nome": "Nome",
            "data": "Data",
            "attivita": "Attivita",
            "inizio": "Inizio",
            "fine": "Fine",
            "ore": "Ore",
            "pausa": "Pausa (min)",
        }
        widths = {
            "codice": 80,
            "nome": 220,
            "data": 100,
            "attivita": 200,
            "inizio": 70,
            "fine": 70,
            "ore": 70,
            "pausa": 80,
        }
        anchors = {
            "codice": "w",
            "nome": "w",
            "data": "center",
            "attivita": "w",
            "inizio": "center",
            "fine": "center",
            "ore": "center",
            "pausa": "center",
        }

        for col in columns:
            self.tree.heading(col, text=headings[col], command=lambda c=col: self._sort_by_column(c))
            self.tree.column(col, width=widths[col], anchor=anchors[col])

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Footer
        footer_frame = tk.Frame(self, bg=COLORS["primary"])
        footer_frame.pack(fill="x", padx=16, pady=(4, 0))

        self.stats_label = tk.Label(
            footer_frame,
            text="",
            font=("Segoe UI", 12, "bold"),
            bg=COLORS["primary"],
            fg="white",
        )
        self.stats_label.pack(side="left", padx=12, pady=10)

    def _parse_date(self, text: str) -> datetime.date | None:
        text = text.strip()
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                return datetime.datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        return None

    def _carica_dati(self) -> None:
        data_da = self._parse_date(self.data_da_var.get())
        data_a = self._parse_date(self.data_a_var.get())

        if not data_da or not data_a:
            messagebox.showwarning(
                "Date non valide",
                "Inserisci date valide nel formato GG/MM/AAAA.",
                parent=self,
            )
            return

        if data_da > data_a:
            data_da, data_a = data_a, data_da

        search = self.search_var.get().strip() or None

        try:
            rows = fetch_attivita_tim_range(data_da, data_a, search=search)
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile recuperare i dati TIM:\n{exc}",
                parent=self,
            )
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        totale_ore = 0.0
        for row in rows:
            codice = row.get("codice_preparatore") or ""
            nome = row.get("nome_preparatore") or ""
            data_rif = row.get("data_riferimento")
            data_str = data_rif.strftime("%d/%m/%Y") if hasattr(data_rif, "strftime") else str(data_rif or "")
            attivita = row.get("attivita") or ""
            inizio = row.get("data_inizio")
            fine = row.get("data_fine")
            inizio_str = inizio.strftime("%H:%M") if hasattr(inizio, "strftime") else str(inizio or "")
            fine_str = fine.strftime("%H:%M") if hasattr(fine, "strftime") else str(fine or "")
            minuti = row.get("durata_minuti") or 0
            ore = float(minuti) / 60.0
            pausa = row.get("pausa") or 0

            totale_ore += ore

            self.tree.insert(
                "", "end",
                values=(
                    codice, nome, data_str, attivita,
                    inizio_str, fine_str, f"{ore:.2f}", str(pausa),
                ),
            )

        n = len(rows)
        self.stats_label.config(
            text=f"Righe: {n} | Totale Ore: {totale_ore:,.2f}"
        )

    def _pulisci_filtri(self) -> None:
        today = datetime.date.today()
        first_day = today.replace(day=1)
        self.data_da_var.set(first_day.strftime("%d/%m/%Y"))
        self.data_a_var.set(today.strftime("%d/%m/%Y"))
        self.search_var.set("")
        self._carica_dati()

    def _sort_by_column(self, col: str) -> None:
        reverse = self._sort_reverse.get(col, False)
        items = list(self.tree.get_children(""))

        col_index = list(self.tree["columns"]).index(col)

        def sort_key(item_id: str):
            val = self.tree.set(item_id, col)
            try:
                return float(val.replace(",", ""))
            except (ValueError, AttributeError):
                return val.lower() if isinstance(val, str) else val

        items.sort(key=sort_key, reverse=reverse)
        for idx, item_id in enumerate(items):
            self.tree.move(item_id, "", idx)

        self._sort_reverse[col] = not reverse
