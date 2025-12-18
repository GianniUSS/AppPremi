"""Dashboard delle performance dei preparatori.

Questa vista replica il mockup "Warehouse Performance" ed è pensata
per essere aperta come modulo standalone dal menù Performance.
"""
import datetime
from contextlib import closing
from decimal import Decimal
from typing import Any, Dict, List, Optional, Tuple, cast

import tkinter as tk
from tkinter import messagebox, ttk

import mysql.connector

from config import COLORS, FONTS, MYSQL_CONFIG
from ui_components import create_button


MONTH_CHOICES: List[Tuple[str, int]] = [
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
MONTH_LABEL_TO_VALUE = {label: value for label, value in MONTH_CHOICES}


class PerformancePreparatoriView(tk.Frame):
    """Vista interattiva per analizzare le performance dei preparatori."""

    def __init__(self, parent: tk.Widget):
        super().__init__(parent, bg=COLORS["background"])
        self.pack(fill="both", expand=True)

        today = datetime.date.today()
        self.anno_var = tk.StringVar(value=str(today.year))
        self.mese_var = tk.StringVar(value=MONTH_CHOICES[today.month - 1][0])
        self.annual_mode_var = tk.BooleanVar(value=False)

        self._current_stats: List[Dict[str, Any]] = []

        self._build_ui()
        self._load_data()

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=COLORS["background"])
        header.pack(fill="x", padx=20, pady=(12, 6))

        tk.Label(
            header,
            text="📈 Performance Preparatori",
            font=FONTS.get("title", ("Segoe UI", 20, "bold")),
            bg=COLORS["background"],
            fg=COLORS.get("primary", "#1f77b4"),
        ).pack(anchor="w")

        controls_frame = tk.Frame(self, bg=COLORS["background"])
        controls_frame.pack(fill="x", padx=20, pady=(0, 10))

        filter_frame = tk.Frame(controls_frame, bg=COLORS["background"], bd=1, relief="groove")
        filter_frame.pack(side="left", fill="x", expand=True)

        self.summary_container = tk.Frame(controls_frame, bg=COLORS["background"], width=420)
        self.summary_container.pack(side="right", padx=(12, 0))

        tk.Label(
            filter_frame,
            text="Anno:",
            font=FONTS.get("label", ("Segoe UI", 11)),
            bg=COLORS["background"],
        ).grid(row=0, column=0, sticky="w", padx=(12, 6), pady=12)

        current_year = datetime.date.today().year
        anni = [str(y) for y in range(current_year - 5, current_year + 2)]
        anno_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.anno_var,
            values=anni,
            width=10,
            state="readonly",
            font=FONTS.get("input", ("Segoe UI", 11)),
        )
        anno_combo.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=8)
        anno_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_data())

        tk.Label(
            filter_frame,
            text="Mese:",
            font=FONTS.get("label", ("Segoe UI", 11)),
            bg=COLORS["background"],
        ).grid(row=0, column=2, sticky="w", padx=(0, 6), pady=12)

        mesi = [label for label, _ in MONTH_CHOICES]
        self.mese_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.mese_var,
            values=mesi,
            width=16,
            state="readonly",
            font=FONTS.get("input", ("Segoe UI", 11)),
        )
        self.mese_combo.grid(row=0, column=3, sticky="w", padx=(0, 16), pady=8)
        self.mese_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_data())

        create_button(
            filter_frame,
            text="Aggiorna",
            command=self._load_data,
            variant="primary",
            width=14,
        ).grid(row=0, column=4, padx=(0, 12), pady=8)

        ttk.Checkbutton(
            filter_frame,
            text="Anno intero",
            variable=self.annual_mode_var,
            command=self._on_period_mode_toggle,
        ).grid(row=0, column=5, padx=(0, 12), pady=8)

        filter_frame.grid_columnconfigure(6, weight=1)

        self.status_label = tk.Label(
            self,
            text="",
            font=FONTS.get("subtitle", ("Segoe UI", 11)),
            bg=COLORS["background"],
            fg=COLORS.get("text_light", "#606060"),
        )
        self.status_label.pack(fill="x", padx=20, pady=(0, 6))

        self._build_dashboard()

    def _on_period_mode_toggle(self) -> None:
        if hasattr(self, "mese_combo"):
            state = "disabled" if self.annual_mode_var.get() else "readonly"
            self.mese_combo.configure(state=state)
        self._load_data()

    def _build_dashboard(self) -> None:
        self.dashboard_frame = tk.Frame(self, bg=COLORS["background"])
        self.dashboard_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.summary_vars = {
            "tot_colli": tk.StringVar(value="0"),
            "colli_hour": tk.StringVar(value="0,00"),
            "ore_totali": tk.StringVar(value="0,00"),
        }

        summary_parent = getattr(self, "summary_container", self.dashboard_frame)
        for child in summary_parent.winfo_children():
            child.destroy()
        summary_outer = tk.Frame(
            summary_parent,
            bg=COLORS["white"],
            padx=8,
            pady=8,
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        summary_outer.pack(fill="both", expand=True)

        summary_frame = tk.Frame(summary_outer, bg=COLORS["white"])
        summary_frame.pack(fill="both", expand=True)
        for col in range(3):
            summary_frame.grid_columnconfigure(col, weight=1)

        def _make_card(idx: int, title: str, var: tk.StringVar) -> None:
            card = tk.Frame(
                summary_frame,
                bg=COLORS["white"],
                padx=10,
                pady=6,
                highlightbackground=COLORS["border"],
                highlightthickness=1,
            )
            card.grid(row=0, column=idx, sticky="nsew", padx=6, pady=2)
            tk.Label(
                card,
                text=title,
                font=FONTS.get("label", ("Segoe UI", 9, "bold")),
                bg=COLORS["white"],
                fg=COLORS.get("text_light", "#777"),
            ).pack(anchor="w")
            tk.Label(
                card,
                textvariable=var,
                font=FONTS.get("title", ("Segoe UI", 16, "bold")),
                bg=COLORS["white"],
                fg=COLORS.get("primary", "#1f77b4"),
            ).pack(anchor="w", pady=(6, 0))

        _make_card(0, "Totale Colli", self.summary_vars["tot_colli"])
        _make_card(1, "Media Colli/Ora", self.summary_vars["colli_hour"])
        _make_card(2, "Ore Lavorate", self.summary_vars["ore_totali"])

        body_frame = tk.Frame(self.dashboard_frame, bg=COLORS["background"])
        body_frame.pack(fill="both", expand=True)

        charts_column = tk.Frame(body_frame, bg=COLORS["background"])
        charts_column.pack(side="left", fill="both", expand=True)

        table_column = tk.Frame(body_frame, bg=COLORS["background"])
        table_column.pack(side="right", fill="y", padx=(12, 0))

        def _make_chart(parent: tk.Frame, title: str, height: int = 220, scrollable: bool = False) -> tk.Canvas:
            section = tk.Frame(
                parent,
                bg=COLORS["white"],
                highlightbackground=COLORS["border"],
                highlightthickness=1,
            )
            section.pack(fill="both", expand=True, pady=(0, 12))
            tk.Label(
                section,
                text=title,
                font=FONTS.get("label", ("Segoe UI", 11, "bold")),
                bg=COLORS["white"],
                fg=COLORS.get("text", "#222"),
            ).pack(anchor="w", padx=12, pady=(10, 0))
            if scrollable:
                container = tk.Frame(section, bg=COLORS["white"])
                container.pack(fill="both", expand=True, padx=10, pady=10)
                canvas = tk.Canvas(
                    container,
                    height=height,
                    bg=COLORS["white"],
                    highlightthickness=0,
                )
                canvas.pack(side="left", fill="both", expand=True)
                scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
                scrollbar.pack(side="right", fill="y")
                canvas.configure(yscrollcommand=scrollbar.set)
                setattr(canvas, "_scrollable", True)
            else:
                canvas = tk.Canvas(
                    section,
                    height=height,
                    bg=COLORS["white"],
                    highlightthickness=0,
                )
                canvas.pack(fill="both", expand=True, padx=10, pady=10)
            return canvas

        self.bar_chart = _make_chart(charts_column, "Colli/Ora per Preparatore", 320, scrollable=True)
        self.line_chart = _make_chart(charts_column, "Colli/Ora - Andamento Giornaliero", 220)
        self.variance_chart = _make_chart(charts_column, "Scostamento Ore Giornaliere", 210)

        table_section = tk.Frame(
            table_column,
            bg=COLORS["white"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        table_section.pack(fill="both", expand=True)

        tk.Label(
            table_section,
            text="Preparatori (tutti gli operatori)",
            font=FONTS.get("label", ("Segoe UI", 11, "bold")),
            bg=COLORS["white"],
            fg=COLORS.get("text", "#222"),
        ).pack(anchor="w", padx=12, pady=(10, 0))

        columns = ("codice", "nome", "tot_colli", "ore", "colli_ora")
        self.tree = ttk.Treeview(table_section, columns=columns, show="headings", height=18)
        headings = {
            "codice": "Codice",
            "nome": "Preparatore",
            "tot_colli": "Tot. Colli",
            "ore": "Ore",
            "colli_ora": "Colli/h",
        }
        for col, title in headings.items():
            self.tree.heading(col, text=title)

        self.tree.column("codice", width=100, anchor="center")
        self.tree.column("nome", width=200, anchor="w")
        self.tree.column("tot_colli", width=110, anchor="center")
        self.tree.column("ore", width=90, anchor="center")
        self.tree.column("colli_ora", width=100, anchor="center")

        scrollbar = ttk.Scrollbar(table_section, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=10)
        scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)

    def _load_data(self) -> None:
        anno_str = self.anno_var.get().strip()
        mese_label = self.mese_var.get().strip()
        annual_mode = self.annual_mode_var.get()
        try:
            anno = int(anno_str)
        except ValueError:
            messagebox.showwarning("Periodo non valido", "Inserisci un anno corretto.", parent=self)
            return

        mese: Optional[int] = None
        if not annual_mode:
            mese = MONTH_LABEL_TO_VALUE.get(mese_label)
            if mese is None:
                messagebox.showwarning("Periodo non valido", "Seleziona un mese valido.", parent=self)
                return

        try:
            stats = self._fetch_operator_stats(anno, mese)
            self._current_stats = stats.copy()
        except Exception as exc:
            messagebox.showerror("Errore", f"Impossibile leggere i dati di produzione:\n{exc}", parent=self)
            self._current_stats = []
            self._reset_dashboard()
            return

        if not stats:
            if mese:
                periodo = f"{mese_label} {anno}"
            else:
                periodo = f"Anno {anno} (tutti i mesi)"
            self.status_label.config(
                text=f"Nessun dato di produzione per {periodo}. Importa i dati per analizzare le performance.",
            )
        else:
            if mese:
                periodo = f"{mese_label} {anno}"
            else:
                periodo = f"Anno {anno} (tutti i mesi)"
            self.status_label.config(
                text=f"Periodo: {periodo} • Operatori analizzati: {len(stats)}",
            )

        self._refresh_dashboard(anno, mese)

    def _refresh_dashboard(self, anno: Optional[int], mese: Optional[int]) -> None:
        stats = self._current_stats or []
        if not stats:
            self._reset_dashboard()
            return

        collis = Decimal("0")
        ore_totali = Decimal("0")
        parsed: List[Dict[str, Any]] = []

        for row in stats:
            totale_colli = Decimal(str(row.get("tot_colli") or 0))
            ore = Decimal(str(row.get("ore") or 0))
            media = Decimal(str(row.get("colli_ora") or 0))

            collis += totale_colli
            ore_totali += ore

            parsed.append(
                {
                    "nome": row.get("nome"),
                    "codice": row.get("codice"),
                    "tot_colli": totale_colli,
                    "ore": ore,
                    "colli_ora": media,
                }
            )

        media_colli_ora = (collis / ore_totali) if ore_totali > 0 else Decimal("0")

        self.summary_vars["tot_colli"].set(self._format_number(collis, 0))
        self.summary_vars["colli_hour"].set(self._format_number(media_colli_ora, 2))
        self.summary_vars["ore_totali"].set(self._format_number(ore_totali, 2))

        for item in self.tree.get_children():
            self.tree.delete(item)

        sorted_rows = sorted(parsed, key=lambda x: x["colli_ora"], reverse=True)
        for row in sorted_rows:
            self.tree.insert(
                "",
                "end",
                values=(
                    row["codice"],
                    row["nome"],
                    self._format_number(row["tot_colli"], 0),
                    self._format_number(row["ore"], 2),
                    self._format_number(row["colli_ora"], 2),
                ),
            )

        bar_data = [(row["nome"] or row["codice"], float(row["colli_ora"])) for row in sorted_rows]
        self._render_bar_chart(self.bar_chart, bar_data)

        timeseries: List[Dict[str, Any]] = []
        if anno:
            timeseries = self._fetch_timeseries(anno, mese)

        line_data = [(entry["label"], entry["colli_ora"]) for entry in timeseries if entry["colli_ora"] > 0]
        self._render_line_chart(self.line_chart, line_data)

        variance_data: List[Tuple[str, float]] = []
        if timeseries:
            avg_hours = sum(entry["ore"] for entry in timeseries) / len(timeseries)
            for entry in timeseries:
                variance_data.append((entry["label"], entry["ore"] - avg_hours))
        self._render_variance_chart(self.variance_chart, variance_data)

    def _reset_dashboard(self) -> None:
        self.summary_vars["tot_colli"].set("0")
        self.summary_vars["colli_hour"].set("0,00")
        self.summary_vars["ore_totali"].set("0,00")
        for widget in (self.tree, self.bar_chart, self.line_chart, self.variance_chart):
            if isinstance(widget, ttk.Treeview):
                for item in widget.get_children():
                    widget.delete(item)
            else:
                widget.delete("all")
                width = int(widget.winfo_width() or widget["width"])
                height = int(widget.winfo_height() or widget["height"])
                widget.create_text(
                    width / 2,
                    height / 2,
                    text="Nessun dato",
                    fill=COLORS.get("text_light", "#777"),
                    font=("Segoe UI", 10, "italic"),
                )

    def _format_number(self, value: Any, decimals: int = 2) -> str:
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            return "-"
        if decimals == 0:
            formatted = f"{numeric:,.0f}"
        else:
            formatted = f"{numeric:,.{decimals}f}"
        return formatted.replace(",", "@").replace(".", ",").replace("@", ".")

    def _fetch_operator_stats(self, anno: int, mese: Optional[int]) -> List[Dict[str, Any]]:
        query = [
            "SELECT",
            "    dp.codice_preparatore AS codice,",
            "    dp.nome_preparatore AS nome,",
            "    SUM(dp.totale_colli) AS tot_colli,",
            "    SUM(dp.ore_tim) AS tot_minuti",
            "FROM dati_produzione dp",
            "WHERE dp.tipo_attivita = 'PICKING'",
            "  AND YEAR(dp.data) = %s",
        ]
        params: List[Any] = [anno]
        if mese is not None:
            query.append("  AND MONTH(dp.data) = %s")
            params.append(mese)
        query.append("GROUP BY dp.codice_preparatore, dp.nome_preparatore")
        sql = "\n".join(query)

        with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
            with closing(conn.cursor(dictionary=True)) as cur:
                cur.execute(sql, params)
                rows = cast(List[Dict[str, Any]], cur.fetchall())

        stats: List[Dict[str, Any]] = []
        for row in rows:
            codice = (row.get("codice") or "").strip() or "-"
            nome = (row.get("nome") or "").strip() or codice
            coll = Decimal(str(row.get("tot_colli") or 0))
            minuti = Decimal(str(row.get("tot_minuti") or 0))
            ore = minuti / Decimal("60") if minuti else Decimal("0")
            media = (coll / ore) if ore > 0 else Decimal("0")
            stats.append(
                {
                    "codice": codice,
                    "nome": nome,
                    "tot_colli": coll,
                    "ore": ore,
                    "colli_ora": media,
                }
            )

        return stats

    def _fetch_timeseries(self, anno: int, mese: Optional[int]) -> List[Dict[str, Any]]:
        if mese is not None:
            query = """
                SELECT
                    DATE(dp.data) AS periodo,
                    SUM(dp.totale_colli) AS tot_colli,
                    SUM(dp.ore_tim) AS tot_minuti
                FROM dati_produzione dp
                WHERE dp.tipo_attivita = 'PICKING'
                  AND YEAR(dp.data) = %s
                  AND MONTH(dp.data) = %s
                GROUP BY periodo
                ORDER BY periodo
            """
            params = (anno, mese)
        else:
            query = """
                SELECT
                    MONTH(dp.data) AS periodo,
                    SUM(dp.totale_colli) AS tot_colli,
                    SUM(dp.ore_tim) AS tot_minuti
                FROM dati_produzione dp
                WHERE dp.tipo_attivita = 'PICKING'
                  AND YEAR(dp.data) = %s
                GROUP BY periodo
                ORDER BY periodo
            """
            params = (anno,)

        try:
            with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
                with closing(conn.cursor(dictionary=True)) as cur:
                    cur.execute(query, params)
                    rows = cast(List[Dict[str, Any]], cur.fetchall())
        except mysql.connector.Error as exc:
            print(f"[PerformancePreparatori] Errore fetch timeseries: {exc}")
            return []

        dataset: List[Dict[str, Any]] = []
        for row in rows:
            coll = Decimal(str(row.get("tot_colli") or 0))
            minuti = Decimal(str(row.get("tot_minuti") or 0))
            ore = (minuti / Decimal("60")) if minuti else Decimal("0")
            media = (coll / ore) if ore > 0 else Decimal("0")
            periodo = row.get("periodo")
            if mese is not None:
                label = periodo.strftime("%d/%m") if hasattr(periodo, "strftime") else str(periodo)
            else:
                try:
                    periodo_int = int(periodo)
                except (TypeError, ValueError):
                    periodo_int = None
                if periodo_int and 1 <= periodo_int <= 12:
                    label = MONTH_CHOICES[periodo_int - 1][0][:3]
                else:
                    label = str(periodo)
            dataset.append(
                {
                    "label": label,
                    "colli": float(coll),
                    "ore": float(ore),
                    "colli_ora": float(media),
                }
            )
        return dataset

    def _render_bar_chart(self, canvas: tk.Canvas, data: List[Tuple[str, float]]) -> None:
        canvas.delete("all")
        width = int(canvas.winfo_width() or canvas["width"])
        height = int(canvas.winfo_height() or canvas["height"])
        scrollable = bool(getattr(canvas, "_scrollable", False))
        if not data:
            canvas.create_text(
                width / 2,
                height / 2,
                text="Nessun dato",
                fill=COLORS.get("text_light", "#777"),
            )
            canvas.configure(scrollregion=(0, 0, width, height))
            return

        # Usa barre orizzontali per evitare sovrapposizioni sui nomi.
        left_padding = 140
        right_padding = 60
        top_padding = 20
        bottom_padding = 20
        available_width = max(width - left_padding - right_padding, 10)

        max_val = max(value for _, value in data) or 1
        if scrollable:
            row_height = 34
            content_height = top_padding + len(data) * row_height + bottom_padding
        else:
            available_height = max(height - top_padding - bottom_padding, 10)
            row_height = available_height / max(len(data), 1)
            content_height = height

        for idx, (label, value) in enumerate(data):
            y_center = top_padding + idx * row_height + row_height / 2
            bar_length = (value / max_val) * available_width
            x0 = left_padding
            x1 = left_padding + bar_length
            y0 = y_center - row_height * 0.35
            y1 = y_center + row_height * 0.35

            canvas.create_rectangle(
                x0,
                y0,
                x1,
                y1,
                fill=COLORS.get("primary", "#1f77b4"),
                width=0,
            )
            canvas.create_text(
                x0 - 6,
                y_center,
                text=(label or "?")[:22],
                anchor="e",
                font=("Segoe UI", 9, "bold"),
                fill=COLORS.get("text", "#222"),
            )
            canvas.create_text(
                x1 + 6,
                y_center,
                text=f"{value:.1f}".replace(".", ","),
                anchor="w",
                font=("Segoe UI", 9),
                fill=COLORS.get("text_light", "#555"),
            )

        canvas.create_line(
            left_padding,
            content_height - bottom_padding,
            width - right_padding,
            content_height - bottom_padding,
            fill=COLORS.get("border", "#ccc"),
        )
        canvas.configure(scrollregion=(0, 0, width, content_height))

    def _render_line_chart(self, canvas: tk.Canvas, data: List[Tuple[str, float]]) -> None:
        canvas.delete("all")
        width = int(canvas.winfo_width() or canvas["width"])
        height = int(canvas.winfo_height() or canvas["height"])
        if not data:
            canvas.create_text(width / 2, height / 2, text="Nessun dato", fill=COLORS.get("text_light", "#777"))
            return

        padding = 40
        max_val = max(value for _, value in data) or 1
        min_val = min(value for _, value in data)
        range_val = max_val - min_val or 1
        plot_width = width - 2 * padding
        plot_height = height - 2 * padding

        points: List[Tuple[float, float, str, float]] = []
        for idx, (label, value) in enumerate(data):
            x = padding + (idx / max(len(data) - 1, 1)) * plot_width
            y = padding + (1 - (value - min_val) / range_val) * plot_height
            points.append((x, y, label, value))

        for idx in range(len(points) - 1):
            canvas.create_line(
                points[idx][0],
                points[idx][1],
                points[idx + 1][0],
                points[idx + 1][1],
                fill=COLORS.get("accent", "#2ca02c"),
                width=2,
            )

        for x, y, label, value in points:
            canvas.create_oval(x - 3, y - 3, x + 3, y + 3, fill=COLORS.get("accent", "#2ca02c"), width=0)
            canvas.create_text(
                x,
                y - 10,
                text=f"{value:.1f}".replace(".", ","),
                font=("Segoe UI", 8),
                fill=COLORS.get("text", "#333"),
            )

        canvas.create_line(padding, height - padding, width - padding, height - padding, fill=COLORS.get("border", "#ccc"))
        for idx, (label, _value) in enumerate(data):
            x = padding + (idx / max(len(data) - 1, 1)) * plot_width
            canvas.create_text(
                x,
                height - padding + 12,
                text=label,
                font=("Segoe UI", 8),
                fill=COLORS.get("text_light", "#666"),
            )

    def _render_variance_chart(self, canvas: tk.Canvas, data: List[Tuple[str, float]]) -> None:
        canvas.delete("all")
        width = int(canvas.winfo_width() or canvas["width"])
        height = int(canvas.winfo_height() or canvas["height"])
        if not data:
            canvas.create_text(width / 2, height / 2, text="Nessun dato", fill=COLORS.get("text_light", "#777"))
            return

        padding = 40
        max_abs = max(abs(value) for _, value in data) or 1
        plot_width = width - 2 * padding
        zero_y = height / 2
        bar_width = plot_width / max(len(data), 1)

        canvas.create_line(padding, zero_y, width - padding, zero_y, fill=COLORS.get("border", "#ccc"))

        for idx, (label, value) in enumerate(data):
            x0 = padding + idx * bar_width + 4
            x1 = x0 + max(10, bar_width - 8)
            bar_height = (abs(value) / max_abs) * (height / 2 - 20)
            if value >= 0:
                y0 = zero_y - bar_height
                y1 = zero_y
                color = "#3ba272"
            else:
                y0 = zero_y
                y1 = zero_y + bar_height
                color = "#e15759"
            canvas.create_rectangle(x0, y0, x1, y1, fill=color, width=0)
            canvas.create_text(
                (x0 + x1) / 2,
                zero_y + 12 if value >= 0 else zero_y + bar_height + 12,
                text=label,
                font=("Segoe UI", 8),
                fill=COLORS.get("text_light", "#555"),
            )
