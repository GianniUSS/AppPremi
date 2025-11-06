"""
Interfaccia principale dell'applicazione con menu laterale
"""
import tkinter as tk
from tkinter import ttk, messagebox

from updater import open_update_dialog
from version import APP_VERSION


class CollapsibleMenuSection(tk.Frame):
    """Sezione del menu laterale apribile e richiudibile."""

    def __init__(
        self,
        master,
        title: str,
        *,
        bg: str,
        content_bg: str,
        text_color: str,
        highlight_color: str,
        expanded: bool = False,
    ) -> None:
        super().__init__(master, bg=bg)
        self._title = title
        self._bg = bg
        self._content_bg = content_bg
        self._text_color = text_color
        self._highlight_color = highlight_color
        self._expanded = False

        self.header_button = tk.Button(
            self,
            text=self._header_text(expanded),
            font=("Segoe UI", 12, "bold"),
            bg=bg,
            fg=text_color,
            activebackground=highlight_color,
            activeforeground=text_color,
            relief="flat",
            bd=0,
            cursor="hand2",
            anchor="w",
            padx=16,
            pady=12,
            command=self.toggle,
        )
        self.header_button.pack(fill=tk.X)

        self.content_frame = tk.Frame(self, bg=content_bg)
        if expanded:
            self.expand()

    def _header_text(self, expanded: bool) -> str:
        prefix = "[-]" if expanded else "[+]"
        return f"{prefix} {self._title}"

    def toggle(self) -> None:
        if self._expanded:
            self.collapse()
        else:
            self.expand()

    def expand(self) -> None:
        if not self._expanded:
            self.content_frame.pack(fill=tk.X)
            self.header_button.config(
                text=self._header_text(True),
                bg=self._highlight_color,
            )
            self._expanded = True

    def collapse(self) -> None:
        if self._expanded:
            self.content_frame.pack_forget()
            self.header_button.config(
                text=self._header_text(False),
                bg=self._bg,
            )
            self._expanded = False

    def add_button(
        self,
        text: str,
        command,
        *,
        font,
        fg: str,
        bg: str,
        hover_bg: str,
    ) -> tk.Button:
        btn = tk.Button(
            self.content_frame,
            text=text,
            font=font,
            bg=bg,
            fg=fg,
            activebackground=hover_bg,
            activeforeground=fg,
            relief="flat",
            bd=0,
            cursor="hand2",
            anchor="w",
            padx=36,
            pady=12,
        )
        btn.configure(command=command)
        btn.pack(fill=tk.X, pady=1)
        return btn


class MainMenu(tk.Tk):
    """Finestra principale con menu laterale"""
    
    def __init__(self):
        super().__init__()
        
        self.title("Gestione Premi di Produzione")
        self.geometry("1400x800")
        try:
            self.state("zoomed")
        except tk.TclError:
            # Alcuni sistemi potrebbero non supportare lo zoom, quindi ignoriamo l'errore
            pass
        
        # Variabile per tracciare il modulo corrente
        self.current_module = None
        self.current_frame = None
        
        # Configura lo stile
        self._setup_style()
        
        # Crea il layout principale
        self._create_layout()
        
    # All'avvio non mostriamo nessun modulo: l'area rimane vuota
    
    def _setup_style(self):
        """Configura lo stile dell'applicazione"""
        style = ttk.Style()
        
        # Stile per i pulsanti del menu
        style.configure(
            "Menu.TButton",
            font=("Segoe UI", 11),
            padding=15,
            relief="flat"
        )
        
        style.map(
            "Menu.TButton",
            background=[("active", "#e0e0e0"), ("!active", "#f5f5f5")]
        )
        
        # Stile per i pulsanti selezionati
        style.configure(
            "MenuActive.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=15,
            relief="flat",
            background="#007acc",
            foreground="white"
        )
    
    def _create_layout(self):
        """Crea il layout principale con menu laterale"""
        
        self.menu_bg = "#2d2d30"
        self.menu_content_bg = "#26262a"
        self.menu_hover_bg = "#37373d"
        self.menu_active_bg = "#007acc"
        self.menu_text_color = "white"
        self.menu_button_font = ("Segoe UI", 11)
        self.menu_button_active_font = ("Segoe UI", 11, "bold")

        # Frame contenitore principale
        main_container = ttk.Frame(self)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # ===== MENU LATERALE SINISTRO =====
        menu_frame = tk.Frame(main_container, bg=self.menu_bg, width=250)
        menu_frame.pack(side=tk.LEFT, fill=tk.Y)
        menu_frame.pack_propagate(False)  # Mantieni larghezza fissa
        
        # Header del menu
        header_frame = tk.Frame(menu_frame, bg="#007acc", height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        app_title = tk.Label(
            header_frame,
            text="Premi di Produzione",
            font=("Segoe UI", 18, "bold"),
            bg="#007acc",
            fg="white"
        )
        app_title.pack(pady=20)
        
        # Separatore
        separator1 = tk.Frame(menu_frame, bg="#3e3e42", height=2)
        separator1.pack(fill=tk.X, pady=10)
        
        self.menu_buttons = {}
        self.button_sections = {}
        self.menu_sections = []

        sections_config = [
            (
                "Importazione e Gestione",
                True,
                [
                    ("import", "Importazione Dati", self._show_importazione),
                    ("gestione", "Gestione Dati Produzione", self._show_gestione_dati),
                    ("anomalie", "Anomalie", self._show_anomalie),
                ],
            ),
            (
                "Configurazioni",
                False,
                [
                    ("premi", "Fasce Premi Attivita", self._show_fasce_premi),
                    ("peso", "Peso Movimenti", self._show_peso_movimenti),
                    ("malus", "Malus-Bonus", self._show_malus_bonus),
                ],
            ),
            (
                "Calcolo Premi",
                False,
                [
                    ("premi_carrellisti", "Premi Carrellisti", self._show_premi_carrellisti),
                    ("premi_preparatori", "Premi Preparatori", self._show_premi_preparatori),
                    ("premi_ricevimento", "Premi Ricevimento", self._show_premi_ricevimento),
                    ("premi_doppia_spunta", "Premi Doppia Spunta", self._show_premi_doppia_spunta),
                ],
            ),
            (
                "Performance",
                False,
                [
                    ("perf_carrellisti", "Performance Carrellisti", self._show_performance_carrellisti),
                    ("perf_preparatori", "Performance Preparatori", self._show_performance_preparatori),
                    ("perf_ricevimento", "Performance Ricevimento", self._show_performance_ricevimento),
                    ("perf_doppia_spunta", "Performance Doppia Spunta", self._show_performance_doppia_spunta),
                ],
            ),
            (
                "Strumenti",
                False,
                [
                    ("aggiorna_app", "Aggiorna Applicazione", self._show_update_dialog),
                ],
            ),
        ]

        for title, expanded, entries in sections_config:
            section = CollapsibleMenuSection(
                menu_frame,
                title,
                bg=self.menu_bg,
                content_bg=self.menu_content_bg,
                text_color=self.menu_text_color,
                highlight_color=self.menu_active_bg,
                expanded=expanded,
            )
            section.pack(fill=tk.X, pady=(0, 6))
            self.menu_sections.append(section)

            for key, label, callback in entries:
                btn = section.add_button(
                    label,
                    callback,
                    font=self.menu_button_font,
                    fg=self.menu_text_color,
                    bg=self.menu_content_bg,
                    hover_bg=self.menu_hover_bg,
                )
                self._register_menu_button(key, btn, section)
        
        # Info in basso
        info_label = tk.Label(
            menu_frame,
            text=f"Versione {APP_VERSION}",
            font=("Segoe UI", 9),
            bg=self.menu_bg,
            fg="#808080"
        )
        info_label.pack(side=tk.BOTTOM, pady=20)
        
        # ===== AREA CONTENUTO PRINCIPALE =====
        self.content_frame = ttk.Frame(main_container)
        self.content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
    
    def _highlight_menu_button(self, active_key):
        """Evidenzia il pulsante del menu attivo"""
        section = self.button_sections.get(active_key)
        if section:
            section.expand()

        for key, btn in self.menu_buttons.items():
            if key == active_key:
                btn.config(
                    bg=self.menu_active_bg,
                    fg=self.menu_text_color,
                    font=self.menu_button_active_font,
                    activebackground=self.menu_active_bg,
                )
            else:
                btn.config(
                    bg=self.menu_content_bg,
                    fg=self.menu_text_color,
                    font=self.menu_button_font,
                    activebackground=self.menu_hover_bg,
                )

    def _register_menu_button(self, key, button, section):
        self.menu_buttons[key] = button
        self.button_sections[key] = section
    
    def _clear_content(self):
        """Pulisce l'area del contenuto"""
        if self.current_frame:
            self.current_frame.destroy()
            self.current_frame = None
        self.current_module = None
    
    def _show_importazione(self):
        """Mostra il modulo di importazione dati"""
        self._clear_content()
        self._highlight_menu_button("import")
        
        # Importa e mostra il modulo di importazione
        try:
            from gui import ImportGUI
            
            self.current_frame = ttk.Frame(self.content_frame)
            self.current_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Forza l'aggiornamento del layout prima di creare l'interfaccia
            self.current_frame.update_idletasks()
            
            # Crea l'interfaccia di importazione nel frame
            import_gui = ImportGUI(self.current_frame)
            
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile caricare il modulo di importazione:\n{e}")
    
    def _show_gestione_dati(self):
        """Mostra il modulo di gestione dati di produzione"""
        self._clear_content()
        self._highlight_menu_button("gestione")
        
        # Importa e mostra il visualizzatore dati
        try:
            from data_viewer import DataViewerApp
            
            # Nessun padding per massimizzare lo spazio disponibile
            self.current_frame = ttk.Frame(self.content_frame)
            self.current_frame.pack(fill=tk.BOTH, expand=True)
            
            # Crea l'interfaccia di visualizzazione nel frame
            viewer = DataViewerApp(self.current_frame)
            
            # Forza aggiornamento completo del layout
            self.current_frame.update()
            self.update_idletasks()
            
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile caricare il modulo di gestione dati:\n{e}")
    
    def _show_fasce_premi(self):
        """Mostra il modulo di gestione delle fasce premio"""
        self._clear_content()
        self._highlight_menu_button("premi")

        try:
            from fasce_premi_view import FascePremiView

            self.current_frame = FascePremiView(self.content_frame)
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile caricare il modulo fasce premi:\n{exc}",
            )

    def _show_peso_movimenti(self):
        """Mostra il modulo gestione peso movimenti"""
        self._clear_content()
        self._highlight_menu_button("peso")

        try:
            from peso_movimenti_view import PesoMovimentiView

            self.current_frame = PesoMovimentiView(self.content_frame)
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile caricare il modulo Peso Movimenti:\n{exc}",
            )

    def _show_malus_bonus(self):
        """Mostra il modulo gestione malus/bonus"""
        self._clear_content()
        self._highlight_menu_button("malus")

        try:
            from malus_bonus_view import MalusBonusView

            self.current_frame = MalusBonusView(self.content_frame)
            self.current_module = None
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile caricare il modulo Malus-Bonus:\n{exc}",
            )

    def _show_update_dialog(self):
        """Apre la finestra di aggiornamento applicazione."""
        self._highlight_menu_button("aggiorna_app")
        open_update_dialog(self)

    def _show_anomalie(self):
        """Mostra il modulo gestione anomalie"""
        self._clear_content()
        self._highlight_menu_button("anomalie")

        try:
            from anomalie_view import AnomalieView

            self.current_frame = ttk.Frame(self.content_frame)
            self.current_frame.pack(fill=tk.BOTH, expand=True)

            self.current_module = AnomalieView(
                parent=self.current_frame,
                use_toplevel=False,
            )
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile caricare il modulo Anomalie:\n{exc}",
            )

    def _show_premi_carrellisti(self):
        """Mostra il modulo calcolo premi carrellisti"""
        self._clear_content()
        self._highlight_menu_button("premi_carrellisti")

        try:
            from premi_carrellisti_view import PremiCarrellistiView

            self.current_frame = PremiCarrellistiView(self.content_frame)
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile caricare il modulo Premi Carrellisti:\n{exc}",
            )

    def _show_premi_preparatori(self):
        """Mostra il modulo calcolo premi preparatori"""
        self._clear_content()
        self._highlight_menu_button("premi_preparatori")

        try:
            from premi_preparatori_view import PremiPreparatoriView

            self.current_frame = PremiPreparatoriView(self.content_frame)
        except Exception as exc:
            messagebox.showerror(
                "Errore",
                f"Impossibile caricare il modulo Premi Preparatori:\n{exc}",
            )

    def _show_premi_ricevimento(self):
        """Mostra il modulo calcolo premi ricevimento"""
        self._clear_content()
        self._highlight_menu_button("premi_ricevimento")

        self.current_frame = ttk.Frame(self.content_frame)
        self.current_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(
            self.current_frame,
            text="🏆 Premi Ricevimento\n\nModulo in sviluppo...",
            font=("Segoe UI", 16),
            fg="#007acc"
        ).pack(expand=True)

    def _show_premi_doppia_spunta(self):
        """Mostra il modulo calcolo premi doppia spunta"""
        self._clear_content()
        self._highlight_menu_button("premi_doppia_spunta")

        self.current_frame = ttk.Frame(self.content_frame)
        self.current_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(
            self.current_frame,
            text="� Premi Doppia Spunta\n\nModulo in sviluppo...",
            font=("Segoe UI", 16),
            fg="#007acc"
        ).pack(expand=True)

    def _show_performance_carrellisti(self):
        """Mostra il modulo performance carrellisti (placeholder)."""
        self._clear_content()
        self._highlight_menu_button("perf_carrellisti")

        self.current_frame = ttk.Frame(self.content_frame)
        self.current_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(
            self.current_frame,
            text="📈 Performance Carrellisti\n\nModulo in sviluppo...",
            font=("Segoe UI", 16),
            fg="#007acc",
        ).pack(expand=True)

    def _show_performance_preparatori(self):
        """Mostra il modulo performance preparatori (placeholder)."""
        self._clear_content()
        self._highlight_menu_button("perf_preparatori")

        self.current_frame = ttk.Frame(self.content_frame)
        self.current_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(
            self.current_frame,
            text="📈 Performance Preparatori\n\nModulo in sviluppo...",
            font=("Segoe UI", 16),
            fg="#007acc",
        ).pack(expand=True)

    def _show_performance_ricevimento(self):
        """Mostra il modulo performance ricevimento (placeholder)."""
        self._clear_content()
        self._highlight_menu_button("perf_ricevimento")

        self.current_frame = ttk.Frame(self.content_frame)
        self.current_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(
            self.current_frame,
            text="📈 Performance Ricevimento\n\nModulo in sviluppo...",
            font=("Segoe UI", 16),
            fg="#007acc",
        ).pack(expand=True)

    def _show_performance_doppia_spunta(self):
        """Mostra il modulo performance Doppia Spunta (placeholder)."""
        self._clear_content()
        self._highlight_menu_button("perf_doppia_spunta")

        self.current_frame = ttk.Frame(self.content_frame)
        self.current_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(
            self.current_frame,
            text="📈 Performance Doppia Spunta\n\nModulo in sviluppo...",
            font=("Segoe UI", 16),
            fg="#007acc",
        ).pack(expand=True)


def main():
    """Avvia l'applicazione principale"""
    # Inizializza il database e crea le tabelle
    from database import ensure_table_and_indexes
    ensure_table_and_indexes()
    
    app = MainMenu()
    app.mainloop()


if __name__ == "__main__":
    main()
