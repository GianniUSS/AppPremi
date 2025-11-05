"""
Interfaccia grafica con Tkinter per l'importazione dati.
"""
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, IntVar
from tkinter import ttk
from tkcalendar import DateEntry
from pathlib import Path
from typing import Optional
import datetime

from config import WINDOW_CONFIG, FONTS, COLORS
from ui_components import create_button
from import_service import ImportService
from data_viewer import DataViewer


class ImportGUI:
    """Classe principale per la gestione dell'interfaccia grafica."""
    
    def __init__(self, parent=None):
        """
        Inizializza l'interfaccia di importazione.
        
        Args:
            parent: Frame parent opzionale. Se None, crea una finestra standalone.
        """
        self.is_standalone = parent is None
        
        if self.is_standalone:
            self.root = tk.Tk()
        else:
            self.root = parent
        
        self.import_service = ImportService()
        self._setup_window()
        self._create_widgets()
        
    def _setup_window(self):
        """Configura la finestra principale o il frame."""
        if self.is_standalone:
            self.root.title(WINDOW_CONFIG["title"])
            W, H = WINDOW_CONFIG["width"], WINDOW_CONFIG["height"]
            sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            x, y = int((sw - W) / 2), int((sh - H) / 2)
            self.root.geometry(f"{W}x{H}+{x}+{y}")
            self.root.configure(bg=COLORS["background"])
            self.root.resizable(False, False)
        # Non configurare background se è embedded (il parent lo gestisce)

        # Configurazione stile moderno (funziona sia standalone che embedded)
        style = ttk.Style()
        style.theme_use("clam")
        
        # Stile per i pulsanti
        style.configure(
            "Primary.TButton",
            font=FONTS["button"],
            padding=(20, 12),
            background=COLORS["primary"],
            foreground="white",
            borderwidth=0,
            focuscolor="none"
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLORS["secondary"])],
            foreground=[("active", "white")]
        )
        
        # Stile per pulsante sfoglia
        style.configure(
            "Secondary.TButton",
            font=FONTS["big"],
            padding=(15, 8),
            background=COLORS["secondary"],
            foreground="white",
            borderwidth=0
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#2980B9")],
            foreground=[("active", "white")]
        )
        
        # Stile per frame
        style.configure("Card.TFrame", background=COLORS["white"], relief="flat")
        style.configure("Main.TFrame", background=COLORS["background"])
        
        # Stile per label
        style.configure("Title.TLabel", background=COLORS["white"], 
                       foreground=COLORS["text_dark"], font=FONTS["title"])
        style.configure("Subtitle.TLabel", background=COLORS["white"], 
                       foreground=COLORS["text_light"], font=FONTS["subtitle"])
        style.configure("Status.TLabel", background=COLORS["background"], 
                       foreground=COLORS["text_dark"], font=FONTS["status"])
        
        # Stile per Entry
        style.configure("Modern.TEntry", fieldbackground="white", 
                       borderwidth=1, relief="solid")

    def _create_widgets(self):
        """Crea tutti i widget dell'interfaccia."""
        # Container principale
        main_frame = ttk.Frame(self.root, style="Main.TFrame")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Header
        header_frame = ttk.Frame(main_frame, style="Card.TFrame")
        header_frame.pack(fill="x", pady=(0, 20))
        
        ttk.Label(
            header_frame,
            text="🗂️ Importazione Dati Produzione",
            style="Title.TLabel"
        ).pack(pady=(20, 5))
        
        ttk.Label(
            header_frame,
            text="Sistema professionale di gestione e importazione dati Excel",
            style="Subtitle.TLabel"
        ).pack(pady=(0, 20))
        
        # Sezione Selezione File
        file_frame = ttk.Frame(main_frame, style="Card.TFrame")
        file_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(
            file_frame,
            text="📄 Seleziona il file Excel da importare",
            style="Title.TLabel",
            font=FONTS["big"]
        ).pack(pady=(15, 10), padx=20, anchor="w")
        
        # Frame per entry e bottone
        file_input_frame = ttk.Frame(file_frame, style="Card.TFrame")
        file_input_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        self.file_entry = ttk.Entry(file_input_frame, width=70, font=FONTS["big"])
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        create_button(
            file_input_frame,
            text="📁 Sfoglia...",
            command=self._browse_file,
            variant="secondary",
            width=14,
        ).pack(side="left")

        # Sezione Tipo Attività
        tipo_frame = ttk.Frame(main_frame, style="Card.TFrame")
        tipo_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(
            tipo_frame,
            text="⚙️ Tipo file/attività",
            style="Title.TLabel",
            font=FONTS["big"]
        ).pack(pady=(15, 10), padx=20, anchor="w")

        opzioni = [
            "Preparatori (PICKING)",
            "Carrellisti (CARRELLO)",
            "Ricevitori (RICEVITORI)",
            "Doppia Spunta",
        ]
        self.tipo_var = StringVar(value=opzioni[0])
        
        tipo_menu = ttk.OptionMenu(tipo_frame, self.tipo_var, opzioni[0], *opzioni)
        tipo_menu.config(width=35)
        tipo_menu.pack(pady=(0, 15), padx=20)

        # Pulsanti Importazione e Visualizzazione
        button_frame = ttk.Frame(main_frame, style="Main.TFrame")
        button_frame.pack(fill="x", pady=(10, 0))
        
        buttons_container = ttk.Frame(button_frame, style="Main.TFrame")
        buttons_container.pack()
        
        create_button(
            buttons_container,
            text="🚀 Avvia Importazione",
            command=self._import_data,
            variant="primary",
            width=24,
        ).pack(side="left", padx=5, pady=10)
        

        # Progress bar e status
        self.progress_var = IntVar(value=0)
        self.progress_bar = ttk.Progressbar(
            main_frame,
            length=700,
            mode="determinate",
            maximum=100,
            variable=self.progress_var
        )
        
        self.status_label = ttk.Label(
            main_frame,
            text="Pronto per l'importazione",
            style="Status.TLabel"
        )

        self.progress_bar.pack(pady=(5, 10), ipady=8)
        self.status_label.pack(pady=5)

    def _browse_file(self):
        """Apre il dialog per selezionare il file Excel."""
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, filename)

    def _import_data(self):
        """Gestisce l'importazione dei dati."""
        file_path = self.file_entry.get().strip()
        if not file_path or not Path(file_path).exists():
            messagebox.showwarning("Attenzione", "Seleziona un file Excel valido.")
            return

        tipo = self.tipo_var.get()
        # La data non è più necessaria (viene estratta automaticamente dal file Excel)
        data_rif = None

        # Esegui importazione
        self.import_service.import_excel(
            file_path=file_path,
            tipo=tipo,
            data_rif=data_rif,
            progress_var=self.progress_var,
            status_label=self.status_label,
            root=self.root
        )
    
    def _open_data_viewer(self):
        """Apre la finestra di visualizzazione dati."""
        try:
            viewer = DataViewer(self.root)
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire il visualizzatore:\n{str(e)}")

    def run(self):
        """Avvia l'interfaccia grafica (solo per modalità standalone)."""
        if self.is_standalone:
            self.root.mainloop()
        # Se è embedded, non fa nulla (il mainloop è gestito dal parent)


def main():
    """Punto di ingresso principale dell'applicazione."""
    app = ImportGUI()
    app.run()


if __name__ == "__main__":
    main()