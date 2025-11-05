"""
Configurazione database e costanti globali.
"""
import os

# ============== CONFIG DB ==============
MYSQL_CONFIG = {
    # Consente override tramite variabili d'ambiente; mantiene i default esistenti per compatibilità.
    "host": os.getenv("TIM_DB_HOST", "172.16.202.141"),
    "user": os.getenv("TIM_DB_USER", "tim_root"),
    "password": os.getenv("TIM_DB_PASSWORD", "Gianni#225524"),
    "database": os.getenv("TIM_DB_NAME", "tim_import"),
}

MYSQL_CONFIG_MAIN = {
    "host": "172.16.202.141",
    "user": "tim_root",
    "password": "Gianni#225524",
    "database": "tim_db"
}

TABLE_NAME = "dati_produzione"

# ============== CONFIGURAZIONE GUI ==============
WINDOW_CONFIG = {
    "width": 900,
    "height": 650,
    "title": "Gestione Dati Produzione"
}

FONTS = {
    "big": ("Segoe UI", 11),
    "title": ("Segoe UI", 14, "bold"),
    "subtitle": ("Segoe UI", 10),
    "status": ("Segoe UI", 10),
    "button": ("Segoe UI", 11, "bold"),
    "label": ("Segoe UI", 10),
    "input": ("Segoe UI", 10)
}

COLORS = {
    "primary": "#2C3E50",      # Blu scuro professionale
    "secondary": "#3498DB",    # Blu accento
    "success": "#27AE60",      # Verde successo
    "danger": "#E74C3C",       # Rosso errore
    "warning": "#F39C12",      # Arancione warning
    "background": "#ECF0F1",   # Grigio chiaro sfondo
    "white": "#FFFFFF",
    "text_dark": "#2C3E50",
    "text_light": "#7F8C8D",
    "border": "#BDC3C7"
}