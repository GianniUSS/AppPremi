# Importazione Dati Produzione

Applicazione per l'importazione di dati di produzione da file Excel verso database MySQL.

## Struttura del Progetto

```
AppEden/
├── main.py              # Script principale di avvio
├── gui.py              # Interfaccia grafica Tkinter
├── import_service.py   # Logica di business per l'importazione
├── parsers.py          # Parser per i diversi tipi di file Excel
├── database.py         # Gestione database e operazioni SQL
├── utils.py            # Funzioni utility e helper
├── config.py           # Configurazioni e costanti
└── requirements.txt    # Dipendenze Python
```

## Moduli

### `config.py`
- Configurazione database (con supporto variabili d'ambiente)
- Impostazioni interfaccia grafica
- Costanti globali

### `database.py`
- Gestione connessioni MySQL con context manager
- Creazione tabelle e indici
- Operazioni CRUD con inserimenti batch
- Aggiornamento penalità per le attività PICKING a partire dalla Doppia Spunta

### `utils.py`
- Funzioni di normalizzazione stringhe
- Ricerca colonne nei DataFrame
- Conversioni tipo sicure
- Preparazione dati per database

### `parsers.py`
- `parse_preparatori()`: Parser per report PICKING
- `parse_carrelisti()`: Parser per report CARRELLO  
- `parse_ricevitori()`: Parser per report RICEVITORI
- `parse_doppia_spunta()`: Parser per report DOPPIA SPUNTA; restituisce anche le penalità aggregate per aggiornare il PICKING

### `import_service.py`
- Logica principale di importazione
- Coordinazione tra parser e database
- Gestione errori e validazioni

### `gui.py`
- Interfaccia grafica con Tkinter
- Gestione eventi utente
- Progress bar e messaggi di stato

## Installazione

1. Installa le dipendenze:
```bash
pip install -r requirements.txt
```

2. (Opzionale) Configura variabili d'ambiente per il database:
```bash
set TIM_DB_HOST=your_host
set TIM_DB_USER=your_user
set TIM_DB_PASSWORD=your_password
set TIM_DB_NAME=your_database
```

## Utilizzo

Avvia l'applicazione:
```bash
python main.py
```

## Migliorie Implementate

- **Modularità**: Codice diviso in moduli specializzati
- **Separazione responsabilità**: GUI, logica business, database separati
- **Gestione sicura database**: Context managers per connessioni
- **Configurazione flessibile**: Variabili d'ambiente con fallback
- **Inserimenti batch**: Performance migliorate
- **Type hints**: Codice più documentato e manutenibile
- **Gestione errori**: Robustezza migliorata