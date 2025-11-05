
import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd
import mysql.connector
from datetime import datetime
from pathlib import Path

def importa_dati():
    file_path = file_entry.get()
    tipo_attivita = tipo_var.get()
    data_selezionata = date_entry.get_date().strftime('%Y-%m-%d')

    if not Path(file_path).exists():
        messagebox.showerror("Errore", "File non trovato.")
        return

    try:
        if tipo_attivita == "Carrellisti":
            df = pd.read_excel(file_path, sheet_name=0, header=2)
            df = df[df['Preparatore'].notna()]
            df = df[df['Preparatore'] != 'TOTALE']

            preparatore_corrente = None
            righe_importazione = []

            for _, row in df.iterrows():
                if pd.notna(row['Preparatore']) and pd.isna(row['Tipo']):
                    preparatore_corrente = row['Preparatore']
                    continue
                if pd.notna(row['Tipo']) and preparatore_corrente:
                    righe_importazione.append({
                        'preparatore': preparatore_corrente,
                        'tipo': row['Tipo'],
                        'numero': row['N°'],
                        'tempo': row['Tempo Lavorato'],
                        'data': data_selezionata,
                        'attivita': 'CARRELLO'
                    })

            df_finale = pd.DataFrame(righe_importazione)
        else:
            messagebox.showerror("Errore", "Tipo attività non supportato al momento.")
            return

        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="",
            database="tim_db"
        )
        cursor = conn.cursor()

        for _, riga in df_finale.iterrows():
            cursor.execute(
                "INSERT INTO produzione (codice_preparatore, tipo, quantita, tempo, data, tipo_attivita) "
                "VALUES (%s, %s, %s, %s, %s, %s)",
                (
                    riga['preparatore'], riga['tipo'], riga['numero'],
                    str(riga['tempo']), riga['data'], riga['attivita']
                )
            )

        conn.commit()
        cursor.close()
        conn.close()

        messagebox.showinfo("Successo", f"✔ Importazione completata:\n{file_path}\nTipo attività: {riga['attivita']}")
    except Exception as e:
        messagebox.showerror("Errore durante l'importazione", str(e))


root = tk.Tk()
root.title("Importazione Produzione - USSolutions")

tk.Label(root, text="Seleziona il file Excel da importare:").pack()
file_entry = tk.Entry(root, width=60)
file_entry.pack()
tk.Button(root, text="Sfoglia...", command=lambda: file_entry.insert(0, filedialog.askopenfilename())).pack()

tk.Label(root, text="Tipo file/attività:").pack()
tipo_var = tk.StringVar(value="Carrellisti")
tk.OptionMenu(root, tipo_var, "Carrellisti").pack()

tk.Label(root, text="Data (solo se il file non contiene la colonna Data):").pack()
date_entry = DateEntry(root, date_pattern='d/m/yy')
date_entry.pack(pady=5)

tk.Button(root, text="Importa in MySQL", bg='green', fg='white', command=importa_dati).pack(pady=10)

root.mainloop()
