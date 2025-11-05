from tkinter import filedialog, messagebox, Tk, Entry, Label, Button
from tkcalendar import DateEntry
import pandas as pd
import mysql.connector
from datetime import datetime
from pathlib import Path

def importa_carrellisti():
    file_path = file_entry.get()
    data_selezionata = date_entry.get_date().strftime('%Y-%m-%d')

    if not Path(file_path).exists():
        messagebox.showerror("Errore", "File non trovato.")
        return

    try:
        # 🔹 Forzo l’engine openpyxl per compatibilità con .xlsx
        df_raw = pd.read_excel(file_path, engine="openpyxl", header=None)

        # Trova la riga con intestazione
        header_row_index = df_raw[df_raw.iloc[:, 0] == 'Preparatore'].index[0]
        df_data = df_raw.iloc[header_row_index + 1:].reset_index(drop=True)
        df_data.columns = ['Preparatore', 'Tipo', 'N°', 'Tempo Lavorato']
        df_data.dropna(how='all', inplace=True)

        records = []
        current_preparatore = None

        for _, row in df_data.iterrows():
            if pd.notna(row['Preparatore']) and pd.isna(row['Tipo']):
                current_preparatore = str(row['Preparatore']).strip()
            elif pd.notna(row['Tipo']) and row['Tipo'] != 'TOTALE' and current_preparatore:
                records.append({
                    'preparatore': current_preparatore,
                    'tipo': row['Tipo'],
                    'quantita': int(row['N°']),
                    'tempo_lavorato': str(row['Tempo Lavorato']),
                    'data': data_selezionata,
                    'attivita': 'CARRELLO'
                })

        df_finale = pd.DataFrame(records)

        conn = mysql.connector.connect(
            host="172.16.202.141",
            user="tim_root",
            password="Gianni#225524",
            database="tim_import"
        )
        cursor = conn.cursor()

        # Elimino righe già presenti per la stessa data e attività
        cursor.execute("DELETE FROM dati_produzione WHERE data = %s AND tipo_attivita = 'CARRELLO'", (data_selezionata,))

        for _, riga in df_finale.iterrows():
            cursor.execute(
                "INSERT INTO dati_produzione (codice_preparatore, tipo, totale_colli, tempo, data, tipo_attivita) "
                "VALUES (%s, %s, %s, %s, %s, %s)",
                (
                    riga['preparatore'], riga['tipo'], riga['quantita'],
                    riga['tempo_lavorato'], riga['data'], riga['attivita']
                )
            )

        conn.commit()
        cursor.close()
        conn.close()

        messagebox.showinfo("Successo", f"✔ Importazione completata:\n{file_path}\nRecord inseriti: {len(df_finale)}")
    except Exception as e:
        messagebox.showerror("Errore durante l'importazione", str(e))

# GUI
root = Tk()
root.title("Importazione Carrellisti - USSolutions")

Label(root, text="Seleziona il file Excel da importare:").pack()
file_entry = Entry(root, width=60)
file_entry.pack()
Button(root, text="Sfoglia...", command=lambda: file_entry.insert(0, filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls")]))).pack()

Label(root, text="Data di riferimento:").pack()
date_entry = DateEntry(root, date_pattern='yyyy-mm-dd')
date_entry.pack(pady=5)

Button(root, text="Importa in MySQL", bg='green', fg='white', command=importa_carrellisti).pack(pady=10)

root.mainloop()
