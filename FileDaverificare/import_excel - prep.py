import pandas as pd
import mysql.connector
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Parametri MySQL
MYSQL_CONFIG = {
    "host": "172.16.202.141",
    "user": "tim_root",
    "password": "Gianni#225524",
    "database": "tim_import"
}

TABLE_NAME = "dati_produzione"

def create_table():
    """Crea la tabella dati_produzione se non esiste"""
    conn = mysql.connector.connect(**MYSQL_CONFIG)
    cursor = conn.cursor()
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
            id INT AUTO_INCREMENT PRIMARY KEY,
            data DATE NOT NULL,
            codice_preparatore VARCHAR(20) NOT NULL,
            nome_preparatore VARCHAR(255),
            totale_colli INT,
            penalita INT DEFAULT 0,
            tipo_attivita VARCHAR(50) NOT NULL,
            UNIQUE KEY uniq_record (data, codice_preparatore, tipo_attivita)
        )
    """)
    conn.commit()
    cursor.close()
    conn.close()

def import_excel_to_mysql(file_path, tipo_attivita):
    """Importa i dati dall'Excel in dati_produzione"""
    try:
        # 1️⃣ Leggo il file Excel
        df = pd.read_excel(file_path, sheet_name=0, header=2)
        df = df[["Data inizio preparazione", "Codice preparatore", "Descrizione preparatore", "N° colli"]]
        df.columns = ["data", "codice_preparatore", "nome_preparatore", "totale_colli"]

        # 2️⃣ Converto la data
        df["data"] = pd.to_datetime(df["data"], format="%Y%m%d", errors="coerce").dt.date

        # 3️⃣ Raggruppo per giorno e preparatore
        df_grouped = df.groupby(
            ["data", "codice_preparatore", "nome_preparatore"], as_index=False
        )["totale_colli"].sum()

        # 4️⃣ Aggiungo colonne nuove
        df_grouped["penalita"] = 0
        df_grouped["tipo_attivita"] = tipo_attivita

        # 5️⃣ Connessione MySQL
        conn = mysql.connector.connect(**MYSQL_CONFIG)
        cursor = conn.cursor()

        for _, row in df_grouped.iterrows():
            sql = f"""
                INSERT INTO {TABLE_NAME} 
                (data, codice_preparatore, nome_preparatore, totale_colli, penalita, tipo_attivita)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    nome_preparatore = VALUES(nome_preparatore),
                    totale_colli = VALUES(totale_colli),
                    penalita = VALUES(penalita),
                    tipo_attivita = VALUES(tipo_attivita)
            """
            cursor.execute(sql, (
                row["data"], 
                row["codice_preparatore"], 
                row["nome_preparatore"], 
                int(row["totale_colli"]),
                int(row["penalita"]),
                row["tipo_attivita"]
            ))
        
        conn.commit()
        cursor.close()
        conn.close()
        messagebox.showinfo("Successo", f"Importazione completata da file:\n{file_path}\nTipo attività: {tipo_attivita}")
    except Exception as e:
        messagebox.showerror("Errore", f"Errore durante l'importazione:\n{str(e)}")

# -------------------------
# GUI
root = tk.Tk()
root.title("Importazione Produzione - USSolutions")
root.geometry("600x250")

selected_file = tk.StringVar()
selected_tipo = tk.StringVar(value="PICKING")

def scegli_file():
    file_path = filedialog.askopenfilename(
        title="Seleziona file Excel",
        filetypes=[("File Excel", "*.xlsx *.xls")]
    )
    if file_path:
        selected_file.set(file_path)

def avvia_import():
    if not selected_file.get():
        messagebox.showwarning("Attenzione", "Devi selezionare un file Excel prima di importare.")
        return
    if not selected_tipo.get():
        messagebox.showwarning("Attenzione", "Devi selezionare un tipo di attività.")
        return
    create_table()
    import_excel_to_mysql(selected_file.get(), selected_tipo.get())

# UI
tk.Label(root, text="Seleziona il file Excel da importare:", font=("Arial", 12)).pack(pady=10)

frame = tk.Frame(root)
frame.pack(pady=5)

tk.Entry(frame, textvariable=selected_file, width=50).pack(side=tk.LEFT, padx=5)
tk.Button(frame, text="Sfoglia...", command=scegli_file).pack(side=tk.LEFT)

# Dropdown tipo attività
tk.Label(root, text="Tipo attività:", font=("Arial", 12)).pack(pady=10)
combo_tipo = ttk.Combobox(root, textvariable=selected_tipo, values=["PICKING", "IMBALLAGGIO", "PACKING", "ALTRO"], state="readonly", width=20)
combo_tipo.pack()

tk.Button(root, text="Importa in MySQL", command=avvia_import, bg="#4CAF50", fg="white", font=("Arial", 12, "bold")).pack(pady=20)

root.mainloop()
