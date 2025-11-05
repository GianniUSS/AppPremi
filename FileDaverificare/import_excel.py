"""
Importazione dati di produzione da Excel verso MySQL con GUI Tkinter.

Migliorie rispetto alla versione precedente:
- Connessioni DB gestite in modo sicuro (context manager) e config tramite variabili d'ambiente (con fallback);
- Parser più robusti (lettura Excel con fallback, parsing date vettorializzato dove possibile);
- Inserimenti in batch (executemany) per migliori prestazioni;
- Riduzione dipendenze da variabili globali nelle funzioni principali;
- Type hints e docstring per maggiore leggibilità e manutenibilità.
"""

# -*- coding: utf-8 -*-
import os
import datetime
from contextlib import closing
from pathlib import Path
from typing import List, Optional

import pandas as pd
import mysql.connector
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, IntVar
from tkinter import ttk
from tkcalendar import DateEntry

# ============== CONFIG DB ==============
MYSQL_CONFIG = {
    # Consente override tramite variabili d'ambiente; mantiene i default esistenti per compatibilità.
    "host": os.getenv("TIM_DB_HOST", "172.16.202.141"),
    "user": os.getenv("TIM_DB_USER", "tim_root"),
    "password": os.getenv("TIM_DB_PASSWORD", "Gianni#225524"),
    "database": os.getenv("TIM_DB_NAME", "tim_import"),
}
TABLE_NAME = "dati_produzione"

# ============== DB INIT ==============
def ensure_table_and_indexes() -> None:
    """Crea la tabella e l'indice unico se non esistono."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                f"""
                CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    data DATE NOT NULL,
                    codice_preparatore VARCHAR(50) NOT NULL,
                    nome_preparatore VARCHAR(255),
                    totale_colli INT,
                    penalita INT DEFAULT 0,
                    tipo_attivita VARCHAR(50) NOT NULL,
                    tipo VARCHAR(20),
                    UNIQUE KEY uniq_record (data, codice_preparatore, tipo_attivita, tipo)
                )
                """
            )
            conn.commit()

# ============== UTILITY ==============
def _norm(s: str) -> str:
    """Normalizza una stringa per confronti più tolleranti."""
    return str(s).lower().replace("°", "").strip()

def find_col(df: pd.DataFrame, candidates: List[List[str]], required: bool = True) -> Optional[str]:
    """Trova una colonna il cui nome contiene tutte le parole di almeno una combinazione.

    candidates: lista di alternative; ciascuna alternativa è una lista di token richiesti.
    Esempio: [["data", "inizio"], ["data"]]
    """
    names = {c: _norm(c) for c in df.columns}
    for col, norm in names.items():
        for option in candidates:
            if all(_norm(part) in norm for part in option):
                return col
    if required:
        need = " | ".join([" + ".join(op) for op in candidates])
        raise ValueError(f"Colonna non trovata (attesa una contenente: {need}).")
    return None

# ============== PARSER PREPARATORI (PICKING) ==============
def parse_preparatori(file_path: str) -> pd.DataFrame:
    """Parsa il report Preparatori (PICKING) e restituisce un DataFrame normalizzato."""
    # Header variabile: tentativo con riga 2, fallback riga 0; engine fallback automatico.
    try:
        df = pd.read_excel(file_path, engine="openpyxl", header=2)
    except Exception:
        try:
            df = pd.read_excel(file_path, engine="openpyxl", header=0)
        except Exception:
            df = pd.read_excel(file_path, header=0)

    df = df.rename(columns=lambda x: str(x).strip())

    col_data = find_col(df, [["data inizio preparazione"], ["data inizio"], ["data"]])
    col_cod = find_col(df, [["codice preparatore"], ["codice"]])
    col_nome = find_col(df, [["descrizione preparatore"], ["descrizione"], ["nome"]])
    col_colli = find_col(df, [["n colli"], ["colli"]])

    sub = df[[col_data, col_cod, col_nome, col_colli]].copy()

    # Parsing date vettorializzato: gestisce anche stringhe 8-caratteri yyyymmdd
    sdate = sub[col_data].astype(str).str.strip()
    sdate = sdate.str.replace(r"^(\d{4})(\d{2})(\d{2})$", r"\\1-\\2-\\3", regex=True)
    sub[col_data] = pd.to_datetime(sdate, errors="coerce").dt.date

    sub = sub.dropna(subset=[col_data, col_cod, col_colli])

    sub.rename(
        columns={
            col_data: "data",
            col_cod: "codice_preparatore",
            col_nome: "nome_preparatore",
            col_colli: "totale_colli",
        },
        inplace=True,
    )

    sub["codice_preparatore"] = sub["codice_preparatore"].astype(str).str.strip()
    sub["nome_preparatore"] = sub["nome_preparatore"].astype(str).str.strip()
    sub["totale_colli"] = pd.to_numeric(sub["totale_colli"], errors="coerce").fillna(0).astype(int)

    out = (
        sub.groupby(["data", "codice_preparatore", "nome_preparatore"], as_index=False)[
            "totale_colli"
        ]
        .sum()
        .assign(penalita=0, tipo_attivita="PICKING", tipo="")
    )
    return out

# ============== PARSER CARRELLISTI (CARRELLO) ==============
def parse_carrelisti(file_path: str, data_rif: datetime.date) -> pd.DataFrame:
    """Parsa il report Carrellisti (CARRELLO) e restituisce un DataFrame normalizzato."""
    try:
        df_raw = pd.read_excel(file_path, engine="openpyxl", header=None)
    except Exception:
        df_raw = pd.read_excel(file_path, header=None)
    col0 = df_raw.iloc[:, 0].astype(str).str.strip().str.lower()
    hits = col0[col0.eq("preparatore")].index.tolist()
    if not hits:
        raise ValueError("Intestazione 'Preparatore | Tipo | N° | Tempo Lavorato' non trovata.")
    header_row = hits[0]

    df = df_raw.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = ["Preparatore", "Tipo", "N°", "Tempo Lavorato"]
    df.dropna(how="all", inplace=True)

    records: List[dict] = []
    current_code = None
    for _, row in df.iterrows():
        prep = (str(row["Preparatore"]).strip() if pd.notna(row["Preparatore"]) else "")
        tipo = (str(row["Tipo"]).strip() if pd.notna(row["Tipo"]) else "")
        nval = row["N°"]

        # Riga intestazione del preparatore: ha "Preparatore" valorizzato e "Tipo" vuoto/NaN
        if prep and (tipo == "" or _norm(tipo) == "nan"):
            current_code = prep
            continue

        if not current_code:
            continue
        if tipo.upper() == "TOTALE":
            continue

        if tipo and not pd.isna(nval):
            try:
                colli = int(float(nval))
            except Exception:
                colli = 0
            records.append({
                "data": data_rif,
                "codice_preparatore": current_code,
                "nome_preparatore": "",
                "totale_colli": colli,
                "penalita": 0,
                "tipo_attivita": "CARRELLO",
                "tipo": tipo
            })

    out = pd.DataFrame(records)
    if out.empty:
        return out
    out = out.groupby(
        ["data", "codice_preparatore", "nome_preparatore", "tipo", "tipo_attivita"],
        as_index=False
    )["totale_colli"].sum()
    return out

# ============== PARSER RICEVITORI (RICEVITORI) ==============
def parse_ricevitori(file_path: str) -> pd.DataFrame:
    """Parsa il report Ricevitori e restituisce un DataFrame normalizzato."""
    try:
        df = pd.read_excel(file_path, engine="openpyxl", header=None)
    except Exception:
        df = pd.read_excel(file_path, header=None)

    records: List[dict] = []
    for row in df.itertuples(index=False):
        # Colonne fisse: B(1), I(8), P(15)
        codice = getattr(row, "_1", None)
        data_raw = getattr(row, "_8", None)
        colli = getattr(row, "_15", None)

        if pd.isna(codice) or pd.isna(data_raw) or pd.isna(colli):
            continue

        try:
            data_str = str(int(data_raw))
            data_fmt = datetime.datetime.strptime(data_str, "%Y%m%d").date()
        except Exception:
            continue

        try:
            colli_int = int(colli)
        except Exception:
            colli_int = 0

        records.append(
            {
                "data": data_fmt,
                "codice_preparatore": str(codice).strip(),
                "nome_preparatore": "",
                "totale_colli": colli_int,
                "penalita": 0,
                "tipo_attivita": "RICEVITORI",
                "tipo": "",
            }
        )

    df_out = pd.DataFrame(records)
    if df_out.empty:
        return df_out

    # somma colli per data e codice
    df_grouped = df_out.groupby(
        ["data", "codice_preparatore", "nome_preparatore", "tipo_attivita", "tipo"],
        as_index=False,
    )["totale_colli"].sum()

    return df_grouped

# ============== IMPORT PRINCIPALE ==============
def importa_excel(
    file_path: str,
    tipo: str,
    data_rif: Optional[datetime.date],
    progress_var: IntVar,
    progress_bar: ttk.Progressbar,
    status_label: ttk.Label,
    root: tk.Tk,
) -> None:
    """Esegue l'import del file Excel nel DB, scegliendo il parser per tipo attività."""
    try:
        status_label.config(text="")
        progress_var.set(0)
        root.update_idletasks()
        ensure_table_and_indexes()

        file_path = (file_path or "").strip()
        if not file_path or not Path(file_path).exists():
            messagebox.showwarning("Attenzione", "Seleziona un file Excel valido.")
            return

        status_label.config(text="Lettura file...")
        progress_var.set(10)
        root.update_idletasks()

        if "Preparatori" in tipo:
            df_grouped = parse_preparatori(file_path)
        elif "Carrellisti" in tipo:
            if not data_rif:
                messagebox.showwarning("Attenzione", "Seleziona la data di riferimento per i Carrellisti.")
                return
            df_grouped = parse_carrelisti(file_path, data_rif)
        else:
            df_grouped = parse_ricevitori(file_path)

        if df_grouped.empty:
            status_label.config(text="Nessun dato importabile.")
            messagebox.showwarning("Attenzione", "Nessun dato valido trovato nel file selezionato.")
            return

        # Preparazione valori per inserimento batch
        values = [
            (
                r["data"],
                str(r["codice_preparatore"]),
                str(r["nome_preparatore"]) if pd.notna(r["nome_preparatore"]) else None,
                int(r["totale_colli"]) if pd.notna(r["totale_colli"]) else 0,
                int(r.get("penalita", 0)),
                r["tipo_attivita"],
                str(r.get("tipo", "")) if pd.notna(r.get("tipo", "")) else "",
            )
            for _, r in df_grouped.iterrows()
        ]

        sql = f"""
            INSERT INTO {TABLE_NAME}
            (data, codice_preparatore, nome_preparatore, totale_colli, penalita, tipo_attivita, tipo)
            VALUES (%s,%s,%s,%s,%s,%s,%s)
            ON DUPLICATE KEY UPDATE
                nome_preparatore = VALUES(nome_preparatore),
                totale_colli     = VALUES(totale_colli),
                penalita         = VALUES(penalita),
                tipo_attivita    = VALUES(tipo_attivita),
                tipo             = VALUES(tipo)
        """

        status_label.config(text="Scrittura su database...")
        progress_var.set(60)
        root.update_idletasks()

        with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
            with closing(conn.cursor()) as cur:
                cur.executemany(sql, values)
                conn.commit()

        progress_var.set(100)
        root.update_idletasks()
        status_label.config(text="Importazione completata ✅")
        messagebox.showinfo("Successo", f"Importazione completata.\nRecord: {len(values)}")

    except Exception as e:
        status_label.config(text="Errore ❌")
        messagebox.showerror("Errore durante l'importazione", str(e))

# ============== GUI ==============
def main() -> None:
    root = tk.Tk()
    root.title("Importazione Produzione - USSolutions")
    W, H = 850, 520
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    x, y = int((sw - W) / 2), int((sh - H) / 2)
    root.geometry(f"{W}x{H}+{x}+{y}")

    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TButton", font=("Arial", 12), padding=(12, 6), width=22)

    big_font = ("Arial", 12)
    title_font = ("Arial", 13, "bold")

    ttk.Label(root, text="Seleziona il file Excel da importare:", font=title_font).pack(pady=10)
    file_entry = ttk.Entry(root, width=90, font=big_font)
    file_entry.pack(pady=5)

    ttk.Button(
        root,
        text="Sfoglia...",
        command=lambda: file_entry.delete(0, "end")
        or file_entry.insert(0, filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")]))
    ).pack(pady=5)

    ttk.Label(root, text="Tipo file/attività:", font=title_font).pack(pady=10)

    opzioni = ["Preparatori (PICKING)", "Carrellisti (CARRELLO)", "Ricevitori (RICEVITORI)"]
    tipo_var = StringVar(value=opzioni[0])
    ttk.OptionMenu(root, tipo_var, opzioni[0], *opzioni).pack(pady=5)

    ttk.Label(root, text="Data di riferimento (per Carrellisti):", font=title_font).pack(pady=10)
    date_entry = DateEntry(root, date_pattern="yyyy-mm-dd", font=big_font)
    date_entry.pack(pady=5)

    progress_var = IntVar(value=0)
    progress_bar = ttk.Progressbar(root, length=650, mode="determinate", maximum=100, variable=progress_var)
    status_label = ttk.Label(root, text="", font=("Arial", 11))

    ttk.Button(
        root,
        text="Importa dati",
        command=lambda: importa_excel(
            file_entry.get(),
            tipo_var.get(),
            date_entry.get_date(),
            progress_var,
            progress_bar,
            status_label,
            root,
        ),
    ).pack(pady=15)

    progress_bar.pack(pady=10, ipady=5)
    status_label.pack(pady=5)

    root.mainloop()


if __name__ == "__main__":
    main()
