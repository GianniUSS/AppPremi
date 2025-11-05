import mysql.connector
from difflib import SequenceMatcher
import csv

# Parametri MySQL
MYSQL_CONFIG_IMPORT = {
    "host": "172.16.202.141",
    "user": "tim_root",
    "password": "Gianni#225524",
    "database": "tim_import"
}

MYSQL_CONFIG_MAIN = {
    "host": "172.16.202.141",
    "user": "tim_root",
    "password": "Gianni#225524",
    "database": "tim"
}

# ID tipo attività per PICKING
PICKING_ID = "00123112341769380464"

# Data da testare (fissata per il momento)
DATA_TARGET = "2025-09-11"

REPORT_FILE = f"update_report_{DATA_TARGET}.csv"

def get_log_preparatori():
    """Legge i dati importati nel DB tim_import.log_preparatori per la data target"""
    conn = mysql.connector.connect(**MYSQL_CONFIG_IMPORT)
    cursor = conn.cursor(dictionary=True)
    cursor.execute("""
        SELECT data, codice_preparatore, nome_preparatore, totale_colli
        FROM log_preparatori
        WHERE data = %s
    """, (DATA_TARGET,))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

def similarity(a, b):
    """Calcola la similarità tra due stringhe"""
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()

def update_attivita():
    log_data = get_log_preparatori()
    if not log_data:
        print(f"⚠️ Nessun dato trovato in log_preparatori per la data {DATA_TARGET}")
        return
    
    conn = mysql.connector.connect(**MYSQL_CONFIG_MAIN)
    cursor = conn.cursor(dictionary=True)

    report_rows = []

    for row in log_data:
        data = row["data"]
        codice_preparatore = row["codice_preparatore"]
        nome_preparatore_excel = row["nome_preparatore"]
        totale_colli = row["totale_colli"]

        esito = ""
        nominativo_db = ""
        sim_percent = 0
        attivita_info = []

        # 1️⃣ Trovo utente_id dal codice gestionale valido per la data
        sql_codice = """
            SELECT utente_id
            FROM codicegestionale
            WHERE codice = %s
              AND tipo_attivita_id = %s
              AND valido_dal <= %s
              AND (valido_al IS NULL OR valido_al >= %s)
            LIMIT 1
        """
        cursor.execute(sql_codice, (codice_preparatore, PICKING_ID, data, data))
        codice_row = cursor.fetchone()

        if not codice_row:
            esito = "Nessun mapping trovato"
            print(f"❌ {esito} per [Excel: {nome_preparatore_excel}] (Codice: {codice_preparatore}) in data {data}")
            report_rows.append([data, codice_preparatore, nome_preparatore_excel, "-", "-", "-", "-", "-", esito])
            continue

        utente_id = codice_row["utente_id"]

        # 2️⃣ Recupero dati anagrafici dell'utente
        sql_utente = "SELECT cognome, nome FROM utente WHERE id = %s"
        cursor.execute(sql_utente, (utente_id,))
        utente_row = cursor.fetchone()
        if utente_row:
            nominativo_db = f"{utente_row['cognome']} {utente_row['nome']}"
        else:
            nominativo_db = "(utente non trovato in tabella utente)"

        # 🔎 Confronto nomi
        sim_score = similarity(nome_preparatore_excel, nominativo_db)
        sim_percent = round(sim_score * 100, 1)
        if sim_percent < 70:
            print(f"⚠️ Nome differente → [Excel: {nome_preparatore_excel}] ↔ [DB: {nominativo_db}] (similarità {sim_percent}%)")

        # 3️⃣ Trovo attività PICKING per quell’utente e data
        sql_attivita = """
            SELECT id, durata
            FROM attivita
            WHERE utente_id = %s
              AND data_riferimento = %s
              AND tipo_attivita_id = %s
        """
        cursor.execute(sql_attivita, (utente_id, data, PICKING_ID))
        attivita_rows = cursor.fetchall()

        if len(attivita_rows) == 0:
            esito = "Nessuna attività trovata"
            print(f"❌ {esito} per [Excel: {nome_preparatore_excel}] (Codice: {codice_preparatore}) ↔ [DB: {nominativo_db}] (utente {utente_id}) in data {data}")
            report_rows.append([data, codice_preparatore, nome_preparatore_excel, nominativo_db, "-", "-", "-", sim_percent, esito])
            continue

        # 4️⃣ Ripartizione proporzionale ai minuti di durata
        total_durata = sum([a["durata"] for a in attivita_rows if a["durata"]])
        if not total_durata:
            esito = "Durate nulle"
            print(f"⚠️ {esito} per [Excel: {nome_preparatore_excel}] (Codice: {codice_preparatore}) ↔ [DB: {nominativo_db}]")
            report_rows.append([data, codice_preparatore, nome_preparatore_excel, nominativo_db, "-", "-", "-", sim_percent, esito])
            continue

        colli_assegnati = 0
        for i, att in enumerate(attivita_rows):
            if i < len(attivita_rows) - 1:
                colli_att = round(totale_colli * att["durata"] / total_durata)
                colli_assegnati += colli_att
            else:
                colli_att = totale_colli - colli_assegnati  # correzione scarto

            update_sql = """
                UPDATE attivita
                SET tag_valore = %s, tag_ultima = 1
                WHERE id = %s
            """
            cursor.execute(update_sql, (colli_att, att["id"]))
            esito = "Aggiornato"
            print(f"{esito}: [Excel: {nome_preparatore_excel}] (Codice: {codice_preparatore}) ↔ [DB: {nominativo_db}] - Attività {att['id']} durata {att['durata']} → {colli_att} colli (similarità {sim_percent}%)")
            attivita_info.append((att["id"], att["durata"], colli_att))

        for att_id, durata, colli in attivita_info:
            report_rows.append([data, codice_preparatore, nome_preparatore_excel, nominativo_db, att_id, durata, colli, sim_percent, esito])

    conn.commit()
    cursor.close()
    conn.close()

    # 📄 Esporta CSV
    with open(REPORT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Data", "Codice_Preparatore", "Nome_Excel", "Nome_DB", "Attivita_ID", "Durata", "Colli_assegnati", "Similarita_%", "Esito"])
        writer.writerows(report_rows)

    print(f"\n📄 Report generato: {REPORT_FILE}")

if __name__ == "__main__":
    update_attivita()
