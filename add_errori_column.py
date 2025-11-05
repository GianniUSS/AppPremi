"""Script per aggiungere la colonna note_errori a dati_produzione"""
import mysql.connector
from config import MYSQL_CONFIG

try:
    conn = mysql.connector.connect(**MYSQL_CONFIG)
    cur = conn.cursor()
    
    # Verifica se la colonna esiste già
    cur.execute("""
        SELECT COUNT(*) 
        FROM information_schema.COLUMNS 
        WHERE TABLE_SCHEMA = DATABASE() 
          AND TABLE_NAME = 'dati_produzione' 
          AND COLUMN_NAME = 'note_errori'
    """)
    result = cur.fetchone()
    
    if result and result[0] == 0:
        # La colonna non esiste, aggiungila
        cur.execute("""
            ALTER TABLE dati_produzione
            ADD COLUMN note_errori TEXT NULL COMMENT 'Errori rilevati durante import'
            AFTER ore_gestionale
        """)
        conn.commit()
        print("✅ Colonna note_errori aggiunta con successo a dati_produzione")
    else:
        print("ℹ️ Colonna note_errori già presente")
    
    conn.close()
    
except Exception as e:
    print(f"❌ Errore: {e}")
    import traceback
    traceback.print_exc()
