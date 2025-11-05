"""
Gestione database: connessioni, creazione tabelle e operazioni CRUD.
"""
import datetime
from contextlib import closing
from typing import Any, Dict, List, Optional, Sequence, Tuple, cast
import mysql.connector
from mysql.connector import errorcode
from config import MYSQL_CONFIG, TABLE_NAME


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
                    ore_tim DECIMAL(10,2) DEFAULT 0 COMMENT 'Ore di lavoro da TIM',
                    ore_gestionale DECIMAL(10,2) DEFAULT 0 COMMENT 'Ore calcolate con sessioni (solo carrellisti)',
                    UNIQUE KEY uniq_record (data, codice_preparatore, tipo_attivita, tipo)
                )
                """
            )
            
            # Migrazione: rinomina tempo -> ore_tim e converti da minuti a ore
            try:
                # Verifica se esiste la vecchia colonna 'tempo'
                cur.execute(
                    f"""
                    SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS 
                    WHERE TABLE_SCHEMA = DATABASE() 
                    AND TABLE_NAME = '{TABLE_NAME}' 
                    AND COLUMN_NAME = 'tempo'
                    """
                )
                if cur.fetchone():
                    # Aggiungi ore_tim
                    cur.execute(
                        f"""
                        ALTER TABLE {TABLE_NAME}
                        ADD COLUMN ore_tim DECIMAL(10,2) DEFAULT 0 COMMENT 'Ore di lavoro da TIM'
                        """
                    )
                    # Converti tempo (minuti) in ore_tim (ore)
                    cur.execute(
                        f"""
                        UPDATE {TABLE_NAME}
                        SET ore_tim = tempo / 60.0
                        WHERE tempo IS NOT NULL
                        """
                    )
                    # Rimuovi la vecchia colonna tempo
                    cur.execute(
                        f"""
                        ALTER TABLE {TABLE_NAME}
                        DROP COLUMN tempo
                        """
                    )
            except mysql.connector.Error:
                pass
            
            # Aggiungi ore_gestionale se non esiste
            try:
                cur.execute(
                    f"""
                    ALTER TABLE {TABLE_NAME}
                    ADD COLUMN ore_gestionale DECIMAL(10,2) DEFAULT 0 COMMENT 'Ore calcolate con sessioni (solo carrellisti)'
                    """
                )
            except mysql.connector.Error:
                pass
            
            # Tabella per le nuove aperture
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS nuove_aperture (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    data_da DATE NOT NULL,
                    data_a DATE NOT NULL,
                    negozio VARCHAR(255) NOT NULL,
                    data_inserimento TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE KEY uniq_apertura (data_da, data_a, negozio)
                )
                """
            )

            # Tabella fasce premi
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS fasce_premi (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    tipo_attivita VARCHAR(50) NOT NULL,
                    valore_riferimento DECIMAL(10,2) NOT NULL,
                    valore_premio DECIMAL(10,5) NOT NULL,
                    unita_riferimento VARCHAR(50) NOT NULL,
                    unita_premio VARCHAR(50) NOT NULL,
                    note VARCHAR(255),
                    UNIQUE KEY uniq_fascia (tipo_attivita, valore_riferimento)
                )
                """
            )

            # Popola dati di default se tabella vuota
            cur.execute("SELECT COUNT(*) FROM fasce_premi")
            count_row = cur.fetchone()
            count = count_row[0] if count_row else 0
            if count == 0:
                default_rows = [
                    ("PICKING", 100, 0.00700, "COLLI/h", "€/COL", None),
                    ("PICKING", 105, 0.00711, "COLLI/h", "€/COL", None),
                    ("PICKING", 110, 0.00722, "COLLI/h", "€/COL", None),
                    ("PICKING", 115, 0.00733, "COLLI/h", "€/COL", None),
                    ("PICKING", 120, 0.00744, "COLLI/h", "€/COL", None),
                    ("PICKING", 125, 0.00755, "COLLI/h", "€/COL", None),
                    ("PICKING", 130, 0.00766, "COLLI/h", "€/COL", None),
                    ("PICKING", 135, 0.00777, "COLLI/h", "€/COL", None),
                    ("PICKING", 140, 0.00789, "COLLI/h", "€/COL", None),
                    ("CARRELLISTI", 18, 0.03000, "Mov/h", "€/Plt", None),
                    ("CARRELLISTI", 20, 0.03630, "Mov/h", "€/Plt", None),
                    ("CARRELLISTI", 22, 0.04392, "Mov/h", "€/Plt", None),
                    ("CARRELLISTI", 24, 0.05314, "Mov/h", "€/Plt", None),
                    ("CARRELLISTI", 26, 0.06430, "Mov/h", "€/Plt", None),
                    ("RICEVITORI", 18, 2.50, "Plt/h", "€/gg", None),
                    ("RICEVITORI", 20, 3.60, "Plt/h", "€/gg", None),
                    ("RICEVITORI", 22, 5.18, "Plt/h", "€/gg", None),
                    ("RICEVITORI", 24, 7.46, "Plt/h", "€/gg", None),
                    ("DOPPIA_SPUNTA", 147, 0.00500, "Colli/h", "€/collo", None),
                    ("DOPPIA_SPUNTA", 160, 0.00525, "Colli/h", "€/collo", None),
                    ("DOPPIA_SPUNTA", 173, 0.00551, "Colli/h", "€/collo", None),
                    ("DOPPIA_SPUNTA", 186, 0.00579, "Colli/h", "€/collo", None),
                    ("DOPPIA_SPUNTA", 199, 0.00608, "Colli/h", "€/collo", None)
                ]
                cur.executemany(
                    """
                    INSERT INTO fasce_premi
                    (tipo_attivita, valore_riferimento, valore_premio, unita_riferimento, unita_premio, note)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    """,
                    default_rows
                )

            # Tabella peso movimenti
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS peso_movimenti (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    tipo_attivita VARCHAR(50) NOT NULL,
                    tipo VARCHAR(10) NOT NULL,
                    peso DECIMAL(10,3) NOT NULL,
                    note VARCHAR(255),
                    UNIQUE KEY uniq_tipo (tipo_attivita, tipo)
                )
                """
            )

            # Allinea schema se la tabella esisteva già senza tipo_attivita
            try:
                cur.execute(
                    """
                    ALTER TABLE peso_movimenti
                    ADD COLUMN tipo_attivita VARCHAR(50) NOT NULL DEFAULT 'CARRELLISTI'
                    """
                )
            except mysql.connector.Error:
                pass
            try:
                cur.execute(
                    """
                    ALTER TABLE peso_movimenti
                    DROP INDEX uniq_tipo
                    """
                )
            except mysql.connector.Error:
                pass
            try:
                cur.execute(
                    """
                    ALTER TABLE peso_movimenti
                    ADD UNIQUE INDEX uniq_tipo (tipo_attivita, tipo)
                    """
                )
            except mysql.connector.Error:
                pass

            cur.execute(
                """
                UPDATE peso_movimenti
                SET tipo_attivita = 'CARRELLISTI'
                WHERE tipo_attivita IS NULL OR tipo_attivita = ''
                """
            )

            cur.execute("SELECT COUNT(*) FROM peso_movimenti")
            peso_count_row = cur.fetchone()
            peso_count = int(peso_count_row[0]) if peso_count_row else 0  # type: ignore[index]
            if peso_count == 0:
                movimento_rows = [
                    ("CARRELLISTI", "ST", 1.0, None),
                    ("CARRELLISTI", "SS", 1.0, None),
                    ("CARRELLISTI", "CM", 1.0, None),
                    ("CARRELLISTI", "AP", 1.3, None),
                ]
                cur.executemany(
                    """
                    INSERT INTO peso_movimenti (tipo_attivita, tipo, peso, note)
                    VALUES (%s, %s, %s, %s)
                    """,
                    movimento_rows,
                )

            # Tabella malus/bonus mensile
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS malus_bonus (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    anno INT NOT NULL,
                    mese INT NOT NULL,
                    importo_rotture DECIMAL(12,2) NOT NULL DEFAULT 0,
                    importo_differenze DECIMAL(12,2) NOT NULL DEFAULT 0,
                    soglia_bonus DECIMAL(12,2) NOT NULL DEFAULT 2500,
                    soglia_rotture DECIMAL(12,2) NOT NULL DEFAULT 2500,
                    soglia_differenze DECIMAL(12,2) NOT NULL DEFAULT 2500,
                    attivita_bonus VARCHAR(255),
                    note VARCHAR(255),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    UNIQUE KEY uniq_mese (anno, mese)
                )
                """
            )

            # Tabella definizioni report export
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS report_templates (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    nome VARCHAR(100) NOT NULL,
                    descrizione VARCHAR(255),
                    sql_template TEXT NOT NULL,
                    attivo BOOLEAN NOT NULL DEFAULT TRUE,
                    attivita VARCHAR(50),
                    categoria VARCHAR(50),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    UNIQUE KEY uniq_nome (nome)
                )
                """
            )

            cur.execute(
                """
                SELECT COLUMN_NAME
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_SCHEMA = DATABASE()
                  AND TABLE_NAME = 'report_templates'
                  AND COLUMN_NAME = 'attivita'
                """
            )
            has_attivita = cur.fetchone() is not None
            if not has_attivita:
                cur.execute(
                    """
                    ALTER TABLE report_templates
                    ADD COLUMN attivita VARCHAR(50) AFTER attivo
                    """
                )

            # Tabella anomalie
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS anomalie (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    tipo_anomalia VARCHAR(50) NOT NULL COMMENT 'CODICE_NON_ABBINATO, ORE_SENZA_PRODUZIONE, PRODUZIONE_SENZA_ORE, DIFFERENZA_60_120, DIFFERENZA_>120',
                    data_rilevamento DATE NOT NULL,
                    anno INT NOT NULL COMMENT 'Anno di competenza',
                    mese INT NOT NULL COMMENT 'Mese di competenza (1-12)',
                    codice_preparatore VARCHAR(50) NOT NULL,
                    nome_preparatore VARCHAR(255),
                    tipo_attivita VARCHAR(50),
                    ore_tim DECIMAL(10,2) COMMENT 'Ore presenti in TIM (per anomalie tipo 2)',
                    dettagli TEXT COMMENT 'Descrizione dettagliata anomalia',
                    stato VARCHAR(20) DEFAULT 'APERTA' COMMENT 'APERTA, VERIFICATA, RISOLTA',
                    note VARCHAR(255),
                    data_creazione TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    data_aggiornamento TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    INDEX idx_tipo (tipo_anomalia),
                    INDEX idx_data (data_rilevamento),
                    INDEX idx_anno_mese (anno, mese),
                    INDEX idx_stato (stato),
                    INDEX idx_codice (codice_preparatore)
                )
                """
            )
            
            # Aggiungi colonne anno e mese se la tabella esiste già
            try:
                cur.execute(
                    """
                    ALTER TABLE anomalie
                    ADD COLUMN anno INT NOT NULL DEFAULT 0 COMMENT 'Anno di competenza'
                    """
                )
            except mysql.connector.Error:
                pass
            
            try:
                cur.execute(
                    """
                    ALTER TABLE anomalie
                    ADD COLUMN mese INT NOT NULL DEFAULT 0 COMMENT 'Mese di competenza (1-12)'
                    """
                )
            except mysql.connector.Error:
                pass
            
            # Aggiungi indice per anno/mese se non esiste
            try:
                cur.execute(
                    """
                    CREATE INDEX idx_anno_mese ON anomalie (anno, mese)
                    """
                )
            except mysql.connector.Error:
                pass

            # Tabella premi carrellisti
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS premi_carrellisti (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    anno INT NOT NULL,
                    mese INT NOT NULL,
                    codice_preparatore VARCHAR(50) NOT NULL,
                    nome_preparatore VARCHAR(255),
                    totale_movimenti DECIMAL(10,2) NOT NULL,
                    ore_lavorate DECIMAL(10,2) NOT NULL,
                    movimenti_ora DECIMAL(10,2) NOT NULL,
                    fascia_raggiunta VARCHAR(50),
                    premio_base DECIMAL(10,2) NOT NULL DEFAULT 0,
                    premio_kpi DECIMAL(10,2) NOT NULL DEFAULT 0,
                    premio_totale DECIMAL(10,2) NOT NULL DEFAULT 0,
                    bonus_applicato BOOLEAN DEFAULT FALSE,
                    data_calcolo TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    note VARCHAR(255),
                    UNIQUE KEY uniq_premio (anno, mese, codice_preparatore),
                    INDEX idx_anno_mese (anno, mese),
                    INDEX idx_codice (codice_preparatore)
                )
                """
            )
            
            # Tabella premi preparatori
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS premi_preparatori (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    anno INT NOT NULL,
                    mese INT NOT NULL,
                    codice_preparatore VARCHAR(50) NOT NULL,
                    nome_preparatore VARCHAR(255),
                    totale_colli DECIMAL(12,0) NOT NULL DEFAULT 0,
                    ore_lavorate DECIMAL(10,2) NOT NULL DEFAULT 0,
                    colli_ora DECIMAL(10,2) NOT NULL DEFAULT 0,
                    fascia_raggiunta VARCHAR(50),
                    premio_base DECIMAL(10,2) NOT NULL DEFAULT 0,
                    penalita_totale DECIMAL(10,2) NOT NULL DEFAULT 0,
                    premio_kpi DECIMAL(10,2) NOT NULL DEFAULT 0,
                    premio_totale DECIMAL(10,2) NOT NULL DEFAULT 0,
                    bonus_applicato BOOLEAN DEFAULT FALSE,
                    data_calcolo TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    note VARCHAR(255),
                    UNIQUE KEY uniq_premio_preparatori (anno, mese, codice_preparatore),
                    INDEX idx_preparatori_anno_mese (anno, mese),
                    INDEX idx_preparatori_codice (codice_preparatore)
                )
                """
            )
            
            # Tabella dettaglio sessioni carrellisti con colonne separate per tipo
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sessioni_carrellisti (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    data DATE NOT NULL,
                    codice_preparatore VARCHAR(50) NOT NULL,
                    numero_riga INT NOT NULL COMMENT 'Progressivo riga nella giornata',
                    ora_inizio_riga TIME NOT NULL COMMENT 'Ora inizio di questa singola riga',
                    ora_fine_riga TIME NOT NULL COMMENT 'Ora fine di questa singola riga',
                    tempo_riga_minuti DECIMAL(10,2) NOT NULL COMMENT 'Durata in minuti di questa singola riga',
                    gap_minuti DECIMAL(10,2) NULL COMMENT 'Minuti tra FINE riga precedente e INIZIO questa (NULL per prima riga)',
                    
                    -- Movimenti per tipo (1 o NULL)
                    movimenti_st INT NULL COMMENT 'Movimenti ST in questa riga',
                    movimenti_ss INT NULL COMMENT 'Movimenti SS in questa riga',
                    movimenti_ap INT NULL COMMENT 'Movimenti AP in questa riga',
                    movimenti_cm INT NULL COMMENT 'Movimenti CM in questa riga',
                    
                    -- Errore (se presente, movimento non conteggiato)
                    errore TEXT NULL COMMENT 'Descrizione errore se movimento non valido',
                    
                    -- Dettagli sessione (solo prima riga della sessione)
                    numero_sessione INT NOT NULL,
                    ora_inizio_sessione TIME NULL COMMENT 'Inizio sessione (solo prima riga)',
                    ora_fine_sessione TIME NULL COMMENT 'Fine sessione (solo prima riga)',
                    tempo_sessione_ore DECIMAL(10,2) NULL COMMENT 'Durata totale sessione (solo prima riga)',
                    totale_righe_sessione INT NULL COMMENT 'Numero righe nella sessione (solo prima riga)',
                    
                    -- Totali giornalieri per tipo (solo prima riga di ogni tipo)
                    ore_gestionale_st DECIMAL(10,2) NULL COMMENT 'Ore totali ST nella giornata (solo prima riga ST)',
                    ore_gestionale_ss DECIMAL(10,2) NULL COMMENT 'Ore totali SS nella giornata (solo prima riga SS)',
                    ore_gestionale_ap DECIMAL(10,2) NULL COMMENT 'Ore totali AP nella giornata (solo prima riga AP)',
                    ore_gestionale_cm DECIMAL(10,2) NULL COMMENT 'Ore totali CM nella giornata (solo prima riga CM)',
                    
                    data_importazione TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    INDEX idx_data_codice (data, codice_preparatore),
                    INDEX idx_data (data),
                    INDEX idx_sessione (data, codice_preparatore, numero_sessione)
                )
                """
            )
            
            conn.commit()

            # Allinea schema esistente con la colonna attività bonus
            _ensure_malus_bonus_schema(cur)
            _ensure_sessioni_carrellisti_schema(cur)


def _ensure_sessioni_carrellisti_schema(cur: Any) -> None:
    """Garantisce la presenza della colonna errore nella tabella sessioni_carrellisti."""
    connection = getattr(cur, "connection", None)
    if connection is None:
        return

    alterations = [
        (
            "movimenti_ss",
            """
                ALTER TABLE sessioni_carrellisti
                ADD COLUMN movimenti_ss INT NULL COMMENT 'Movimenti SS in questa riga'
                AFTER movimenti_st
            """
        ),
        (
            "errore",
            """
                ALTER TABLE sessioni_carrellisti
                ADD COLUMN errore TEXT NULL COMMENT 'Descrizione errore se movimento non valido'
                AFTER movimenti_cm
            """
        ),
        (
            "ore_gestionale_ss",
            """
                ALTER TABLE sessioni_carrellisti
                ADD COLUMN ore_gestionale_ss DECIMAL(10,2) NULL COMMENT 'Ore totali SS nella giornata (solo prima riga SS)'
                AFTER ore_gestionale_st
            """
        ),
    ]

    for column_name, alter_sql in alterations:
        try:
            cur.execute(
                """
                    SELECT COUNT(*)
                    FROM information_schema.COLUMNS
                    WHERE TABLE_SCHEMA = DATABASE()
                      AND TABLE_NAME = 'sessioni_carrellisti'
                      AND COLUMN_NAME = %s
                """,
                (column_name,),
            )
            result = cur.fetchone()

            if result and result[0] == 0:
                cur.execute(alter_sql)
                connection.commit()
                print(f"[OK] Colonna '{column_name}' aggiunta a sessioni_carrellisti")
        except Exception as exc:
            print(
                f"[WARN] Errore durante l'aggiornamento schema sessioni_carrellisti (colonna {column_name}): {exc}"
            )


def _ensure_malus_bonus_schema(cur: Any) -> None:
    """Garantisce la presenza delle colonne opzionali nella tabella malus_bonus."""
    connection = getattr(cur, "connection", None)
    alterations = [
        (
            "attivita_bonus",
            """
            ALTER TABLE malus_bonus
            ADD COLUMN attivita_bonus VARCHAR(255) NULL
            """,
            False,
        ),
        (
            "soglia_rotture",
            """
            ALTER TABLE malus_bonus
            ADD COLUMN soglia_rotture DECIMAL(12,2) NULL
            """,
            True,
        ),
        (
            "soglia_differenze",
            """
            ALTER TABLE malus_bonus
            ADD COLUMN soglia_differenze DECIMAL(12,2) NULL
            """,
            True,
        ),
    ]

    for column, statement, copy_from_legacy in alterations:
        try:
            cur.execute(statement)
            if connection is not None:
                connection.commit()
            if copy_from_legacy:
                cur.execute(
                    f"UPDATE malus_bonus SET {column} = soglia_bonus WHERE {column} IS NULL"
                )
                if connection is not None:
                    connection.commit()
        except mysql.connector.Error as exc:
            errno = getattr(exc, "errno", None)
            if errno in {errorcode.ER_DUP_FIELDNAME, errorcode.ER_NO_SUCH_TABLE}:
                if copy_from_legacy and errno != errorcode.ER_NO_SUCH_TABLE:
                    try:
                        cur.execute(
                            f"UPDATE malus_bonus SET {column} = soglia_bonus WHERE {column} IS NULL"
                        )
                        if connection is not None:
                            connection.commit()
                    except mysql.connector.Error:
                        pass
                continue
            raise


def insert_batch_data(values: List[Tuple]) -> int:
    """
    Inserisce i dati in batch nel database usando upsert.
    
    Args:
        values: Lista di tuple con i dati da inserire (inclusi ore_tim e ore_gestionale)
        
    Returns:
        Numero di record inseriti/aggiornati
    """
    if not values:
        return 0
        
    sql = f"""
        INSERT INTO {TABLE_NAME}
        (data, codice_preparatore, nome_preparatore, totale_colli, penalita, tipo_attivita, tipo, ore_tim, ore_gestionale)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
        ON DUPLICATE KEY UPDATE
            nome_preparatore = VALUES(nome_preparatore),
            totale_colli     = VALUES(totale_colli),
            penalita         = VALUES(penalita),
            tipo_attivita    = VALUES(tipo_attivita),
            tipo             = VALUES(tipo),
            ore_tim          = VALUES(ore_tim),
            ore_gestionale   = VALUES(ore_gestionale)
    """

    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            try:
                cur.executemany(sql, values)
                conn.commit()
                return len(values)
            except Exception as e:
                print("\n[ERROR] Errore durante insert_batch_data:")
                print(f"   Errore: {e}")
                print("\n   Tentativo di inserimento riga per riga per trovare il record problematico...")
                
                # Prova a inserire riga per riga per trovare il problema
                for i, val in enumerate(values):
                    try:
                        cur.execute(sql.replace("VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)", "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)"), val)
                        conn.commit()
                    except Exception as row_error:
                        print(f"\n[FAIL] Errore alla riga {i}:")
                        print(f"   Valori: {val}")
                        print(f"   Tipi: {[type(v).__name__ for v in val]}")
                        print(f"   Errore: {row_error}")
                        raise  # Rilancia l'errore originale
                
                return len(values)


def update_penalita_picking(values: List[Tuple[datetime.date, str, int]]) -> int:
    """Aggiorna la penalità delle attività PICKING per data e codice preparatore."""

    if not values:
        return 0

    sql = f"""
        UPDATE {TABLE_NAME}
        SET penalita = %s
        WHERE data = %s
          AND codice_preparatore = %s
          AND tipo_attivita = %s
    """

    params = [(pen, data, codice, "PICKING") for data, codice, pen in values]

    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.executemany(sql, params)
            conn.commit()
            return cur.rowcount


def save_nuove_aperture(data_da: str, data_a: str, negozi: List[str]) -> int:
    """
    Salva le nuove aperture nel database.
    Prima elimina i negozi esistenti per il periodo, poi inserisce i nuovi.
    
    Args:
        data_da: Data inizio periodo (formato YYYY-MM-DD)
        data_a: Data fine periodo (formato YYYY-MM-DD)
        negozi: Lista dei nomi negozi selezionati
        
    Returns:
        Numero di negozi salvati
    """
    if not negozi:
        return 0
    
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            # Elimina i negozi esistenti per questo periodo
            cur.execute(
                """
                DELETE FROM nuove_aperture
                WHERE data_da = %s AND data_a = %s
                """,
                (data_da, data_a)
            )
            
            # Inserisce i nuovi negozi
            sql = """
                INSERT INTO nuove_aperture (data_da, data_a, negozio)
                VALUES (%s, %s, %s)
            """
            params = [(data_da, data_a, negozio) for negozio in negozi]
            cur.executemany(sql, params)
            conn.commit()
            return len(negozi)


def load_nuove_aperture(data_da: str, data_a: str) -> List[str]:
    """
    Carica le nuove aperture dal database per un periodo specifico.
    
    Args:
        data_da: Data inizio periodo (formato YYYY-MM-DD)
        data_a: Data fine periodo (formato YYYY-MM-DD)
        
    Returns:
        Lista dei nomi negozi salvati per il periodo
    """
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                """
                SELECT negozio
                FROM nuove_aperture
                WHERE data_da = %s AND data_a = %s
                ORDER BY negozio
                """,
                (data_da, data_a)
            )
            rows = cur.fetchall()
            return [str(row[0]) for row in rows]  # type: ignore[index]


def fetch_fasce_premi(tipo: Optional[str] = None) -> List[Dict[str, Any]]:
    """Recupera le fasce premio, opzionalmente filtrate per tipo attività."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            if tipo:
                cur.execute(
                    """
                    SELECT id, tipo_attivita, valore_riferimento, valore_premio,
                           unita_riferimento, unita_premio, note
                    FROM fasce_premi
                    WHERE tipo_attivita = %s
                    ORDER BY valore_riferimento
                    """,
                    (tipo,)
                )
            else:
                cur.execute(
                    """
                    SELECT id, tipo_attivita, valore_riferimento, valore_premio,
                           unita_riferimento, unita_premio, note
                    FROM fasce_premi
                    ORDER BY tipo_attivita, valore_riferimento
                    """
                )
            return cur.fetchall()


def fetch_pesi_movimenti(tipo_attivita: Optional[str] = None) -> List[Dict[str, Any]]:
    """Restituisce il peso dei movimenti per ogni tipo e attività."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            if tipo_attivita:
                cur.execute(
                    """
                    SELECT id, tipo_attivita, tipo, peso, note
                    FROM peso_movimenti
                    WHERE tipo_attivita = %s
                    ORDER BY tipo
                    """,
                    (tipo_attivita,),
                )
            else:
                cur.execute(
                    """
                    SELECT id, tipo_attivita, tipo, peso, note
                    FROM peso_movimenti
                    ORDER BY tipo_attivita, tipo
                    """
                )
            result = cur.fetchall()
            return cast(List[Dict[str, Any]], result)


def insert_fascia_premio(
    tipo_attivita: str,
    valore_riferimento: float,
    valore_premio: float,
    unita_riferimento: str,
    unita_premio: str,
    note: Optional[str] = None,
) -> int:
    """Inserisce una nuova fascia premio."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                """
                INSERT INTO fasce_premi
                (tipo_attivita, valore_riferimento, valore_premio, unita_riferimento, unita_premio, note)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (
                    tipo_attivita,
                    valore_riferimento,
                    valore_premio,
                    unita_riferimento,
                    unita_premio,
                    note,
                ),
            )
            conn.commit()
            last_id = cur.lastrowid
            return cast(int, last_id) if last_id is not None else 0


def insert_peso_movimento(
    tipo_attivita: str,
    tipo: str,
    peso: float,
    note: Optional[str] = None,
) -> int:
    """Inserisce un nuovo peso movimento."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                """
                INSERT INTO peso_movimenti (tipo_attivita, tipo, peso, note)
                VALUES (%s, %s, %s, %s)
                """,
                (tipo_attivita, tipo, peso, note),
            )
            conn.commit()
            last_id = cur.lastrowid
            return cast(int, last_id) if last_id is not None else 0


def update_fascia_premio(
    fascia_id: int,
    tipo_attivita: str,
    valore_riferimento: float,
    valore_premio: float,
    unita_riferimento: str,
    unita_premio: str,
    note: Optional[str] = None,
) -> None:
    """Aggiorna una fascia premio esistente."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                """
                UPDATE fasce_premi
                SET tipo_attivita = %s,
                    valore_riferimento = %s,
                    valore_premio = %s,
                    unita_riferimento = %s,
                    unita_premio = %s,
                    note = %s
                WHERE id = %s
                """,
                (
                    tipo_attivita,
                    valore_riferimento,
                    valore_premio,
                    unita_riferimento,
                    unita_premio,
                    note,
                    fascia_id,
                ),
            )
            conn.commit()


def update_peso_movimento(
    peso_id: int,
    tipo_attivita: str,
    tipo: str,
    peso: float,
    note: Optional[str] = None,
) -> None:
    """Aggiorna un peso movimento esistente."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                """
                UPDATE peso_movimenti
                SET tipo_attivita = %s,
                    tipo = %s,
                    peso = %s,
                    note = %s
                WHERE id = %s
                """,
                (tipo_attivita, tipo, peso, note, peso_id),
            )
            conn.commit()


def delete_fascia_premio(fascia_id: int) -> None:
    """Elimina una fascia premio."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute("DELETE FROM fasce_premi WHERE id = %s", (fascia_id,))
            conn.commit()


def delete_peso_movimento(peso_id: int) -> None:
    """Elimina un peso movimento."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute("DELETE FROM peso_movimenti WHERE id = %s", (peso_id,))
            conn.commit()


def upsert_malus_bonus(
    anno: int,
    mese: int,
    importo_rotture: float,
    importo_differenze: float,
    soglia_rotture: float,
    soglia_differenze: float,
    attivita_bonus: Optional[str] = None,
    note: Optional[str] = None,
) -> None:
    """Inserisce o aggiorna il record malus/bonus per un mese."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            _ensure_malus_bonus_schema(cur)
            legacy_soglia = max(soglia_rotture, soglia_differenze)
            cur.execute(
                """
                INSERT INTO malus_bonus (
                    anno, mese, importo_rotture, importo_differenze, soglia_bonus, soglia_rotture, soglia_differenze, attivita_bonus, note
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    importo_rotture = VALUES(importo_rotture),
                    importo_differenze = VALUES(importo_differenze),
                    soglia_bonus = VALUES(soglia_bonus),
                    soglia_rotture = VALUES(soglia_rotture),
                    soglia_differenze = VALUES(soglia_differenze),
                    attivita_bonus = VALUES(attivita_bonus),
                    note = VALUES(note)
                """,
                (
                    anno,
                    mese,
                    importo_rotture,
                    importo_differenze,
                    legacy_soglia,
                    soglia_rotture,
                    soglia_differenze,
                    attivita_bonus,
                    note,
                ),
            )
            conn.commit()


def fetch_malus_bonus(anno: Optional[int] = None) -> List[Dict[str, Any]]:
    """Recupera i record malus/bonus, opzionalmente filtrati per anno."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            _ensure_malus_bonus_schema(cur)
            if anno:
                cur.execute(
                    """
                    SELECT id, anno, mese, importo_rotture, importo_differenze, soglia_bonus, soglia_rotture, soglia_differenze, attivita_bonus, note
                    FROM malus_bonus
                    WHERE anno = %s
                    ORDER BY mese
                    """,
                    (anno,),
                )
            else:
                cur.execute(
                    """
                    SELECT id, anno, mese, importo_rotture, importo_differenze, soglia_bonus, soglia_rotture, soglia_differenze, attivita_bonus, note
                    FROM malus_bonus
                    ORDER BY anno DESC, mese DESC
                    """
                )
            result = cur.fetchall()
            return cast(List[Dict[str, Any]], result)


def get_malus_bonus(anno: int, mese: int) -> Optional[Dict[str, Any]]:
    """Restituisce il record malus/bonus per anno e mese."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            _ensure_malus_bonus_schema(cur)
            cur.execute(
                """
                SELECT id, anno, mese, importo_rotture, importo_differenze, soglia_bonus, soglia_rotture, soglia_differenze, attivita_bonus, note
                FROM malus_bonus
                WHERE anno = %s AND mese = %s
                LIMIT 1
                """,
                (anno, mese),
            )
            row = cur.fetchone()
            return cast(Optional[Dict[str, Any]], row)


def delete_malus_bonus(record_id: int) -> None:
    """Elimina un record malus/bonus."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            _ensure_malus_bonus_schema(cur)
            cur.execute("DELETE FROM malus_bonus WHERE id = %s", (record_id,))
            conn.commit()


# ============== GESTIONE ANOMALIE ==============


def insert_anomalia(
    tipo_anomalia: str,
    data_rilevamento: datetime.date,
    codice_preparatore: str,
    nome_preparatore: Optional[str] = None,
    tipo_attivita: Optional[str] = None,
    ore_tim: Optional[float] = None,
    dettagli: Optional[str] = None,
    note: Optional[str] = None,
) -> int:
    """Inserisce una nuova anomalia."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            # Estrae anno e mese dalla data di rilevamento
            anno = data_rilevamento.year
            mese = data_rilevamento.month
            
            cur.execute(
                """
                INSERT INTO anomalie (
                    tipo_anomalia, data_rilevamento, anno, mese, codice_preparatore, nome_preparatore,
                    tipo_attivita, ore_tim, dettagli, note
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    tipo_anomalia,
                    data_rilevamento,
                    anno,
                    mese,
                    codice_preparatore,
                    nome_preparatore,
                    tipo_attivita,
                    ore_tim,
                    dettagli,
                    note,
                ),
            )
            conn.commit()
            last_id = cur.lastrowid
            return cast(int, last_id) if last_id is not None else 0


def fetch_anomalie(
    tipo_anomalia: Optional[str | List[str]] = None,
    stato: Optional[str] = None,
    data_da: Optional[datetime.date] = None,
    data_a: Optional[datetime.date] = None,
    tipo_attivita: Optional[str] = None,
    codice_preparatore: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Recupera le anomalie con filtri opzionali. tipo_anomalia può essere una stringa o una lista."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            conditions: List[str] = []
            params: List[Any] = []

            if tipo_anomalia:
                if isinstance(tipo_anomalia, list):
                    # Selezione multipla
                    placeholders = ", ".join(["%s"] * len(tipo_anomalia))
                    conditions.append(f"tipo_anomalia IN ({placeholders})")
                    params.extend(tipo_anomalia)
                else:
                    # Singola selezione
                    conditions.append("tipo_anomalia = %s")
                    params.append(tipo_anomalia)
            if stato:
                conditions.append("stato = %s")
                params.append(stato)
            if tipo_attivita:
                conditions.append("tipo_attivita = %s")
                params.append(tipo_attivita)
            if codice_preparatore:
                conditions.append("codice_preparatore = %s")
                params.append(codice_preparatore)
            if data_da:
                conditions.append("data_rilevamento >= %s")
                params.append(data_da)
            if data_a:
                conditions.append("data_rilevamento <= %s")
                params.append(data_a)

            where_clause = " WHERE " + " AND ".join(conditions) if conditions else ""
            query = f"""
                SELECT id, tipo_anomalia, data_rilevamento, anno, mese, codice_preparatore, nome_preparatore,
                       tipo_attivita, ore_tim, dettagli, stato, note, data_creazione, data_aggiornamento
                FROM anomalie
                {where_clause}
                ORDER BY data_rilevamento ASC, anno ASC, mese ASC, data_creazione ASC
            """
            cur.execute(query, params)
            result = cur.fetchall()
            return cast(List[Dict[str, Any]], result)


def fetch_report_templates(
    attivi_solo: bool = True,
    attivita: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Restituisce la lista dei report configurati per gli export."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            base_query = (
                "SELECT id, nome, descrizione, sql_template, attivo, attivita, categoria, "
                "created_at, updated_at FROM report_templates"
            )
            conditions: List[str] = []
            params: List[Any] = []

            if attivi_solo:
                conditions.append("attivo = TRUE")

            if attivita:
                conditions.append("(attivita IS NULL OR attivita = '' OR attivita = 'TUTTE' OR attivita = %s)")
                params.append(attivita)
            else:
                conditions.append("(attivita IS NULL OR attivita = '' OR attivita = 'TUTTE')")

            query = base_query
            if conditions:
                query += " WHERE " + " AND ".join(conditions)
            query += " ORDER BY nome ASC"

            cur.execute(query, params)
            result = cur.fetchall() or []
            return cast(List[Dict[str, Any]], result)


def execute_custom_query(query: str, params: Sequence[Any]) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Esegue una query arbitraria e restituisce righe e intestazioni."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            cur.execute(query, tuple(params))
            rows = cur.fetchall() or []
            columns = [col[0] for col in cur.description] if cur.description else []
            return cast(List[Dict[str, Any]], rows), columns


def update_anomalia_stato(anomalia_id: int, nuovo_stato: str, note: Optional[str] = None) -> None:
    """Aggiorna lo stato di un'anomalia."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            if note:
                cur.execute(
                    """
                    UPDATE anomalie
                    SET stato = %s, note = %s
                    WHERE id = %s
                    """,
                    (nuovo_stato, note, anomalia_id),
                )
            else:
                cur.execute(
                    """
                    UPDATE anomalie
                    SET stato = %s
                    WHERE id = %s
                    """,
                    (nuovo_stato, anomalia_id),
                )
            conn.commit()


# ========== GESTIONE PREMI CARRELLISTI ==========

def save_premi_carrellisti(anno: int, mese: int, premi: List[Dict[str, Any]]) -> None:
    """Salva i premi carrellisti per un dato mese. Se esistono già, li sovrascrive."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            # Prima elimina i premi esistenti per quel mese
            cur.execute(
                "DELETE FROM premi_carrellisti WHERE anno = %s AND mese = %s",
                (anno, mese)
            )
            
            # Inserisci i nuovi premi
            for premio in premi:
                cur.execute(
                    """
                    INSERT INTO premi_carrellisti 
                    (anno, mese, codice_preparatore, nome_preparatore, totale_movimenti, 
                     ore_lavorate, movimenti_ora, fascia_raggiunta, premio_base, 
                     premio_kpi, premio_totale, bonus_applicato, note)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        anno,
                        mese,
                        premio.get("codice"),
                        premio.get("nome"),
                        premio.get("tot_movimenti"),
                        premio.get("ore"),
                        premio.get("mov_ora"),
                        premio.get("fascia"),
                        premio.get("premio_base"),
                        premio.get("premio_kpi"),
                        premio.get("premio_totale"),
                        premio.get("bonus_applicato", False),
                        premio.get("note"),
                    )
                )
            conn.commit()


def fetch_premi_carrellisti(
    anno: Optional[int] = None,
    mese: Optional[int] = None,
    codice_preparatore: Optional[str] = None
) -> List[Dict[str, Any]]:
    """Recupera i premi carrellisti filtrati per anno/mese/codice."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            conditions = []
            params: List[Any] = []
            
            if anno:
                conditions.append("anno = %s")
                params.append(anno)
            if mese:
                conditions.append("mese = %s")
                params.append(mese)
            if codice_preparatore:
                conditions.append("codice_preparatore = %s")
                params.append(codice_preparatore)
            
            where_clause = " WHERE " + " AND ".join(conditions) if conditions else ""
            query = f"""
                SELECT id, anno, mese, codice_preparatore, nome_preparatore,
                       totale_movimenti, ore_lavorate, movimenti_ora, fascia_raggiunta,
                       premio_base, premio_kpi, premio_totale, bonus_applicato,
                       data_calcolo, note
                FROM premi_carrellisti
                {where_clause}
                ORDER BY premio_totale DESC
            """
            cur.execute(query, params)
            result = cur.fetchall()
            return cast(List[Dict[str, Any]], result)


def delete_premi_carrellisti(anno: int, mese: int) -> None:
    """Elimina i premi carrellisti per un dato mese."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                "DELETE FROM premi_carrellisti WHERE anno = %s AND mese = %s",
                (anno, mese)
            )
            conn.commit()


def save_premi_preparatori(anno: int, mese: int, premi: List[Dict[str, Any]]) -> None:
    """Salva i premi preparatori per un mese specifico sovrascrivendo quelli esistenti."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                "DELETE FROM premi_preparatori WHERE anno = %s AND mese = %s",
                (anno, mese),
            )

            for premio in premi:
                cur.execute(
                    """
                    INSERT INTO premi_preparatori (
                        anno, mese, codice_preparatore, nome_preparatore,
                        totale_colli, ore_lavorate, colli_ora, fascia_raggiunta,
                        premio_base, penalita_totale, premio_kpi, premio_totale,
                        bonus_applicato, note
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        anno,
                        mese,
                        premio.get("codice"),
                        premio.get("nome"),
                        premio.get("tot_colli"),
                        premio.get("ore"),
                        premio.get("colli_ora"),
                        premio.get("fascia"),
                        premio.get("premio_base"),
                        premio.get("penalita"),
                        premio.get("premio_kpi"),
                        premio.get("premio_totale"),
                        premio.get("bonus_applicato", False),
                        premio.get("note"),
                    ),
                )
            conn.commit()


def fetch_premi_preparatori(
    anno: Optional[int] = None,
    mese: Optional[int] = None,
    codice_preparatore: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Restituisce i premi preparatori filtrati per anno/mese/codice."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            conditions: List[str] = []
            params: List[Any] = []

            if anno:
                conditions.append("anno = %s")
                params.append(anno)
            if mese:
                conditions.append("mese = %s")
                params.append(mese)
            if codice_preparatore:
                conditions.append("codice_preparatore = %s")
                params.append(codice_preparatore)

            where_clause = " WHERE " + " AND ".join(conditions) if conditions else ""

            query = f"""
                SELECT
                    id, anno, mese, codice_preparatore, nome_preparatore,
                    totale_colli, ore_lavorate, colli_ora, fascia_raggiunta,
                    premio_base, penalita_totale, premio_kpi, premio_totale,
                    bonus_applicato, data_calcolo, note
                FROM premi_preparatori
                {where_clause}
                ORDER BY premio_totale DESC
            """
            cur.execute(query, params)
            result = cur.fetchall()
            return cast(List[Dict[str, Any]], result)


def delete_premi_preparatori(anno: int, mese: int) -> None:
    """Elimina i premi preparatori per un determinato mese."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                "DELETE FROM premi_preparatori WHERE anno = %s AND mese = %s",
                (anno, mese),
            )
            conn.commit()


def delete_anomalia(anomalia_id: int) -> None:
    """Elimina un'anomalia."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute("DELETE FROM anomalie WHERE id = %s", (anomalia_id,))
            conn.commit()


def clear_anomalie_by_date(data_rilevamento: datetime.date, tipo_anomalia: Optional[str] = None) -> int:
    """Elimina anomalie per una data specifica (utile prima di rigenerare)."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            if tipo_anomalia:
                cur.execute(
                    "DELETE FROM anomalie WHERE data_rilevamento = %s AND tipo_anomalia = %s",
                    (data_rilevamento, tipo_anomalia),
                )
            else:
                cur.execute(
                    "DELETE FROM anomalie WHERE data_rilevamento = %s",
                    (data_rilevamento,),
                )
            conn.commit()
            return cur.rowcount


def save_sessioni_carrellisti(sessioni: List[Dict[str, Any]]) -> None:
    """Salva i dettagli delle sessioni carrellisti nel database con colonne separate per tipo.
    Include gap_minuti che indica i minuti tra la FINE della riga precedente e l'INIZIO di questa."""
    if not sessioni:
        print("[WARN] save_sessioni_carrellisti: nessuna sessione da salvare")
        return
    
    print(f"[INFO] save_sessioni_carrellisti: ricevute {len(sessioni)} sessioni")
    
    try:
        with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
            # Disabilita autocommit per gestire transazione manuale
            conn.autocommit = False
            
            with closing(conn.cursor()) as cur:
                # Allinea lo schema nel caso l'istanza non abbia ancora le nuove colonne SS
                _ensure_sessioni_carrellisti_schema(cur)

                # Prima elimina le sessioni esistenti per le stesse date/codici
                date_codici = set((s['data'], s['codice_preparatore']) for s in sessioni)
                print(f"[INFO] Eliminazione sessioni esistenti per {len(date_codici)} combinazioni data/codice...")
                deleted_count = 0
                for data, codice in date_codici:
                    cur.execute(
                        "DELETE FROM sessioni_carrellisti WHERE data = %s AND codice_preparatore = %s",
                        (data, codice)
                    )
                    deleted_count += cur.rowcount
                print(f"  [INFO] Eliminate {deleted_count} righe esistenti")

                print(f"[INFO] Preparazione insert di {len(sessioni)} righe...")
                
                # Inserisci le nuove righe con colonne separate per tipo
                cur.executemany(
                """
                INSERT INTO sessioni_carrellisti 
                (data, codice_preparatore, numero_riga, ora_inizio_riga, ora_fine_riga, tempo_riga_minuti, gap_minuti,
                 movimenti_st, movimenti_ss, movimenti_ap, movimenti_cm, errore,
                 numero_sessione, ora_inizio_sessione, ora_fine_sessione, tempo_sessione_ore,
                 totale_righe_sessione,
                 ore_gestionale_st, ore_gestionale_ss, ore_gestionale_ap, ore_gestionale_cm)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                [
                    (
                        s['data'],
                        s['codice_preparatore'],
                        s['numero_riga'],
                        s['ora_inizio_riga'],
                        s['ora_fine_riga'],
                        s['tempo_riga_minuti'],
                        s['gap_minuti'],
                        s.get('movimenti_st'),
                        s.get('movimenti_ss'),
                        s.get('movimenti_ap'),
                        s.get('movimenti_cm'),
                        s.get('errore'),
                        s['numero_sessione'],
                        s['ora_inizio_sessione'],
                        s['ora_fine_sessione'],
                        s['tempo_sessione_ore'],
                        s['totale_righe_sessione'],
                        s.get('ore_gestionale_st'),
                        s.get('ore_gestionale_ss'),
                        s.get('ore_gestionale_ap'),
                        s.get('ore_gestionale_cm')
                    )
                    for s in sessioni
                ]
                )
                rows_affected = cur.rowcount
                print(f"[INFO] Righe inserite: {rows_affected}")
                
                # Commit solo se tutto è andato bene
                conn.commit()
                print(f"[OK] Salvate {rows_affected} righe in sessioni_carrellisti (transazione completata)")
    except Exception as e:
        import traceback
        print(f"[ERROR] Errore in save_sessioni_carrellisti: {e}")
        print(traceback.format_exc())
        # Rollback in caso di errore per ripristinare i dati eliminati
        try:
            if 'conn' in locals():
                conn.rollback()
                print("[INFO] Rollback eseguito - dati precedenti ripristinati")
        except:
            pass
        raise


def fetch_sessioni_carrellisti(
    data_da: Optional[datetime.date] = None,
    data_a: Optional[datetime.date] = None,
    codice_preparatore: Optional[str] = None
) -> List[Dict[str, Any]]:
    """Recupera i dettagli delle sessioni carrellisti, includendo colonne per tipo."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            conditions = []
            params = []
            
            if data_da:
                conditions.append("data >= %s")
                params.append(data_da)
            if data_a:
                conditions.append("data <= %s")
                params.append(data_a)
            if codice_preparatore:
                conditions.append("codice_preparatore = %s")
                params.append(codice_preparatore)
            
            where_clause = " WHERE " + " AND ".join(conditions) if conditions else ""
            query = f"""
                SELECT
                    id,
                    data,
                    codice_preparatore,
                    numero_riga,
                    ora_inizio_riga,
                    ora_fine_riga,
                    tempo_riga_minuti,
                    gap_minuti,
                    movimenti_st,
                    movimenti_ss,
                    movimenti_ap,
                    movimenti_cm,
                    errore,
                    numero_sessione,
                    ora_inizio_sessione,
                    ora_fine_sessione,
                    tempo_sessione_ore,
                    totale_righe_sessione,
                    ore_gestionale_st,
                    ore_gestionale_ss,
                    ore_gestionale_ap,
                    ore_gestionale_cm,
                    data_importazione,
                    CASE
                        WHEN movimenti_st IS NOT NULL THEN 'ST'
                        WHEN movimenti_ss IS NOT NULL THEN 'SS'
                        WHEN movimenti_ap IS NOT NULL THEN 'AP'
                        WHEN movimenti_cm IS NOT NULL THEN 'CM'
                        ELSE NULL
                    END AS tipo,
                    COALESCE(movimenti_st, movimenti_ss, movimenti_ap, movimenti_cm) AS movimenti,
                    CASE
                        WHEN movimenti_st IS NOT NULL THEN ore_gestionale_st
                        WHEN movimenti_ss IS NOT NULL THEN ore_gestionale_ss
                        WHEN movimenti_ap IS NOT NULL THEN ore_gestionale_ap
                        WHEN movimenti_cm IS NOT NULL THEN ore_gestionale_cm
                        ELSE NULL
                    END AS ore_gestionale
                FROM sessioni_carrellisti
                {where_clause}
                ORDER BY data, codice_preparatore, numero_sessione, numero_riga
            """
            cur.execute(query, params)
            result = cur.fetchall()
            return cast(List[Dict[str, Any]], result)


def delete_sessioni_carrellisti(data: datetime.date, codice_preparatore: Optional[str] = None) -> None:
    """Elimina le sessioni carrellisti per una data (e opzionalmente un codice)."""
    with closing(mysql.connector.connect(**MYSQL_CONFIG)) as conn:
        with closing(conn.cursor()) as cur:
            if codice_preparatore:
                cur.execute(
                    "DELETE FROM sessioni_carrellisti WHERE data = %s AND codice_preparatore = %s",
                    (data, codice_preparatore)
                )
            else:
                cur.execute(
                    "DELETE FROM sessioni_carrellisti WHERE data = %s",
                    (data,)
                )
            conn.commit()