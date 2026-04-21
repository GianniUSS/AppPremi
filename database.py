"""
Gestione database: connessioni, creazione tabelle e operazioni CRUD.
"""
import datetime
import random
from contextlib import closing
from typing import Any, Dict, List, Optional, Sequence, Tuple, cast
import mysql.connector
from mysql.connector import errorcode
from mysql.connector.pooling import MySQLConnectionPool
from config import MYSQL_CONFIG, MYSQL_CONFIG_MAIN, TABLE_NAME

# ---------------------------------------------------------------------------
# Connection pool — evita di aprire una nuova connessione TCP per ogni query.
# Le connessioni vengono riusate: risparmio di ~30-100 ms a chiamata.
# ---------------------------------------------------------------------------
def _create_pool(name: str, config: dict, size: int = 5) -> Optional[MySQLConnectionPool]:
    """Crea il pool; se il server non è raggiungibile ritorna None (gestione graceful)."""
    try:
        return MySQLConnectionPool(pool_name=name, pool_size=size, **config)
    except Exception:
        return None

_pool_tim: Optional[MySQLConnectionPool] = _create_pool("appeden_tim", MYSQL_CONFIG)
_pool_main: Optional[MySQLConnectionPool] = _create_pool("appeden_main", MYSQL_CONFIG_MAIN)


def _get_conn():
    """Restituisce una connessione dal pool TIM (o ne apre una nuova come fallback)."""
    if _pool_tim:
        return _pool_tim.get_connection()
    return mysql.connector.connect(**MYSQL_CONFIG)


def _get_conn_main():
    """Restituisce una connessione dal pool MAIN (o ne apre una nuova come fallback)."""
    if _pool_main:
        return _pool_main.get_connection()
    return mysql.connector.connect(**MYSQL_CONFIG_MAIN)


def _column_exists(cur: Any, table_name: str, column_name: str) -> bool:
    """Return True if the given column exists in the specified table."""
    cur.execute(
        """
        SELECT COUNT(*)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
            AND TABLE_NAME = %s
            AND COLUMN_NAME = %s
        """,
        (table_name, column_name),
    )
    result = cur.fetchone()
    if not result:
        return False
    if isinstance(result, dict):
        return bool(next(iter(result.values()), 0))
    return bool(result[0])


def _tim_column_exists(cur: Any, table_name: str, column_name: str) -> bool:
    """Return True if the given column exists in the TIM schema."""
    cur.execute(
        """
        SELECT COUNT(*)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
            AND TABLE_NAME = %s
            AND COLUMN_NAME = %s
        """,
        (table_name, column_name),
    )
    result = cur.fetchone()
    if not result:
        return False
    if isinstance(result, dict):
        return bool(next(iter(result.values()), 0))
    return bool(result[0])


def _generate_tim_long_id(prefix: str = "001") -> int:
    """Genera un id numerico stile legacy con prefisso a 3 cifre."""
    now = datetime.datetime.now()
    date_part = now.strftime("%y%m%d")  # 6 cifre
    ms_part = int(now.timestamp() * 1000) % 10_000_000  # 7 cifre
    prog_part = random.randint(0, 99)  # 2 cifre
    rand_part = random.randint(0, 9)  # 1 cifra
    id_str = f"{prefix}{date_part}{ms_part:07d}{prog_part:02d}{rand_part}"
    return int(id_str)


def ensure_table_and_indexes() -> None:
    """Crea la tabella e l'indice unico se non esistono."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                f"""
                CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    data DATE NOT NULL,
                    codice_preparatore VARCHAR(50) NOT NULL,
                    nome_preparatore VARCHAR(255),
                    totale_colli INT,
                    penalita_eccesso INT DEFAULT 0,
                    penalita_difetto INT DEFAULT 0,
                    penalita_totale DECIMAL(10,2) DEFAULT 0 COMMENT 'Penalità totale (abs(diff/uvc) sommato)',
                    tipo_attivita VARCHAR(50) NOT NULL,
                    tipo VARCHAR(20),
                    ore_tim DECIMAL(10,2) DEFAULT 0 COMMENT 'MINUTI di lavoro da TIM',
                    ore_gestionale DECIMAL(10,2) DEFAULT 0 COMMENT 'MINUTI calcolati con sessioni (solo carrellisti)',
                    UNIQUE KEY uniq_record (data, codice_preparatore, tipo_attivita, tipo)
                )
                """
            )
            
            # Migrazione colonne penalità separate
            for column_name, definition in (
                ("penalita_eccesso", f"ALTER TABLE {TABLE_NAME} ADD COLUMN penalita_eccesso INT DEFAULT 0 AFTER totale_colli"),
                ("penalita_difetto", f"ALTER TABLE {TABLE_NAME} ADD COLUMN penalita_difetto INT DEFAULT 0 AFTER penalita_eccesso"),
            ):
                try:
                    if not _column_exists(cur, TABLE_NAME, column_name):
                        cur.execute(definition)
                except mysql.connector.Error:
                    pass

            if _column_exists(cur, TABLE_NAME, "penalita"):
                try:
                    cur.execute(
                        f"""
                        UPDATE {TABLE_NAME}
                        SET penalita_eccesso = GREATEST(penalita, 0),
                            penalita_difetto = GREATEST(-penalita, 0)
                        WHERE penalita IS NOT NULL
                        """
                    )
                    cur.execute(f"ALTER TABLE {TABLE_NAME} DROP COLUMN penalita")
                except mysql.connector.Error:
                    pass
            
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
                        ADD COLUMN ore_tim DECIMAL(10,2) DEFAULT 0 COMMENT 'MINUTI di lavoro da TIM'
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
                    ADD COLUMN ore_gestionale DECIMAL(10,2) DEFAULT 0 COMMENT 'MINUTI calcolati con sessioni (solo carrellisti)'
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
                    ore_tim DECIMAL(10,2) COMMENT 'MINUTI presenti in TIM (per anomalie tipo 2)',
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

            # Tabella premi ricevitori
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS premi_ricevitori (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    anno INT NOT NULL,
                    mese INT NOT NULL,
                    codice_preparatore VARCHAR(50) NOT NULL,
                    nome_preparatore VARCHAR(255),
                    giorni_lavorati INT NOT NULL DEFAULT 0,
                    giorni_in_premio DECIMAL(10,2) NOT NULL DEFAULT 0,
                    totale_pallet DECIMAL(12,0) NOT NULL DEFAULT 0,
                    ore_lavorate DECIMAL(10,2) NOT NULL DEFAULT 0,
                    plt_ora_medio DECIMAL(10,2) NOT NULL DEFAULT 0,
                    fascia_raggiunta VARCHAR(50),
                    premio_base DECIMAL(10,2) NOT NULL DEFAULT 0,
                    premio_kpi DECIMAL(10,2) NOT NULL DEFAULT 0,
                    premio_totale DECIMAL(10,2) NOT NULL DEFAULT 0,
                    bonus_applicato BOOLEAN DEFAULT FALSE,
                    data_calcolo TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    note VARCHAR(255),
                    UNIQUE KEY uniq_premio_ricevitori (anno, mese, codice_preparatore),
                    INDEX idx_ricevitori_anno_mese (anno, mese),
                    INDEX idx_ricevitori_codice (codice_preparatore)
                )
                """
            )

            # Migrazione: giorni_in_premio da INT a DECIMAL(10,2)
            try:
                cur.execute(
                    """
                    ALTER TABLE premi_ricevitori
                    MODIFY COLUMN giorni_in_premio DECIMAL(10,2) NOT NULL DEFAULT 0
                    """
                )
            except mysql.connector.Error:
                pass
            
            # Tabella premi doppia spunta
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS premi_doppia_spunta (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    anno INT NOT NULL,
                    mese INT NOT NULL,
                    codice_preparatore VARCHAR(50) NOT NULL,
                    nome_preparatore VARCHAR(255),
                    totale_colli DECIMAL(12,0) NOT NULL DEFAULT 0,
                    colli_nuove_aperture DECIMAL(12,0) NOT NULL DEFAULT 0,
                    ore_lavorate DECIMAL(10,2) NOT NULL DEFAULT 0,
                    colli_ora DECIMAL(10,2) NOT NULL DEFAULT 0,
                    penalita_eccesso DECIMAL(10,2) NOT NULL DEFAULT 0,
                    penalita_difetto DECIMAL(10,2) NOT NULL DEFAULT 0,
                    errori_difetto INT NOT NULL DEFAULT 0,
                    penalita_totale DECIMAL(10,2) NOT NULL DEFAULT 0,
                    fascia_raggiunta VARCHAR(50),
                    premio_base DECIMAL(10,2) NOT NULL DEFAULT 0,
                    premio_kpi DECIMAL(10,2) NOT NULL DEFAULT 0,
                    premio_totale DECIMAL(10,2) NOT NULL DEFAULT 0,
                    bonus_applicato BOOLEAN DEFAULT FALSE,
                    data_calcolo TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    note VARCHAR(255),
                    UNIQUE KEY uniq_premio_doppia (anno, mese, codice_preparatore),
                    INDEX idx_doppia_anno_mese (anno, mese),
                    INDEX idx_doppia_codice (codice_preparatore)
                )
                """
            )

            # Garantisce l'esistenza delle colonne anche su installazioni precedenti
            doppia_columns_alter = [
                "ADD COLUMN colli_nuove_aperture DECIMAL(12,0) NOT NULL DEFAULT 0 AFTER totale_colli",
                "ADD COLUMN ore_lavorate DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER colli_nuove_aperture",
                "ADD COLUMN colli_ora DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER ore_lavorate",
                "ADD COLUMN penalita_eccesso DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER colli_ora",
                "ADD COLUMN penalita_difetto DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER penalita_eccesso",
                "ADD COLUMN errori_difetto INT NOT NULL DEFAULT 0 AFTER penalita_difetto",
                "ADD COLUMN penalita_totale DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER errori_difetto",
                "ADD COLUMN fascia_raggiunta VARCHAR(50) NULL AFTER penalita_totale",
                "ADD COLUMN premio_base DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER fascia_raggiunta",
                "ADD COLUMN premio_kpi DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER premio_base",
                "ADD COLUMN premio_totale DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER premio_kpi",
                "ADD COLUMN bonus_applicato BOOLEAN DEFAULT FALSE AFTER premio_totale",
                "ADD COLUMN note VARCHAR(255) NULL AFTER bonus_applicato",
                "ADD COLUMN data_calcolo TIMESTAMP DEFAULT CURRENT_TIMESTAMP AFTER note",
            ]

            for alter in doppia_columns_alter:
                try:
                    cur.execute(f"ALTER TABLE premi_doppia_spunta {alter}")
                except mysql.connector.Error:
                    pass

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
            
            # Tabella dettaglio sessioni doppia spunta
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sessioni_doppia_spunta (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    data DATE NOT NULL,
                    codice_preparatore VARCHAR(50) NOT NULL COMMENT 'Codice preparatore (prepa)',
                    altro_codice VARCHAR(50) NULL COMMENT 'Altro codice dalla colonna Prep',
                    numero_riga INT NOT NULL COMMENT 'Progressivo riga nella giornata',
                    excel_row INT NULL COMMENT 'Indice riga originale del file (pandas index)',
                    ora_inizio TIME NULL COMMENT 'Ora inizio attività',
                    ora_fine TIME NULL COMMENT 'Ora fine attività',
                    tempo_minuti DECIMAL(10,2) NULL COMMENT 'Durata in minuti',
                    n_colli INT NOT NULL COMMENT 'Numero colli calcolato (qtasped/uvc arrotondato per eccesso)',
                    uvc INT NULL COMMENT 'Valore UVC dalla colonna N',
                    qtasped INT NULL COMMENT 'Valore qtasped dalla colonna M',
                    codpro VARCHAR(50) NULL COMMENT 'Codice prodotto',
                    lotto VARCHAR(50) NULL COMMENT 'Lotto',
                    tipo VARCHAR(100) NULL COMMENT 'Tipo cliente o RS',
                    diff DECIMAL(10,2) NULL COMMENT 'Valore differenza originale',
                    penalita_eccesso INT NOT NULL DEFAULT 0 COMMENT 'Penalità per errori in eccesso',
                    penalita_difetto INT NOT NULL DEFAULT 0 COMMENT 'Penalità per errori in difetto',
                    penalita DECIMAL(10,2) NULL COMMENT 'Penalità assoluta (abs(diff/uvc))',
                    
                    -- Totali giornalieri (solo prima riga)
                    totale_colli_giornata INT NULL COMMENT 'Totale colli nella giornata (solo prima riga)',
                    ore_gestionale_giornaliere DECIMAL(10,2) NULL COMMENT 'Ore totali nella giornata (solo prima riga)',
                    
                    data_importazione TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    INDEX idx_data_codice (data, codice_preparatore),
                    INDEX idx_data (data)
                )
                """
            )
            
            conn.commit()

            # Allinea schema esistente con la colonna attività bonus
            _ensure_malus_bonus_schema(cur)
            _ensure_sessioni_carrellisti_schema(cur)
            _ensure_sessioni_doppia_spunta_schema(cur)


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


def _ensure_sessioni_doppia_spunta_schema(cur: Any) -> None:
    """Garantisce la presenza delle colonne uvc, qtasped, codpro, lotto, excel_row, altro_codice, diff e penalita nella tabella sessioni_doppia_spunta."""
    connection = getattr(cur, "connection", None)
    if connection is None:
        return

    alterations = [
        (
            "uvc",
            """
                ALTER TABLE sessioni_doppia_spunta
                ADD COLUMN uvc INT NULL COMMENT 'Valore UVC dalla colonna N'
                AFTER n_colli
            """
        ),
        (
            "excel_row",
            """
                ALTER TABLE sessioni_doppia_spunta
                ADD COLUMN excel_row INT NULL COMMENT 'Indice riga originale del file (pandas index)'
                AFTER numero_riga
            """
        ),
        (
            "altro_codice",
            """
                ALTER TABLE sessioni_doppia_spunta
                ADD COLUMN altro_codice VARCHAR(50) NULL COMMENT 'Altro codice dalla colonna Prep'
                AFTER codice_preparatore
            """
        ),
        (
            "qtasped",
            """
                ALTER TABLE sessioni_doppia_spunta
                ADD COLUMN qtasped INT NULL COMMENT 'Valore qtasped dalla colonna M'
                AFTER uvc
            """
        ),
        (
            "codpro",
            """
                ALTER TABLE sessioni_doppia_spunta
                ADD COLUMN codpro VARCHAR(50) NULL COMMENT 'Codice prodotto'
                AFTER qtasped
            """
        ),
        (
            "lotto",
            """
                ALTER TABLE sessioni_doppia_spunta
                ADD COLUMN lotto VARCHAR(50) NULL COMMENT 'Lotto'
                AFTER codpro
            """
        ),
        (
            "penalita_eccesso",
            f"""
                ALTER TABLE sessioni_doppia_spunta
                ADD COLUMN penalita_eccesso INT NOT NULL DEFAULT 0 COMMENT 'Penalità per errori in eccesso'
            """
        ),
        (
            "penalita_difetto",
            f"""
                ALTER TABLE sessioni_doppia_spunta
                ADD COLUMN penalita_difetto INT NOT NULL DEFAULT 0 COMMENT 'Penalità per errori in difetto'
            """
        ),
        (
            "penalita",
            f"""
                ALTER TABLE sessioni_doppia_spunta
                ADD COLUMN penalita DECIMAL(10,2) NULL COMMENT 'Penalità assoluta (abs(diff/uvc))'
                AFTER penalita_difetto
            """
        ),
        (
            "diff",
            f"""
                ALTER TABLE sessioni_doppia_spunta
                ADD COLUMN diff DECIMAL(10,2) NULL COMMENT 'Valore differenza originale'
                AFTER tipo
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
                      AND TABLE_NAME = 'sessioni_doppia_spunta'
                      AND COLUMN_NAME = %s
                """,
                (column_name,),
            )
            result = cur.fetchone()

            if result and result[0] == 0:
                cur.execute(alter_sql)
                connection.commit()
                print(f"[OK] Colonna '{column_name}' aggiunta a sessioni_doppia_spunta")
        except Exception as exc:
            print(
                f"[WARN] Errore durante l'aggiornamento schema sessioni_doppia_spunta (colonna {column_name}): {exc}"
            )

    if _column_exists(cur, "sessioni_doppia_spunta", "penalita"):
        try:
            cur.execute(
                """
                UPDATE sessioni_doppia_spunta
                SET penalita_eccesso = GREATEST(penalita, 0),
                    penalita_difetto = GREATEST(-penalita, 0)
                WHERE penalita IS NOT NULL
                """
            )
            connection.commit()
            cur.execute("ALTER TABLE sessioni_doppia_spunta DROP COLUMN penalita")
            connection.commit()
        except Exception as exc:
            print(f"[WARN] Impossibile migrare/droppare colonna penalita in sessioni_doppia_spunta: {exc}")


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
    Inserisce i dati in batch nel database.

    Prima elimina tutti i record esistenti con le stesse combinazioni
    (data, tipo_attivita) presenti nei dati in arrivo, poi inserisce
    i nuovi record.  In questo modo un file con meno righe rispetto
    al caricamento precedente non lascia record "fantasma".

    Args:
        values: Lista di tuple con i dati da inserire
                (data, codice_preparatore, nome_preparatore, totale_colli,
                 penalita_eccesso, penalita_difetto, penalita_totale,
                 tipo_attivita, tipo, ore_tim, ore_gestionale)

    Returns:
        Numero di record inseriti
    """
    if not values:
        return 0

    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            try:
                # ── Salva ore_tim esistenti prima della cancellazione ──
                keys_to_delete = {(v[0], v[7]) for v in values}  # (data, tipo_attivita)
                saved_ore_tim: Dict[Tuple, float] = {}
                for data_val, tipo_att in keys_to_delete:
                    cur.execute(
                        f"SELECT data, codice_preparatore, tipo_attivita, ore_tim "
                        f"FROM {TABLE_NAME} "
                        f"WHERE data = %s AND tipo_attivita = %s AND ore_tim > 0",
                        (data_val, tipo_att),
                    )
                    for row in cur.fetchall():
                        key = (row[0], row[1], row[2])  # data, codice, tipo_att
                        saved_ore_tim[key] = float(row[3])

                if saved_ore_tim:
                    print(f"[OK] Salvati {len(saved_ore_tim)} valori ore_tim esistenti da preservare")

                # ── Elimina i vecchi record per le stesse (data, tipo_attivita) ──
                for data_val, tipo_att in keys_to_delete:
                    cur.execute(
                        f"DELETE FROM {TABLE_NAME} "
                        f"WHERE data = %s AND tipo_attivita = %s",
                        (data_val, tipo_att),
                    )
                deleted = cur.rowcount
                print(f"[OK] Eliminati {deleted} record vecchi per {len(keys_to_delete)} combinazioni (data, tipo_attivita)")

                # ── Inserisce i nuovi record ──
                sql = f"""
                    INSERT INTO {TABLE_NAME}
                    (data, codice_preparatore, nome_preparatore, totale_colli,
                     penalita_eccesso, penalita_difetto, penalita_totale,
                     tipo_attivita, tipo, ore_tim, ore_gestionale)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """
                cur.executemany(sql, values)

                # ── Ripristina ore_tim dove era già presente ──
                if saved_ore_tim:
                    restored = 0
                    for (data_val, codice, tipo_att), ore in saved_ore_tim.items():
                        cur.execute(
                            f"UPDATE {TABLE_NAME} SET ore_tim = %s "
                            f"WHERE data = %s AND codice_preparatore = %s "
                            f"AND tipo_attivita = %s AND ore_tim = 0",
                            (ore, data_val, codice, tipo_att),
                        )
                        restored += cur.rowcount
                    print(f"[OK] Ripristinati {restored}/{len(saved_ore_tim)} valori ore_tim")

                conn.commit()
                return len(values)

            except Exception as e:
                conn.rollback()
                print(f"\n[ERROR] Errore durante insert_batch_data: {e}")
                raise


def update_penalita_picking(values: List[Tuple[datetime.date, str, float]]) -> int:
    """Aggiorna la penalità totale delle attività PICKING per data e codice preparatore.

    La data di rilevamento (dalla doppia spunta) potrebbe non coincidere con una
    data in cui il picker ha lavorato. Quando l'UPDATE per la data esatta non trova
    righe, la penalità viene sommata e assegnata alla prima data disponibile del
    mese per quel codice picking.

    Nota: i campi legacy penalita_eccesso/penalita_difetto vengono azzerati.
    """

    if not values:
        return 0

    # Determina il mese di riferimento dalle date
    all_dates = [d for (d, _, _) in values]
    anno = all_dates[0].year
    mese = all_dates[0].month
    primo_giorno = datetime.date(anno, mese, 1)
    if mese == 12:
        ultimo_giorno = datetime.date(anno + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        ultimo_giorno = datetime.date(anno, mese + 1, 1) - datetime.timedelta(days=1)

    # UPDATE penalità PICKING per data e codice preparatore
    sql_update = f"""
        UPDATE {TABLE_NAME}
        SET penalita_totale = penalita_totale + %s,
            penalita_eccesso = 0,
            penalita_difetto = 0
        WHERE data = %s
          AND codice_preparatore = %s
          AND tipo_attivita = 'PICKING'
    """

    # Trova la prima data PICKING valida nel mese per un codice
    # Preferisce date con ore_tim > 0 (senza anomalie PRODUZIONE_SENZA_ORE)
    sql_find_date = f"""
        SELECT data AS first_date
        FROM {TABLE_NAME}
        WHERE codice_preparatore = %s
          AND tipo_attivita = 'PICKING'
          AND data BETWEEN %s AND %s
        ORDER BY (ore_tim > 0) DESC, data ASC
        LIMIT 1
    """

    # Azzera TUTTE le penalità PICKING del mese
    reset_sql = (
        f"UPDATE {TABLE_NAME} "
        "SET penalita_totale = 0, penalita_eccesso = 0, penalita_difetto = 0 "
        "WHERE tipo_attivita = 'PICKING' AND data BETWEEN %s AND %s"
    )

    with closing(_get_conn()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            # Reset tutte le penalità PICKING del mese
            cur.execute(reset_sql, (primo_giorno, ultimo_giorno))

            updated_total = 0
            # Penalità non assegnate: (codice -> penalità accumulata)
            unassigned: Dict[str, float] = {}

            for data, codice, penalita_totale in values:
                pen = float(penalita_totale or 0)
                if pen == 0:
                    continue

                # Prova ad assegnare alla data esatta
                cur.execute(sql_update, (pen, data, codice))
                if cur.rowcount > 0:
                    updated_total += cur.rowcount
                else:
                    # Accumula per assegnazione successiva
                    unassigned[codice] = unassigned.get(codice, 0) + pen

            # Assegna le penalità rimaste alla prima data disponibile nel mese
            for codice, pen in unassigned.items():
                cur.execute(sql_find_date, (codice, primo_giorno, ultimo_giorno))
                row = cur.fetchone()
                first_date = row.get("first_date") if row else None
                if first_date:
                    cur.execute(sql_update, (pen, first_date, codice))
                    if cur.rowcount > 0:
                        updated_total += cur.rowcount
                        print(f"[INFO] Penalita PICKING {codice}: {pen} assegnata a {first_date} (data rilevamento non disponibile)")
                else:
                    print(f"[WARN] Penalita PICKING {codice}: {pen} NON assegnata (nessun record PICKING nel mese)")

            conn.commit()
    return updated_total


def get_penalita_picking_from_sessioni(dates: List[datetime.date]) -> List[Tuple[datetime.date, str, float]]:
    """Ritorna le penalità totali da applicare ai record PICKING.

    Regola: il codice PICKING è sempre il primo codice a sinistra del trattino in `altro_codice`.
    """

    if not dates:
        return []

    placeholders = ",".join(["%s"] * len(dates))
    sql = f"""
        SELECT
            data,
            UPPER(TRIM(SUBSTRING_INDEX(altro_codice, '-', 1))) AS codice_picking,
            ROUND(SUM(COALESCE(penalita, 0)), 2) AS penalita_totale
        FROM sessioni_doppia_spunta
        WHERE data IN ({placeholders})
                    AND penalita <> 0
                    AND qtasped IS NOT NULL
                    AND qtasped <> 0
          AND altro_codice IS NOT NULL
          AND TRIM(altro_codice) <> ''
        GROUP BY data, codice_picking
    """

    with closing(_get_conn()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            cur.execute(sql, dates)
            rows = cast(List[Dict[str, Any]], cur.fetchall())

    out: List[Tuple[datetime.date, str, float]] = []
    for r in rows:
        codice = (r.get("codice_picking") or "").strip().upper()
        if not codice:
            continue
        out.append((r["data"], codice, float(r.get("penalita_totale") or 0)))
    return out


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
    
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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


def fetch_all_nuove_aperture() -> List[Dict[str, Any]]:
    """Restituisce tutti i periodi registrati di nuove aperture."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            cur.execute(
                """
                SELECT data_da, data_a, negozio
                FROM nuove_aperture
                ORDER BY negozio, data_da
                """
            )
            return cast(List[Dict[str, Any]], cur.fetchall())


def fetch_fasce_premi(tipo: Optional[str] = None) -> List[Dict[str, Any]]:
    """Recupera le fasce premio, opzionalmente filtrate per tipo attività."""
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute("DELETE FROM fasce_premi WHERE id = %s", (fascia_id,))
            conn.commit()


def delete_peso_movimento(peso_id: int) -> None:
    """Elimina un peso movimento."""
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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
    """Inserisce una nuova anomalia. Se esiste già una ELIMINATA per la stessa combinazione, la salta."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            # Estrae anno e mese dalla data di rilevamento
            anno = data_rilevamento.year
            mese = data_rilevamento.month

            # Se esiste già un record ELIMINATA per la stessa combinazione, non re-inserire
            cur.execute(
                """
                SELECT id FROM anomalie
                WHERE tipo_anomalia = %s
                  AND data_rilevamento = %s
                  AND codice_preparatore = %s
                  AND COALESCE(tipo_attivita, '') = COALESCE(%s, '')
                  AND stato = 'ELIMINATA'
                """,
                (tipo_anomalia, data_rilevamento, codice_preparatore, tipo_attivita),
            )
            if cur.fetchone():
                return 0  # Record già eliminato dall'utente, non rigenerare
            
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
    with closing(_get_conn()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            conditions: List[str] = ["COALESCE(a.stato, 'APERTA') != 'ELIMINATA'"]
            params: List[Any] = []

            if tipo_anomalia:
                if isinstance(tipo_anomalia, list):
                    placeholders = ", ".join(["%s"] * len(tipo_anomalia))
                    conditions.append(f"a.tipo_anomalia IN ({placeholders})")
                    params.extend(tipo_anomalia)
                else:
                    conditions.append("a.tipo_anomalia = %s")
                    params.append(tipo_anomalia)
            if stato:
                conditions.append("a.stato = %s")
                params.append(stato)
            if tipo_attivita:
                conditions.append("a.tipo_attivita = %s")
                params.append(tipo_attivita)
            if codice_preparatore:
                conditions.append("a.codice_preparatore = %s")
                params.append(codice_preparatore)
            if data_da:
                conditions.append("a.data_rilevamento >= %s")
                params.append(data_da)
            if data_a:
                conditions.append("a.data_rilevamento <= %s")
                params.append(data_a)

            where_clause = " WHERE " + " AND ".join(conditions) if conditions else ""
            query = f"""
                SELECT a.id, a.tipo_anomalia, a.data_rilevamento, a.anno, a.mese,
                       a.codice_preparatore, a.nome_preparatore, a.tipo_attivita,
                       a.ore_tim, a.dettagli, a.stato, a.note,
                       a.data_creazione, a.data_aggiornamento,
                       COALESCE(SUM(dp.totale_colli), 0) AS totale_colli
                FROM anomalie a
                LEFT JOIN dati_produzione dp ON
                    dp.data = a.data_rilevamento
                    AND UPPER(dp.codice_preparatore) = UPPER(a.codice_preparatore)
                    AND dp.tipo_attivita = a.tipo_attivita
                {where_clause}
                GROUP BY a.id, a.tipo_anomalia, a.data_rilevamento, a.anno, a.mese,
                         a.codice_preparatore, a.nome_preparatore, a.tipo_attivita,
                         a.ore_tim, a.dettagli, a.stato, a.note,
                         a.data_creazione, a.data_aggiornamento
                ORDER BY a.data_rilevamento ASC, a.anno ASC, a.mese ASC, a.data_creazione ASC
            """
            cur.execute(query, params)
            result = cur.fetchall()
            return cast(List[Dict[str, Any]], result)


def fetch_anomalia_by_id(anomalia_id: int) -> Optional[Dict[str, Any]]:
    """Recupera una singola anomalia per ID."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            cur.execute(
                """
                SELECT id, tipo_anomalia, data_rilevamento, anno, mese, codice_preparatore, nome_preparatore,
                       tipo_attivita, ore_tim, dettagli, stato, note, data_creazione, data_aggiornamento
                FROM anomalie
                WHERE id = %s
                """,
                (anomalia_id,),
            )
            row = cur.fetchone()
            return cast(Optional[Dict[str, Any]], row)


def fetch_attivita_tim_range(
    data_da: datetime.date,
    data_a: datetime.date,
    search: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Recupera le attività TIM per un intervallo date, con ricerca opzionale per codice/nome."""
    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            conditions = ["a.data_riferimento BETWEEN %s AND %s"]
            params: List[Any] = [data_da, data_a]

            if search:
                like = f"%{search}%"
                conditions.append(
                    "(LOWER(cg.codice) LIKE LOWER(%s)"
                    " OR LOWER(CONCAT(u.cognome, ' ', u.nome)) LIKE LOWER(%s))"
                )
                params.extend([like, like])

            where_clause = " AND ".join(conditions)

            query = f"""
                SELECT
                    cg.codice AS codice_preparatore,
                    CONCAT(u.cognome, ' ', u.nome) AS nome_preparatore,
                    a.data_riferimento,
                    ta.descrizione AS attivita,
                    a.data_inizio,
                    a.data_fine,
                    COALESCE((SELECT SUM(p.durata) FROM pausa p WHERE p.attivita_id = a.id), 0) AS pausa,
                    GREATEST(
                        TIMESTAMPDIFF(MINUTE, a.data_inizio, a.data_fine)
                          - COALESCE((SELECT SUM(p.durata) FROM pausa p WHERE p.attivita_id = a.id), 0),
                        0
                    ) AS durata_minuti
                FROM attivita a
                JOIN tipoattivita ta ON a.tipo_attivita_id = ta.id
                JOIN utente u ON a.utente_id = u.id
                JOIN codicegestionale cg ON cg.utente_id = u.id
                    AND a.data_riferimento BETWEEN cg.valido_dal AND COALESCE(cg.valido_al, '9999-12-31')
                WHERE {where_clause}
                ORDER BY a.data_riferimento, cg.codice, a.data_inizio
            """
            cur.execute(query, params)
            rows = cur.fetchall() or []
            return cast(List[Dict[str, Any]], rows)


def fetch_attivita_tim(
    codice_preparatore: str,
    data_riferimento: datetime.date,
) -> List[Dict[str, Any]]:
    """Recupera le attività TIM con inizio/fine e durata (minuti) per codice e data."""
    if not codice_preparatore:
        return []

    codice = codice_preparatore.strip().lower()
    if not codice:
        return []

    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            query = """
                SELECT
                    a.id AS attivita_id,
                    a.utente_id,
                    ta.id AS tipo_attivita_id,
                    ta.descrizione AS attivita,
                    a.data_inizio,
                    a.data_fine,
                    a.data_riferimento,
                    COALESCE((SELECT SUM(p.durata) FROM pausa p WHERE p.attivita_id = a.id), 0) AS pausa,
                    GREATEST(
                        TIMESTAMPDIFF(MINUTE, a.data_inizio, a.data_fine)
                          - COALESCE((SELECT SUM(p.durata) FROM pausa p WHERE p.attivita_id = a.id), 0),
                        0
                    ) AS durata_minuti
                FROM attivita a
                JOIN tipoattivita ta ON a.tipo_attivita_id = ta.id
                JOIN utente u ON a.utente_id = u.id
                JOIN codicegestionale cg ON cg.utente_id = u.id
                WHERE LOWER(cg.codice) = %s
                  AND a.data_riferimento = %s
                  AND %s BETWEEN cg.valido_dal AND COALESCE(cg.valido_al, '9999-12-31')
                ORDER BY a.data_inizio
            """
            cur.execute(query, (codice, data_riferimento, data_riferimento))
            rows = cur.fetchall() or []
            return cast(List[Dict[str, Any]], rows)


def fetch_tipi_attivita_tim() -> List[Dict[str, Any]]:
    """Recupera l'elenco dei tipi attività da TIM."""
    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            cur.execute(
                """
                SELECT id, descrizione
                FROM tipoattivita
                ORDER BY descrizione
                """
            )
            rows = cur.fetchall() or []
            return cast(List[Dict[str, Any]], rows)


def fetch_pause_by_attivita_ids(
    attivita_ids: List[int],
) -> Dict[int, List[Dict[str, Any]]]:
    """Recupera le pause raggruppate per attivita_id."""
    if not attivita_ids:
        return {}

    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            placeholders = ", ".join(["%s"] * len(attivita_ids))
            cur.execute(
                f"""
                SELECT
                    p.id AS pausa_id,
                    p.attivita_id,
                    p.descrizione,
                    p.data_inizio,
                    p.data_fine,
                    p.durata,
                    p.pagata
                FROM pausa p
                WHERE p.attivita_id IN ({placeholders})
                ORDER BY p.attivita_id, p.data_inizio
                """,
                tuple(attivita_ids),
            )
            rows = cur.fetchall() or []
            result: Dict[int, List[Dict[str, Any]]] = {}
            for row in rows:
                aid = int(row["attivita_id"])
                if aid not in result:
                    result[aid] = []
                result[aid].append(row)
            return result


def update_attivita_tim(attivita_id: int, tipo_attivita_id: int) -> None:
    """Aggiorna il tipo attività di una singola attività su TIM."""
    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                """
                UPDATE attivita
                SET tipo_attivita_id = %s
                WHERE id = %s
                """,
                (tipo_attivita_id, attivita_id),
            )
            conn.commit()


def split_attivita_tim(
    original_id: int,
    part1_fine: datetime.datetime,
    part1_tipo_attivita_id: Optional[int] = None,
) -> int:
    """Spezza un'attività TIM in due parti.

    La riga originale (Parte 1) viene aggiornata con data_fine = part1_fine.
    Se part1_tipo_attivita_id è specificato, la Parte 1 cambia tipo attività.
    La Parte 2 (nuovo record) mantiene il tipo attività originale.
    La durata viene ricalcolata per entrambe le parti.
    Le pause vengono riassegnate al record corretto in base agli orari.

    Ritorna l'id della nuova riga inserita.
    """
    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            try:
                cur.execute(
                    """SELECT * FROM attivita WHERE id = %s""",
                    (original_id,),
                )
                original = cur.fetchone()
                if not original:
                    raise ValueError(f"Attività con id={original_id} non trovata.")

                orig_inizio = original["data_inizio"]
                orig_fine = original["data_fine"]

                if part1_fine <= orig_inizio:
                    raise ValueError(
                        "La fine della Parte 1 deve essere successiva "
                        "all'inizio dell'attività originale."
                    )
                if part1_fine >= orig_fine:
                    raise ValueError(
                        "La fine della Parte 1 deve essere precedente "
                        "alla fine dell'attività originale."
                    )

                # ── Recupera le pause legate all'attività originale ──
                cur.execute(
                    """SELECT * FROM pausa WHERE attivita_id = %s
                       ORDER BY data_inizio""",
                    (original_id,),
                )
                pause_rows = cur.fetchall() or []

                # Suddividi le pause: quelle che iniziano PRIMA di part1_fine
                # restano sulla Parte 1, le altre vanno sulla Parte 2
                pause_part1 = []
                pause_part2 = []
                for p in pause_rows:
                    p_inizio = p.get("data_inizio")
                    if p_inizio and p_inizio >= part1_fine:
                        pause_part2.append(p)
                    else:
                        pause_part1.append(p)

                # ── Calcola pausa e durata per ciascuna parte ──
                # Solo le pause NON pagate (es. pranzo) vengono sottratte
                # dalla durata. Le pause pagate non incidono sul calcolo.
                pausa_min_part1 = sum(
                    int(p.get("durata") or 0)
                    for p in pause_part1
                    if not p.get("pagata")
                )
                pausa_min_part2 = sum(
                    int(p.get("durata") or 0)
                    for p in pause_part2
                    if not p.get("pagata")
                )

                durata_part1 = max(
                    int((part1_fine - orig_inizio).total_seconds() // 60) - pausa_min_part1,
                    0,
                )
                durata_part2 = max(
                    int((orig_fine - part1_fine).total_seconds() // 60) - pausa_min_part2,
                    0,
                )

                # ── Determina tipo attività per Parte 1 (se cambiato) ──
                p1_tipo_id = part1_tipo_attivita_id or original.get("tipo_attivita_id")
                p1_descrizione = original.get("descrizione")
                if part1_tipo_attivita_id and part1_tipo_attivita_id != original.get("tipo_attivita_id"):
                    cur.execute(
                        "SELECT descrizione FROM tipoattivita WHERE id = %s",
                        (part1_tipo_attivita_id,),
                    )
                    tipo_row = cur.fetchone()
                    if tipo_row:
                        p1_descrizione = tipo_row.get("descrizione")

                # ── UPDATE record originale (Parte 1) ──
                cur.execute(
                    """
                    UPDATE attivita
                    SET data_fine = %s,
                        durata = %s,
                        pausa = %s,
                        tipo_attivita_id = %s,
                        descrizione = %s,
                        chiusura_giornata = 0
                    WHERE id = %s
                    """,
                    (part1_fine, durata_part1, pausa_min_part1,
                     p1_tipo_id, p1_descrizione, original_id),
                )

                # ── INSERT nuovo record (Parte 2) copiando tutti i campi ──
                new_id = _generate_tim_long_id("001")
                cur.execute(
                    """
                    INSERT INTO attivita
                        (id, utente_id, buonoattivitaextra_id,
                         location_id, descrizione, data_riferimento,
                         data_inizio, data_fine,
                         codice_terminale, terminale_id,
                         squadra_id, reparto_id, tipo_attivita_id,
                         tag_produzione_id, tag_valore, tag_note,
                         tag_aggiornamento, tag_ultima,
                         mezzo_id, device_id, codice_gestionale_id,
                         durata, pausa, note, stato_id,
                         apertura_giornata, chiusura_giornata,
                         origine, attivita_id_provenienza,
                         utente_modifica)
                    VALUES (%s, %s, %s,
                            %s, %s, %s,
                            %s, %s,
                            %s, %s,
                            %s, %s, %s,
                            %s, %s, %s,
                            %s, %s,
                            %s, %s, %s,
                            %s, %s, %s, %s,
                            %s, %s,
                            %s, %s,
                            %s)
                    """,
                    (
                        new_id,
                        original["utente_id"],
                        original.get("buonoattivitaextra_id"),
                        original.get("location_id"),
                        original.get("descrizione"),
                        original.get("data_riferimento"),
                        part1_fine,                       # data_inizio = fine parte 1
                        orig_fine,                        # data_fine   = fine originale
                        original.get("codice_terminale"),
                        original.get("terminale_id"),
                        original.get("squadra_id"),
                        original.get("reparto_id"),
                        original.get("tipo_attivita_id"),
                        original.get("tag_produzione_id"),
                        original.get("tag_valore"),
                        original.get("tag_note"),
                        original.get("tag_aggiornamento"),
                        original.get("tag_ultima"),
                        original.get("mezzo_id"),
                        original.get("device_id"),
                        original.get("codice_gestionale_id"),
                        durata_part2,
                        pausa_min_part2,
                        original.get("note"),
                        original.get("stato_id"),
                        0,                                # apertura_giornata = 0
                        original.get("chiusura_giornata"),
                        original.get("origine"),
                        original_id,                      # attivita_id_provenienza
                        original.get("utente_modifica"),
                    ),
                )

                # ── Sposta le pause della Parte 2 sul nuovo record ──
                if pause_part2:
                    pause_ids = [p["id"] for p in pause_part2]
                    placeholders = ", ".join(["%s"] * len(pause_ids))
                    cur.execute(
                        f"""
                        UPDATE pausa
                        SET attivita_id = %s
                        WHERE id IN ({placeholders})
                        """,
                        [new_id] + pause_ids,
                    )

                conn.commit()
                return new_id

            except ValueError:
                conn.rollback()
                raise
            except Exception:
                conn.rollback()
                raise


def ripristina_attivita_tim(selected_id: int) -> None:
    """Ripristina un'attività spezzata riunendo le due parti.

    Accetta l'id di una qualsiasi delle due parti (originale o derivata).
    Trova la coppia tramite attivita_id_provenienza, riporta il record
    originale (Parte 1) alla data_fine della Parte 2, sposta le pause
    della Parte 2 sul record originale, ricalcola durata e pausa,
    e infine elimina il record derivato (Parte 2).
    """
    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            try:
                # ── Trova la coppia originale + derivato ──
                # Caso 1: l'utente ha selezionato il record derivato (Parte 2)
                cur.execute(
                    "SELECT * FROM attivita WHERE id = %s",
                    (selected_id,),
                )
                selected = cur.fetchone()
                if not selected:
                    raise ValueError(f"Attività con id={selected_id} non trovata.")

                provenienza = selected.get("attivita_id_provenienza")

                if provenienza:
                    # Il record selezionato È la Parte 2
                    derivato = selected
                    cur.execute(
                        "SELECT * FROM attivita WHERE id = %s",
                        (provenienza,),
                    )
                    originale = cur.fetchone()
                    if not originale:
                        raise ValueError(
                            f"Record originale (id={provenienza}) non trovato."
                        )
                else:
                    # Caso 2: l'utente ha selezionato il record originale (Parte 1)
                    originale = selected
                    cur.execute(
                        "SELECT * FROM attivita WHERE attivita_id_provenienza = %s "
                        "ORDER BY data_inizio LIMIT 1",
                        (selected_id,),
                    )
                    derivato = cur.fetchone()
                    if not derivato:
                        raise ValueError(
                            "Nessun record derivato trovato. "
                            "Questa attività non è stata spezzata."
                        )

                orig_id = originale["id"]
                deriv_id = derivato["id"]

                # ── Sposta tutte le pause del derivato sull'originale ──
                cur.execute(
                    "UPDATE pausa SET attivita_id = %s WHERE attivita_id = %s",
                    (orig_id, deriv_id),
                )

                # ── Ricalcola pausa (solo non pagate) e durata ──
                cur.execute(
                    "SELECT durata, pagata FROM pausa WHERE attivita_id = %s",
                    (orig_id,),
                )
                pause_rows = cur.fetchall() or []
                pausa_totale = sum(
                    int(p.get("durata") or 0)
                    for p in pause_rows
                    if not p.get("pagata")
                )

                orig_inizio = originale["data_inizio"]
                deriv_fine = derivato["data_fine"]
                durata_totale = max(
                    int((deriv_fine - orig_inizio).total_seconds() // 60) - pausa_totale,
                    0,
                )

                # ── Ripristina il record originale ──
                cur.execute(
                    """
                    UPDATE attivita
                    SET data_fine = %s,
                        durata = %s,
                        pausa = %s,
                        tipo_attivita_id = %s,
                        descrizione = %s,
                        chiusura_giornata = %s
                    WHERE id = %s
                    """,
                    (
                        deriv_fine,
                        durata_totale,
                        pausa_totale,
                        derivato.get("tipo_attivita_id"),
                        derivato.get("descrizione"),
                        derivato.get("chiusura_giornata"),
                        orig_id,
                    ),
                )

                # ── Elimina il record derivato ──
                cur.execute(
                    "DELETE FROM attivita WHERE id = %s",
                    (deriv_id,),
                )

                conn.commit()

            except ValueError:
                conn.rollback()
                raise
            except Exception:
                conn.rollback()
                raise


def fetch_report_templates(
    attivi_solo: bool = True,
    attivita: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Restituisce la lista dei report configurati per gli export."""
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            cur.execute(query, tuple(params))
            rows = cur.fetchall() or []
            columns = [col[0] for col in cur.description] if cur.description else []
            return cast(List[Dict[str, Any]], rows), columns


def update_anomalia_stato(anomalia_id: int, nuovo_stato: str, note: Optional[str] = None) -> None:
    """Aggiorna lo stato di un'anomalia."""
    with closing(_get_conn()) as conn:
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


def resolve_anomalie_codice_non_abbinato(codice_preparatore: str, note: Optional[str] = None) -> None:
    """Imposta RISOLTA per tutte le anomalie CODICE_NON_ABBINATO di un codice."""
    codice = (codice_preparatore or "").strip().upper()
    if not codice:
        return
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            if note:
                cur.execute(
                    """
                    UPDATE anomalie
                    SET stato = 'RISOLTA', note = %s
                    WHERE tipo_anomalia = 'CODICE_NON_ABBINATO'
                      AND codice_preparatore = %s
                    """,
                    (note, codice),
                )
            else:
                cur.execute(
                    """
                    UPDATE anomalie
                    SET stato = 'RISOLTA'
                    WHERE tipo_anomalia = 'CODICE_NON_ABBINATO'
                      AND codice_preparatore = %s
                    """,
                    (codice,),
                )
            conn.commit()
            conn.commit()


# ========== GESTIONE PREMI CARRELLISTI ==========

def save_premi_carrellisti(anno: int, mese: int, premi: List[Dict[str, Any]]) -> None:
    """Salva i premi carrellisti per un dato mese. Se esistono già, li sovrascrive."""
    with closing(_get_conn()) as conn:
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
    codice_preparatore: Optional[str] = None,
    search: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Recupera i premi carrellisti filtrati per anno/mese/codice o ricerca libera."""
    with closing(_get_conn()) as conn:
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
            if search:
                like = f"%{search}%"
                conditions.append(
                    "(LOWER(codice_preparatore) LIKE LOWER(%s)"
                    " OR LOWER(nome_preparatore) LIKE LOWER(%s))"
                )
                params.extend([like, like])
            
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
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                "DELETE FROM premi_carrellisti WHERE anno = %s AND mese = %s",
                (anno, mese)
            )
            conn.commit()


def save_premi_ricevitori(anno: int, mese: int, premi: List[Dict[str, Any]]) -> None:
    """Salva i premi ricevimento per il mese selezionato sovrascrivendo quelli esistenti."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                "DELETE FROM premi_ricevitori WHERE anno = %s AND mese = %s",
                (anno, mese),
            )

            insert_sql = """
                INSERT INTO premi_ricevitori (
                    anno, mese, codice_preparatore, nome_preparatore,
                    giorni_lavorati, giorni_in_premio, totale_pallet, ore_lavorate,
                    plt_ora_medio, fascia_raggiunta, premio_base, premio_kpi,
                    premio_totale, bonus_applicato, note
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """

            for premio in premi:
                cur.execute(
                    insert_sql,
                    (
                        anno,
                        mese,
                        premio.get("codice"),
                        premio.get("nome"),
                        premio.get("giorni_lavorati"),
                        premio.get("giorni_in_premio"),
                        premio.get("tot_pallet"),
                        premio.get("ore"),
                        premio.get("plt_ora"),
                        premio.get("fascia"),
                        premio.get("premio_base"),
                        premio.get("premio_kpi"),
                        premio.get("premio_totale"),
                        premio.get("bonus_applicato", False),
                        premio.get("note"),
                    ),
                )

            conn.commit()


def fetch_premi_ricevitori(
    anno: Optional[int] = None,
    mese: Optional[int] = None,
    codice_preparatore: Optional[str] = None,
    search: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Recupera i premi ricevimento filtrati per anno/mese/codice o ricerca libera."""
    with closing(_get_conn()) as conn:
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
            if search:
                like = f"%{search}%"
                conditions.append(
                    "(LOWER(codice_preparatore) LIKE LOWER(%s)"
                    " OR LOWER(nome_preparatore) LIKE LOWER(%s))"
                )
                params.extend([like, like])

            where_clause = " WHERE " + " AND ".join(conditions) if conditions else ""

            query = f"""
                SELECT
                    id, anno, mese, codice_preparatore, nome_preparatore,
                    giorni_lavorati, giorni_in_premio, totale_pallet, ore_lavorate,
                    plt_ora_medio, fascia_raggiunta, premio_base, premio_kpi,
                    premio_totale, bonus_applicato, data_calcolo, note
                FROM premi_ricevitori
                {where_clause}
                ORDER BY premio_totale DESC
            """

            cur.execute(query, params)
            result = cur.fetchall()
            return cast(List[Dict[str, Any]], result)


def delete_premi_ricevitori(anno: int, mese: int) -> None:
    """Elimina i premi ricevimento per un mese specifico."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                "DELETE FROM premi_ricevitori WHERE anno = %s AND mese = %s",
                (anno, mese),
            )
            conn.commit()


def save_premi_preparatori(anno: int, mese: int, premi: List[Dict[str, Any]]) -> None:
    """Salva i premi preparatori per un mese specifico sovrascrivendo quelli esistenti."""
    with closing(_get_conn()) as conn:
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
    search: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Restituisce i premi preparatori filtrati per anno/mese/codice o ricerca libera."""
    with closing(_get_conn()) as conn:
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
            if search:
                like = f"%{search}%"
                conditions.append(
                    "(LOWER(codice_preparatore) LIKE LOWER(%s)"
                    " OR LOWER(nome_preparatore) LIKE LOWER(%s))"
                )
                params.extend([like, like])

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
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                "DELETE FROM premi_preparatori WHERE anno = %s AND mese = %s",
                (anno, mese),
            )
            conn.commit()


def save_premi_doppia_spunta(anno: int, mese: int, premi: List[Dict[str, Any]]) -> None:
    """Salva i premi Doppia Spunta per un dato mese sovrascrivendo gli esistenti."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                "DELETE FROM premi_doppia_spunta WHERE anno = %s AND mese = %s",
                (anno, mese),
            )

            for premio in premi:
                cur.execute(
                    """
                    INSERT INTO premi_doppia_spunta (
                        anno, mese, codice_preparatore, nome_preparatore,
                        totale_colli, colli_nuove_aperture, ore_lavorate, colli_ora,
                        penalita_eccesso, penalita_difetto, errori_difetto, penalita_totale,
                        fascia_raggiunta, premio_base, premio_kpi, premio_totale,
                        bonus_applicato, note
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        anno,
                        mese,
                        premio.get("codice"),
                        premio.get("nome"),
                        premio.get("tot_colli"),
                        premio.get("colli_nuove_aperture", 0),
                        premio.get("ore"),
                        premio.get("colli_ora"),
                        premio.get("penalita_eccesso", premio.get("penalita", 0)),
                        premio.get("penalita_difetto", 0),
                        premio.get("errori_difetto", 0),
                        premio.get("penalita"),
                        premio.get("fascia"),
                        premio.get("premio_base"),
                        premio.get("premio_kpi"),
                        premio.get("premio_totale"),
                        premio.get("bonus_applicato", False),
                        premio.get("note"),
                    ),
                )
            conn.commit()


def fetch_premi_doppia_spunta(
    anno: Optional[int] = None,
    mese: Optional[int] = None,
    codice_preparatore: Optional[str] = None,
    search: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Recupera i premi Doppia Spunta filtrati per anno/mese/codice o ricerca libera."""
    with closing(_get_conn()) as conn:
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
            if search:
                like = f"%{search}%"
                conditions.append(
                    "(LOWER(codice_preparatore) LIKE LOWER(%s)"
                    " OR LOWER(nome_preparatore) LIKE LOWER(%s))"
                )
                params.extend([like, like])

            where_clause = " WHERE " + " AND ".join(conditions) if conditions else ""

            query = f"""
                SELECT
                    id, anno, mese, codice_preparatore, nome_preparatore,
                    totale_colli, colli_nuove_aperture, ore_lavorate, colli_ora,
                    penalita_eccesso, penalita_difetto, errori_difetto, penalita_totale,
                    fascia_raggiunta, premio_base, premio_kpi, premio_totale,
                    bonus_applicato, data_calcolo, note
                FROM premi_doppia_spunta
                {where_clause}
                ORDER BY premio_totale DESC
            """
            cur.execute(query, params)
            result = cur.fetchall()
            return cast(List[Dict[str, Any]], result)


def delete_premi_doppia_spunta(anno: int, mese: int) -> None:
    """Elimina i premi Doppia Spunta per un determinato mese."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute(
                "DELETE FROM premi_doppia_spunta WHERE anno = %s AND mese = %s",
                (anno, mese),
            )
            conn.commit()


def delete_anomalia(anomalia_id: int) -> None:
    """Soft-delete un'anomalia (imposta stato = 'ELIMINATA')."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            cur.execute("UPDATE anomalie SET stato = 'ELIMINATA' WHERE id = %s", (anomalia_id,))
            conn.commit()


def clear_anomalie_by_date(data_rilevamento: datetime.date, tipo_anomalia: Optional[str] = None) -> int:
    """Elimina anomalie per una data specifica (utile prima di rigenerare). Preserva le ELIMINATA."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            if tipo_anomalia:
                cur.execute(
                    "DELETE FROM anomalie WHERE data_rilevamento = %s AND tipo_anomalia = %s AND COALESCE(stato, 'APERTA') != 'ELIMINATA'",
                    (data_rilevamento, tipo_anomalia),
                )
            else:
                cur.execute(
                    "DELETE FROM anomalie WHERE data_rilevamento = %s AND COALESCE(stato, 'APERTA') != 'ELIMINATA'",
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
        with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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
    with closing(_get_conn()) as conn:
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


def save_sessioni_doppia_spunta(sessioni: List[Dict[str, Any]]) -> None:
    """Salva i dettagli delle sessioni doppia spunta nel database."""
    if not sessioni:
        print("[WARN] save_sessioni_doppia_spunta: nessuna sessione da salvare")
        return
    
    print(f"[INFO] save_sessioni_doppia_spunta: ricevute {len(sessioni)} righe")
    try:
        unique_codes = {s.get('codice_preparatore') for s in sessioni}
        sample = sessioni[:2]
        print(f"[DEBUG] sessioni_doppia_spunta codici={len(unique_codes)} sample={sample}")
    except Exception:
        pass
    
    try:
        with closing(_get_conn()) as conn:
            conn.autocommit = False
            
            with closing(conn.cursor()) as cur:
                # Garantisce che la tabella abbia sempre le nuove colonne prima di scrivere
                try:
                    _ensure_sessioni_doppia_spunta_schema(cur)
                except Exception as schema_exc:
                    print(f"[WARN] Impossibile aggiornare schema sessioni_doppia_spunta prima dell'inserimento: {schema_exc}")

                # Salvaguardia: aggiunge 'altro_codice' se ancora assente
                try:
                    cur.execute(
                        """
                            SELECT COUNT(*)
                            FROM information_schema.COLUMNS
                            WHERE TABLE_SCHEMA = DATABASE()
                              AND TABLE_NAME = 'sessioni_doppia_spunta'
                              AND COLUMN_NAME = 'altro_codice'
                        """
                    )
                    res = cur.fetchone()
                    if res and res[0] == 0:
                        print("[INFO] Colonna 'altro_codice' assente: la aggiungo ora")
                        cur.execute(
                            """
                                ALTER TABLE sessioni_doppia_spunta
                                ADD COLUMN altro_codice VARCHAR(50) NULL COMMENT 'Altro codice dalla colonna Prep'
                                AFTER codice_preparatore
                            """
                        )
                        conn.commit()
                        print("[OK] Colonna 'altro_codice' aggiunta (salvaguardia)")
                except Exception as exc:
                    print(f"[WARN] Salvaguardia altro_codice fallita: {exc}")

                # Salvaguardia: aggiunge penalita_eccesso/difetto/penalita/diff se mancanti
                extra_cols = {
                    "penalita_eccesso": "ALTER TABLE sessioni_doppia_spunta ADD COLUMN penalita_eccesso INT NOT NULL DEFAULT 0 COMMENT 'Penalità per errori in eccesso'",
                    "penalita_difetto": "ALTER TABLE sessioni_doppia_spunta ADD COLUMN penalita_difetto INT NOT NULL DEFAULT 0 COMMENT 'Penalità per errori in difetto'",
                    "penalita": "ALTER TABLE sessioni_doppia_spunta ADD COLUMN penalita DECIMAL(10,2) NULL COMMENT 'Penalità assoluta (abs(diff/uvc))' AFTER penalita_difetto",
                    "diff": "ALTER TABLE sessioni_doppia_spunta ADD COLUMN diff DECIMAL(10,2) NULL COMMENT 'Valore differenza originale' AFTER tipo",
                }
                for col_name, alter_sql in extra_cols.items():
                    try:
                        cur.execute(
                            """
                                SELECT COUNT(*)
                                FROM information_schema.COLUMNS
                                WHERE TABLE_SCHEMA = DATABASE()
                                  AND TABLE_NAME = 'sessioni_doppia_spunta'
                                  AND COLUMN_NAME = %s
                            """,
                            (col_name,),
                        )
                        res = cur.fetchone()
                        if res and res[0] == 0:
                            print(f"[INFO] Colonna '{col_name}' assente: la aggiungo ora")
                            cur.execute(alter_sql)
                            conn.commit()
                            print(f"[OK] Colonna '{col_name}' aggiunta (salvaguardia)")
                    except Exception as exc:
                        print(f"[WARN] Salvaguardia {col_name} fallita: {exc}")

                # Elimina le sessioni esistenti per le stesse date/codici
                date_codici = set((s['data'], s['codice_preparatore']) for s in sessioni)
                print(f"[INFO] Eliminazione sessioni esistenti per {len(date_codici)} combinazioni data/codice...")
                deleted_count = 0
                for data, codice in date_codici:
                    cur.execute(
                        "DELETE FROM sessioni_doppia_spunta WHERE data = %s AND codice_preparatore = %s",
                        (data, codice)
                    )
                    deleted_count += cur.rowcount
                print(f"  [INFO] Eliminate {deleted_count} righe esistenti")

                print(f"[INFO] Preparazione insert di {len(sessioni)} righe...")
                
                # Inserisci le nuove righe in batch per evitare timeout
                BATCH_SIZE = 1000
                total_inserted = 0
                
                for i in range(0, len(sessioni), BATCH_SIZE):
                    batch = sessioni[i:i + BATCH_SIZE]
                    cur.executemany(
                    """
                    INSERT INTO sessioni_doppia_spunta 
                    (data, codice_preparatore, altro_codice, numero_riga, excel_row, ora_inizio, ora_fine, tempo_minuti,
                     n_colli, uvc, qtasped, codpro, lotto, tipo, diff, penalita_eccesso, penalita_difetto, penalita,
                     totale_colli_giornata, ore_gestionale_giornaliere)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    [
                        (
                            s['data'],
                            s['codice_preparatore'],
                            s.get('altro_codice'),
                            s['numero_riga'],
                            s.get('excel_row'),
                            s.get('ora_inizio'),
                            s.get('ora_fine'),
                            s.get('tempo_minuti'),
                            s['n_colli'],
                            s.get('uvc'),
                            s.get('qtasped'),
                            s.get('codpro'),
                            s.get('lotto'),
                            s.get('tipo'),
                            s.get('diff'),
                            s.get('penalita_eccesso', 0),
                            s.get('penalita_difetto', 0),
                            s.get('penalita'),
                            s.get('totale_colli_giornata'),
                            s.get('ore_gestionale_giornaliere')
                        )
                        for s in batch
                    ]
                    )
                    total_inserted += cur.rowcount
                    print(f"  [INFO] Batch {i//BATCH_SIZE + 1}: inserite {cur.rowcount} righe (totale: {total_inserted}/{len(sessioni)})")
                
                conn.commit()
                print(f"[OK] Salvate {total_inserted} righe in sessioni_doppia_spunta (transazione completata)")
    except Exception as e:
        import traceback
        print(f"[ERROR] Errore in save_sessioni_doppia_spunta: {e}")
        print(traceback.format_exc())
        try:
            if 'conn' in locals():
                conn.rollback()
                print("[INFO] Rollback eseguito - dati precedenti ripristinati")
        except:
            pass
        raise


def fetch_sessioni_doppia_spunta(
    data_da: Optional[datetime.date] = None,
    data_a: Optional[datetime.date] = None,
    codice_preparatore: Optional[str] = None
) -> List[Dict[str, Any]]:
    """Recupera i dettagli delle sessioni doppia spunta."""
    with closing(_get_conn()) as conn:
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
                  SELECT id, data, codice_preparatore, numero_riga,
                      excel_row, ora_inizio, ora_fine, tempo_minuti, n_colli, tipo,
                      altro_codice, diff,
                      penalita_eccesso, penalita_difetto, penalita,
                       totale_colli_giornata, ore_gestionale_giornaliere,
                       data_importazione
                FROM sessioni_doppia_spunta
                {where_clause}
                ORDER BY data, codice_preparatore, numero_riga
            """
            cur.execute(query, params)
            result = cur.fetchall()
            return cast(List[Dict[str, Any]], result)


def delete_sessioni_doppia_spunta(data: datetime.date, codice_preparatore: Optional[str] = None) -> None:
    """Elimina le sessioni doppia spunta per una data (e opzionalmente un codice)."""
    with closing(_get_conn()) as conn:
        with closing(conn.cursor()) as cur:
            if codice_preparatore:
                cur.execute(
                    "DELETE FROM sessioni_doppia_spunta WHERE data = %s AND codice_preparatore = %s",
                    (data, codice_preparatore)
                )
            else:
                cur.execute(
                    "DELETE FROM sessioni_doppia_spunta WHERE data = %s",
                    (data,)
                )
            conn.commit()


def search_utenti_tim(nominativo: str) -> List[Dict[str, Any]]:
    """Cerca utenti TIM per nominativo (nome/cognome)."""
    nome = (nominativo or "").strip()
    if not nome:
        return []

    like = f"%{nome}%"
    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            has_matricola = _tim_column_exists(cur, "utente", "matricola")
            has_badge = _tim_column_exists(cur, "utente", "badge")

            extra_fields: List[str] = []
            if has_matricola:
                extra_fields.append("matricola")
            if has_badge:
                extra_fields.append("badge")

            extra_select = (", " + ", ".join(extra_fields)) if extra_fields else ""

            query = f"""
                SELECT id, nome, cognome{extra_select}
                FROM utente
                WHERE nome LIKE %s OR cognome LIKE %s OR CONCAT(cognome, ' ', nome) LIKE %s
                ORDER BY cognome, nome
            """
            cur.execute(query, (like, like, like))
            rows = cur.fetchall() or []
            return cast(List[Dict[str, Any]], rows)


def fetch_codici_gestionali_by_utente(utente_id: int) -> List[Dict[str, Any]]:
    """Recupera i codici gestionali associati a un utente TIM."""
    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor(dictionary=True)) as cur:
            has_reparto = _tim_column_exists(cur, "codicegestionale", "reparto")
            reparto_select = "cg.reparto" if has_reparto else "NULL AS reparto"

            query = f"""
                SELECT
                    cg.id AS record_id,
                    cg.tipo_attivita_id,
                    COALESCE(cg.descrizione, cg.codice) AS descrizione,
                    ta.descrizione AS attivita,
                    cg.codice,
                    cg.valido_dal,
                    cg.valido_al,
                    {reparto_select}
                FROM codicegestionale cg
                LEFT JOIN tipoattivita ta ON cg.tipo_attivita_id = ta.id
                WHERE cg.utente_id = %s
                ORDER BY cg.valido_dal DESC, cg.codice
            """
            cur.execute(query, (utente_id,))
            rows = cur.fetchall() or []
            return cast(List[Dict[str, Any]], rows)


def update_codice_gestionale_tim(
    record_id: int,
    tipo_attivita_id: int,
    codice: str,
    valido_dal: datetime.date,
    valido_al: Optional[datetime.date],
    descrizione: Optional[str] = None,
    reparto: Optional[str] = None,
) -> None:
    """Aggiorna un record in codicegestionale su TIM."""
    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor()) as cur:
            has_descrizione = _tim_column_exists(cur, "codicegestionale", "descrizione")
            has_reparto = _tim_column_exists(cur, "codicegestionale", "reparto")

            set_parts: List[str] = [
                "tipo_attivita_id = %s",
                "codice = %s",
                "valido_dal = %s",
                "valido_al = %s",
            ]
            params: List[Any] = [tipo_attivita_id, codice, valido_dal, valido_al]

            if has_descrizione:
                set_parts.append("descrizione = %s")
                params.append(descrizione)
            if has_reparto:
                set_parts.append("reparto = %s")
                params.append(reparto)

            params.append(record_id)

            query = f"""
                UPDATE codicegestionale
                SET {', '.join(set_parts)}
                WHERE id = %s
            """
            cur.execute(query, params)
            conn.commit()


def insert_codice_gestionale_tim(
    utente_id: int,
    tipo_attivita_id: int,
    codice: str,
    valido_dal: datetime.date,
    valido_al: Optional[datetime.date],
    descrizione: Optional[str] = None,
    reparto: Optional[str] = None,
) -> None:
    """Inserisce un nuovo record in codicegestionale su TIM."""
    with closing(_get_conn_main()) as conn:
        with closing(conn.cursor()) as cur:
            has_id = _tim_column_exists(cur, "codicegestionale", "id")
            has_descrizione = _tim_column_exists(cur, "codicegestionale", "descrizione")
            has_reparto = _tim_column_exists(cur, "codicegestionale", "reparto")

            columns: List[str] = []
            values: List[Any] = []

            if has_id:
                columns.append("id")
                values.append(_generate_tim_long_id("001"))

            columns.extend([
                "utente_id",
                "tipo_attivita_id",
                "codice",
                "valido_dal",
                "valido_al",
            ])
            values.extend([utente_id, tipo_attivita_id, codice, valido_dal, valido_al])

            if has_descrizione:
                columns.append("descrizione")
                values.append(descrizione)
            if has_reparto:
                columns.append("reparto")
                values.append(reparto)

            placeholders = ", ".join(["%s"] * len(columns))
            query = f"""
                INSERT INTO codicegestionale ({', '.join(columns)})
                VALUES ({placeholders})
            """
            cur.execute(query, values)
            conn.commit()