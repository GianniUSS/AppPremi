# E:/Progetti/AppEden/.venv/Scripts/python.exe -c "
# import pandas as pd
# path = r'E:/Progetti/AppEden/DaImportare/agosto/LOG Produttività carrellisti (SGB) (agosto).xlsx'
# df = pd.read_excel(path, engine='openpyxl', header=None)
# # trova i blocchi per a96
# mask_code = df[1].astype(str).str.strip().eq('a96')
# idx_list = df.index[mask_code]
# if not len(idx_list):
#     raise SystemExit('Codice a96 non trovato')
# start = idx_list[0] + 1       # righe dati subito dopo l'intestazione
# end = idx_list[0] + 1
# while end < len(df) and pd.isna(df.loc[end,1]):  # continua finche non parte un altro codice/totale
#     end += 1
# bloc = df.loc[start:end-1, [2,4,5,6,7]].copy()
# bloc.columns = ['Tipo','NumeroTrasporto','OraInizio','OraFine','Durata']
# bloc = bloc[bloc['OraInizio'].notna()]          # solo righe movimento
# bloc['Data'] = '2025-08-04'                      # la sezione e quella
# print(bloc)
# """"
"""
Parser per i diversi tipi di file Excel (Preparatori, Carrellisti, Ricevitori, Doppia Spunta).
"""
import datetime
import logging
import numpy as np
from dataclasses import dataclass
from typing import List, Optional
import pandas as pd
from utils import find_column, normalize_string, safe_int_conversion

logger = logging.getLogger(__name__)

@dataclass
class DoppiaSpuntaResult:
    """Risultato del parsing Doppia Spunta con dati per DB e penalità picking."""

    records: pd.DataFrame
    penalita_picking: pd.DataFrame



def _read_excel_with_fallbacks(file_path: str, header_candidates: List[int]) -> pd.DataFrame:
    """Utility interna per leggere un file Excel provando più header."""
    errors: List[str] = []
    for header_row in header_candidates:
        try:
            df = pd.read_excel(file_path, engine="openpyxl", header=header_row)
            if not df.empty and len(df.columns) > 3:
                logger.debug(
                    "File letto con header=%s, righe=%d, colonne=%s",
                    header_row,
                    len(df),
                    list(df.columns),
                )
                return df
        except Exception as exc:
            errors.append(f"Header {header_row}: {exc}")
            continue

    try:
        df = pd.read_excel(file_path, header=header_candidates[0])
        if not df.empty:
            logger.debug(
                "File letto con engine automatico, righe=%d, colonne=%s",
                len(df),
                list(df.columns),
            )
            return df
    except Exception as exc:
        errors.append(str(exc))

    raise ValueError(
        "Impossibile leggere il file Excel. Tentativi falliti: " + "; ".join(errors)
    )


def parse_preparatori(file_path: str) -> pd.DataFrame:
    """Parsa il report Preparatori (PICKING) e restituisce un DataFrame normalizzato."""
    # Header variabile: tentativo con riga 2, fallback riga 0; engine fallback automatico.
    df = None
    errors = []
    
    # Prova diverse strategie di lettura
    # NOTA: Alcuni file Excel hanno header complessi o righe vuote all'inizio
    # Prima: leggi senza header per ispezionare le prime righe
    try:
        df_inspect = pd.read_excel(file_path, engine="openpyxl", header=None, nrows=5)
        logger.debug("Prime 5 righe del file senza header per %s", file_path)
        for idx, row in df_inspect.iterrows():
            logger.debug("  Riga %s: %s", idx, list(row.values)[:10])
        
        # Trova quale riga contiene le intestazioni (cerca "data", "codice", etc.)
        header_row = None
        for idx, row in df_inspect.iterrows():
            row_str = ' '.join([str(v).lower() for v in row.values if pd.notna(v)])
            if 'data' in row_str and ('codice' in row_str or 'preparatore' in row_str):
                header_row = idx
                logger.debug("Trovata riga header: %s", idx)
                break
        
        if header_row is None:
            # Fallback: usa riga 2 se non trovata
            header_row = 2
            logger.debug("Header non trovato automaticamente, uso riga %s", header_row)
    except Exception as e:
        logger.debug("Errore ispezione file: %s, uso header=2", e)
        header_row = 2
    
    # Ora leggi con l'header corretto
    for h_row in [header_row, 2, 1, 0, 3]:
        try:
            df = pd.read_excel(file_path, engine="openpyxl", header=h_row)
            if not df.empty and len(df.columns) > 3:
                logger.debug(
                    "File letto con header=%s, righe=%d, colonne=%d",
                    h_row,
                    len(df),
                    len(df.columns),
                )
                # Verifica se le colonne NON sono tutte "Unnamed"
                unnamed_count = sum(1 for col in df.columns if str(col).startswith('Unnamed'))
                if unnamed_count < len(df.columns) * 0.5:  # Meno del 50% "Unnamed"
                    logger.debug(
                        "Header valido con %d colonne Unnamed su %d",
                        unnamed_count,
                        len(df.columns),
                    )
                    break
                else:
                    logger.debug(
                        "Troppe colonne Unnamed (%d/%d), provo altro header",
                        unnamed_count,
                        len(df.columns),
                    )
        except Exception as e:
            errors.append(f"Header {h_row}: {str(e)}")
            continue
    
    if df is None or df.empty:
        # Ultimo tentativo senza specificare engine
        try:
            df = pd.read_excel(file_path, header=0)
            logger.debug(
                "File letto con engine automatico, righe=%d, colonne=%s",
                len(df),
                list(df.columns),
            )
        except Exception as e:
            raise ValueError(f"Impossibile leggere il file Excel. Errori: {'; '.join(errors + [str(e)])}")

    df = df.rename(columns=lambda x: str(x).strip())

    # Debug: mostra tutte le colonne disponibili con rappresentazione esatta
    logger.debug("Colonne disponibili nel file: %s", list(df.columns))
    
    try:
        col_data = find_column(df, [["data inizio preparazione"], ["data inizio"], ["data"]])
        logger.debug("Colonna data trovata: %s", col_data)
    except ValueError as e:
        logger.error("Errore colonna data: %s", e)
        raise
        
    try:
        col_cod = find_column(df, [["codice preparatore"], ["codice"]])
        logger.debug("Colonna codice trovata: %s", col_cod)
    except ValueError as e:
        logger.error("Errore colonna codice: %s", e)
        raise
        
    try:
        col_nome = find_column(df, [["descrizione preparatore"], ["descrizione"], ["nome"]])  
        logger.debug("Colonna nome trovata: %s", col_nome)
    except ValueError as e:
        logger.debug("Colonna nome non trovata: %s", e)
        # Nome può essere opzionale, proviamo senza
        col_nome = None
        
    try:
        col_colli = find_column(df, [["n colli"], ["colli"], ["totale colli"], ["pezzi"]])
        logger.debug("Colonna colli trovata: %s", col_colli)
    except ValueError as e:
        logger.error("Errore colonna colli: %s", e)
        raise

    col_tempo = None
    try:
        col_tempo = find_column(
            df,
            [["tempo", "lavorato"], ["durata"], ["tempo", "lavoro"]],
            required=False,
        )
        if col_tempo:
            logger.debug("Colonna tempo lavorato trovata: %s", col_tempo)
        else:
            logger.warning("Colonna tempo lavorato non trovata - ore_gestionale impostato a 0")
    except ValueError:
        # Non dovrebbe arrivare qui (required=False), ma gestiamo per sicurezza
        logger.warning("Colonna tempo lavorato assente - ore_gestionale impostato a 0")
        col_tempo = None

    # Seleziona colonne, gestendo il caso di col_nome opzionale
    columns_to_copy = [col_data, col_cod, col_colli]
    if col_nome:
        columns_to_copy.append(col_nome)
    if col_tempo:
        columns_to_copy.append(col_tempo)

    sub = df[columns_to_copy].copy()

    if not col_nome:
        sub['nome_temp'] = ""
    if col_tempo:
        sub['tempo_lavorato_raw'] = sub[col_tempo]
        sub.drop(columns=[col_tempo], inplace=True)
    else:
        sub['tempo_lavorato_raw'] = 0

    # Parsing date vettorializzato: gestisce anche stringhe 8-caratteri yyyymmdd
    sdate = sub[col_data].astype(str).str.strip()
    sdate = sdate.str.replace(r"^(\d{4})(\d{2})(\d{2})$", r"\1-\2-\3", regex=True)
    sub[col_data] = pd.to_datetime(sdate, errors="coerce").dt.date

    sub = sub.dropna(subset=[col_data, col_cod, col_colli])
    logger.debug("Dati processati: %d righe valide", len(sub))

    # Rinomina colonne in base a quelle disponibili
    rename_dict = {
        col_data: "data",
        col_cod: "codice_preparatore", 
        col_colli: "totale_colli",
    }
    
    if col_nome:
        rename_dict[col_nome] = "nome_preparatore"
    else:
        rename_dict['nome_temp'] = "nome_preparatore"
        
    sub.rename(columns=rename_dict, inplace=True)

    def _parse_tempo_lavorato(value) -> float:
        """Converte il tempo lavorato in ore decimali."""
        if pd.isna(value):
            return 0.0
        if isinstance(value, datetime.timedelta):
            return value.total_seconds() / 3600
        if isinstance(value, datetime.time):
            return (
                value.hour
                + value.minute / 60
                + value.second / 3600
            )
        if isinstance(value, datetime.datetime):
            return (
                value.hour
                + value.minute / 60
                + value.second / 3600
            )
        if isinstance(value, (int, float)):
            # Excel memorizza il tempo come frazione di giorno
            return float(value) * 24
        if isinstance(value, str):
            value = value.strip()
            if not value:
                return 0.0
            try:
                td = pd.to_timedelta(value)
                if pd.notna(td):
                    return td.total_seconds() / 3600
            except ValueError:
                pass
            # Tentativo di parsing manuale hh:mm[:ss]
            try:
                parts = value.split(":")
                if len(parts) >= 2:
                    hours = int(parts[0])
                    minutes = int(parts[1])
                    seconds = int(parts[2]) if len(parts) > 2 else 0
                    return hours + minutes / 60 + seconds / 3600
            except ValueError:
                return 0.0
        return 0.0

    sub["codice_preparatore"] = sub["codice_preparatore"].astype(str).str.strip()
    sub["nome_preparatore"] = sub["nome_preparatore"].astype(str).str.strip()
    sub["totale_colli"] = pd.to_numeric(sub["totale_colli"], errors="coerce").fillna(0).astype(int)
    sub["ore_gestionale"] = sub["tempo_lavorato_raw"].apply(_parse_tempo_lavorato).astype(float)
    sub.drop(columns=["tempo_lavorato_raw"], inplace=True)

    out = (
        sub.groupby(["data", "codice_preparatore", "nome_preparatore"], as_index=False)
        .agg({"totale_colli": "sum", "ore_gestionale": "sum"})
        .assign(penalita=0, tipo_attivita="PICKING", tipo="")
    )
    out["ore_gestionale"] = out["ore_gestionale"].round(2)
    return out


def parse_carrelisti(file_path: str, data_rif: datetime.date) -> pd.DataFrame:
    """
    Parsa il report Carrellisti (CARRELLO) con logica sessioni:
    - Legge Ora inizio e Ora fine
    - Calcola gap tra inizi consecutivi
    - Chiude sessione se gap > 15 minuti
    - Proporziona tempo sessione per tipo
    - Restituisce DataFrame con ore_gestionale calcolato
    """
    
    # Prova a leggere il file con diverse strategie di header
    df = None
    errors = []
    
    # Prima: ispeziona le prime righe per trovare l'header
    try:
        df_inspect = pd.read_excel(file_path, engine="openpyxl", header=None, nrows=10)
        # Cerca la riga che contiene "Preparatore" o "Data" + "Tipo" + "Ora"
        header_row = None
        for idx, row in df_inspect.iterrows():
            row_str = ' '.join([str(v).lower() for v in row.values if pd.notna(v)])
            # Cerca header con ora inizio/fine
            if 'ora' in row_str and ('inizio' in row_str or 'fine' in row_str):
                header_row = idx
                break
            elif 'preparatore' in row_str or ('data' in row_str and 'tipo' in row_str):
                header_row = idx
                break
        
        if header_row is None:
            raise ValueError("Header non trovato nel file carrellisti")
            
        # Leggi con l'header trovato
        df = pd.read_excel(file_path, engine="openpyxl", header=header_row)
    except Exception as e:
        logger.error("Errore lettura file carrellisti: %s", e)
        raise ValueError(f"Impossibile leggere il file carrellisti: {e}")
    
    # Normalizza i nomi delle colonne
    df.columns = [str(c).strip() if pd.notna(c) else f"Col_{i}" for i, c in enumerate(df.columns)]
    df = df.dropna(how="all")
    
    # Cerca le colonne necessarie
    data_col = None
    prep_col = None
    tipo_col = None
    num_col = None
    ora_inizio_col = None
    ora_fine_col = None
    errore_col = None
    
    for col in df.columns:
        col_norm = normalize_string(col)
        if col_norm == "data":
            data_col = col
        elif "preparatore" in col_norm:
            prep_col = col
        elif col_norm == "tipo":
            tipo_col = col
        elif "n°" in col_norm or "ntrasporto" in col_norm or "trasporto" in col_norm:
            num_col = col
        elif "ora" in col_norm and "inizio" in col_norm:
            ora_inizio_col = col
        elif "ora" in col_norm and "fine" in col_norm:
            ora_fine_col = col
        elif "errore" in col_norm:
            errore_col = col
    
    if not prep_col or not tipo_col:
        raise ValueError(f"Colonne necessarie non trovate. Preparatore={prep_col}, Tipo={tipo_col}")
    
    if not ora_inizio_col or not ora_fine_col:
        logger.debug("Colonne Ora inizio/fine non trovate, uso la logica senza sessioni")
        # Fallback alla logica vecchia (senza ore_gestionale)
        return _parse_carrellisti_old_logic(df, data_col, prep_col, tipo_col, num_col, data_rif)
    
    # Funzione per convertire ora in datetime.time
    def parse_time(val):
        """Converte vari formati di ora in datetime.time"""
        if pd.isna(val):
            return None
        
        # Se è già datetime.time
        if isinstance(val, datetime.time):
            return val
        
        # Se è datetime.datetime
        if isinstance(val, datetime.datetime):
            return val.time()
        
        # Se è stringa formato HH:MM o HH:MM:SS
        if isinstance(val, str):
            val = val.strip()
            try:
                parts = val.split(':')
                if len(parts) >= 2:
                    hour = int(parts[0])
                    minute = int(parts[1])
                    second = int(parts[2]) if len(parts) > 2 else 0
                    return datetime.time(hour, minute, second)
            except:
                pass
        
        # Se è un float (Excel time format: frazione di giorno)
        if isinstance(val, (int, float)):
            try:
                # CASO 1: Se il valore è < 1, è una frazione di giorno Excel (0.5 = 12:00)
                if 0 <= val < 1:
                    total_seconds = int(val * 24 * 3600)
                    hours = total_seconds // 3600
                    minutes = (total_seconds % 3600) // 60
                    seconds = total_seconds % 60
                    return datetime.time(hours % 24, minutes, seconds)
                
                # CASO 2: Se il valore è >= 1, è formato HH.MM (es. 07.53 = 07:53)
                # Es: 07.53 → ore=7, minuti=53
                # Es: 14.09 → ore=14, minuti=09
                else:
                    ore = int(val)
                    minuti_decimale = (val - ore) * 100
                    minuti = int(round(minuti_decimale))
                    
                    # Valida che sia un orario sensato
                    if 0 <= ore < 24 and 0 <= minuti < 60:
                        return datetime.time(ore, minuti, 0)
                    else:
                        return None
            except:
                pass
        
        return None
    
    # Converti le colonne data
    if data_col:
        def convert_date(val):
            if pd.isna(val):
                return None
            if isinstance(val, (datetime.datetime, datetime.date)):
                return val if isinstance(val, datetime.date) else val.date()
            if isinstance(val, (int, float)):
                val_str = str(int(val))
                if len(val_str) == 8:
                    try:
                        return datetime.date(int(val_str[0:4]), int(val_str[4:6]), int(val_str[6:8]))
                    except:
                        pass
            if isinstance(val, str):
                val_str = val.strip()
                if len(val_str) == 8 and val_str.isdigit():
                    try:
                        return datetime.date(int(val_str[0:4]), int(val_str[4:6]), int(val_str[6:8]))
                    except:
                        pass
            try:
                return pd.to_datetime(val).date()
            except:
                return None

        parsed_dates = df[data_col].apply(convert_date)
        df["_is_new_date_section"] = parsed_dates.notna()
        df[data_col] = parsed_dates.ffill()
    
    # Salva una copia dei valori originali delle colonne orario per eventuali debug
    df["_raw_ora_inizio"] = df[ora_inizio_col]
    df["_raw_ora_fine"] = df[ora_fine_col]

    # Converti le colonne orario
    df[ora_inizio_col] = df[ora_inizio_col].apply(parse_time)
    df[ora_fine_col] = df[ora_fine_col].apply(parse_time)
    
    DEBUG_CODICE = "a96"

    # STEP 1: Leggi tutti i dati e organizza per (data, preparatore)
    all_rows = []
    current_code = None
    
    for idx, row in df.iterrows():
        prep = str(row[prep_col]).strip() if pd.notna(row[prep_col]) else ""
        tipo = str(row[tipo_col]).strip() if pd.notna(row[tipo_col]) else ""
        
        if prep.upper() == "TOTALE" or tipo.upper() == "TOTALE":
            continue
        
        # Determina la data
        if data_col and pd.notna(row[data_col]):
            record_date = row[data_col]
        else:
            record_date = data_rif
        
        # Reset se la riga contiene una nuova data (inizio sezione)
        if data_col and bool(row.get("_is_new_date_section", False)):
            current_code = None

        # Riga intestazione preparatore
        if prep and (not tipo or tipo.upper() in ["", "NAN", "ST", "AP", "CM"]):
            current_code = prep
            continue
        
        if not current_code:
            continue
        
        # Riga dati
        if tipo and tipo.upper() != "NAN":
            ora_inizio = row[ora_inizio_col] if pd.notna(row[ora_inizio_col]) else None
            ora_fine = row[ora_fine_col] if pd.notna(row[ora_fine_col]) else None
            
            if not ora_inizio or not ora_fine:
                continue

            if ora_inizio == ora_fine:
                if prep.lower() == DEBUG_CODICE:
                    logger.debug(
                        "Riga con stessa ora di inizio/fine rilevata per %s %s (%s -> %s)",
                        prep,
                        tipo,
                        ora_inizio,
                        ora_fine,
                    )
            
            # Leggi la colonna errore (se esiste)
            errore = None
            if errore_col and pd.notna(row[errore_col]):
                errore = str(row[errore_col]).strip()
            
            # IMPORTANTE: Per i carrellisti, ogni riga = 1 movimento
            # La colonna N°trasporto è un CODICE IDENTIFICATIVO, non una quantità!
            # Se c'è un errore, il movimento NON viene conteggiato (colli = 0)
            colli = 0 if errore else 1
            
            all_rows.append({
                "data": record_date,
                "codice": current_code,
                "tipo": tipo,
                "colli": colli,
                "ora_inizio": ora_inizio,
                "ora_fine": ora_fine,
                "errore": errore,
                "raw_ora_inizio": row.get("_raw_ora_inizio"),
                "raw_ora_fine": row.get("_raw_ora_fine"),
            })
    
    if "_is_new_date_section" in df.columns:
        df.drop(columns=["_is_new_date_section"], inplace=True)

    # STEP 1.5: RAGGRUPPA righe duplicate (stesso data, codice, ora_inizio, ora_fine, tipo, errore)
    # PRIMA di calcolare le sessioni!
    righe_consolidate = {}
    for row in all_rows:
        chiave = (row["data"], row["codice"], row["ora_inizio"], row["ora_fine"], row["tipo"], row.get("errore"))
        if chiave not in righe_consolidate:
            righe_consolidate[chiave] = {
                "data": row["data"],
                "codice": row["codice"],
                "tipo": row["tipo"],
                "ora_inizio": row["ora_inizio"],
                "ora_fine": row["ora_fine"],
                "colli": 0,
                "errore": row.get("errore"),
                "raw_ora_inizio": row.get("raw_ora_inizio"),
                "raw_ora_fine": row.get("raw_ora_fine"),
            }
        righe_consolidate[chiave]["colli"] += row["colli"]
    
    # Converti il dizionario in lista
    all_rows = list(righe_consolidate.values())
    # STEP 2: Calcola sessioni con logica gap > 15 minuti
    records_finali = []
    sessioni_dettaglio = []  # Per salvare dettagli sessioni nel DB
    # Raggruppa per (data, codice)
    from itertools import groupby
    
    all_rows_sorted = sorted(all_rows, key=lambda x: (x["data"], x["codice"], x["ora_inizio"]))
    
    for (data, codice), group in groupby(all_rows_sorted, key=lambda x: (x["data"], x["codice"])):
        righe_giorno = list(group)
        righe_giorno.sort(key=lambda x: x["ora_inizio"])
        
        is_debug_code = codice.lower() == DEBUG_CODICE

        if is_debug_code:
            logger.debug("Dettaglio sessioni per codice %s in data %s", DEBUG_CODICE, data)

            def fmt_minutes(total_minutes: Optional[float]) -> str:
                if total_minutes is None:
                    return "-"
                sign = "-" if total_minutes < 0 else ""
                total = abs(total_minutes)
                minutes = int(round(total))
                hours, mins = divmod(minutes, 60)
                if hours:
                    return f"{sign}{hours}:{mins:02d}h"
                return f"{sign}{mins:02d}'"

            def fmt_raw(raw_val) -> str:
                if raw_val is None:
                    return ""
                if isinstance(raw_val, float):
                    return f"{raw_val:.12g}"
                if isinstance(raw_val, datetime.datetime):
                    return raw_val.strftime("%H:%M:%S")
                if isinstance(raw_val, datetime.time):
                    return raw_val.strftime("%H:%M:%S")
                return str(raw_val)

            valid_rows = sum(1 for r in righe_giorno if not r.get("errore"))
            error_rows = len(righe_giorno) - valid_rows
            logger.debug(
                "Righe totali=%d | valide=%d | con errore=%d",
                len(righe_giorno),
                valid_rows,
                error_rows,
            )
            logger.debug("  #  Sess Tipo  Inizio Fine   Dur  Gap  Mov Stato RawIn RawFi Note")

            prev_fine_dt = None
            current_session = 1
            for idx_riga, riga in enumerate(righe_giorno, 1):
                dt_inizio = datetime.datetime.combine(datetime.date.today(), riga["ora_inizio"])
                dt_fine = datetime.datetime.combine(datetime.date.today(), riga["ora_fine"])
                if dt_fine < dt_inizio:
                    dt_fine += datetime.timedelta(days=1)

                durata_min = (dt_fine - dt_inizio).total_seconds() / 60
                gap_min = None
                note_parts: List[str] = []

                if prev_fine_dt is not None:
                    raw_gap = (dt_inizio - prev_fine_dt).total_seconds() / 60
                    if raw_gap > 15:
                        current_session += 1
                        note_parts.append("NEW session")
                        gap_min = raw_gap
                    else:
                        if raw_gap < 0:
                            note_parts.append("OVERLAP")
                        gap_min = max(0, raw_gap)

                sess_label = f"S{current_session}"
                gap_desc = fmt_minutes(gap_min)
                durata_desc = fmt_minutes(durata_min)

                errore_val = riga.get("errore") or ""
                stato = "SKIP" if errore_val else "OK"
                if errore_val:
                    note_parts.append(f"Errore={errore_val}")

                raw_in_str = fmt_raw(riga.get("raw_ora_inizio"))
                raw_fi_str = fmt_raw(riga.get("raw_ora_fine"))

                note = ", ".join(note_parts)
                ora_inizio_str = riga["ora_inizio"].strftime("%H:%M")
                ora_fine_str = riga["ora_fine"].strftime("%H:%M")

                logger.debug(
                    "%3d  %-4s%-4s%-7s%-6s%5s%6s%5d  %-4s %-10s %-10s %s",
                    idx_riga,
                    sess_label,
                    riga["tipo"],
                    ora_inizio_str,
                    ora_fine_str,
                    durata_desc,
                    gap_desc,
                    riga["colli"],
                    stato,
                    raw_in_str,
                    raw_fi_str,
                    note,
                )

                if prev_fine_dt is None:
                    prev_fine_dt = dt_fine
                else:
                    prev_fine_dt = max(prev_fine_dt, dt_fine)

            logger.debug("Fine dettaglio sessioni per %s in data %s", DEBUG_CODICE, data)
        
        # Dividi in sessioni
        sessioni = []
        sessione_corrente: List[dict] = []
        session_end_dt: Optional[datetime.datetime] = None

        for riga in righe_giorno:
            dt_inizio = datetime.datetime.combine(datetime.date.today(), riga["ora_inizio"])
            dt_fine = datetime.datetime.combine(datetime.date.today(), riga["ora_fine"])
            if dt_fine < dt_inizio:
                dt_fine += datetime.timedelta(days=1)

            if not sessione_corrente:
                sessione_corrente = [riga]
                session_end_dt = dt_fine
                continue

            gap_minuti = (dt_inizio - session_end_dt).total_seconds() / 60 if session_end_dt else None

            if gap_minuti is not None and gap_minuti > 15:
                sessioni.append(sessione_corrente)
                sessione_corrente = [riga]
                session_end_dt = dt_fine
            else:
                sessione_corrente.append(riga)
                if session_end_dt:
                    session_end_dt = max(session_end_dt, dt_fine)

        if sessione_corrente:
            sessioni.append(sessione_corrente)
        
        # STEP 3: Calcola tempo totale di TUTTE le sessioni e conta movimenti per tipo
        tempo_totale_giornata_ore = 0.0
        movimenti_per_tipo = {}  # Conta movimenti validi per tipo
        errori_per_tipo_giornata = {}  # Conta errori per tipo
        righe_con_sessione = []  # Lista di TUTTE le righe con info sessione
        
        for num_sess, righe_sess in enumerate(sessioni, 1):
            # Tempo totale sessione: da primo inizio a ultima fine
            primo_inizio = righe_sess[0]["ora_inizio"]
            ultima_fine = righe_sess[-1]["ora_fine"]
            
            dt_inizio = datetime.datetime.combine(datetime.date.today(), primo_inizio)
            dt_fine = datetime.datetime.combine(datetime.date.today(), ultima_fine)
            
            # Gestisci caso di passaggio mezzanotte
            if dt_fine < dt_inizio:
                if is_debug_code:
                    logger.debug(
                        "Fine (%s) precedente all'inizio (%s) per codice %s in data %s",
                        ultima_fine,
                        primo_inizio,
                        codice,
                        data,
                    )
                dt_fine += datetime.timedelta(days=1)
            
            tempo_sessione_min = (dt_fine - dt_inizio).total_seconds() / 60
            tempo_sessione_ore = tempo_sessione_min / 60.0
            tempo_totale_giornata_ore += tempo_sessione_ore
            
            # Conta movimenti per tipo in questa sessione
            # Separa movimenti validi da errori
            for riga in righe_sess:
                tipo = riga["tipo"]
                colli = riga["colli"]
                errore = riga.get("errore")
                
                if errore:
                    # Conta gli errori separatamente
                    if tipo not in errori_per_tipo_giornata:
                        errori_per_tipo_giornata[tipo] = 0
                    errori_per_tipo_giornata[tipo] += 1
                else:
                    # Conta solo i movimenti validi
                    if tipo not in movimenti_per_tipo:
                        movimenti_per_tipo[tipo] = 0
                    movimenti_per_tipo[tipo] += colli
            
            # Salva OGNI RIGA con i dettagli della sessione
            for idx_riga, riga in enumerate(righe_sess):
                righe_con_sessione.append({
                    'riga': riga,
                    'numero_sessione': num_sess,
                    'ora_inizio_sessione': primo_inizio,
                    'ora_fine_sessione': ultima_fine,
                    'tempo_sessione_ore': tempo_sessione_ore,
                    'totale_righe_sessione': len(righe_sess)
                })

        if is_debug_code:
            logger.debug(
                "Sessioni calcolate=%d | Tempo totale=%.2fh",
                len(sessioni),
                tempo_totale_giornata_ore,
            )
            for idx_sess, righe_sess in enumerate(sessioni, 1):
                start = righe_sess[0]["ora_inizio"].strftime("%H:%M")
                end = righe_sess[-1]["ora_fine"].strftime("%H:%M")
                validi_sess = sum(1 for r in righe_sess if not r.get("errore"))
                errori_sess = sum(1 for r in righe_sess if r.get("errore"))
                logger.debug(
                    "  S%s: %s->%s | righe valide=%d | errori=%d",
                    idx_sess,
                    start,
                    end,
                    validi_sess,
                    errori_sess,
                )
            logger.debug("Fine riepilogo sessioni per %s in data %s", codice, data)
        
        # STEP 4: Proporziona il tempo totale in base ai movimenti per tipo
        totale_movimenti = sum(movimenti_per_tipo.values())
        
        if totale_movimenti == 0:
            continue
        
        # Calcola ore gestionale per tipo in base ai movimenti totali
        ore_gestionale_per_tipo = {}
        for tipo, movimenti in movimenti_per_tipo.items():
            proporzione = movimenti / totale_movimenti
            ore_gestionale = tempo_totale_giornata_ore * proporzione
            ore_gestionale_per_tipo[tipo] = ore_gestionale
            
            # Crea note_errori se ci sono errori per questo tipo
            note_errori = None
            if tipo in errori_per_tipo_giornata:
                num_errori = errori_per_tipo_giornata[tipo]
                note_errori = f"{num_errori} movimento{'i' if num_errori > 1 else ''} con errore"
            
            records_finali.append({
                "data": data,
                "codice_preparatore": codice,
                "nome_preparatore": "",
                "totale_colli": movimenti,
                "penalita": 0,
                "tipo_attivita": "CARRELLISTI",
                "tipo": tipo,
                "ore_tim": 0.0,
                "ore_gestionale": round(ore_gestionale, 2),
                "note_errori": note_errori,
            })
        
        # Salva OGNI riga con i dettagli della sessione
        # ESCLUDE le righe con errore (non vanno in sessioni_carrellisti)
        numero_riga_giornata = 0
        sessione_corrente = None
        tipo_gia_visto_in_giornata = set()
        fine_riga_precedente_dt: Optional[datetime.datetime] = None
        
        for riga_info in righe_con_sessione:
            riga = riga_info['riga']
            errore = riga.get('errore')
            
            # SALTA le righe con errore - non vanno in sessioni_carrellisti
            if errore:
                continue
            
            numero_riga_giornata += 1
            tipo = riga['tipo']
            num_sess = riga_info['numero_sessione']
            ora_inizio = riga['ora_inizio']
            ora_fine = riga['ora_fine']
            conteggio_movimenti = riga['colli']  # Già consolidato!
            
            # Calcola tempo di questa riga
            dt_inizio_riga = datetime.datetime.combine(datetime.date.today(), ora_inizio)
            dt_fine_riga = datetime.datetime.combine(datetime.date.today(), ora_fine)
            if dt_fine_riga < dt_inizio_riga:
                dt_fine_riga += datetime.timedelta(days=1)
            tempo_riga_minuti = (dt_fine_riga - dt_inizio_riga).total_seconds() / 60
            
            # Calcola gap: tempo tra FINE riga precedente e INIZIO riga corrente (NULL per prima riga o nuova sessione)
            # IMPORTANTE: gap si azzera anche a inizio nuova sessione
            if fine_riga_precedente_dt is not None and num_sess == sessione_corrente:
                raw_gap = (dt_inizio_riga - fine_riga_precedente_dt).total_seconds() / 60
                gap_minuti = round(max(0.0, raw_gap), 2)
            else:
                gap_minuti = None
            
            # I totali vanno solo sulla PRIMA riga di ogni sessione
            if num_sess != sessione_corrente:
                # Nuova sessione - metti i totali
                sessione_corrente = num_sess
                ora_inizio_sess = riga_info['ora_inizio_sessione']
                ora_fine_sess = riga_info['ora_fine_sessione']
                tempo_sess = round(riga_info['tempo_sessione_ore'], 2)
                tot_righe = riga_info['totale_righe_sessione']
                fine_riga_precedente_dt = dt_fine_riga
            else:
                # Riga successiva della stessa sessione - NULL
                ora_inizio_sess = None
                ora_fine_sess = None
                tempo_sess = None
                tot_righe = None
                fine_riga_precedente_dt = max(fine_riga_precedente_dt, dt_fine_riga)
            
            # Movimenti: conteggio nella colonna corrispondente al tipo, NULL nelle altre
            movimenti_st = conteggio_movimenti if tipo == 'ST' else None
            movimenti_ss = conteggio_movimenti if tipo == 'SS' else None
            movimenti_ap = conteggio_movimenti if tipo == 'AP' else None
            movimenti_cm = conteggio_movimenti if tipo == 'CM' else None
            
            # ore_gestionale per tipo va solo sulla PRIMA occorrenza del tipo nella giornata
            if tipo not in tipo_gia_visto_in_giornata:
                ore_gest_st = round(ore_gestionale_per_tipo.get('ST', 0), 2) if tipo == 'ST' else None
                ore_gest_ss = round(ore_gestionale_per_tipo.get('SS', 0), 2) if tipo == 'SS' else None
                ore_gest_ap = round(ore_gestionale_per_tipo.get('AP', 0), 2) if tipo == 'AP' else None
                ore_gest_cm = round(ore_gestionale_per_tipo.get('CM', 0), 2) if tipo == 'CM' else None
                tipo_gia_visto_in_giornata.add(tipo)
            else:
                ore_gest_st = None
                ore_gest_ss = None
                ore_gest_ap = None
                ore_gest_cm = None
            
            sessioni_dettaglio.append({
                'data': data,
                'codice_preparatore': codice,
                'numero_riga': numero_riga_giornata,
                'ora_inizio_riga': ora_inizio,
                'ora_fine_riga': ora_fine,
                'tempo_riga_minuti': round(tempo_riga_minuti, 2),
                'gap_minuti': gap_minuti,
                'movimenti_st': movimenti_st,
                'movimenti_ss': movimenti_ss,
                'movimenti_ap': movimenti_ap,
                'movimenti_cm': movimenti_cm,
                'errore': riga.get('errore'),
                'numero_sessione': num_sess,
                'ora_inizio_sessione': ora_inizio_sess,
                'ora_fine_sessione': ora_fine_sess,
                'tempo_sessione_ore': tempo_sess,
                'totale_righe_sessione': tot_righe,
                'ore_gestionale_st': ore_gest_st,
                'ore_gestionale_ss': ore_gest_ss,
                'ore_gestionale_ap': ore_gest_ap,
                'ore_gestionale_cm': ore_gest_cm
            })
            
            # fine_riga_precedente_dt già aggiornato nella logica qui sopra
    
    # Salva le sessioni nel database
    if sessioni_dettaglio:
        try:
            from database import save_sessioni_carrellisti
            save_sessioni_carrellisti(sessioni_dettaglio)
        except Exception as e:
            logger.exception("Errore nel salvataggio delle sessioni: %s", e)
    
    out = pd.DataFrame(records_finali)
    if out.empty:
        return out
    
    # Raggruppa per sommare colli e ore dello stesso tipo
    out = out.groupby(
        ["data", "codice_preparatore", "nome_preparatore", "tipo", "tipo_attivita"],
        as_index=False
    ).agg({
        "totale_colli": "sum",
        "penalita": "first",
        "ore_tim": "sum",
        "ore_gestionale": "sum"
    })
    
    # Assicura che totale_colli sia int e penalita sia int
    out["totale_colli"] = out["totale_colli"].astype(int)
    out["penalita"] = out["penalita"].astype(int)
    
    # Debug: mostra i primi record
    return out


def _parse_carrellisti_old_logic(
    df: pd.DataFrame,
    data_col: Optional[str],
    prep_col: str,
    tipo_col: str,
    num_col: Optional[str],
    data_rif: datetime.date
) -> pd.DataFrame:
    """Logica vecchia senza sessioni (fallback se mancano Ora inizio/fine)"""
    
    records: List[dict] = []
    current_code = None
    skipped_count = 0
    
    for idx, row in df.iterrows():
        prep = str(row[prep_col]).strip() if pd.notna(row[prep_col]) else ""
        tipo = str(row[tipo_col]).strip() if pd.notna(row[tipo_col]) else ""
        
        if prep.upper() == "TOTALE" or tipo.upper() == "TOTALE":
            continue
        
        # Determina la data
        if data_col and pd.notna(row[data_col]):
            record_date = row[data_col]
        else:
            record_date = data_rif
        
        # Riga intestazione
        if prep and (not tipo or tipo.upper() in ["", "NAN"]):
            current_code = prep
            continue
        
        if not current_code:
            skipped_count += 1
            continue
        
        if tipo and tipo.upper() != "NAN":
            # IMPORTANTE: Per i carrellisti, ogni riga = 1 movimento
            # La colonna N°trasporto è un CODICE IDENTIFICATIVO, non una quantità!
            colli = 1
            
            records.append({
                "data": record_date,
                "codice_preparatore": current_code,
                "nome_preparatore": "",
                "totale_colli": colli,
                "penalita": 0,
                "tipo_attivita": "CARRELLISTI",
                "tipo": tipo,
                "ore_tim": 0.0,
                "ore_gestionale": 0.0,
            })
    
    logger.debug(
        "Logica carrellisti senza sessioni: record creati=%d, righe saltate=%d",
        len(records),
        skipped_count,
    )
    
    out = pd.DataFrame(records)
    if out.empty:
        return out
    
    out = out.groupby(
        ["data", "codice_preparatore", "nome_preparatore", "tipo", "tipo_attivita"],
        as_index=False
    )["totale_colli"].sum()
    out["penalita"] = 0
    out["ore_tim"] = 0.0
    out["ore_gestionale"] = 0.0
    
    # Assicura che totale_colli e penalita siano int
    out["totale_colli"] = out["totale_colli"].astype(int)
    out["penalita"] = out["penalita"].astype(int)
    
    return out


def _find_specific_prep_column(df: pd.DataFrame) -> Optional[str]:
    """Restituisce la colonna che contiene i codici tipo `Q57-Q57` evitando `prepa/prepc`."""
    for col in df.columns:
        norm = normalize_string(col)
        if "prep" in norm and norm not in {"prepa", "prepc"}:
            series = df[col].astype(str).str.strip()
            if series.str.contains("-", na=False).any() or series.str.match(r"^[A-Z]\d+", na=False).any():
                return col
    return None


def parse_doppia_spunta(file_path: str) -> DoppiaSpuntaResult:
    """Parsa il report Doppia Spunta restituendo dati per DB e penalità PICKING."""
    df = _read_excel_with_fallbacks(file_path, [1, 0, 2, 3])
    df = df.rename(columns=lambda x: str(x).strip())

    logger.debug("Colonne disponibili Doppia Spunta: %s", list(df.columns))

    # Data
    try:
        col_data = find_column(df, [["data", "spunta"], ["data"]])
        logger.debug("Colonna data trovata: %s", col_data)
    except ValueError as e:
        logger.error("Errore nella ricerca della colonna data: %s", e)
        raise

    # Codice preparatore dalla colonna G (es. 'prepa')
    try:
        col_codice_preparatore = find_column(df, [["prepa"], ["codice"]])
        logger.debug("Colonna codice preparatore trovata: %s", col_codice_preparatore)
    except ValueError as e:
        logger.error("Errore nella ricerca del codice preparatore: %s", e)
        raise

    # Codice preparatore PICKING (prespo) dalla colonna Q
    col_codice_picking = _find_specific_prep_column(df)
    if not col_codice_picking:
        try:
            col_codice_picking = find_column(df, [["prep"]])
        except ValueError as exc:
            logger.error("Colonna codice PICKING non trovata")
            raise ValueError(
                "Colonna codice preparatore PICKING (es. 'Q57-Q57') non trovata."
            ) from exc
    logger.debug("Colonna codice PICKING trovata: %s", col_codice_picking)

    # Tipo valorizzato con la colonna E ('rs' o descrizione cliente)
    col_tipo = find_column(df, [["rs"], ["tipo"], ["cliente"]], required=False)
    if not col_tipo:
        logger.debug("Colonna tipo non trovata, uso valore vuoto")
        # Invece di sollevare errore, crea una colonna tipo vuota
        df["_tipo_temp"] = ""
        col_tipo = "_tipo_temp"
    else:
        logger.debug("Colonna tipo trovata: %s", col_tipo)

    # Penalità (differenze o colonne con 'penal')
    try:
        col_penalita = find_column(df, [["diffe"], ["penal"], ["differenze"]])
        logger.debug("Colonna penalità trovata: %s", col_penalita)
    except ValueError:
        logger.debug("Colonna penalità non trovata, uso valore 0")
        col_penalita = None

    # Quantità da sommare (preferiamo UVC, altrimenti qtaspunta, righe)
    col_quantita: Optional[str] = None
    for candidate in (["uvc"], ["qta", "spunta"], ["qtaspunta"], ["righe"], ["quantita"]):
        try:
            col_quantita = find_column(df, [candidate], required=False)
        except ValueError:
            col_quantita = None
        if col_quantita:
            logger.debug("Colonna quantità trovata: %s", col_quantita)
            break
    
    if not col_quantita:
        logger.debug("Nessuna colonna quantità trovata, uso valore 0")

    subset_cols = {
        "data": col_data,
        "codice_preparatore": col_codice_preparatore,
        "tipo": col_tipo,
        "codice_picking": col_codice_picking,
    }
    if col_penalita:
        subset_cols["penalita"] = col_penalita
    if col_quantita:
        subset_cols["quantita"] = col_quantita

    sub = df[list(subset_cols.values())].copy()
    sub.columns = list(subset_cols.keys())

    # Normalizzazione data
    sdate = sub["data"].astype(str).str.strip()
    sdate = sdate.str.replace(r"^(\d{4})(\d{2})(\d{2})$", r"\1-\2-\3", regex=True)
    sub["data"] = pd.to_datetime(sdate, errors="coerce").dt.date

    sub["codice_preparatore"] = sub["codice_preparatore"].astype(str).str.strip().str.upper()
    sub["tipo"] = sub["tipo"].astype(str).str.strip()
    sub.loc[sub["tipo"].str.lower() == "nan", "tipo"] = ""
    sub["nome_preparatore"] = ""

    # Normalizzazione codice PICKING (colonna Q)
    sub["codice_picking"] = (
        sub["codice_picking"].astype(str).str.strip().str.split("-").str[0].str.upper()
    )

    if "quantita" in sub.columns:
        uvc_val = pd.to_numeric(sub["quantita"], errors="coerce").fillna(0).abs()
    else:
        uvc_val = pd.Series(0, index=sub.index, dtype=float)

    if "penalita" in sub.columns:
        penalita_raw = pd.to_numeric(sub["penalita"], errors="coerce").fillna(0).abs()
    else:
        penalita_raw = pd.Series(0, index=sub.index, dtype=float)

    sub["totale_colli"] = uvc_val.astype(int)

    # Calcola l'errore come |diffe| / UVC; se UVC è 0 manteniamo l'errore grezzo
    penalita_ratio = penalita_raw.div(uvc_val.replace(0, np.nan)).fillna(penalita_raw)
    sub["penalita"] = penalita_ratio.astype(float)

    sub = sub.dropna(subset=["data", "codice_preparatore", "codice_picking"])
    
    logger.debug("Righe valide dopo pulizia: %d", len(sub))

    # Creo DataFrame per penalità PICKING
    penalita_picking = (
        sub.groupby(["data", "codice_picking"], as_index=False)["penalita"]
        .sum()
    )
    penalita_picking = penalita_picking.rename(columns={"codice_picking": "codice_preparatore"})
    penalita_picking["penalita"] = np.ceil(penalita_picking["penalita"]).astype(int)

    grouped = (
        sub.drop(columns=["codice_picking"])
        .groupby(["data", "codice_preparatore", "nome_preparatore", "tipo"], as_index=False)[
            ["totale_colli", "penalita"]
        ]
        .sum()
        .assign(tipo_attivita="DOPPIA_SPUNTA")
    )
    grouped["totale_colli"] = grouped["totale_colli"].astype(int)
    grouped["penalita"] = np.ceil(grouped["penalita"]).astype(int)

    # Riorganizzo le colonne per coerenza con il resto dei parser
    grouped = grouped[[
        "data",
        "codice_preparatore",
        "nome_preparatore",
        "totale_colli",
        "penalita",
        "tipo_attivita",
        "tipo",
    ]]

    penalita_picking["tipo_attivita"] = "PICKING"

    return DoppiaSpuntaResult(records=grouped, penalita_picking=penalita_picking)


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

        colli_int = safe_int_conversion(colli)

        records.append({
            "data": data_fmt,
            "codice_preparatore": str(codice).strip(),
            "nome_preparatore": "",
            "totale_colli": colli_int,
            "penalita": 0,
            "tipo_attivita": "RICEVITORI",
            "tipo": "",
        })

    df_out = pd.DataFrame(records)
    if df_out.empty:
        return df_out

    # somma colli per data e codice
    df_grouped = df_out.groupby(
        ["data", "codice_preparatore", "nome_preparatore", "tipo_attivita", "tipo"],
        as_index=False,
    )["totale_colli"].sum()

    return df_grouped