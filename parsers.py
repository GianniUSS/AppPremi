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
            df.attrs["source_header_row"] = header_row
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
        df.attrs["source_header_row"] = header_candidates[0]
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
            df.attrs["source_header_row"] = 0
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
    # Converti ore in MINUTI moltiplicando per 60
    sub["ore_gestionale"] = (sub["tempo_lavorato_raw"].apply(_parse_tempo_lavorato) * 60).astype(float)
    sub.drop(columns=["tempo_lavorato_raw"], inplace=True)

    out = (
        sub.groupby(["data", "codice_preparatore", "nome_preparatore"], as_index=False)
        .agg({"totale_colli": "sum", "ore_gestionale": "sum"})
        .assign(
            penalita_eccesso=0,
            penalita_difetto=0,
            tipo_attivita="PICKING",
            tipo="",
        )
    )
    out["ore_gestionale"] = out["ore_gestionale"].round(2)  # MINUTI con 2 decimali
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
            minuti_gestionale = tempo_totale_giornata_ore * 60 * proporzione  # Converti ore in MINUTI
            ore_gestionale_per_tipo[tipo] = minuti_gestionale
            
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
                "penalita_eccesso": 0,
                "penalita_difetto": 0,
                "tipo_attivita": "CARRELLISTI",
                "tipo": tipo,
                "ore_tim": 0.0,
                "ore_gestionale": round(minuti_gestionale, 2),  # SALVA MINUTI
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
        "penalita_eccesso": "sum",
        "penalita_difetto": "sum",
        "ore_tim": "sum",
        "ore_gestionale": "sum"
    })
    
    # Assicura che totale_colli sia int e inizializza penalità separate
    out["totale_colli"] = out["totale_colli"].astype(int)
    out["penalita_eccesso"] = out["penalita_eccesso"].astype(int)
    out["penalita_difetto"] = out["penalita_difetto"].astype(int)
    
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
                "penalita_eccesso": 0,
                "penalita_difetto": 0,
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
    out["penalita_eccesso"] = 0
    out["penalita_difetto"] = 0
    out["ore_tim"] = 0.0
    out["ore_gestionale"] = 0.0
    
    # Assicura che totale_colli e penalita siano int
    out["totale_colli"] = out["totale_colli"].astype(int)
    out["penalita_eccesso"] = out["penalita_eccesso"].astype(int)
    out["penalita_difetto"] = out["penalita_difetto"].astype(int)
    
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
    header_row = int(df.attrs.get("source_header_row", 0))  # riga intestazione usata in pandas (0-based)

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
    
    # Codice prodotto e lotto
    col_codpro = find_column(df, [["codpro"], ["cod", "prod"], ["prodotto"]], required=False)
    col_lotto = find_column(df, [["lotto"]], required=False)
    
    if col_codpro:
        logger.debug("Colonna codpro trovata: %s", col_codpro)
    else:
        logger.debug("Colonna codpro non trovata")
    
    if col_lotto:
        logger.debug("Colonna lotto trovata: %s", col_lotto)
    else:
        logger.debug("Colonna lotto non trovata")

    # Penalità: cerchiamo diff/uvc e ECCESSO/DIFETTO
    col_diff_uvc = None
    col_eccesso_difetto = None
    
    try:
        # Cerca esattamente "diff/uvc" come unica colonna
        col_diff_uvc = find_column(df, [["diff/uvc"], ["diff", "/", "uvc"]], required=False)
        if col_diff_uvc:
            norm = normalize_string(col_diff_uvc)
            if "diff" in norm and "uvc" in norm:
                logger.debug("Colonna diff/uvc trovata: %s", col_diff_uvc)
            else:
                logger.debug(
                    "Match parziale su %s per diff/uvc (norm=%s), ignoro e uso fallback",
                    col_diff_uvc,
                    norm,
                )
                col_diff_uvc = None
    except ValueError:
        pass
    
    try:
        # Cerca esattamente "ECCESSO/DIFETTO"
        col_eccesso_difetto = find_column(df, [["eccesso/difetto"], ["eccesso", "/", "difetto"]], required=False)
        if col_eccesso_difetto:
            norm = normalize_string(col_eccesso_difetto)
            if "eccesso" in norm or "difetto" in norm:
                logger.debug("Colonna ECCESSO/DIFETTO trovata: %s", col_eccesso_difetto)
            else:
                logger.debug(
                    "Match parziale su %s per ECCESSO/DIFETTO (norm=%s), ignoro",
                    col_eccesso_difetto,
                    norm,
                )
                col_eccesso_difetto = None
    except ValueError:
        pass
    
    # Cerca sempre la colonna differenze (diffe) per il calcolo penalità
    col_penalita = None
    try:
        col_penalita = find_column(df, [["diffe"], ["penal"], ["differenze"]], required=False)
        if col_penalita:
            logger.debug("Colonna differenze (diffe) trovata: %s", col_penalita)
    except ValueError:
        logger.debug("Colonna penalità non trovata, uso valore 0")
        col_penalita = None

    # Quantità: qtasped (colonna M) e uvc (colonna N)
    # n_colli = qtasped / uvc (arrotondato per eccesso, 0/0 = 0)
    col_qtasped: Optional[str] = None
    col_uvc: Optional[str] = None
    
    try:
        col_qtasped = find_column(df, [["qta", "sped"], ["qtasped"]], required=False)
        if col_qtasped:
            logger.debug("Colonna qtasped trovata: %s (posizione: %d)", col_qtasped, df.columns.get_loc(col_qtasped))
    except ValueError:
        col_qtasped = None
    
    try:
        col_uvc = find_column(df, [["uvc"]], required=False)
        if col_uvc:
            logger.debug("Colonna uvc trovata: %s (posizione: %d)", col_uvc, df.columns.get_loc(col_uvc))
    except ValueError:
        col_uvc = None
    
    # DEBUG opzionale: elenco colonne
    logger.debug("Colonne Excel indicizzate: %s", [f"[{idx}] {col}" for idx, col in enumerate(df.columns)])
    
    if not col_qtasped:
        logger.warning("Colonna qtasped non trovata, uso valore 0")
    if not col_uvc:
        logger.warning("Colonna uvc non trovata, uso valore 0")
    
    # DEBUG: Mostra i primi valori per verificare l'ordine
    if col_qtasped and col_uvc:
        logger.debug("VERIFICA COLONNE - Prime 3 righe:")
        for i in range(min(3, len(df))):
            logger.debug(f"  Riga {i}: qtasped={df[col_qtasped].iloc[i]}, uvc={df[col_uvc].iloc[i]}")
        
        # Verifica se le colonne sono invertite controllando la posizione
        qtasped_idx = df.columns.get_loc(col_qtasped)
        uvc_idx = df.columns.get_loc(col_uvc)
        logger.debug(f"Posizione colonne: qtasped={qtasped_idx} (dovrebbe essere colonna M), uvc={uvc_idx} (dovrebbe essere colonna N)")
        
        # Se qtasped è DOPO uvc, sono invertite!
        if qtasped_idx > uvc_idx:
            logger.warning("ATTENZIONE: Le colonne qtasped e uvc sembrano invertite! Scambio...")
            col_qtasped, col_uvc = col_uvc, col_qtasped
            logger.debug(f"Dopo scambio: qtasped={col_qtasped}, uvc={col_uvc}")

    subset_cols = {
        "data": col_data,
        "codice_preparatore": col_codice_preparatore,
        "tipo": col_tipo,
        "codice_picking": col_codice_picking,
    }
    if col_diff_uvc:
        subset_cols["diff_uvc"] = col_diff_uvc
    if col_eccesso_difetto:
        subset_cols["eccesso_difetto"] = col_eccesso_difetto
    if col_penalita:
        subset_cols["penalita"] = col_penalita
    if col_qtasped:
        subset_cols["qtasped"] = col_qtasped
    if col_uvc:
        subset_cols["uvc"] = col_uvc
    if col_codpro:
        subset_cols["codpro"] = col_codpro
    if col_lotto:
        subset_cols["lotto"] = col_lotto

    # CRITICO: Prima di fare il subset, salviamo gli indici originali di df
    # Creiamo una colonna temporanea in df con l'indice originale
    df_with_idx = df.copy()
    df_with_idx['_temp_original_row'] = df_with_idx.index

    sub = df_with_idx[list(subset_cols.values()) + ['_temp_original_row']].copy()
    # Rinomina colonne selezionate
    col_rename = {v: k for k, v in subset_cols.items()}
    col_rename['_temp_original_row'] = '_original_excel_row'
    sub.rename(columns=col_rename, inplace=True)
    
    def _ensure_series(frame: pd.DataFrame, column_name: str) -> Optional[pd.Series]:
        """Restituisce una Series anche se esistono colonne duplicate."""
        if column_name not in frame.columns:
            return None
        col = frame[column_name]
        if isinstance(col, pd.DataFrame):
            return col.iloc[:, 0]
        return col

    # Ora salviamo i valori originali usando gli indici originali memorizzati
    # Quando copiamo insieme alle colonne, l'allineamento è già garantito
    uvc_series = _ensure_series(sub, 'uvc')
    if uvc_series is not None:
        sub['_uvc_original'] = uvc_series.copy()

    qtasped_series = _ensure_series(sub, 'qtasped')
    if qtasped_series is not None:
        sub['_qtasped_original'] = qtasped_series.copy()
    
    # Salva valori originali della colonna diffe PRIMA di qualsiasi conversione
    penalita_series = _ensure_series(sub, 'penalita')
    diff_uvc_series_orig = _ensure_series(sub, 'diff_uvc')
    if penalita_series is not None:
        # Mantieni i valori ESATTAMENTE come sono nel file, senza conversioni
        sub['_diffe_original'] = penalita_series
        logger.debug(f"Salvati {len(sub)} valori diffe originali da colonna 'penalita' (primi 3 non-null: {penalita_series.dropna().head(3).tolist()})")
    elif diff_uvc_series_orig is not None:
        sub['_diffe_original'] = diff_uvc_series_orig
        logger.debug(f"Salvati {len(sub)} valori diffe originali da colonna 'diff_uvc'")
    else:
        sub['_diffe_original'] = 0
        logger.debug("Nessuna colonna diffe trovata, uso 0")

    # Log diagnostico per DT7 2025-08-18 nelle prime 5 righe rilevate
    debug_mask = (
        sub['codice_preparatore'].astype(str).str.upper() == 'DT7'
    ) & (sub['data'].astype(str) == '2025-08-18')
    if debug_mask.any():
        sample = sub[debug_mask].head(5)
        for ridx, r in sample.iterrows():
            logger.debug(
                "[DT7 TRACE] idx=%s qtasped=%s uvc=%s tipo=%s",
                ridx,
                r.get('_qtasped_original'),
                r.get('_uvc_original'),
                r.get('tipo'),
            )

    # Normalizzazione data
    sdate = sub["data"].astype(str).str.strip()
    sdate = sdate.str.replace(r"^(\d{4})(\d{2})(\d{2})$", r"\1-\2-\3", regex=True)
    sub["data"] = pd.to_datetime(sdate, errors="coerce").dt.date

    sub["codice_preparatore"] = sub["codice_preparatore"].astype(str).str.strip().str.upper()
    sub["tipo"] = sub["tipo"].astype(str).str.strip()
    sub.loc[sub["tipo"].str.lower() == "nan", "tipo"] = ""
    sub["nome_preparatore"] = ""

    # Mantiene il valore originale della colonna Prep come "altro codice"
    raw_prep = sub["codice_picking"].astype(str).str.strip().str.upper()
    raw_prep = raw_prep.where(raw_prep.str.lower() != "nan", "")
    sub["altro_codice"] = raw_prep

    # Normalizzazione codice PICKING (colonna Q)
    # Regola richiesta: considera SOLO il primo codice (sinistra del trattino).
    parts = raw_prep.str.split("-")
    sub["codice_picking"] = parts.str[0].str.strip().str.upper()

    # Calcola n_colli = qtasped / uvc (arrotondato per eccesso)
    # Caso 0/0 = 0
    if qtasped_series is not None and uvc_series is not None:
        qtasped_val = pd.to_numeric(qtasped_series, errors="coerce").fillna(0)
        uvc_val = pd.to_numeric(uvc_series, errors="coerce").fillna(0)
        
        # Calcola il rapporto gestendo divisione per zero
        colli_val = pd.Series(0, index=sub.index, dtype=float)
        for idx in sub.index:
            sped = qtasped_val.loc[idx]
            uvc = uvc_val.loc[idx]
            
            if sped == 0 and uvc == 0:
                # 0/0 = 0
                colli_val.loc[idx] = 0
            elif uvc == 0:
                # x/0 = 0 (per evitare divisione per zero)
                colli_val.loc[idx] = 0
            else:
                # Calcola rapporto e arrotonda per eccesso
                rapporto = sped / uvc
                colli_val.loc[idx] = np.ceil(rapporto)
    elif qtasped_series is not None:
        colli_val = pd.to_numeric(qtasped_series, errors="coerce").fillna(0)
        colli_val = np.ceil(colli_val)
    elif uvc_series is not None:
        # Se abbiamo solo uvc, usiamo il valore arrotondato per eccesso
        colli_val = pd.to_numeric(uvc_series, errors="coerce").fillna(0)
        colli_val = np.ceil(colli_val)
    else:
        colli_val = pd.Series(0, index=sub.index, dtype=float)

    # Calcola penalità usando diff/uvc ed ECCESSO/DIFETTO
    diff_uvc_series = _ensure_series(sub, "diff_uvc")
    eccesso_difetto_series = _ensure_series(sub, "eccesso_difetto")
    penalita_series = _ensure_series(sub, "penalita")

    # Usa i valori originali salvati per diff
    if '_diffe_original' in sub.columns:
        sub["diff_original"] = pd.to_numeric(sub['_diffe_original'], errors="coerce").fillna(0)
    else:
        sub["diff_original"] = 0

    # Penalità (errori) per preparatori: regola richiesta
    # Se diff è 0 -> penalità 0, altrimenti abs(diff / uvc) se uvc>0, altrimenti abs(diff)
    penalita_raw = pd.Series(0, index=sub.index, dtype=float)
    diff_val = sub["diff_original"].copy()
    
    if uvc_series is not None:
        uvc_val = pd.to_numeric(uvc_series, errors="coerce").fillna(0)
    else:
        uvc_val = pd.Series(0, index=sub.index, dtype=float)

    # Per ogni riga: se diff=0 -> penalità=0, altrimenti calcola abs(diff/uvc) o abs(diff)
    for idx in sub.index:
        d = diff_val.loc[idx]
        if d == 0:
            penalita_raw.loc[idx] = 0
        else:
            u = uvc_val.loc[idx]
            if u > 0:
                penalita_raw.loc[idx] = abs(d / u)
            else:
                penalita_raw.loc[idx] = abs(d)
    
    logger.debug("Calcolo penalità: se diff=0 -> 0, altrimenti abs(diff/uvc) con fallback abs(diff) quando uvc=0")

    sub["totale_colli"] = colli_val.astype(int)

    # Penalità: separa eccesso/difetto mantenendo valori positivi; conserva anche il valore assoluto grezzo
    penalita_int = penalita_raw.round().astype(int)
    sub["penalita_eccesso"] = penalita_int.clip(lower=0).astype(int)
    sub["penalita_difetto"] = (-penalita_int.clip(upper=0)).astype(int)
    sub["penalita_abs"] = penalita_raw.astype(float)

    # Elimina righe senza chiavi minime
    sub = sub.dropna(subset=["data", "codice_preparatore", "codice_picking"])
    sub = sub[sub["codice_picking"].astype(str).str.strip() != ""]
    
    logger.debug("Righe valide dopo pulizia: %d", len(sub))

    # Creo DataFrame per penalità PICKING - usa penalita_abs (float) invece di int
    penalita_picking = (
        sub.groupby(["data", "codice_picking"], as_index=False)[["penalita_abs"]].sum()
    )
    penalita_picking = penalita_picking.rename(
        columns={"codice_picking": "codice_preparatore", "penalita_abs": "penalita_totale"}
    )
    # Arrotonda a 2 decimali per coerenza
    penalita_picking["penalita_totale"] = penalita_picking["penalita_totale"].round(2)

    # Raggruppa per preparatori Doppia Spunta - usa penalita_abs (float) invece di int
    grouped = (
        sub.drop(columns=["codice_picking"])
        .groupby(["data", "codice_preparatore", "nome_preparatore", "tipo"], as_index=False)[
            ["totale_colli", "penalita_abs"]
        ]
        .sum()
        .assign(tipo_attivita="DOPPIA_SPUNTA")
    )
    grouped["totale_colli"] = grouped["totale_colli"].astype(int)
    # Arrotonda penalità a 2 decimali e rinomina
    grouped["penalita_totale"] = grouped["penalita_abs"].round(2)
    grouped = grouped.drop(columns=["penalita_abs"])
    
    # Crea penalita_eccesso e penalita_difetto a 0 per compatibilità con il database
    grouped["penalita_eccesso"] = 0
    grouped["penalita_difetto"] = 0

    # Riorganizzo le colonne per coerenza con il resto dei parser
    grouped = grouped[[
        "data",
        "codice_preparatore",
        "nome_preparatore",
        "totale_colli",
        "penalita_eccesso",
        "penalita_difetto",
        "tipo_attivita",
        "tipo",
    ]]

    penalita_picking["tipo_attivita"] = "PICKING"
    
    # Prepara dettagli per sessioni_doppia_spunta
    sessioni_dettaglio = []
    for (data, codice), group in sub.groupby(['data', 'codice_preparatore']):
        # Ordina per indice originale per mantenere l'ordine del file
        group = group.sort_index()
        
        # Calcola totale giornaliero per questo preparatore
        totale_colli_giornata = int(group['totale_colli'].sum())
        
        # Per ora non abbiamo orari nella doppia spunta, quindi li lasciamo NULL
        # Se in futuro li aggiungiamo, possiamo calcolare ore_gestionale
        ore_gestionale_giornaliere = None
        
        for idx, (row_idx, row) in enumerate(group.iterrows(), start=1):
            # Solo la prima riga contiene i totali giornalieri
            # Recupera i valori originali di uvc e qtasped PRIMA della conversione numerica
            uvc_val = None
            qtasped_val = None
            altro_codice_val = None
            
            # Usa i valori originali dalla riga corrente del gruppo
            if '_uvc_original' in row.index:
                uvc_raw = row['_uvc_original']
                if pd.notna(uvc_raw):
                    try:
                        uvc_val = int(float(uvc_raw))
                    except:
                        uvc_val = None
            
            if '_qtasped_original' in row.index:
                qtasped_raw = row['_qtasped_original']
                if pd.notna(qtasped_raw):
                    try:
                        qtasped_val = int(float(qtasped_raw))
                    except:
                        qtasped_val = None

            excel_row_db = None
            if pd.notna(row_idx):
                try:
                    # Formula: indice pandas + header_row (0-based) + 2 (per passare alla riga Excel 1-based)
                    excel_row_db = int(row_idx) + header_row + 2
                except Exception:
                    excel_row_db = None

            if codice.upper() == 'DT7' and str(data) == '2025-08-18' and idx in {225, 230, 232}:
                logger.debug(
                    "[DT7 TRACE] data=%s idx=%s excel_row=%s qtasped=%s uvc=%s codpro=%s lotto=%s",
                    data,
                    idx,
                    excel_row_db,
                    qtasped_val,
                    uvc_val,
                    row.get('codpro'),
                    row.get('lotto'),
                )
            # Recupera codpro e lotto se disponibili
            codpro_val = None
            lotto_val = None
            diff_val = None
            
            if 'altro_codice' in row.index:
                altro_raw = row['altro_codice']
                if pd.notna(altro_raw):
                    altro_codice_val = str(altro_raw).strip() or None
            
            # Recupera il valore ORIGINALE della colonna diffe dal dataframe (senza conversioni)
            if '_diffe_original' in row.index:
                diff_raw = row['_diffe_original']
                # Accetta qualsiasi valore non-NaN, anche 0
                if pd.notna(diff_raw):
                    try:
                        # Converti a float, accettando anche 0
                        val = float(diff_raw)
                        diff_val = val
                    except (ValueError, TypeError) as e:
                        # Se fallisce la conversione, salva None
                        diff_val = None
                        if idx <= 5:  # Log solo per le prime 5 righe problematiche
                            logger.debug(f"Impossibile convertire diff_raw={diff_raw} (tipo={type(diff_raw).__name__}): {e}")
                else:
                    diff_val = None
            
            if 'codpro' in row.index:
                codpro_raw = row['codpro']
                if pd.notna(codpro_raw):
                    codpro_val = str(codpro_raw).strip()
            
            if 'lotto' in row.index:
                lotto_raw = row['lotto']
                if pd.notna(lotto_raw):
                    lotto_val = str(lotto_raw).strip()

            penalita_val = None
            if 'penalita_abs' in row.index:
                try:
                    val = row['penalita_abs']
                    if pd.notna(val):
                        # Mantieni il valore float SENZA arrotondamenti
                        penalita_val = float(val)
                        if idx <= 5:  # Debug per le prime 5 righe
                            logger.debug(f"Riga {idx}: penalita_abs={val} -> penalita_val={penalita_val}")
                except Exception as e:
                    penalita_val = None
                    logger.debug(f"Errore conversione penalita riga {idx}: {e}")

            sessioni_dettaglio.append({
                'data': data,
                'codice_preparatore': codice,
                'numero_riga': idx,
                'excel_row': excel_row_db,
                'ora_inizio': None,
                'ora_fine': None,
                'tempo_minuti': None,
                'n_colli': int(row['totale_colli']),
                'uvc': uvc_val,
                'qtasped': qtasped_val,
                'codpro': codpro_val,
                'lotto': lotto_val,
                'tipo': row['tipo'] if row['tipo'] else None,
                'diff': diff_val,
                'altro_codice': altro_codice_val,
                'penalita': penalita_val,
                'penalita_eccesso': int(row['penalita_eccesso']),
                'penalita_difetto': int(row['penalita_difetto']),
                'totale_colli_giornata': totale_colli_giornata if idx == 1 else None,
                'ore_gestionale_giornaliere': ore_gestionale_giornaliere if idx == 1 else None
            })
    
    # Salva i dettagli nel database
    if sessioni_dettaglio:
        try:
            print(f"[DEBUG] sessioni_doppia_spunta da salvare: {len(sessioni_dettaglio)} (prime 2): {sessioni_dettaglio[:2]}")
        except Exception:
            pass
        try:
            from database import save_sessioni_doppia_spunta
            save_sessioni_doppia_spunta(sessioni_dettaglio)
        except Exception as e:
            import traceback
            print(f"[ERROR] Errore nel salvataggio sessioni doppia spunta: {e}")
            print(traceback.format_exc())

    # RECUPERA penalita_totale dalla tabella sessioni_doppia_spunta
    # Per i record DOPPIA_SPUNTA usa i valori dal database
    # Per i record PICKING lascia i valori già calcolati (sono basati su codice_picking, non salvato nelle sessioni)
    try:
        import mysql.connector
        from config import MYSQL_CONFIG
        from contextlib import closing
        
        conn = mysql.connector.connect(**MYSQL_CONFIG)
        cursor = conn.cursor(dictionary=True)
        
        # Ottieni le date uniche dai dati raggruppati
        date_uniche = grouped['data'].unique()
        date_str = ', '.join([f"'{d}'" for d in date_uniche])
        
        # Query compatibile: usa le date specifiche invece di subquery con LIMIT
        query = f"""
            SELECT data, codice_preparatore, SUM(penalita) as penalita_totale
            FROM sessioni_doppia_spunta
            WHERE data IN ({date_str})
            GROUP BY data, codice_preparatore
        """
        cursor.execute(query)
        penalita_db = cursor.fetchall()
        cursor.close()
        conn.close()
        
        # Crea un dizionario per mappare (data, codice) -> penalita_totale
        penalita_map = {}
        for row in penalita_db:
            key = (str(row['data']), str(row['codice_preparatore']))
            penalita_map[key] = float(row['penalita_totale'] or 0)
        
        # Aggiorna SOLO grouped (DOPPIA_SPUNTA) con i valori reali dal database
        grouped['penalita_totale'] = grouped.apply(
            lambda row: penalita_map.get((str(row['data']), str(row['codice_preparatore'])), 0.0),
            axis=1
        )
        
        # Per penalita_picking, lascia i valori calcolati da penalita_abs
        # perché sono basati su codice_picking che non è disponibile nelle sessioni
        logger.info(f"Penalità DOPPIA_SPUNTA recuperate dal DB per {len(penalita_map)} combinazioni")
        logger.info(f"Penalità PICKING usano valori calcolati ({len(penalita_picking)} record)")
        
        # Log di debug
        for key, val in list(penalita_map.items())[:3]:
            logger.info(f"  DS: {key[0]} - {key[1]}: penalità={val}")
        for idx, row in penalita_picking.head(3).iterrows():
            logger.info(f"  PICK: {row['data']} - {row['codice_preparatore']}: penalità={row['penalita_totale']}")
            
    except Exception as e:
        logger.warning(f"Impossibile recuperare penalità dal DB, uso valori calcolati: {e}")
        import traceback
        logger.warning(traceback.format_exc())

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
            "penalita_eccesso": 0,
            "penalita_difetto": 0,
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
    )[["totale_colli", "penalita_eccesso", "penalita_difetto"]].sum()

    return df_grouped