"""
Funzioni utility per il parsing e la normalizzazione dei dati.
"""
from typing import List, Optional
import pandas as pd


def normalize_string(s: str) -> str:
    """Normalizza una stringa per confronti più tolleranti."""
    return str(s).lower().replace("°", "").strip()


def find_column(df: pd.DataFrame, candidates: List[List[str]], required: bool = True) -> Optional[str]:
    """
    Trova una colonna il cui nome contiene tutte le parole di almeno una combinazione.

    Args:
        df: DataFrame pandas
        candidates: lista di alternative; ciascuna alternativa è una lista di token richiesti.
                   Esempio: [["data", "inizio"], ["data"]]
        required: se True, solleva eccezione se non trova la colonna

    Returns:
        Nome della colonna trovata o None se non richiesta

    Raises:
        ValueError: se required=True e nessuna colonna è trovata
    """
    names = {c: normalize_string(c) for c in df.columns}
    
    # Prima prova: ricerca esatta con tutte le parole
    for col, norm in names.items():
        for option in candidates:
            if all(normalize_string(part) in norm for part in option):
                return col
    
    # Seconda prova: ricerca parziale (almeno una parola chiave)
    for col, norm in names.items():
        for option in candidates:
            for part in option:
                if normalize_string(part) in norm:
                    print(f"⚠️ Trovata colonna parziale '{col}' per '{part}'")
                    return col
    
    if required:
        available_cols = list(df.columns)
        need = " | ".join([" + ".join(op) for op in candidates])
        raise ValueError(f"Colonna non trovata (attesa una contenente: {need}).\nColonne disponibili: {available_cols}")
    
    return None


def safe_int_conversion(value, default: int = 0) -> int:
    """Converte un valore in intero gestendo errori."""
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default


def prepare_dataframe_for_db(df: pd.DataFrame) -> List[tuple]:
    """
    Prepara un DataFrame per l'inserimento nel database.
    
    Args:
        df: DataFrame con le colonne standard
        
    Returns:
        Lista di tuple pronte per l'inserimento batch
    """
    if df.empty:
        return []
        
    values = []
    for idx, r in df.iterrows():
        # Gestisci ore_tim e ore_gestionale se presenti
        ore_tim = 0.0
        ore_gestionale = 0.0
        
        if "ore_tim" in r and pd.notna(r["ore_tim"]):
            ore_tim = float(r["ore_tim"])
        
        if "ore_gestionale" in r and pd.notna(r["ore_gestionale"]):
            ore_gestionale = float(r["ore_gestionale"])
        
        # Converti esplicitamente totale_colli in int
        try:
            totale_colli = int(r["totale_colli"]) if pd.notna(r["totale_colli"]) else 0
        except (ValueError, OverflowError) as e:
            print(f"❌ ERRORE conversione totale_colli alla riga {idx}:")
            print(f"   Valore: {r['totale_colli']} (tipo: {type(r['totale_colli'])})")
            print(f"   Record completo: {r.to_dict()}")
            print(f"   Errore: {e}")
            totale_colli = 0
        
        try:
            penalita = int(r.get("penalita", 0))
        except (ValueError, OverflowError) as e:
            print(f"❌ ERRORE conversione penalita alla riga {idx}: {e}")
            penalita = 0
        
        tupla = (
            r["data"],
            str(r["codice_preparatore"]),
            str(r["nome_preparatore"]) if pd.notna(r["nome_preparatore"]) else None,
            totale_colli,
            penalita,
            r["tipo_attivita"],
            str(r.get("tipo", "")) if pd.notna(r.get("tipo", "")) else "",
            ore_tim,
            ore_gestionale,
        )
        
        # Log della tupla
        if idx < 5:  # Log solo le prime 5
            print(f"  Tupla {idx}: data={tupla[0]}, codice={tupla[1]}, colli={tupla[3]} ({type(tupla[3])}), ore_gest={tupla[8]}")
        
        values.append(tupla)
    
    print(f"\n✅ Totale tuple preparate: {len(values)}")
    return values