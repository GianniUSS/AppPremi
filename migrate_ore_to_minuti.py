"""
Script di migrazione: converte ore_tim e ore_gestionale da ORE a MINUTI
IMPORTANTE: Esegui questo script UNA SOLA VOLTA prima di usare la nuova versione!
"""
import mysql.connector
from config import MYSQL_CONFIG

def migrate_ore_to_minuti():
    """Converte tutti i valori ore_tim e ore_gestionale da ore a minuti"""
    conn = mysql.connector.connect(**MYSQL_CONFIG)
    cur = conn.cursor()
    
    try:
        print("🔄 Migrazione ORE → MINUTI")
        print("=" * 60)
        
        # 1. Converti dati_produzione
        print("\n📊 Conversione tabella dati_produzione...")
        cur.execute("""
            UPDATE dati_produzione 
            SET ore_tim = ore_tim * 60,
                ore_gestionale = ore_gestionale * 60
            WHERE ore_tim < 100 OR ore_gestionale < 100
        """)
        updated_prod = cur.rowcount
        print(f"   ✅ Aggiornati {updated_prod} record")
        
        # 2. Converti anomalie
        print("\n⚠️  Conversione tabella anomalie...")
        cur.execute("""
            UPDATE anomalie 
            SET ore_tim = ore_tim * 60
            WHERE ore_tim IS NOT NULL AND ore_tim < 100
        """)
        updated_anom = cur.rowcount
        print(f"   ✅ Aggiornati {updated_anom} record")
        
        # 3. Converti sessioni_carrellisti (ore_gestionale per tipo)
        print("\n🚚 Conversione tabella sessioni_carrellisti...")
        cur.execute("""
            UPDATE sessioni_carrellisti 
            SET ore_gestionale_st = ore_gestionale_st * 60,
                ore_gestionale_ss = ore_gestionale_ss * 60,
                ore_gestionale_ap = ore_gestionale_ap * 60,
                ore_gestionale_cm = ore_gestionale_cm * 60
            WHERE (ore_gestionale_st IS NOT NULL AND ore_gestionale_st < 100)
               OR (ore_gestionale_ss IS NOT NULL AND ore_gestionale_ss < 100)
               OR (ore_gestionale_ap IS NOT NULL AND ore_gestionale_ap < 100)
               OR (ore_gestionale_cm IS NOT NULL AND ore_gestionale_cm < 100)
        """)
        updated_sess_car = cur.rowcount
        print(f"   ✅ Aggiornati {updated_sess_car} record")
        
        # 4. Converti sessioni_doppia_spunta
        print("\n✅ Conversione tabella sessioni_doppia_spunta...")
        cur.execute("""
            UPDATE sessioni_doppia_spunta 
            SET ore_gestionale_giornaliere = ore_gestionale_giornaliere * 60
            WHERE ore_gestionale_giornaliere IS NOT NULL 
              AND ore_gestionale_giornaliere < 100
        """)
        updated_sess_ds = cur.rowcount
        print(f"   ✅ Aggiornati {updated_sess_ds} record")
        
        conn.commit()
        
        print("\n" + "=" * 60)
        print("✅ MIGRAZIONE COMPLETATA CON SUCCESSO!")
        print(f"   • dati_produzione: {updated_prod} record")
        print(f"   • anomalie: {updated_anom} record")
        print(f"   • sessioni_carrellisti: {updated_sess_car} record")
        print(f"   • sessioni_doppia_spunta: {updated_sess_ds} record")
        print("\n⚠️  IMPORTANTE: Non eseguire questo script nuovamente!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ ERRORE: {e}")
        conn.rollback()
        raise
    finally:
        cur.close()
        conn.close()

if __name__ == "__main__":
    risposta = input("⚠️  Questa migrazione convertirà TUTTI i dati da ore a minuti.\n"
                    "   Sei sicuro di voler procedere? (SI/no): ")
    if risposta.upper() == "SI":
        migrate_ore_to_minuti()
    else:
        print("❌ Migrazione annullata")
