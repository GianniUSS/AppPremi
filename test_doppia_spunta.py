"""
Script di test per verificare l'importazione della doppia spunta.
"""
import sys
from pathlib import Path

def test_doppia_spunta(file_path):
    """Test del parser doppia spunta."""
    try:
        from parsers import parse_doppia_spunta
        
        print("=" * 60)
        print("TEST IMPORTAZIONE DOPPIA SPUNTA")
        print("=" * 60)
        print(f"\nFile: {file_path}")
        print("\nInizio parsing...")
        
        result = parse_doppia_spunta(file_path)
        
        print("\n" + "=" * 60)
        print("RISULTATI")
        print("=" * 60)
        
        print(f"\n✓ Records per database: {len(result.records)}")
        print(f"✓ Penalità PICKING: {len(result.penalita_picking)}")
        
        if not result.records.empty:
            print("\n--- Prime 5 righe Records ---")
            print(result.records.head())
            print(f"\nColonne: {list(result.records.columns)}")
        
        if not result.penalita_picking.empty:
            print("\n--- Prime 5 righe Penalità PICKING ---")
            print(result.penalita_picking.head())
            print(f"\nColonne: {list(result.penalita_picking.columns)}")
        
        print("\n" + "=" * 60)
        print("✓ TEST COMPLETATO CON SUCCESSO")
        print("=" * 60)
        
    except Exception as e:
        print("\n" + "=" * 60)
        print("❌ ERRORE DURANTE IL TEST")
        print("=" * 60)
        print(f"\nErrore: {type(e).__name__}")
        print(f"Messaggio: {str(e)}")
        import traceback
        print("\n--- Stack Trace ---")
        traceback.print_exc()
        print("=" * 60)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python test_doppia_spunta.py <percorso_file_excel>")
        print("\nOppure trascina un file Excel su questo script.")
        input("\nPremi INVIO per uscire...")
    else:
        test_doppia_spunta(sys.argv[1])
        input("\nPremi INVIO per uscire...")
