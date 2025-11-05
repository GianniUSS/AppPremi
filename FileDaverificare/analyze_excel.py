import pandas as pd
import sys

def analyze_excel(file_path):
    print(f"🔎 Analisi file: {file_path}\n")

    # Carico senza header → tutte le righe diventano dati
    df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

    print("Prime 10 righe del file (senza header):")
    print(df_raw.head(10))
    print("\n---\n")

    # Carico provando con le prime 5 righe come intestazione
    for i in range(5):
        try:
            df = pd.read_excel(file_path, sheet_name=0, header=i)
            print(f"Header ipotizzato alla riga {i}:")
            print(df.head(3))  # prime 3 righe di dati
            print("Colonne trovate:", list(df.columns))
            print("\n---\n")
        except Exception as e:
            print(f"Errore caricando con header={i}: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("❌ Devi passare il file Excel da analizzare.\nEsempio:\n  python analyze_excel.py file.xlsx")
        sys.exit(1)

    analyze_excel(sys.argv[1])
