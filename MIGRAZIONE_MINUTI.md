# 🔄 MIGRAZIONE: ORE → MINUTI

## ✅ Modifiche Completate

### 1. Database
- ✅ Aggiornati commenti schema per indicare che `ore_tim` e `ore_gestionale` contengono **MINUTI**
- ✅ Struttura tabelle invariata (DECIMAL(10,2) supporta minuti con decimali)

### 2. Import/Sync
- ✅ **Sync TIM**: Salva minuti direttamente (no conversione /60)
- ✅ **Parser Preparatori**: Converte ore → minuti (* 60)
- ✅ **Parser Carrellisti**: Converte ore → minuti (* 60)  
- ✅ **Anomalie**: Salvano valori in minuti

### 3. Visualizzazione
- ✅ **data_viewer.py**: Converte minuti → ore (/60) nel display
- ✅ **anomalie_view.py**: Converte minuti → ore (/60) nel display
- ✅ **Totali/statistiche**: Somma in minuti, converte solo per display

## 🚀 PROCEDURA DI AGGIORNAMENTO

### Passo 1: BACKUP
```bash
# OBBLIGATORIO! Backup del database prima di procedere
mysqldump -h 172.16.202.141 -u tim_root -p tim_import > backup_pre_migrazione.sql
```

### Passo 2: MIGRAZIONE DATI
```powershell
# Esegui lo script di migrazione (converte dati esistenti da ore a minuti)
python migrate_ore_to_minuti.py
```

**ATTENZIONE**: Questo script:
- Moltiplica per 60 tutti i valori `ore_tim` e `ore_gestionale` < 100
- Aggiorna `dati_produzione`, `anomalie`, `sessioni_carrellisti`, `sessioni_doppia_spunta`
- Va eseguito **UNA SOLA VOLTA**!

### Passo 3: TEST
1. Apri l'applicazione
2. Visualizza dati esistenti → devono mostrare ORE corrette (es. 8.00h)
3. Fai un piccolo import di prova
4. Verifica che le somme siano precise (no arrotondamenti)

### Passo 4: SYNC COMPLETO
- Esegui sync TIM con un mese completo
- Verifica che le somme tornino esatte (480 min = 8.00h)

## 📊 ESEMPIO DI VERIFICA

### Prima (ORE con arrotondamenti)
```
Record 1: 2.67h (160.2 min)
Record 2: 2.67h (160.2 min)  
Record 3: 2.66h (159.6 min)
TOTALE: 8.00h (480.0 min) ✅ MA SOMMA = 8.00h da 2.67+2.67+2.66 = 8.00h per fortuna!
```

### Dopo (MINUTI precisi)
```
Record 1: 160.20 min → display: 2.67h
Record 2: 160.20 min → display: 2.67h
Record 3: 159.60 min → display: 2.66h
TOTALE: 480.00 min → display: 8.00h ✅ SEMPRE ESATTO!
```

## ⚠️ NOTE IMPORTANTI

1. **Non eseguire** `migrate_ore_to_minuti.py` più di una volta!
2. **Tutti i nuovi import** salveranno automaticamente in minuti
3. **Le view** convertono automaticamente per mostrare ore
4. **I calcoli interni** lavorano sempre in minuti (massima precisione)

## 🐛 TROUBLESHOOTING

### Problema: Valori troppo grandi dopo migrazione
**Causa**: Script eseguito più volte
**Soluzione**: Ripristina da backup e riesegui UNA volta

### Problema: Export mostra minuti invece di ore
**Causa**: Export non aggiornato
**Soluzione**: Verifica che l'export divida per 60

### Problema: Somme non tornano
**Causa**: Mix di dati vecchi (ore) e nuovi (minuti)
**Soluzione**: Riparti da backup pulito ed esegui migrazione completa

## ✅ VANTAGGI OTTENUTI

- **Precisione assoluta**: No arrotondamenti, no centesimi persi
- **Somme sempre corrette**: 480.00 min = esattamente 8.00h
- **Compatibilità TIM**: Stesso formato (minuti)
- **Proporzioni esatte**: Distribuzione senza perdite
- **Performance**: Calcoli integer più veloci

## 📝 CHECKLIST

- [ ] Backup database completato
- [ ] Script migrazione eseguito (UNA volta)
- [ ] Test visualizzazione dati esistenti
- [ ] Test import nuovo file
- [ ] Verifica somme/totali
- [ ] Sync TIM completo
- [ ] Export report verificato
- [ ] Premi calcolati correttamente
