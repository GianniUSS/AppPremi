# Analisi Confronto Excel vs App - AppEden

## Struttura Excel (agosto.xlsx)

17 fogli: dati grezzi (LOG), timbrature (TIMBR_*), calcolo premi (PREMIO *), riepilogo (Totale).

---

## 1. PREMIO PIK (PICKING)

### Formule Excel (foglio "PREMIO PIK")
- **Ore**: `F4 = SUMIFS(TIMBR_COL!D:D, TIMBR_COL!$A:$A, $A4)` — ore dal foglio timbrature (gia in ORE)
- **Colli**: `G4 = SUMIFS(TIMBR_COL!G:G, TIMBR_COL!$A:$A, $A4)` — somma colli per codice
- **Colli/h**: `H4 = G4/F4` — divisione diretta (no conversione minuti)
- **Premio base**: `I4 = IFERROR(IF(H4>=$Q$19, $R$19, VLOOKUP(FLOOR.MATH(H4,5), $Q$11:$R$19, 2, FALSE)) * G4, "-")`
  - Se colli/h < 100 → "-" (nessun premio)
  - Il FLOOR.MATH(H4,5) arrotonda PER DIFETTO al multiplo di 5
  - Cerca nella tabella fasce e moltiplica per colli totali
- **Penalita DS** (colonne K-N):
  - `K4 = SUMIFS('ERRORI DS'!$B:$B, 'ERRORI DS'!$A:$A, A4)` — colli controllati
  - `L4 = SUMIFS('ERRORI DS'!C:C, 'ERRORI DS'!A:A, A4)` — errori DS
  - `M4 = L4/K4` — % errori
  - `N4 = IF(I4<>"-", IF(M4<=$N$1, 0, (-L4+ROUND(K4*$N$1,0))*$N$2), "-")`
    - Soglia errori $N$1 = 0.025% = 0.00025
    - Penalita per errore $N$2 = 1 EUR
    - Se %errori <= soglia → 0 penalita
    - Altrimenti: penalita = (-errori + arrotondamento(colli_controllati * soglia)) * 1€
- **KPI Bonus**:
  - `J1 = 0.15` (15%)
  - `J2 = IF(T6="NO","NO",IF(T7="NO","NO","OK"))` — serve che ENTRAMBE le soglie siano OK
  - `T6 = IF(R6<=S6,"OK","NO")` — R6=importo_rotture(2549.87), S6=soglia_rotture(2500) → NO
  - `T7 = IF(R7<=S7,"OK","NO")` — R7=importo_differenze(1221.46), S7=soglia_diff(2000) → OK
  - `J4 = IF(I4="-", "-", IF($J$2="NO", 0, IF($J$2="OK", I4*$J$1, "ERR")))` — 15% del premio_base
- **Premio totale**: `O4 = IF(SUM(I4:J4,N4)<=0, "-", SUM(I4:J4,N4))` — somma premio_base + kpi + penalita_ds

### Fasce Excel PICKING
| Soglia (Colli/h) | €/Collo |
|---|---|
| 100 | 0.00700 |
| 105 | 0.00711 |
| 110 | 0.00722 |
| 115 | 0.00733 |
| 120 | 0.00744 |
| 125 | 0.00755 |
| 130 | 0.00766 |
| 135 | 0.00777 |
| 140 | 0.00789 |

Crescita: +1.5% per fascia (R9=0.015), step: 5 colli/h (Q9=5)

### Differenze App vs Excel — PICKING

#### D1: PENALITA DOPPIA SPUNTA — LOGICA COMPLETAMENTE DIVERSA
- **Excel**: La penalita PIK viene dalla % errori DS del singolo preparatore. Formula: se %errori > 0.025%, penalita = (-errori + round(colli_controllati * 0.00025)) * 1€. E una penalita per errore di doppia spunta, NON viene dalla tabella penalita_totale.
- **App**: Usa `penalita_totale` dal DB (campo `dati_produzione.penalita_totale`), che e la somma di penalita_eccesso + penalita_difetto dalla produzione. NON calcola penalita dalla % errori DS.
- **Impatto**: La penalita nell'app e completamente diversa da quella dell'Excel. Nell'Excel e legata agli errori di doppia spunta, nell'app e legata alle penalita di produzione (eccesso/difetto colli).

#### D2: KPI BONUS — FORMULA CORRETTA MA APPLICAZIONE DIVERSA
- **Excel**: `J4 = IF(I4="-", "-", IF($J$2="NO", 0, IF($J$2="OK", I4*$J$1, "ERR")))` — il 15% si applica al **premio_base** (I4), NON al premio_netto (dopo penalita)
- **App**: `premio_kpi = (premio_netto * bonus_perc)` — il 15% si applica al **premio_netto** (dopo sottrazione penalita)
- **Impatto**: Quando ci sono penalita, il bonus KPI sara diverso. Nell'Excel il bonus e sul lordo, nell'app sul netto.

#### D3: PREMIO TOTALE — FORMULA DIVERSA
- **Excel**: `O4 = SUM(I4, J4, N4)` se >0, altrimenti "-" — somma premio_base + bonus_kpi + penalita_ds (N4 e negativa)
- **App**: `premio_totale = premio_netto + premio_kpi` dove `premio_netto = premio_base - penalita` (floor a 0)
- **Impatto**: Nell'Excel le penalita vengono SOMMATE (sono negative), nell'app vengono SOTTRATTE e poi il netto viene portato a 0 se negativo. Il risultato e simile ma il floor a 0 del premio_netto nell'app e diverso dal check <=0 del totale nell'Excel.

#### D4: LOOKUP FASCE — FLOOR vs ITERATE
- **Excel**: `VLOOKUP(FLOOR.MATH(H4,5), tabella_fasce)` — arrotonda per difetto al multiplo di 5 prima del lookup
- **App**: Itera tutte le fasce e prende l'ultima per cui `colli_ora >= soglia` (crescente)
- **Impatto**: Equivalente nel risultato (entrambi trovano la fascia giusta), ma l'arrotondamento a .01 nell'app potrebbe creare casi limite diversi.

#### D5: ORE — POSSIBILE DIFFERENZA DI CONVERSIONE
- **Excel**: Le ore nel foglio TIMBR_COL colonna D sono GIA in ore (valori tipo 126.48, 90.83 ecc.)
- **App**: Il DB salva ore_tim in MINUTI e converte `ore_tim / 60`
- **Impatto**: Se i dati vengono importati correttamente (convertiti in minuti durante l'import), nessun problema. Ma bisogna verificare il parser di importazione.

---

## 2. PREMIO CAR (CARRELLISTI)

### Formule Excel (foglio "PREMIO CAR")
- **Ore**: `E4 = SUMIFS(TIMBR_CAR!$D:$D, TIMBR_CAR!$A:$A, $A4)` — ore dal foglio timbrature
- **Movimenti**: `F4 = SUMIFS(TIMBR_CAR!$G:$G, TIMBR_CAR!$A:$A, $A4)` — movimenti PESATI dal foglio timbrature
- **Mov/h**: `G4 = F4/E4`
- **Premio base**: `H4 = IFERROR(IF(G4>=$L$15, $M$15, VLOOKUP(FLOOR.MATH(G4,2), $L$11:$M$15, 2, FALSE))*F4, "-")`
  - FLOOR.MATH al multiplo di 2 (step fasce = 2)
- **KPI**: Stessa logica del PIK, con soglie condivise (M6='PREMIO PIK'!R6, M7='PREMIO PIK'!R7)
- **Premio totale**: `J4 = IF(SUM(H4:I4)<=0, "-", SUM(H4:I4))` — NO penalita DS per carrellisti

### Fasce Excel CARRELLISTI
| Soglia (Mov/h) | €/Plt |
|---|---|
| 18 | 0.03000 |
| 20 | 0.03630 |
| 22 | 0.04392 |
| 24 | 0.05314 |
| 26 | 0.06430 |

Crescita: +21% per fascia (M9=0.21), step: 2 mov/h (L9=2)

### Differenze App vs Excel — CARRELLISTI

#### D6: MOVIMENTI GIA PESATI NELL'EXCEL vs PESO NELL'APP
- **Excel**: La colonna F (MOV) nel foglio PREMIO CAR prende i dati da TIMBR_CAR!$G che contiene gia i movimenti pesati (la colonna G di TIMBR_CAR e "MOV pes" = movimenti * peso)
- **App**: L'app applica i pesi runtime: `movimenti_pesati = colli * peso` dove il peso viene dalla tabella `peso_movimenti`
- **Impatto**: Dovrebbe dare lo stesso risultato SE i pesi sono allineati. Ma nell'Excel il peso e applicato a livello di riga giornaliera nel foglio timbrature, mentre nell'app e applicato per tipo (ST, SS, CM, AP) a livello aggregato per mese.

#### D7: NESSUNA PENALITA PER CARRELLISTI
- **Excel**: `J4 = IF(SUM(H4:I4)<=0, "-", SUM(H4:I4))` — nessuna colonna penalita
- **App**: `premio_totale = premio_base + premio_kpi` (nessuna penalita) — **COERENTE**

#### D8: ORE UTILIZZATE
- **Excel**: Usa le ore dal foglio timbrature (TIMBR_CAR!D = ore)
- **App**: Usa `ore_tim` dal DB (convertito da minuti) — la query carica `ore_tim` non `ore_gestionale`
- **Nota**: L'Excel sembra usare ore gestionali (timbrature), l'app query su `ore_tim`. Nella prima analisi del codice l'app usa `ore_tim` per carrellisti. Verificare se i dati gestionali e TIM coincidono.

---

## 3. PREMIO DS (DOPPIA SPUNTA)

### Formule Excel (foglio "PREMIO DS")
- **Errori difetto**: `E4 = SUMIF(TIMBR_DS!$B:$B, B4, TIMBR_DS!$I:$I)` — errori in difetto
- **Ore DS**: `F4 = SUMIFS(TIMBR_DS!$E:$E, TIMBR_DS!$A:$A, $A4)` — ore DS
- **Spunte**: `G4 = SUMIFS(TIMBR_DS!$H:$H, TIMBR_DS!$A:$A, $A4)` — colli controllati (spunte)
- **Colli/h**: `H4 = G4/(F4-(E4*$E$1))` dove **$E$1 = 5/60 = 0.08333h**
  - IMPORTANTE: le ore vengono COMPENSATE sottraendo 5 minuti per ogni errore in difetto
- **Premio base**: `I4 = IFERROR(IF(H4>=$L$15, $M$15, IF(AND(H4<$L$15,H4>=$L$14), $M$14, ...)) * G4, "-")`
  - Usa IF annidati anziche VLOOKUP (ma stessa logica)
- **Premio totale**: `J4 = IF(I4<=0, "-", I4)` — nessun KPI bonus per DS nell'Excel!

### Fasce Excel DOPPIA SPUNTA
| Soglia (Colli/h) | €/collo |
|---|---|
| 147 | 0.00500 |
| 160 | 0.00525 |
| 173 | 0.00551 |
| 186 | 0.00579 |
| 199 | 0.00608 |

Crescita: +5% per fascia (M9=0.05), step: 13 colli/h (L9=ROUND(100/7.5,0)=13)

### Differenze App vs Excel — DOPPIA SPUNTA

#### D9: COMPENSAZIONE ORE — FORMULA DIVERSA
- **Excel**: `H4 = G4/(F4-(E4*$E$1))` — sottrae 5 minuti (0.0833h) PER OGNI ERRORE IN DIFETTO dalle ore totali
- **App**: `ore_effettive_comp = ore_effettive - (penalita_difetto * minuti_comp / 60)` — sottrae `penalita_difetto * minuti_compensazione / 60`
- **Impatto**: Nell'Excel E4 e il CONTEGGIO degli errori in difetto, nell'app `penalita_difetto` e un IMPORTO in euro. Se penalita_difetto != conteggio errori, i risultati saranno diversi. PERO l'app ha anche `errori_difetto = SUM(CASE WHEN dp.penalita_difetto > 0 THEN 1 ELSE 0 END)` che e un conteggio di GIORNI con errori, non di errori singoli.

#### D10: NESSUN KPI BONUS PER DS NELL'EXCEL
- **Excel**: `J4 = IF(I4<=0, "-", I4)` — premio totale = premio base, NESSUN bonus KPI
- **App**: L'app calcola e applica il bonus KPI anche per DS (15% se soglie rispettate)
- **Impatto**: L'app potrebbe dare premi piu alti per DS rispetto all'Excel.

#### D11: NUOVE APERTURE — SOLO NELL'APP
- **Excel**: Non c'e logica di esclusione colli per nuove aperture
- **App**: Esclude i colli relativi a nuove aperture dalla query DS
- **Impatto**: L'app potrebbe dare meno colli validi rispetto all'Excel.

---

## 4. PREMIO RICEV (RICEVITORI)

### Formule Excel (foglio "PREMIO RICEV")
- **Ore**: `E4 = SUMIFS(TIMBR_RICEV!$E:$E, TIMBR_RICEV!$A:$A, $A4)` — ore ricevimento
- **Giorni**: `F4 = E4/$F$2` dove **$F$2 = 8** (ore/giorno)
- **Media Plt/h**: `K4 = K2/K3` dove K2=totale_pallet(10629), K3=ore_ricevimento(655.36) → **media DI SQUADRA** (16.22 plt/h)
- **Premio base**: `G4 = IFERROR(IF($K$4>=$J$14, $K$14, VLOOKUP(FLOOR.MATH($K$4,2), $J$11:$K$14, 2, FALSE)) * F4, "-")`
  - Usa la media DI SQUADRA ($K$4) per determinare la fascia
  - Moltiplica il premio per i GIORNI INDIVIDUALI (F4)
- **Premio totale**: `H4 = IF(G4<=0, "-", G4)` — premio = premio base (no KPI nell'Excel)

### Fasce Excel RICEVITORI
| Soglia (Plt/h) | €/gg |
|---|---|
| 18 | 2.50 |
| 20 | 3.60 |
| 22 | 5.184 |
| 24 | 7.465 |

Crescita: +44% per fascia (K9=0.44), step: 2 plt/h (J9=2)

### Differenze App vs Excel — RICEVITORI

#### D12: MEDIA DI SQUADRA vs INDIVIDUALE — COERENTE
- **Excel**: Usa $K$4 (media di squadra: totale pallet / ore totali) per la fascia
- **App**: Usa `media_squadra = totale_pallet_squadra / ore_mag_squadra` — **COERENTE**

#### D13: NESSUN KPI BONUS PER RICEVITORI NELL'EXCEL
- **Excel**: `H4 = IF(G4<=0, "-", G4)` — premio totale = premio base, nessun bonus
- **App**: L'app calcola e applica il bonus KPI 15% anche per ricevitori
- **Impatto**: L'app potrebbe dare premi piu alti per RICEV rispetto all'Excel.

#### D14: FONTE ORE — DIVERSA
- **Excel**: Ore da foglio timbrature locali (TIMBR_RICEV)
- **App**: Ore calcolate dal DB TIM (gestionale esterno) con `TIMESTAMPDIFF(MINUTE, data_inizio, data_fine) - pausa`
- **Impatto**: Possibili differenze nei minuti calcolati se ci sono discrepanze tra timbrature Excel e dati TIM.

#### D15: CALCOLO GIORNI
- **Excel**: `F4 = E4/8` — ore individuali / 8
- **App**: `giorni_in_premio = ore / ore_giorno` (parametro UI, default 8) — **COERENTE** se ore_giorno=8

---

## 5. FOGLIO TOTALE

### Excel
- Raccoglie premi da tutti i fogli PREMIO con VLOOKUP
- `K4 = SUM(G4:J4)` — somma tutti i premi
- Verifica di quadratura: `N7 = SUM(N3:N6)`, `O7 = IF(N7=K2,"ok","CHK")`

### App
- L'app non ha un foglio "Totale" equivalente visibile, ma ogni vista premi e separata

---

## RIEPILOGO DIVERGENZE CRITICHE

| # | Area | Severita | Descrizione |
|---|---|---|---|
| D1 | PIK Penalita | **ALTA** | App usa penalita_totale (produzione). Excel usa % errori DS con formula specifica (soglia 0.025%, 1€/errore) |
| D2 | PIK Bonus KPI | MEDIA | App applica 15% al netto (dopo penalita). Excel applica 15% al lordo (premio_base) |
| D3 | PIK Premio totale | MEDIA | Formula di aggregazione diversa (floor a 0 vs check <=0) |
| D9 | DS Compensazione | **ALTA** | App usa penalita_difetto (€), Excel usa conteggio errori difetto. Grandezze diverse! |
| D10 | DS Bonus KPI | MEDIA | App applica bonus KPI a DS. Excel NO. |
| D11 | DS Nuove aperture | BASSA | App esclude colli nuove aperture. Excel no. |
| D13 | RIC Bonus KPI | MEDIA | App applica bonus KPI a RICEV. Excel NO. |
| D14 | RIC Fonte ore | MEDIA | App usa TIM (gestionale esterno). Excel usa timbrature locali. |
| D6 | CAR Pesi | BASSA | Stessa logica ma applicata in punti diversi della pipeline |

## NOTA SULLE FASCE PREMIO

Le fasce default dell'app (database.py) corrispondono ESATTAMENTE a quelle dell'Excel:
- PICKING: 100-140 step 5, base 0.007, crescita 1.5% ✓
- CARRELLISTI: 18-26 step 2, base 0.03, crescita 21% ✓
- RICEVITORI: 18-24 step 2, base 2.50, crescita 44% ✓ (ma solo 4 fasce vs potenziali 5 nell'app)
- DOPPIA_SPUNTA: 147-199 step 13, base 0.005, crescita 5% ✓
