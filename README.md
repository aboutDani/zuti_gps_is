
Report Assegnazione Coordinate GPS ad Avvisi IS (ZUTI_GPS_IS) -> [svolto in autonomia su linguaggio ABAP per gestionale SAP]

Questo report gestisce l'allineamento e l'assegnazione delle coordinate GPS agli avvisi di interruzione di servizio (Tipo Avviso IS). Poiché la Cartografia associa le coordinate agli oggetti di allacciamento e gli avvisi IS contengono le sedi tecniche, il report esegue una riconciliazione basata sul confronto degli indirizzi per esteso.

🎛️ Schermata di Selezione (Selection Screen)

Il report prevede un blocco iniziale per la scelta della modalità operativa, configurato tramite Radio Button. I parametri di input variano dinamicamente in base alla fase selezionata.

Blocco Principale: Selezione Modalità

• [ ] Fase 1: Valorizzare tabella ZGPS_APPO
• [ ] Fase 2: Valorizzare avvisi IS

PARAMETRI IN INPUT - FASE 1 (Valorizzazione Tabella di Appoggio)

Da mostrare solo se è selezionata la Funzione 1.

Campo	Parametro SAP / Tipo	Obbligatorio	Valore di Default / Note
Percorso File Input	PARAMETER (Stringa)	Sì	Percorso di rete (AL11 / Application Server)
Flag Cancellazione Dati	PARAMETER (Checkbox)	Sì	Pre-fleggato ('X'). Se attivo, svuota la tabella


PARAMETRI IN INPUT - FASE 2 (Associazione Coordinate ad Avvisi)

Da mostrare solo se è selezionata la Funzione 2.

Campo	Parametro SAP / Tipo	Obbligatorio	Valore di Default / Note
Anno	ZRQTI_M2HEAD-ANNO (PARAMETER)	Sì	Selezione singola
Numero avviso	ZRQTI_M2HEAD-QMNUM (SELECT-OPTIONS)	No	Selezione multipla


🗄️ Struttura Tabella di Appoggio: ZGPS_APPO

Prima dello sviluppo del report, è necessaria la creazione della tabella trasparente ZGPS_APPO nel Dizionario Dati (SE11) con i seguenti campi:

Campo	Elemento Dati (Rif. Campo)	Chiave	Descrizione
MANDT	MANDT	Sì	Mandante
HAUS	EVBS-HAUS	Sì	Oggetto di allacciamento
CITY1	ADRC-CITY1	-	Località
STREET	ADRC-STREET	-	Via
HOUSE_NUM1	ADRC-HOUSE_NUM1	-	Civico
HOUSE_NUM2	ADRC-HOUSE_NUM2	-	Estensione civico
ZZGEOINTX	VIQMEL-ZZGEOINTX	-	Coordinata X
ZZGEOINTY	VIQMEL-ZZGEOINTY	-	Coordinata Y


⚙️ Logica di Elaborazione (Processing Logic)


🚀 ESECUZIONE FASE 1: Caricamento File e Normalizzazione Indirizzi

1. Reset Tabella (Se Flag Attivo): Se il "Flag di cancellazione dati" è valorizzato, esegui un'istruzione di DELETE FROM zgps_appo per azzerare la tabella di appoggio.
2. Lettura File: Il report acquisisce il file .xlsx dal percorso specificato (struttura standard con riga di intestazione da saltare).
3. Loop di Elaborazione: Per ciascuna riga del file, partendo dalla seconda, analizza il valore della Colonna B (OGG_ALL).
4. Decodifica Indirizzo SAP: Richiama il modulo di funzione standard IS-U:abap
CALL FUNCTION 'ISU_ADDRESS_PROVIDE'
  EXPORTING
    x_address_type = 'C'
    x_haus         = <valore_colonna_B>
  IMPORTING
    y_ieadrc       = ls_ieadrc
  EXCEPTIONS
    not_found      = 1
    OTHERS         = 2.
Usa il codice con cautela.
5. Scrittura su DB: Inserisci o aggiorna un record nella tabella ZGPS_APPO mappando i dati estratti:
	• ZGPS_APPO-HAUS = Oggetto di allacciamento (Colonna B)
	• ZGPS_APPO-CITY1 = Y_IEADRC-CITY1
	• ZGPS_APPO-STREET = Y_IEADRC-STREET
	• ZGPS_APPO-HOUSE_NUM1 = Y_IEADRC-HOUSE_NUM1
	• ZGPS_APPO-HOUSE_NUM2 = Y_IEADRC-STR_SUPPL1 (Nota di mappatura da specifica)
	• ZGPS_APPO-ZZGEOINTX = Colonna J del file Excel
	• ZGPS_APPO-ZZGEOINTY = Colonna I del file Excel

🔄 ESECUZIONE FASE 2: Riconciliazione e Assegnazione ad Avvisi IS

1. Controllo Tabella Vuota: Verifica preliminare su ZGPS_APPO. Se non contiene record, interrompi l'esecuzione con il messaggio di errore:❌ "Devi prima eseguire il passo 1: tabella ZGPS_APPO vuota"
2. Selezione Avvisi (ZRQTI_M2HEAD): Estrai la lista degli avvisi filtrando per ANNO (obbligatorio) e QMNUM (se valorizzato in selezione).
3. Loop Avvisi: Per ogni avviso estratto:
	• Filtro Avanzamento: Accedi a QMEL (o VIQMEL). Se il campo QMEL-GEOINTX è già valorizzato, l'avviso viene saltato.
	• Recupero Chiavi Ubicazione: Da VIQMEL, estrai i campi ILOAN (Ubicazione) e TPLNR (Sede Tecnica).
	• Recupero Codice Indirizzo (ILOA): Accedi alla tabella ILOA dove ILOA-ILOAN = VIQMEL-ILOAN e ILOA-TPLNR = VIQMEL-TPLNR per prelevare il codice indirizzo ILOA-ADRNR.
	• Lettura Indirizzo Reale (ADRC): Con il codice ADRNR, accedi a ADRC ed estrai: CITY1, STREET, HOUSE_NUM1, HOUSE_NUM2.
4. Confronto e Match: Cerca nella tabella ZGPS_APPO un record che corrisponda esattamente per:
	• CITY1 (Località)
	• STREET (Via)
	• HOUSE_NUM1 (Civico)
	• HOUSE_NUM2 (Estensione civico)
5. Aggiornamento Tabelle: Se viene trovato un match esatto, aggiorna i campi delle coordinate sulle tabelle degli Avvisi e della Qualità Tecnica:
	• Coordinata X: Aggiorna QMEL-GEOINTX e ZRQTI_M2HEAD-ORIGX
	• Coordinata Y: Aggiorna QMEL-GEOINTY e ZRQTI_M2HEAD-ORIGY
(Se il match non esiste, l'avviso viene saltato senza sollevare errori bloccanti, passando al successivo).

📊 Log di Output (ALV / Report finale)

Al termine dell'elaborazione della Fase 2, il programma deve esporre un log finale (preferibilmente in formato ALV Grid per permettere filtri e ordinamenti) contenente il dettaglio di tutti gli avvisi elaborati:

Colonna GitHub / ALV	Descrizione	Tabella / Campo Sorgente	Note
Numero avviso	Numero Avviso IS	ZRQTI_M2HEAD-QMNUM	-
Località	Località trovata su ADRC	ADRC-CITY1	-
Via	Via trovata su ADRC	ADRC-STREET	-
Civico	Civico trovato su ADRC	ADRC-HOUSE_NUM1	-
Estensione civico	Estensione trovata su ADRC	ADRC-HOUSE_NUM2	-
Coordinata X	Coordinata X assegnata	ZGPS_APPO-ZZGEOINTX	Vuota se nessun match
Coordinata Y	Coordinata Y assegnata	ZGPS_APPO-ZZGEOINTY	Vuota se nessun match
