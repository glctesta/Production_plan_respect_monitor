# Specifica funzionale e tecnica — Web app di monitoraggio rispetto piano di produzione

## Obiettivo
Realizzare una web application utilizzabile via browser per monitorare in modo continuo l'osservanza del piano di produzione giornaliero, verificando per ogni prodotto e per ogni fase produttiva se l'avanzamento reale è coerente con il pianificato.

L'applicazione dovrà:
- acquisire il piano giornaliero dal file Excel più recente presente in `T:\Planning`;
- usare come riferimento il tab `Last Phase`;
- eseguire automaticamente rilevazioni periodiche della produzione;
- confrontare quantità pianificate e quantità prodotte per ordine/fase/giorno;
- evidenziare in tempo reale i rischi di mancato rispetto del piano;
- inviare email automatiche di warning quando necessario.

## Contesto operativo
La giornata produttiva di riferimento dura **15 ore**, dalle **07:30** alle **23:30**.

A partire dalle **07:30 di ogni giorno lavorativo**, il programma deve avviarsi in automatico ed eseguire il ciclo di elaborazione **ogni 10 minuti**, ma l'intervallo deve essere **parametrizzabile**.

## Architettura attesa
Claude Code dovrà generare un programma composto indicativamente da questi moduli:
- backend Python;
- interfaccia web fruibile via browser;
- scheduler configurabile;
- modulo lettura Excel;
- modulo accesso database SQL Server;
- modulo motore di controllo KPI/trend;
- modulo invio email automatiche;
- file di configurazione centralizzato.

## Vincoli importanti
- I file per le connessioni al server sono già presenti nella root del progetto.
- I file Python per l’invio automatico delle email sono già presenti nella root del progetto.
- Il programma dovrà riutilizzare tali file, senza duplicare logiche già disponibili, ove possibile.
- L’applicazione dovrà essere pensata per uso continuo in ambiente produttivo.
- L’intervallo di polling, i percorsi, gli orari e le soglie dovranno essere configurabili.

## Funzionamento generale
Ad ogni ciclo schedulato il programma dovrà eseguire le seguenti macro-attività:
1. individuare il file Excel più recente nella cartella `T:\Planning`;
2. leggere il tab `Last Phase` e costruire il piano giornaliero per ordine/fase;
3. eseguire la query di snapshot della produzione reale;
4. analizzare le righe non ancora controllate di `traceability_rs.dbo.ShapShots` (`IsChecked = 0`);
5. confrontare produzione reale vs piano del giorno in funzione dell’ora corrente;
6. aggiornare dashboard/browser;
7. inviare eventuali email di warning;
8. marcare come controllate le righe elaborate (`IsChecked = 1`).

## Scheduling
Lo scheduler dovrà prevedere almeno i seguenti parametri:
- ora inizio giornata lavorativa: default `07:30`;
- ora fine giornata lavorativa: default `23:30`;
- intervallo di esecuzione: default `10 minuti`;
- calendario giorni lavorativi;
- eventuale modalità manuale “Esegui ora”.

Comportamento richiesto:
- il job deve partire solo nei giorni lavorativi;
- dalle 07:30 in avanti deve ciclare con periodicità configurabile;
- fuori fascia oraria non deve generare nuove elaborazioni automatiche, salvo lancio manuale;
- l’orario effettivo di esecuzione deve essere tracciato nei log.

## Sorgente dati Excel
Il programma deve usare come riferimento principale il file Excel presente in `T:\Planning`.

Regole:
- se nella cartella sono presenti più file Excel, va utilizzato **il più recente**;
- il foglio da utilizzare è **`Last Phase`**;
- il file selezionato e la data/ora di ultima modifica devono essere mostrati anche in interfaccia, per trasparenza operativa.

## Parsing del foglio `Last Phase`
Mappatura campi:
- **Colonna C** = numero ordine;
- **Colonna E** = fase di lavoro / macchina;
- **Colonna M in poi** = giorni di produzione;
- **Dalla riga 2 in poi** = quantità pianificate.

Regole di pulizia dati:
- se il numero ordine in colonna C contiene il carattere `•` all’inizio, tale carattere deve essere rimosso;
- il valore ordine va normalizzato tramite trim degli spazi;
- le intestazioni dalla colonna M in poi rappresentano i giorni di produzione e vanno convertite in date coerenti;
- ogni cella quantità da M2 in avanti rappresenta la quantità pianificata per una specifica combinazione **ordine + fase + giorno**.

Struttura logica attesa dopo il parsing:
- `OrderNumber`
- `MachineName` / `PhaseNameExcel`
- `ProductionDate`
- `PlannedQtyDay`

## Risoluzione Order ID
Per ogni numero ordine letto dall’Excel, il programma deve ricavare `IdOrder` e `ProductCode` tramite query su `Traceability_rs.dbo.Orders`.

Query di riferimento:
```sql
Select IdOrder, ProductCode
from Traceability_rs.dbo.Orders o
inner join traceability_rs.dbo.Products p on o.idproduct = p.idproduct
where Ordernumber = ?
```

Requisiti:
- il parametro `?` è il valore pulito della colonna C;
- se l’ordine non viene trovato, il record deve essere tracciato come anomalia di mapping;
- tali anomalie devono essere visibili in interfaccia e nei log.

## Risoluzione IdPhase
Per ogni fase/macchina letta dall’Excel, il programma deve ricavare l’`IdPhase` usando la seguente query di riferimento:

```sql
SELECT p.idphase AS phaseIdIntrasa
FROM [TraceabilityPlanning_RS].[dbo].[Machine]
INNER JOIN phase ON Machine.PhaseId = Phase.Phaseid
LEFT JOIN traceability_rs.dbo.phases p ON phase.PhaseName COLLATE DATABASE_DEFAULT = p.PhaseName
WHERE Machine.MachineName = ?
```

Requisiti:
- il parametro `?` è il valore presente in colonna E;
- se non esiste corrispondenza, la riga deve essere segnalata come anomalia;
- il mapping macchina/fase deve essere gestito con attenzione a maiuscole, spazi e differenze di collazione, senza alterare in modo improprio i dati originali.

## Snapshot di produzione
Ad ogni esecuzione il programma dovrà lanciare questa query di inserimento snapshot:

```sql
insert INTO ShapShots
SELECT  Idorder,
        IdPhase,
        isnull(Value,0) as QtyProcessed,
        Getdate() AS SnapShotTime,0
FROM (
    SELECT orders.idorder,
           Orders.ordernumber,
           Phases.IdPhase,
           cast(getdate() as date) AS Period,
           COUNT(DISTINCT dbo.BoardLabels(Scannings.IDBoard)) AS [Value],
           Phases.Phaseorder
    FROM Scannings
    INNER JOIN OrderPhases ON Scannings.IDOrderPhase = OrderPhases.IDOrderPhase
    INNER JOIN Orders ON OrderPhases.IDOrder = Orders.IDOrder
    INNER JOIN Phases ON OrderPhases.IDPhase = Phases.IDPhase
    INNER JOIN Products ON Orders.IDProduct = Products.IDProduct
    INNER JOIN Boards ON Boards.IDBoard = Scannings.IDBoard
    WHERE Scannings.ScanTimeFinish BETWEEN
        CAST(CAST(GETDATE() AS DATE) AS DATETIME) + CAST('07:30:00' AS DATETIME) AND
        CAST(CAST(GETDATE() + 1 AS DATE) AS DATETIME) + CAST('07:30:00' AS DATETIME)
        AND IsPass = 1
    GROUP BY orders.idorder, Orders.ordernumber, Phases.IdPhase, phases.PhaseOrder
) AS AllData;
```

Note implementative:
- il campo finale `0` rappresenta `IsChecked = 0`;
- la finestra temporale di raccolta dati è dalle `07:30` del giorno corrente alle `07:30` del giorno successivo, come da query attuale;
- l’orario di riferimento della giornata operativa va comunque esposto come parametro configurabile, anche se la query base viene inizialmente mantenuta coerente con l’impostazione corrente.

## Dati da leggere per il controllo
Dopo l’inserimento snapshot, il programma deve leggere dalla tabella `traceability_rs.dbo.ShapShots` **solo** le righe con:

```sql
IsChecked = 0
```

Le righe recuperate dovranno poi essere confrontate con il piano estratto dall’Excel del giorno corrente.

## Logica di verifica avanzamento
Per ogni riga snapshot, il programma deve determinare se, mantenendo il trend attuale, la produzione riuscirà a raggiungere la quantità pianificata entro la fine della giornata.

### Dati di input al calcolo
Per ogni combinazione `OrderNumber / IdOrder / IdPhase / giorno` servono:
- quantità pianificata giornaliera (`PlannedQtyDay`);
- quantità prodotta fino al momento dello snapshot (`QtyDone`);
- ora corrente dello snapshot;
- tempo trascorso dall’inizio della giornata produttiva;
- tempo residuo fino alle 23:30.

### Regola temporale
La giornata produttiva dura 15 ore:
- inizio: `07:30`
- fine: `23:30`

Il programma deve calcolare:
- minuti trascorsi da inizio turno al momento dello snapshot;
- percentuale di giornata trascorsa;
- produzione teorica attesa a quell’ora;
- proiezione di fine giornata con trend corrente.

### Formula attesa
Esempio logico minimo:
- `elapsed_minutes = minuti trascorsi da 07:30`
- `total_minutes = 900`
- `expected_by_now = PlannedQtyDay * (elapsed_minutes / total_minutes)`
- `projected_end_of_day = QtyDone / (elapsed_minutes / total_minutes)` se `elapsed_minutes > 0`
- `gap_projected = PlannedQtyDay - projected_end_of_day`

Claude Code dovrà implementare la formula in modo robusto, gestendo:
- snapshot vicinissimi all’inizio giornata;
- divisioni per zero;
- quantità pianificate nulle;
- casi oltre fine turno;
- casi prima dell’inizio turno.

## Regole colore stato
Per ogni riga a video, accanto a `Qty Done`, deve comparire una pallina di stato.

Regole richieste:
- **Verde**: andamento coerente con il raggiungimento del piano;
- **Gialla (Warning)**: esiste rischio di non raggiungere il piano a fine giornata;
- **Rossa**: se la differenza negativa stimata è **maggiore di 10 pezzi**.

Interpretazione operativa:
- se la proiezione di fine giornata è inferiore al pianificato, almeno stato **giallo**;
- se lo scostamento negativo stimato supera 10 pezzi, stato **rosso**;
- in tutti gli altri casi, stato **verde**.

È opportuno che Claude Code preveda soglie configurabili, con default:
- `warning_threshold = deficit > 0`
- `critical_threshold = deficit > 10`

## Dati da mostrare nella UI
La dashboard browser deve mostrare almeno queste colonne:
- `OrderNumber`
- `ProductCode`
- `Phase`
- `Planning QTY/Day`
- `Qty Done (pallina)`

Suggerimenti UI richiesti:
- refresh automatico coerente con il polling backend oppure via API;
- chiara leggibilità da monitor reparto/ufficio;
- uso di colori netti e intuitivi;
- ordinamento con priorità alle anomalie più critiche.

## Evidenza ordini fuori piano Excel
Se dalla tabella `ShapShots` emerge un `OrderNumber` che **non esiste nel file Excel di riferimento**, tale situazione deve essere evidenziata in modo speciale.

Requisiti UI obbligatori:
- la riga deve apparire **prima di tutte le altre**;
- tutta la riga deve essere **ROSSA**;
- il testo deve essere **BOLD**;
- la riga deve **lampeggiare**.

Questa categoria segnala produzione presente in traceability ma non prevista nel piano Excel giornaliero.

## Email automatiche
Il programma deve inviare email automatiche quando esistono righe con pallina:
- gialla;
- rossa.

### Recupero destinatari
I destinatari devono essere letti dal valore di configurazione avente attributo:
- `sys_email_planning_warning`

nel campo `[Value]` della tabella:
- `traceability_rs.dbo.settings`

Caratteristiche:
- il campo può contenere **più indirizzi email separati da `;`**;
- il programma deve effettuare parsing, pulizia e validazione minima degli indirizzi.

### Contenuto email
Le email devono essere:
- in **inglese**;
- con tono **più aggressivo**, ma sempre **professionale**;
- più incisive se nel corso della giornata le quantità in rosso aumentano.

Requisito funzionale importante:
- il sistema dovrebbe riconoscere se la situazione sta peggiorando durante la giornata;
- in tal caso il testo dell’email deve diventare progressivamente più urgente.

È consigliato prevedere livelli di severità email, ad esempio:
- livello 1: presenza di warning gialli;
- livello 2: presenza di rossi;
- livello 3: numero/entità dei rossi in aumento rispetto al ciclo precedente.

## Aggiornamento ShapShots
Al termine di ogni elaborazione, tutte le righe di `ShapShots` effettivamente processate in quel ciclo devono essere aggiornate impostando:

```sql
IsChecked = 1
```

Vincoli:
- l’update deve avvenire solo dopo completamento corretto delle verifiche;
- è preferibile usare transazioni o comunque un meccanismo che eviti doppie lavorazioni o perdita di righe;
- il sistema deve lavorare sempre sulle sole righe `IsChecked = 0`.

## Requisiti tecnici suggeriti
Claude Code dovrà proporre una soluzione con:
- backend Python, ad esempio FastAPI o Flask;
- frontend semplice ma efficace, servito dal backend;
- accesso SQL Server via driver appropriato (`pyodbc` o equivalente);
- lettura Excel con `openpyxl` o `pandas`;
- scheduler interno (`APScheduler`) oppure servizio Windows/scheduled worker;
- file `.env` o file `config.yaml/json` per i parametri.

## Configurazioni richieste
Prevedere parametri configurabili almeno per:
- `PLANNING_FOLDER = T:\Planning`
- `PLANNING_SHEET = Last Phase`
- `WORKDAY_START = 07:30`
- `WORKDAY_END = 23:30`
- `POLL_MINUTES = 10`
- `RED_DEFICIT_THRESHOLD = 10`
- `SETTINGS_EMAIL_ATTRIBUTE = sys_email_planning_warning`
- `ENABLE_EMAILS = true/false`
- `ENABLE_BLINKING_ALERTS = true/false`

## Requisiti di robustezza
Il programma deve gestire correttamente:
- cartella `T:\Planning` non raggiungibile;
- nessun file Excel presente;
- file Excel corrotto o foglio `Last Phase` assente;
- intestazioni date non valide da colonna M in poi;
- ordini non trovati nel DB;
- fasi/macchine non mappate;
- DB temporaneamente non disponibile;
- timeout query;
- invio email fallito;
- duplicazioni o incoerenze nei dati.

Ogni anomalia deve essere registrata nei log e, se utile, mostrata in una sezione dedicata della UI.

## Logging e audit
Prevedere log applicativi con almeno queste informazioni:
- avvio/fine ciclo;
- file Excel selezionato;
- numero righe piano lette;
- numero snapshot inseriti;
- numero righe controllate;
- numero verdi/gialle/rosse;
- numero ordini fuori piano Excel;
- email inviate / fallite;
- errori di mapping ordine/fase.

## Dashboard web — requisiti minimi
La UI browser deve includere almeno:
- tabella principale stato produzione giornaliero;
- evidenza visiva di warning e critical;
- indicazione ultimo aggiornamento;
- file Excel sorgente attualmente usato;
- contatori sintetici (verdi, gialli, rossi, fuori piano);
- eventuale sezione anomalie/log essenziali.

## Ordinamento suggerito in tabella
Ordine di visualizzazione consigliato:
1. ordini presenti in `ShapShots` ma assenti in Excel;
2. righe rosse;
3. righe gialle;
4. righe verdi.

All’interno dei gruppi, ordinare idealmente per:
- fase,
- order number,
- maggior deficit stimato.

## API suggerite
Claude Code può organizzare il backend con endpoint del tipo:
- `GET /api/status` → stato corrente dashboard;
- `POST /api/run-now` → esecuzione manuale ciclo;
- `GET /api/config` → configurazione runtime non sensibile;
- `GET /api/health` → health check;
- `GET /` → dashboard HTML.

## Modello dati interno suggerito
Per ogni riga visualizzata, il backend dovrebbe costruire un record logico contenente almeno:
- `OrderNumber`
- `IdOrder`
- `ProductCode`
- `Phase`
- `IdPhase`
- `PlanningDate`
- `PlannedQtyDay`
- `QtyDone`
- `SnapshotTime`
- `ExpectedByNow`
- `ProjectedEndQty`
- `ProjectedDeficit`
- `StatusColor`
- `InExcelPlan`
- `IsOutOfPlan`
- `NeedsEmailAlert`

## Regole funzionali da implementare con priorità alta
1. uso dell’Excel più recente in `T:\Planning`;
2. parsing corretto del tab `Last Phase`;
3. mapping ordine e fase verso DB;
4. insert snapshot automatico;
5. lettura snapshot `IsChecked = 0`;
6. confronto con piano giornaliero in base all’ora corrente;
7. assegnazione pallina verde/gialla/rossa;
8. evidenza speciale per ordini non presenti nel piano Excel;
9. invio email automatiche a destinatari multipli;
10. update finale `IsChecked = 1`.

## Richiesta finale per Claude Code
Genera un progetto completo, pronto per esecuzione locale/intranet, con:
- struttura cartelle chiara;
- codice backend Python ben separato per moduli;
- interfaccia web chiara e leggibile;
- file configurazione parametrico;
- logging robusto;
- commenti essenziali nei punti delicati;
- istruzioni di avvio;
- eventuale script iniziale per setup ambiente.

## Criteri di accettazione
Il lavoro sarà considerato corretto se:
- il sistema seleziona davvero l’Excel più recente da `T:\Planning`;
- legge correttamente il foglio `Last Phase`;
- ricava correttamente `IdOrder`, `ProductCode` e `IdPhase`;
- inserisce gli snapshot in `ShapShots`;
- controlla esclusivamente snapshot con `IsChecked = 0`;
- valuta il trend temporale sulle 15 ore di lavoro;
- assegna correttamente verde/giallo/rosso;
- segnala in rosso bold lampeggiante gli ordini fuori piano Excel;
- invia email ai destinatari configurati quando necessario;
- aggiorna `IsChecked = 1` dopo l’elaborazione;
- espone tutto via browser in modo chiaro.

## Nota finale
Se alcuni dettagli implementativi non sono definiti nei file già presenti in root, Claude Code dovrà integrarli senza rompere le utility esistenti, privilegiando riuso, configurabilità e robustezza operativa.
