const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        ImageRun, Header, AlignmentType, HeadingLevel, BorderStyle,
        WidthType, ShadingType, LevelFormat, PageBreak, PageNumber,
        Footer } = require("docx");

const logoData = fs.readFileSync("Logo.png");

// Colors
const BLUE = "1B4F72";
const LIGHT_BLUE = "D6EAF8";
const GREEN = "27AE60";
const YELLOW = "F39C12";
const RED = "E74C3C";
const GRAY = "F2F3F4";
const WHITE = "FFFFFF";

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

function headerCell(text, width) {
    return new TableCell({
        borders, width: { size: width, type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: cellMargins,
        children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: WHITE, font: "Arial", size: 20 })] })]
    });
}

function dataCell(text, width, fill) {
    return new TableCell({
        borders, width: { size: width, type: WidthType.DXA },
        shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
        margins: cellMargins,
        children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 20 })] })]
    });
}

function h1(text) {
    return new Paragraph({
        spacing: { before: 300, after: 200 },
        children: [new TextRun({ text, bold: true, size: 36, color: BLUE, font: "Arial" })]
    });
}

function h2(text) {
    return new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [new TextRun({ text, bold: true, size: 28, color: BLUE, font: "Arial" })]
    });
}

function h3(text) {
    return new Paragraph({
        spacing: { before: 180, after: 80 },
        children: [new TextRun({ text, bold: true, size: 24, color: "2C3E50", font: "Arial" })]
    });
}

function para(text) {
    return new Paragraph({
        spacing: { after: 80 },
        children: [new TextRun({ text, size: 22, font: "Arial" })]
    });
}

function bullet(text, bold_prefix) {
    const children = [];
    if (bold_prefix) {
        children.push(new TextRun({ text: bold_prefix, bold: true, size: 22, font: "Arial" }));
        children.push(new TextRun({ text, size: 22, font: "Arial" }));
    } else {
        children.push(new TextRun({ text, size: 22, font: "Arial" }));
    }
    return new Paragraph({
        spacing: { after: 40 },
        indent: { left: 360, hanging: 180 },
        children: [new TextRun({ text: "\u2022 ", size: 22, font: "Arial" }), ...children]
    });
}

function numbered(num, text) {
    return new Paragraph({
        spacing: { after: 40 },
        indent: { left: 360, hanging: 360 },
        children: [
            new TextRun({ text: `${num}. `, bold: true, size: 22, font: "Arial" }),
            new TextRun({ text, size: 22, font: "Arial" })
        ]
    });
}

function colorTable(rows, colWidths) {
    const tableWidth = colWidths.reduce((a, b) => a + b, 0);
    return new Table({
        width: { size: tableWidth, type: WidthType.DXA },
        columnWidths: colWidths,
        rows: rows.map((row, ri) => new TableRow({
            children: row.map((cell, ci) => {
                if (ri === 0) return headerCell(cell, colWidths[ci]);
                return dataCell(cell, colWidths[ci], ri % 2 === 0 ? GRAY : undefined);
            })
        }))
    });
}

function pageBreakPara() {
    return new Paragraph({ children: [new PageBreak()] });
}

// ===================== CONTENT GENERATORS PER LANGUAGE =====================

function generateSection(lang) {
    const t = translations[lang];
    return [
        h1(t.title),
        para(t.subtitle),
        new Paragraph({ spacing: { after: 100 }, children: [] }),

        h2(t.s1_title),
        para(t.s1_text),
        h3(t.s1_features),
        ...t.s1_bullets.map(b => bullet(b)),

        h2(t.s2_title),
        h3(t.s2_excel_title),
        ...t.s2_excel_bullets.map(b => bullet(b)),
        h3(t.s2_db_title),
        ...t.s2_db_bullets.map(b => bullet(b)),
        h3(t.s2_planning_title),
        ...t.s2_planning_bullets.map(b => bullet(b)),

        h2(t.s3_title),
        h3(t.s3_auto_title),
        ...t.s3_auto_bullets.map(b => bullet(b)),
        h3(t.s3_cycle_title),
        ...t.s3_cycle_steps.map((s, i) => numbered(i + 1, s)),
        h3(t.s3_manual_title),
        para(t.s3_manual_text),

        h2(t.s4_title),
        ...t.s4_formulas.map(b => bullet(b)),

        h2(t.s5_title),
        colorTable(t.s5_table, [2000, 2200, 5160]),

        h2(t.s6_title),
        bullet(t.s6_yellow),
        bullet(t.s6_blue),

        h2(t.s7_title),
        h3(t.s7_when_title),
        ...t.s7_when_bullets.map(b => bullet(b)),
        h3(t.s7_cooldown_title),
        ...t.s7_cooldown_bullets.map(b => bullet(b)),
        h3(t.s7_severity_title),
        ...t.s7_severity_bullets.map(b => bullet(b)),
        h3(t.s7_recipients_title),
        ...t.s7_recipients_bullets.map(b => bullet(b)),
        h3(t.s7_content_title),
        ...t.s7_content_bullets.map(b => bullet(b)),

        h2(t.s8_title),
        colorTable(t.s8_table, [2200, 7160]),

        h2(t.s9_title),
        para(t.s9_intro),
        ...t.s9_params.map(b => bullet(b)),
    ];
}

// ===================== TRANSLATIONS =====================

const translations = {

// ---- ENGLISH ----
en: {
    lang_label: "ENGLISH",
    title: "Production Plan Monitor",
    subtitle: "User Guide",
    s1_title: "1. What the Application Does",
    s1_text: "The Production Plan Monitor is a web application that continuously monitors compliance with the daily production plan. Every 10 minutes (configurable), it compares actual production quantities recorded in the traceability system (SQL Server) against planned quantities from the Excel planning file. Results are displayed in a real-time dashboard at http://localhost:8085.",
    s1_features: "Key Features:",
    s1_bullets: [
        "Real-time comparison of planned vs actual production",
        "Color-coded status indicators (green / yellow / red)",
        "Automatic detection of orders produced but not in the daily plan",
        "Historical context analysis (past 2 working days and future 3 working days)",
        "Automatic email alerts with escalating urgency",
        "Auto-scrolling dashboard designed for shop floor monitors"
    ],
    s2_title: "2. Data Sources",
    s2_excel_title: "Excel Planning File",
    s2_excel_bullets: [
        "Location: T:\\Planning (the most recent .xlsx file is automatically selected)",
        "Sheet used: PlanningMachine (configurable)",
        "Column C: Order Number (bullet character automatically removed)",
        "Column E: Machine / Phase name",
        "Columns M onwards: Production dates with planned quantities per day",
        "File is re-checked every 30 minutes; new versions are loaded automatically"
    ],
    s2_db_title: "SQL Server Database (Traceability_rs)",
    s2_db_bullets: [
        "Orders table: resolves Order Number to IdOrder and ProductCode",
        "Phases table: resolves Machine Name to IdPhase, with PhaseOrder for sorting",
        "Scannings table: counts actual production (07:30 to 07:30 next day)",
        "ShapShots table: periodic snapshots; only IsChecked=0 rows are processed",
        "Settings table: email recipients (attribute: sys_email_planning_warning)"
    ],
    s2_planning_title: "TraceabilityPlanning_RS Database",
    s2_planning_bullets: [
        "Machine + Phase tables: maps machine names from Excel to phase IDs"
    ],
    s3_title: "3. Execution Timing",
    s3_auto_title: "Automatic Schedule",
    s3_auto_bullets: [
        "Working hours: 07:30 to 23:30 (configurable)",
        "Polling interval: every 10 minutes (configurable)",
        "Scheduler starts automatically at application launch",
        "Outside working hours, automatic cycles are suspended",
        "First cycle runs immediately at startup"
    ],
    s3_cycle_title: "Each Cycle Performs:",
    s3_cycle_steps: [
        "Checks for the most recent Excel file (every 30 min, cached otherwise)",
        "Parses the planning sheet and filters for today",
        "Connects to SQL Server",
        "Inserts a production snapshot (INSERT INTO ShapShots)",
        "Reads unchecked snapshots (IsChecked = 0)",
        "Compares actual production vs plan using time-based projection",
        "Assigns status: Green (on track), Yellow (behind), Red (deficit > 10 pcs)",
        "Detects out-of-plan orders",
        "For out-of-plan orders: checks past 2 and future 3 working days",
        "Updates the dashboard",
        "Sends email alerts if needed (respecting cooldown)",
        "Marks processed snapshots as IsChecked = 1"
    ],
    s3_manual_title: "Manual Execution",
    s3_manual_text: "The \"Run Now\" button triggers an immediate cycle, bypassing the time window check.",
    s4_title: "4. Projection Formula",
    s4_formulas: [
        "Production day = 900 minutes (07:30 to 23:30)",
        "elapsed_minutes = minutes since 07:30",
        "fraction = elapsed_minutes / 900",
        "Expected by now = Planned Qty x fraction",
        "Projected end of day = Qty Done / fraction",
        "Projected deficit = Planned Qty - Projected end of day",
        "All values are rounded to integers (whole pieces)"
    ],
    s5_title: "5. Status Colors",
    s5_table: [
        ["Color", "Meaning", "Condition"],
        ["Green", "On track", "Projected deficit = 0"],
        ["Yellow", "Warning", "Projected deficit between 1 and 10 pieces"],
        ["Red", "Critical", "Projected deficit > 10 pieces"],
        ["Red blinking", "Out of Plan", "Order in production but not in today's Excel plan"]
    ],
    s6_title: "6. Star Indicators",
    s6_yellow: "Yellow star: order not in today's plan but scheduled in next 3 working days",
    s6_blue: "Blue star: order not in today's plan but was scheduled in past 2 working days (delayed)",
    s7_title: "7. Email Warnings",
    s7_when_title: "When Emails Are Sent:",
    s7_when_bullets: ["When orders have Yellow or Red status", "When out-of-plan orders are detected"],
    s7_cooldown_title: "Cooldown Rules:",
    s7_cooldown_bullets: ["Yellow alerts: max 1 email every 2 hours", "Red alerts: max 1 email every 1 hour", "Cooldown resets each new day"],
    s7_severity_title: "Severity Levels:",
    s7_severity_bullets: ["Level 1 (moderate): only yellow warnings", "Level 2 (firm): red alerts or out-of-plan orders", "Level 3 (urgent): red alerts INCREASED vs previous cycle"],
    s7_recipients_title: "Recipients:",
    s7_recipients_bullets: ["Read from: traceability_rs.dbo.settings (attribute = sys_email_planning_warning)", "Multiple emails separated by semicolon are supported"],
    s7_content_title: "Email Content:",
    s7_content_bullets: ["Written in English, professional but firm tone", "Summary counters (green/yellow/red/out-of-plan)", "Detailed table of problematic orders", "Star context for out-of-plan orders", "Source Excel file name and timestamp"],
    s8_title: "8. Dashboard Columns",
    s8_table: [
        ["Column", "Description"],
        ["Order Number", "Production order identifier"],
        ["Product Code", "Product code from the database"],
        ["Phase", "Production phase name"],
        ["Planning QTY/Day", "Planned quantity for today"],
        ["Qty Done", "Actual produced (with status dot and star if applicable)"],
        ["Expected Now", "Pieces expected by current time"],
        ["Projected End", "Projected total at end of day"],
        ["Deficit", "Projected shortfall in pieces"]
    ],
    s9_title: "9. Configuration (config.yaml)",
    s9_intro: "All parameters are in config.yaml:",
    s9_params: [
        "planning.folder: Path to Excel files (default: T:\\Planning)",
        "planning.sheet: Sheet name (default: PlanningMachine)",
        "workday.start / end: Working hours (default: 07:30 / 23:30)",
        "polling.interval_minutes: Cycle interval (default: 10)",
        "thresholds.red_deficit: Pieces for red (default: 10)",
        "email.enabled: Enable/disable emails",
        "email.yellow_cooldown_minutes: Yellow cooldown (default: 120)",
        "email.red_cooldown_minutes: Red cooldown (default: 60)",
        "server.port: Web server port (default: 8085)"
    ]
},

// ---- ROMANIAN ----
ro: {
    lang_label: "ROMANA",
    title: "Production Plan Monitor",
    subtitle: "Ghid de Utilizare",
    s1_title: "1. Ce Face Aplicatia",
    s1_text: "Production Plan Monitor este o aplicatie web care monitorizeaza continuu respectarea planului zilnic de productie. La fiecare 10 minute (configurabil), compara cantitatile reale de productie inregistrate in sistemul de trasabilitate (SQL Server) cu cantitatile planificate din fisierul Excel. Rezultatele sunt afisate intr-un dashboard in timp real la http://localhost:8085.",
    s1_features: "Functionalitati principale:",
    s1_bullets: [
        "Comparatie in timp real intre productia planificata si cea realizata",
        "Indicatori de stare cu culori (verde / galben / rosu)",
        "Detectarea automata a comenzilor produse dar neplanificate",
        "Analiza contextului istoric (2 zile lucratoare inapoi si 3 zile inainte)",
        "Alerte email automate cu urgenta progresiva",
        "Dashboard cu defilare automata, proiectat pentru monitoare de productie"
    ],
    s2_title: "2. Surse de Date",
    s2_excel_title: "Fisierul Excel de Planificare",
    s2_excel_bullets: [
        "Locatie: T:\\Planning (cel mai recent fisier .xlsx este selectat automat)",
        "Foaia utilizata: PlanningMachine (configurabil)",
        "Coloana C: Numar comanda (caracterul bullet este eliminat automat)",
        "Coloana E: Numele masinii / fazei",
        "Coloanele M+: Datele de productie cu cantitatile planificate pe zi",
        "Fisierul este reverificat la fiecare 30 minute; versiunile noi sunt incarcate automat"
    ],
    s2_db_title: "Baza de Date SQL Server (Traceability_rs)",
    s2_db_bullets: [
        "Tabela Orders: rezolva numarul comenzii in IdOrder si ProductCode",
        "Tabela Phases: rezolva numele masinii in IdPhase, cu PhaseOrder pentru sortare",
        "Tabela Scannings: numara productia reala (07:30 - 07:30 ziua urmatoare)",
        "Tabela ShapShots: snapshot-uri periodice; doar randurile cu IsChecked=0 sunt procesate",
        "Tabela Settings: destinatarii email (atribut: sys_email_planning_warning)"
    ],
    s2_planning_title: "Baza de Date TraceabilityPlanning_RS",
    s2_planning_bullets: [
        "Tabelele Machine + Phase: mapeaza numele masinilor din Excel in ID-uri de faze"
    ],
    s3_title: "3. Programarea Executiei",
    s3_auto_title: "Programare Automata",
    s3_auto_bullets: [
        "Ore de lucru: 07:30 - 23:30 (configurabil)",
        "Interval de polling: la fiecare 10 minute (configurabil)",
        "Scheduler-ul porneste automat la lansarea aplicatiei",
        "In afara orelor de lucru, ciclurile automate sunt suspendate",
        "Primul ciclu ruleaza imediat la pornire"
    ],
    s3_cycle_title: "Fiecare Ciclu Executa:",
    s3_cycle_steps: [
        "Verifica cel mai recent fisier Excel (la 30 min, altfel din cache)",
        "Parseaza foaia de planificare si filtreaza pentru ziua curenta",
        "Se conecteaza la SQL Server",
        "Insereaza un snapshot de productie (INSERT INTO ShapShots)",
        "Citeste snapshot-urile neverificate (IsChecked = 0)",
        "Compara productia reala cu planul folosind proiectia temporala",
        "Atribuie starea: Verde (conform), Galben (intarziere), Rosu (deficit > 10 buc)",
        "Detecteaza comenzile in afara planului",
        "Pentru comenzi in afara planului: verifica 2 zile anterioare si 3 zile viitoare",
        "Actualizeaza dashboard-ul",
        "Trimite alerte email daca e necesar (respectand cooldown-ul)",
        "Marcheaza snapshot-urile procesate ca IsChecked = 1"
    ],
    s3_manual_title: "Executie Manuala",
    s3_manual_text: "Butonul \"Run Now\" declanseaza un ciclu imediat, fara restrictie de interval orar.",
    s4_title: "4. Formula de Proiectie",
    s4_formulas: [
        "Ziua de productie = 900 minute (07:30 - 23:30)",
        "minute_scurse = minute de la 07:30",
        "fractie = minute_scurse / 900",
        "Asteptat acum = Cantitate Planificata x fractie",
        "Proiectie sfarsit de zi = Cantitate Realizata / fractie",
        "Deficit proiectat = Cantitate Planificata - Proiectie sfarsit de zi",
        "Toate valorile sunt rotunjite la numere intregi (piese intregi)"
    ],
    s5_title: "5. Culorile de Stare",
    s5_table: [
        ["Culoare", "Semnificatie", "Conditie"],
        ["Verde", "Conform", "Deficit proiectat = 0"],
        ["Galben", "Avertisment", "Deficit proiectat intre 1 si 10 piese"],
        ["Rosu", "Critic", "Deficit proiectat > 10 piese"],
        ["Rosu clipitor", "In afara planului", "Comanda in productie dar nu in planul Excel"]
    ],
    s6_title: "6. Indicatorii Stea",
    s6_yellow: "Stea galbena: comanda nu e in planul de azi dar e programata in urmatoarele 3 zile lucratoare",
    s6_blue: "Stea albastra: comanda nu e in planul de azi dar era programata in ultimele 2 zile lucratoare (intarziere)",
    s7_title: "7. Avertismente Email",
    s7_when_title: "Cand se Trimit Emailuri:",
    s7_when_bullets: ["Cand exista comenzi cu starea Galben sau Rosu", "Cand sunt detectate comenzi in afara planului"],
    s7_cooldown_title: "Reguli de Cooldown:",
    s7_cooldown_bullets: ["Alerte galbene: max 1 email la 2 ore", "Alerte rosii: max 1 email la 1 ora", "Cooldown-ul se reseteaza in fiecare zi noua"],
    s7_severity_title: "Niveluri de Severitate:",
    s7_severity_bullets: ["Nivel 1 (moderat): doar avertismente galbene", "Nivel 2 (ferm): alerte rosii sau comenzi in afara planului", "Nivel 3 (urgent): alertele rosii au CRESCUT fata de ciclul anterior"],
    s7_recipients_title: "Destinatari:",
    s7_recipients_bullets: ["Cititi din: traceability_rs.dbo.settings (atribut = sys_email_planning_warning)", "Mai multe adrese email separate prin punct si virgula sunt acceptate"],
    s7_content_title: "Continutul Emailului:",
    s7_content_bullets: ["Scris in engleza, ton profesional dar ferm", "Contoare sumar (verde/galben/rosu/in afara planului)", "Tabel detaliat cu comenzile problematice", "Context stea pentru comenzi in afara planului", "Numele fisierului Excel sursa si timestamp"],
    s8_title: "8. Coloanele Dashboard-ului",
    s8_table: [
        ["Coloana", "Descriere"],
        ["Order Number", "Identificator comanda de productie"],
        ["Product Code", "Codul produsului din baza de date"],
        ["Phase", "Numele fazei de productie"],
        ["Planning QTY/Day", "Cantitatea planificata pentru azi"],
        ["Qty Done", "Cantitatea realizata (cu indicator de stare si stea daca e cazul)"],
        ["Expected Now", "Piese asteptate pana la ora curenta"],
        ["Projected End", "Total proiectat la sfarsitul zilei"],
        ["Deficit", "Deficit proiectat in piese"]
    ],
    s9_title: "9. Configurare (config.yaml)",
    s9_intro: "Toti parametrii sunt in config.yaml:",
    s9_params: [
        "planning.folder: Calea catre fisierele Excel (implicit: T:\\Planning)",
        "planning.sheet: Numele foii (implicit: PlanningMachine)",
        "workday.start / end: Ore de lucru (implicit: 07:30 / 23:30)",
        "polling.interval_minutes: Interval ciclu (implicit: 10)",
        "thresholds.red_deficit: Piese pentru rosu (implicit: 10)",
        "email.enabled: Activeaza/dezactiveaza emailuri",
        "email.yellow_cooldown_minutes: Cooldown galben (implicit: 120)",
        "email.red_cooldown_minutes: Cooldown rosu (implicit: 60)",
        "server.port: Port server web (implicit: 8085)"
    ]
},

// ---- ITALIAN ----
it: {
    lang_label: "ITALIANO",
    title: "Production Plan Monitor",
    subtitle: "Guida Utente",
    s1_title: "1. Cosa Fa l'Applicazione",
    s1_text: "Il Production Plan Monitor e' un'applicazione web che monitora continuamente il rispetto del piano di produzione giornaliero. Ogni 10 minuti (configurabile), confronta le quantita' reali registrate nel sistema di tracciabilita' (SQL Server) con le quantita' pianificate dal file Excel. I risultati vengono visualizzati in una dashboard in tempo reale su http://localhost:8085.",
    s1_features: "Funzionalita' principali:",
    s1_bullets: [
        "Confronto in tempo reale tra produzione pianificata e reale",
        "Indicatori di stato con colori (verde / giallo / rosso)",
        "Rilevamento automatico degli ordini prodotti ma non pianificati",
        "Analisi del contesto storico (2 giorni lavorativi precedenti e 3 successivi)",
        "Email di allarme automatiche con urgenza crescente",
        "Dashboard con scorrimento automatico, progettata per monitor di reparto"
    ],
    s2_title: "2. Fonti dei Dati",
    s2_excel_title: "File Excel di Pianificazione",
    s2_excel_bullets: [
        "Posizione: T:\\Planning (il file .xlsx piu' recente viene selezionato automaticamente)",
        "Foglio utilizzato: PlanningMachine (configurabile)",
        "Colonna C: Numero ordine (il carattere bullet viene rimosso automaticamente)",
        "Colonna E: Nome macchina / fase",
        "Colonne M+: Date di produzione con quantita' pianificate per giorno",
        "Il file viene ricontrollato ogni 30 minuti; le nuove versioni vengono caricate automaticamente"
    ],
    s2_db_title: "Database SQL Server (Traceability_rs)",
    s2_db_bullets: [
        "Tabella Orders: risolve il numero ordine in IdOrder e ProductCode",
        "Tabella Phases: risolve il nome macchina in IdPhase, con PhaseOrder per l'ordinamento",
        "Tabella Scannings: conta la produzione reale (07:30 - 07:30 giorno successivo)",
        "Tabella ShapShots: snapshot periodici; solo le righe con IsChecked=0 vengono elaborate",
        "Tabella Settings: destinatari email (attributo: sys_email_planning_warning)"
    ],
    s2_planning_title: "Database TraceabilityPlanning_RS",
    s2_planning_bullets: [
        "Tabelle Machine + Phase: mappa i nomi macchina dall'Excel agli ID fase nel sistema"
    ],
    s3_title: "3. Tempistiche di Esecuzione",
    s3_auto_title: "Programmazione Automatica",
    s3_auto_bullets: [
        "Orario di lavoro: 07:30 - 23:30 (configurabile)",
        "Intervallo di polling: ogni 10 minuti (configurabile)",
        "Lo scheduler si avvia automaticamente al lancio dell'applicazione",
        "Fuori orario lavorativo, i cicli automatici sono sospesi",
        "Il primo ciclo viene eseguito immediatamente all'avvio"
    ],
    s3_cycle_title: "Ogni Ciclo Esegue:",
    s3_cycle_steps: [
        "Verifica il file Excel piu' recente (ogni 30 min, altrimenti dalla cache)",
        "Parsa il foglio di pianificazione e filtra per la data odierna",
        "Si connette a SQL Server",
        "Inserisce uno snapshot di produzione (INSERT INTO ShapShots)",
        "Legge gli snapshot non verificati (IsChecked = 0)",
        "Confronta la produzione reale con il piano usando la proiezione temporale",
        "Assegna lo stato: Verde (conforme), Giallo (ritardo), Rosso (deficit > 10 pz)",
        "Rileva gli ordini fuori piano",
        "Per gli ordini fuori piano: controlla 2 giorni precedenti e 3 giorni successivi",
        "Aggiorna la dashboard",
        "Invia email di allarme se necessario (rispettando il cooldown)",
        "Marca gli snapshot elaborati come IsChecked = 1"
    ],
    s3_manual_title: "Esecuzione Manuale",
    s3_manual_text: "Il pulsante \"Run Now\" attiva un ciclo immediato, senza controllo della fascia oraria.",
    s4_title: "4. Formula di Proiezione",
    s4_formulas: [
        "Giornata produttiva = 900 minuti (07:30 - 23:30)",
        "minuti_trascorsi = minuti dalle 07:30",
        "frazione = minuti_trascorsi / 900",
        "Atteso adesso = Quantita' Pianificata x frazione",
        "Proiezione fine giornata = Quantita' Prodotta / frazione",
        "Deficit proiettato = Quantita' Pianificata - Proiezione fine giornata",
        "Tutti i valori sono arrotondati a numeri interi (pezzi interi)"
    ],
    s5_title: "5. Colori di Stato",
    s5_table: [
        ["Colore", "Significato", "Condizione"],
        ["Verde", "Conforme", "Deficit proiettato = 0"],
        ["Giallo", "Attenzione", "Deficit proiettato tra 1 e 10 pezzi"],
        ["Rosso", "Critico", "Deficit proiettato > 10 pezzi"],
        ["Rosso lampeggiante", "Fuori Piano", "Ordine in produzione ma non nel piano Excel"]
    ],
    s6_title: "6. Indicatori Stella",
    s6_yellow: "Stella gialla: ordine non nel piano di oggi ma programmato nei prossimi 3 giorni lavorativi",
    s6_blue: "Stella blu: ordine non nel piano di oggi ma era programmato nei 2 giorni lavorativi precedenti (ritardo)",
    s7_title: "7. Email di Avviso",
    s7_when_title: "Quando Vengono Inviate le Email:",
    s7_when_bullets: ["Quando ci sono ordini con stato Giallo o Rosso", "Quando vengono rilevati ordini fuori piano"],
    s7_cooldown_title: "Regole di Cooldown:",
    s7_cooldown_bullets: ["Allarmi gialli: max 1 email ogni 2 ore", "Allarmi rossi: max 1 email ogni 1 ora", "Il cooldown si resetta ogni nuovo giorno"],
    s7_severity_title: "Livelli di Severita':",
    s7_severity_bullets: ["Livello 1 (moderato): solo avvisi gialli", "Livello 2 (fermo): allarmi rossi o ordini fuori piano", "Livello 3 (urgente): allarmi rossi AUMENTATI rispetto al ciclo precedente"],
    s7_recipients_title: "Destinatari:",
    s7_recipients_bullets: ["Letti da: traceability_rs.dbo.settings (attributo = sys_email_planning_warning)", "Piu' indirizzi email separati da punto e virgola sono supportati"],
    s7_content_title: "Contenuto Email:",
    s7_content_bullets: ["Scritte in inglese, tono professionale ma deciso", "Contatori sintetici (verde/giallo/rosso/fuori piano)", "Tabella dettagliata degli ordini problematici", "Contesto stella per ordini fuori piano", "Nome file Excel sorgente e timestamp"],
    s8_title: "8. Colonne della Dashboard",
    s8_table: [
        ["Colonna", "Descrizione"],
        ["Order Number", "Identificativo dell'ordine di produzione"],
        ["Product Code", "Codice prodotto dal database"],
        ["Phase", "Nome della fase di produzione"],
        ["Planning QTY/Day", "Quantita' pianificata per oggi"],
        ["Qty Done", "Quantita' prodotta (con pallina di stato e stella se applicabile)"],
        ["Expected Now", "Pezzi attesi all'ora corrente"],
        ["Projected End", "Totale proiettato a fine giornata"],
        ["Deficit", "Deficit proiettato in pezzi"]
    ],
    s9_title: "9. Configurazione (config.yaml)",
    s9_intro: "Tutti i parametri sono nel file config.yaml:",
    s9_params: [
        "planning.folder: Percorso file Excel (default: T:\\Planning)",
        "planning.sheet: Nome foglio (default: PlanningMachine)",
        "workday.start / end: Orario lavorativo (default: 07:30 / 23:30)",
        "polling.interval_minutes: Intervallo ciclo (default: 10)",
        "thresholds.red_deficit: Pezzi per rosso (default: 10)",
        "email.enabled: Abilita/disabilita email",
        "email.yellow_cooldown_minutes: Cooldown giallo (default: 120)",
        "email.red_cooldown_minutes: Cooldown rosso (default: 60)",
        "server.port: Porta server web (default: 8085)"
    ]
},

// ---- SWEDISH ----
sv: {
    lang_label: "SVENSKA",
    title: "Production Plan Monitor",
    subtitle: "Anvandarguide",
    s1_title: "1. Vad Applikationen Gor",
    s1_text: "Production Plan Monitor ar en webbapplikation som kontinuerligt overvakar efterlevnaden av den dagliga produktionsplanen. Var 10:e minut (konfigurerbart) jamfor den verkliga produktionen registrerad i sparbarhetssystemet (SQL Server) med planerade kvantiteter fran Excel-planeringsfilen. Resultaten visas i en realtids-dashboard pa http://localhost:8085.",
    s1_features: "Huvudfunktioner:",
    s1_bullets: [
        "Realtidsjamforelse av planerad mot faktisk produktion",
        "Fargkodade statusindikatorer (gron / gul / rod)",
        "Automatisk identifiering av ordrar som produceras men inte finns i dagplanen",
        "Historisk kontextanalys (2 arbetsdagar bakot och 3 framat)",
        "Automatiska e-postvarningar med eskalerande braskande",
        "Autoscrollande dashboard designad for verkstadsskadarm"
    ],
    s2_title: "2. Datakalor",
    s2_excel_title: "Excel-planeringsfil",
    s2_excel_bullets: [
        "Plats: T:\\Planning (den senaste .xlsx-filen valjs automatiskt)",
        "Blad som anvands: PlanningMachine (konfigurerbart)",
        "Kolumn C: Ordernummer (punkt-tecken tas bort automatiskt)",
        "Kolumn E: Maskin-/fasnamn",
        "Kolumner M+: Produktionsdatum med planerade kvantiteter per dag",
        "Filen kontrolleras var 30:e minut; nya versioner laddas automatiskt"
    ],
    s2_db_title: "SQL Server-databas (Traceability_rs)",
    s2_db_bullets: [
        "Orders-tabell: kopplar ordernummer till IdOrder och ProductCode",
        "Phases-tabell: kopplar maskinnamn till IdPhase, med PhaseOrder for sortering",
        "Scannings-tabell: raknar faktisk produktion (07:30 till 07:30 nasta dag)",
        "ShapShots-tabell: periodiska ogonblicksbilder; bara IsChecked=0-rader bearbetas",
        "Settings-tabell: e-postmottagare (attribut: sys_email_planning_warning)"
    ],
    s2_planning_title: "TraceabilityPlanning_RS-databas",
    s2_planning_bullets: [
        "Machine + Phase-tabeller: mappar maskinnamn fran Excel till fas-ID:n"
    ],
    s3_title: "3. Korningsschema",
    s3_auto_title: "Automatiskt Schema",
    s3_auto_bullets: [
        "Arbetstid: 07:30 till 23:30 (konfigurerbart)",
        "Pollingintervall: var 10:e minut (konfigurerbart)",
        "Schemallaggaren startar automatiskt vid applikationsstart",
        "Utanfor arbetstid ar automatiska cykler avstangda",
        "Forsta cykeln kors omedelbart vid start"
    ],
    s3_cycle_title: "Varje Cykel Utfor:",
    s3_cycle_steps: [
        "Kontrollerar senaste Excel-filen (var 30 min, annars fran cache)",
        "Tolkar planeringsbladet och filtrerar for dagens datum",
        "Ansluter till SQL Server",
        "Infogar en produktions-snapshot (INSERT INTO ShapShots)",
        "Laser okontrollerade snapshots (IsChecked = 0)",
        "Jamfor faktisk produktion mot plan med tidsbaserad projektion",
        "Tilldelar status: Gron (i fas), Gul (efter), Rod (underskott > 10 st)",
        "Identifierar ordrar utanfor planen",
        "For ordrar utanfor planen: kontrollerar 2 dagar bakot och 3 framat",
        "Uppdaterar dashboarden",
        "Skickar e-postvarningar vid behov (respekterar cooldown)",
        "Markerar bearbetade snapshots som IsChecked = 1"
    ],
    s3_manual_title: "Manuell Korning",
    s3_manual_text: "Knappen \"Run Now\" utloser en omedelbar cykel utan tidsfonsterkontroll.",
    s4_title: "4. Projektionsformel",
    s4_formulas: [
        "Produktionsdag = 900 minuter (07:30 till 23:30)",
        "forlupna_minuter = minuter sedan 07:30",
        "andel = forlupna_minuter / 900",
        "Forvantad nu = Planerad kvantitet x andel",
        "Projektion dagsslut = Producerat / andel",
        "Projicerat underskott = Planerad kvantitet - Projektion dagsslut",
        "Alla varden avrundas till heltal (hela stycken)"
    ],
    s5_title: "5. Statusfarger",
    s5_table: [
        ["Farg", "Betydelse", "Villkor"],
        ["Gron", "I fas", "Projicerat underskott = 0"],
        ["Gul", "Varning", "Projicerat underskott mellan 1 och 10 stycken"],
        ["Rod", "Kritisk", "Projicerat underskott > 10 stycken"],
        ["Rod blinkande", "Utanfor plan", "Order i produktion men inte i dagens Excel-plan"]
    ],
    s6_title: "6. Stjarnindikatorer",
    s6_yellow: "Gul stjarna: order inte i dagens plan men schemallagd inom nasta 3 arbetsdagar",
    s6_blue: "Bla stjarna: order inte i dagens plan men var schemallagd inom senaste 2 arbetsdagar (forsening)",
    s7_title: "7. E-postvarningar",
    s7_when_title: "Nar E-post Skickas:",
    s7_when_bullets: ["Nar ordrar har Gul eller Rod status", "Nar ordrar utanfor planen upptacks"],
    s7_cooldown_title: "Cooldown-regler:",
    s7_cooldown_bullets: ["Gula varningar: max 1 e-post var 2:a timme", "Roda varningar: max 1 e-post varje timme", "Cooldown aterstalls varje ny dag"],
    s7_severity_title: "Allvarlighetsnivaoer:",
    s7_severity_bullets: ["Niva 1 (mattlig): bara gula varningar", "Niva 2 (bestaamd): roda larm eller ordrar utanfor plan", "Niva 3 (bradskande): roda larm har OKAT jamfort med foregaende cykel"],
    s7_recipients_title: "Mottagare:",
    s7_recipients_bullets: ["Lases fran: traceability_rs.dbo.settings (attribut = sys_email_planning_warning)", "Flera e-postadresser separerade med semikolon stods"],
    s7_content_title: "E-postinnehall:",
    s7_content_bullets: ["Skrivet pa engelska, professionell men bestaamd ton", "Sammanfattande rakneverk (gron/gul/rod/utanfor plan)", "Detaljerad tabell over problematiska ordrar", "Stjaernkontext for ordrar utanfor plan", "Kall-Excels filnamn och tidstaempel"],
    s8_title: "8. Dashboard-kolumner",
    s8_table: [
        ["Kolumn", "Beskrivning"],
        ["Order Number", "Produktionsorderidentifierare"],
        ["Product Code", "Produktkod fran databasen"],
        ["Phase", "Produktionsfasens namn"],
        ["Planning QTY/Day", "Planerad kvantitet for idag"],
        ["Qty Done", "Faktiskt producerat (med statusindikator och stjarna vid behov)"],
        ["Expected Now", "Stycken forvantade vid nuvarande tid"],
        ["Projected End", "Projicerad total vid dagens slut"],
        ["Deficit", "Projicerat underskott i stycken"]
    ],
    s9_title: "9. Konfiguration (config.yaml)",
    s9_intro: "Alla parametrar finns i config.yaml:",
    s9_params: [
        "planning.folder: Sokvag till Excel-filer (standard: T:\\Planning)",
        "planning.sheet: Bladnamn (standard: PlanningMachine)",
        "workday.start / end: Arbetstid (standard: 07:30 / 23:30)",
        "polling.interval_minutes: Cykelintervall (standard: 10)",
        "thresholds.red_deficit: Stycken for rod (standard: 10)",
        "email.enabled: Aktivera/inaktivera e-post",
        "email.yellow_cooldown_minutes: Gul cooldown (standard: 120)",
        "email.red_cooldown_minutes: Rod cooldown (standard: 60)",
        "server.port: Webbserverport (standard: 8085)"
    ]
}

};

// ===================== BUILD DOCUMENT =====================

const langOrder = ["ro", "en", "it", "sv"];
const allSections = [];

for (let i = 0; i < langOrder.length; i++) {
    const lang = langOrder[i];
    const t = translations[lang];

    const sectionChildren = [];

    // Language separator
    sectionChildren.push(new Paragraph({
        spacing: { after: 300 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `[ ${t.lang_label} ]`, bold: true, size: 28, color: BLUE, font: "Arial" })]
    }));

    sectionChildren.push(...generateSection(lang));

    allSections.push({
        properties: {
            page: {
                size: { width: 11906, height: 16838 },
                margin: { top: 1800, right: 1200, bottom: 1200, left: 1200 }
            }
        },
        headers: {
            default: new Header({
                children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new ImageRun({
                        type: "png",
                        data: logoData,
                        transformation: { width: 120, height: 50 },
                        altText: { title: "Logo", description: "Company Logo", name: "Logo" }
                    })]
                })]
            })
        },
        footers: {
            default: new Footer({
                children: [new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({ text: "Production Plan Monitor - ", size: 16, color: "888888", font: "Arial" }),
                        new TextRun({ text: t.subtitle, size: 16, color: "888888", font: "Arial" }),
                        new TextRun({ text: "  |  Page ", size: 16, color: "888888", font: "Arial" }),
                        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "888888", font: "Arial" })
                    ]
                })]
            })
        },
        children: sectionChildren
    });
}

const doc = new Document({
    styles: {
        default: { document: { run: { font: "Arial", size: 22 } } }
    },
    sections: allSections
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync("Production_Plan_Monitor_Guide.docx", buffer);
    console.log("Document created: Production_Plan_Monitor_Guide.docx (" + buffer.length + " bytes)");
});
