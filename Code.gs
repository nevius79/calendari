/**
 * Code.gs - Funzioni Principali di Sincronizzazione
 *
 * ✅ MIGLIORAMENTO FASE 1 - SICUREZZA:
 * - Credenziali rimosse dal codice
 * - Configurazione centralizzata in Config.gs
 * - Sistema sicuro con Properties Service
 */

function sincronizzaCalendariPerCorso() {
  // ✅ SICUREZZA: Usa configurazione centralizzata invece di CONFIG hardcoded
  const CONFIG = getConfig();
  const credentials = getCredentials(); // ✅ SICUREZZA: Credenziali dal Properties Service

  // Contatori globali e per corso
  const counters = {
    global: { creati: 0, aggiornati: 0, eliminati: 0, invariati: 0, errori: 0 },
    perCorso: new Map()
  };

  // Cache e strutture dati
  const calendariCreati = new Map();
  const eventiAPIperCalendario = new Map();
  const corsoColorMap = new Map();

  // Utility functions
  const toSeconds = date => Math.floor(date.getTime() / 1000);
  const formatDate = date => Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  // Inizializzazione
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Rileva se l'esecuzione è manuale o automatica
  let ui = null;
  let isAutomatic = false;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    // Esecuzione automatica - non possiamo usare UI
    isAutomatic = true;
    Logger.log("Esecuzione automatica rilevata");
  }

  // Setup foglio mappatura anni accademici
  function setupMappingSheetLocal() {
    let mappingSheet = ss.getSheetByName(CONFIG.sheets.mappingYears);
    if (!mappingSheet) {
      mappingSheet = ss.insertSheet(CONFIG.sheets.mappingYears);

      // Header
      mappingSheet.getRange(1, 1, 1, 3).setValues([["Codice", "Anno Accademico", "Note"]])
        .setFontWeight("bold")
        .setBackground("#1a73e8")
        .setFontColor("#ffffff");

      // Esempi predefiniti
      const currentYear = new Date().getFullYear();
      const defaultMappings = [];

      // Genera mappature per gli ultimi 5 anni e i prossimi 2
      for (let i = -5; i <= 2; i++) {
        const year = currentYear + i;
        defaultMappings.push([
          year.toString(),
          `${year}/${year + 1}`,
          i === 0 ? "Anno corrente" : ""
        ]);
      }

      mappingSheet.getRange(2, 1, defaultMappings.length, 3).setValues(defaultMappings);

      // Formattazione
      mappingSheet.getRange(2, 1, defaultMappings.length, 3)
        .setBorder(true, true, true, true, true, true)
        .setHorizontalAlignment("center");

      mappingSheet.setColumnWidth(1, 100);
      mappingSheet.setColumnWidth(2, 150);
      mappingSheet.setColumnWidth(3, 200);
    }

    return mappingSheet;
  }

  // Ottieni mappatura anni accademici
  function getYearMappingLocal() {
    const mappingSheet = setupMappingSheetLocal();
    const data = mappingSheet.getDataRange().getValues();
    const mapping = new Map();

    // Salta header
    for (let i = 1; i < data.length; i++) {
      const [codice, descrizione] = data[i];
      if (codice && descrizione) {
        mapping.set(codice.toString(), descrizione.toString());
      }
    }

    return mapping;
  }

  // Converti codice anno in descrizione
  function convertYearCodeLocal(codice, yearMapping) {
    if (!codice) return "Anno Sconosciuto";

    const codiceStr = codice.toString();

    // Controlla se esiste nella mappatura
    if (yearMapping.has(codiceStr)) {
      return yearMapping.get(codiceStr);
    }

    // Se non trovato, genera automaticamente
    const year = parseInt(codiceStr);
    if (!isNaN(year) && year > 2000 && year < 2100) {
      return `${year}/${year + 1}`;
    }

    return `Anno ${codiceStr}`;
  }

  // Crea o ottieni foglio progressi con design migliorato
  function setupProgressSheet() {
    let progressSheet = ss.getSheetByName(CONFIG.sheets.progress);
    if (!progressSheet) {
      progressSheet = ss.insertSheet(CONFIG.sheets.progress);
    }

    progressSheet.clear();
    progressSheet.setColumnWidths(1, 4, 150);

    // Header
    progressSheet.getRange(1, 1, 1, 4).merge()
      .setValue("🔄 SINCRONIZZAZIONE IN CORSO")
      .setFontSize(16)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#1a73e8")
      .setFontColor("#ffffff");

    // Progress bar area
    progressSheet.getRange(3, 1).setValue("Progresso:");
    progressSheet.getRange(3, 2, 1, 3).merge();

    // Status message
    progressSheet.getRange(5, 1).setValue("Stato:");
    progressSheet.getRange(5, 2, 1, 3).merge();

    // Details area
    progressSheet.getRange(7, 1).setValue("Dettagli:");
    progressSheet.getRange(8, 1, 1, 4).merge();

    // Time info
    progressSheet.getRange(10, 1).setValue("Tempo trascorso:");
    progressSheet.getRange(10, 2).setValue("0:00");
    progressSheet.getRange(11, 1).setValue("Tempo stimato:");
    progressSheet.getRange(11, 2).setValue("Calcolo...");

    // Formatting
    progressSheet.getRange(3, 1).setFontWeight("bold");
    progressSheet.getRange(5, 1).setFontWeight("bold");
    progressSheet.getRange(7, 1).setFontWeight("bold");
    progressSheet.getRange(10, 1).setFontWeight("bold");
    progressSheet.getRange(11, 1).setFontWeight("bold");

    return progressSheet;
  }

  // Aggiorna progress bar visuale
  function updateProgressBar(progressSheet, percentage, message, details = "") {
    const barLength = 20;
    const filled = Math.round(barLength * percentage / 100);
    const empty = barLength - filled;

    const progressBar = "█".repeat(filled) + "▒".repeat(empty);
    const progressText = `${progressBar} ${percentage}%`;

    progressSheet.getRange(3, 2).setValue(progressText)
      .setFontFamily("Courier New")
      .setFontSize(12);

    progressSheet.getRange(5, 2).setValue(message)
      .setFontSize(11)
      .setFontColor("#5f6368");

    if (details) {
      progressSheet.getRange(8, 1).setValue(details)
        .setFontSize(10)
        .setFontColor("#5f6368")
        .setWrap(true);
    }

    // Colora la barra in base alla percentuale
    let barColor = "#4285F4"; // Blu
    if (percentage >= 75) barColor = "#0F9D58"; // Verde
    else if (percentage >= 50) barColor = "#F4B400"; // Giallo
    else if (percentage >= 25) barColor = "#FF7043"; // Arancione

    progressSheet.getRange(3, 2).setFontColor(barColor);

    SpreadsheetApp.flush();
  }

  // Timer per tempo trascorso
  const startTime = new Date();
  function updateTimer(progressSheet) {
    const elapsed = new Date() - startTime;
    const minutes = Math.floor(elapsed / 60000);
    const seconds = Math.floor((elapsed % 60000) / 1000);
    progressSheet.getRange(10, 2).setValue(`${minutes}:${seconds.toString().padStart(2, '0')}`);
  }

  // Setup dashboard
  function setupDashboard() {
    let dashboard = ss.getSheetByName(CONFIG.sheets.dashboard);
    if (!dashboard) {
      dashboard = ss.insertSheet(CONFIG.sheets.dashboard);
    }

    dashboard.clear();
    dashboard.getRange(1, 1, 1, 7).merge()
      .setValue("📊 DASHBOARD SINCRONIZZAZIONE")
      .setFontSize(18)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#f8f9fa");

    dashboard.getRange(3, 1).setValue("Codice Corso");
    dashboard.getRange(3, 2).setValue("Corso");
    dashboard.getRange(3, 3).setValue("Anno Accademico");
    dashboard.getRange(3, 4).setValue("✅ Creati");
    dashboard.getRange(3, 5).setValue("🔄 Aggiornati");
    dashboard.getRange(3, 6).setValue("❌ Eliminati");
    dashboard.getRange(3, 7).setValue("⏭️ Invariati");

    dashboard.getRange(3, 1, 1, 7)
      .setFontWeight("bold")
      .setBackground("#e8eaed")
      .setHorizontalAlignment("center");

    return dashboard;
  }

  // Setup foglio mappatura corsi
  function setupCorsiMappingSheetLocal() {
    let mappingSheet = ss.getSheetByName(CONFIG.sheets.mappingCourses);
    if (!mappingSheet) {
      mappingSheet = ss.insertSheet(CONFIG.sheets.mappingCourses);

      // Header
      mappingSheet.getRange(1, 1, 1, 4).setValues([["Codice Corso", "Descrizione Breve", "Descrizione Completa", "Note"]])
        .setFontWeight("bold")
        .setBackground("#0f9d58")
        .setFontColor("#ffffff");

      // Istruzioni
      mappingSheet.getRange(2, 1, 1, 4).merge()
        .setValue("⚠️ Compila questa tabella per evitare calendari 'Corso Sconosciuto'. La colonna 'Descrizione Breve' verrà usata per il nome del calendario.")
        .setFontSize(10)
        .setFontStyle("italic")
        .setBackground("#fff3cd");

      // Formattazione colonne
      mappingSheet.setColumnWidth(1, 120);
      mappingSheet.setColumnWidth(2, 250);
      mappingSheet.setColumnWidth(3, 350);
      mappingSheet.setColumnWidth(4, 200);

      // Bordi
      mappingSheet.getRange(1, 1, 2, 4).setBorder(true, true, true, true, true, true);
    }

    return mappingSheet;
  }

  // Ottieni mappatura corsi
  function getCorsiMappingLocal() {
    const mappingSheet = setupCorsiMappingSheetLocal();
    const data = mappingSheet.getDataRange().getValues();
    const mapping = new Map();

    // Salta header e istruzioni
    for (let i = 2; i < data.length; i++) {
      const [codice, descrizioneBreve, descrizioneCompleta] = data[i];
      if (codice && descrizioneBreve) {
        mapping.set(codice.toString(), {
          breve: descrizioneBreve.toString(),
          completa: descrizioneCompleta ? descrizioneCompleta.toString() : descrizioneBreve.toString()
        });
      }
    }

    return mapping;
  }

  try {
    // Inizializza UI
    const progressSheet = setupProgressSheet();
    const dashboard = setupDashboard();
    const yearMapping = getYearMappingLocal();
    const corsiMapping = getCorsiMappingLocal();

    updateProgressBar(progressSheet, 0, "Inizializzazione...", "Preparazione ambiente di lavoro");

    // Setup foglio principale
    const sheet = ss.getSheetByName(CONFIG.sheets.main);
    sheet.clear();
    const headers = ["Corso", "Anno Accademico", "Semestre", "Titolo", "Docente", "Aula", "Data Inizio", "Data Fine", "Calendario", "Stato"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setBackground("#1a73e8")
      .setFontColor("#ffffff");

    // ✅ SICUREZZA: Autenticazione con credenziali sicure
    updateProgressBar(progressSheet, 5, "Autenticazione in corso...", "Connessione ai server SPD");

    const loginResponse = UrlFetchApp.fetch(CONFIG.api.loginUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        userName: credentials.user,  // ✅ Da Properties Service
        password: credentials.pass   // ✅ Da Properties Service
      }),
      muteHttpExceptions: true
    });

    const loginData = JSON.parse(loginResponse.getContentText());
    const token = loginData.token;

    if (!token) {
      throw new Error("Autenticazione fallita - Token non ottenuto");
    }

    updateProgressBar(progressSheet, 10, "Autenticazione completata", "Token ottenuto con successo");

    // Recupera TUTTI i dati con una singola chiamata API
    updateProgressBar(progressSheet, 15, "Recupero dati corsi...", "Download di tutti i dati dei corsi SPD");

    const response = UrlFetchApp.fetch(CONFIG.api.dataUrl, {
      method: "get",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Errore nel recupero dati: HTTP ${response.getResponseCode()}`);
    }

    const allData = JSON.parse(response.getContentText());
    updateProgressBar(progressSheet, 25, "Dati recuperati", `Elaborazione di ${allData.length} record`);

    const oggi = new Date();
    oggi.setHours(0, 0, 0, 0);
    const tra180giorni = new Date();
    tra180giorni.setDate(oggi.getDate() + CONFIG.sync.daysInFuture); // ✅ Da configurazione

    // Cache degli eventi esistenti per calendario
    const eventiCachePerCalendario = new Map();

    // Batch per inserimento righe
    const rowsToInsert = [];

    // Raggruppa i dati per codice corso e anno accademico
    const dataByCorsoAnno = new Map();
    const codiceToCorsoNome = new Map(); // Mappa codice corso -> nome più corto
    const corsoToCode = new Map(); // Mappa per memorizzare corso -> codice corso

    // Prima passa: determina il nome più corto per ogni codice corso
    const corsiSconosciuti = new Set();

    allData.forEach(lezione => {
      const corso = lezione.tb_decorso?.trim() || "";
      const codiceCorso = lezione.tb_codcorso?.toString() || "NOCODE";

      // Usa la mappatura se disponibile
      if (corsiMapping.has(codiceCorso)) {
        const mappedNome = corsiMapping.get(codiceCorso).breve;
        codiceToCorsoNome.set(codiceCorso, mappedNome);
      } else if (corso && corso !== "Corso Sconosciuto") {
        // Se non c'è mappatura, usa la logica del nome più corto
        if (!codiceToCorsoNome.has(codiceCorso)) {
          codiceToCorsoNome.set(codiceCorso, corso);
        } else {
          const nomeEsistente = codiceToCorsoNome.get(codiceCorso);
          if (corso.length < nomeEsistente.length) {
            codiceToCorsoNome.set(codiceCorso, corso);
          }
        }
      } else {
        // Corso sconosciuto - lo segnaleremo all'utente
        corsiSconosciuti.add(codiceCorso);
      }

      if (corso) {
        corsoToCode.set(corso, codiceCorso);
      }
    });

    // Seconda passa: raggruppa per codice corso e anno
    allData.forEach(lezione => {
      const inizio = new Date(lezione.tb_dallehh);
      const inizioSoloData = new Date(inizio);
      inizioSoloData.setHours(0, 0, 0, 0);

      // Salta se la lezione è nel passato
      if (inizioSoloData < oggi) return;

      const codiceCorso = lezione.tb_codcorso?.toString() || "NOCODE";
      const codiceAnno = lezione.co_codfamcorso || "";
      const annoAccademico = convertYearCodeLocal(codiceAnno, yearMapping);
      const key = `${codiceCorso}___${annoAccademico}`;

      if (!dataByCorsoAnno.has(key)) {
        dataByCorsoAnno.set(key, []);
      }
      dataByCorsoAnno.get(key).push(lezione);
    });

    // Processa dati raggruppati
    let processedGroups = 0;
    const totalGroups = dataByCorsoAnno.size;

    dataByCorsoAnno.forEach((lezioni, corsoAnnoKey) => {
      processedGroups++;
      const progressPercentage = 30 + (processedGroups / totalGroups * 50);

      const [codiceCorso, annoAccademico] = corsoAnnoKey.split("___");
      const corso = codiceToCorsoNome.get(codiceCorso) || `Corso ${codiceCorso}`;

      // Salta se è un corso sconosciuto senza mappatura
      if (!codiceToCorsoNome.has(codiceCorso)) {
        Logger.log(`⚠️ Corso sconosciuto saltato: codice ${codiceCorso}`);
        counters.global.errori++;
        return;
      }

      updateProgressBar(
        progressSheet,
        Math.round(progressPercentage),
        `Elaborazione ${corso} - ${annoAccademico}`,
        `Gruppo ${processedGroups} di ${totalGroups}`
      );

      updateTimer(progressSheet);

      // Nome calendario con codice corso, nome e anno accademico
      const nomeCalendario = `${CONFIG.calendarPrefix}${codiceCorso} - ${corso} - ${annoAccademico}`; // ✅ Prefisso da config
      const chiaveCalendario = corsoAnnoKey;

      // Inizializza contatori per corso se necessario
      if (!counters.perCorso.has(chiaveCalendario)) {
        counters.perCorso.set(chiaveCalendario, {
          codiceCorso: codiceCorso,
          corso: corso,
          annoAccademico: annoAccademico,
          creati: 0,
          aggiornati: 0,
          eliminati: 0,
          invariati: 0
        });
      }

      // Assegna colore al codice corso (non al nome)
      if (!corsoColorMap.has(codiceCorso)) {
        const colorIndex = corsoColorMap.size % CONFIG.courseColors.length;
        corsoColorMap.set(codiceCorso, CONFIG.courseColors[colorIndex]);
      }

      // Ottieni o crea calendario
      let calendario = calendariCreati.get(chiaveCalendario);
      if (!calendario) {
        const calendari = CalendarApp.getCalendarsByName(nomeCalendario);
        if (calendari.length > 0) {
          calendario = calendari[0];
        } else {
          calendario = CalendarApp.createCalendar(nomeCalendario);
          calendario.setColor(corsoColorMap.get(codiceCorso));
        }
        calendariCreati.set(chiaveCalendario, calendario);
        eventiAPIperCalendario.set(chiaveCalendario, new Set());

        // Carica eventi esistenti in cache
        const eventiEsistenti = calendario.getEvents(oggi, tra180giorni);
        const cache = new Map();
        eventiEsistenti.forEach(evento => {
          const key = `${evento.getTitle()}___${formatDate(evento.getStartTime())}`;
          if (!cache.has(key)) cache.set(key, []);
          cache.get(key).push(evento);
        });
        eventiCachePerCalendario.set(chiaveCalendario, cache);
      }

      // Processa lezioni per questo corso/anno
      lezioni.forEach(lezione => {
        try {
          const titoloModulo = lezione.tb_desmodulo?.trim() || "Lezione SPD";
          const docente = `${lezione.tb_nomdoc?.trim() || ""} ${lezione.tb_cogdoc?.trim() || ""}`.trim() || "Docente da definire";
          const aula = lezione.aula?.trim() || lezione.sede?.trim() || "Online";
          const inizio = new Date(lezione.tb_dallehh);
          const fine = new Date(lezione.tb_allehh);
          const semestre = lezione.tb_hhsemestre || "-";

          const apiKey = `${titoloModulo}___${inizio.toISOString()}___${fine.toISOString()}`;
          eventiAPIperCalendario.get(chiaveCalendario).add(apiKey);

          // Cerca in cache
          const cache = eventiCachePerCalendario.get(chiaveCalendario);
          const cacheKey = `${titoloModulo}___${formatDate(inizio)}`;
          const eventiInCache = cache.get(cacheKey) || [];

          const newStartSec = toSeconds(inizio);
          const newEndSec = toSeconds(fine);

          const eventoEsistente = eventiInCache.find(ev =>
            toSeconds(ev.getStartTime()) === newStartSec &&
            toSeconds(ev.getEndTime()) === newEndSec &&
            ev.getDescription() === docente &&
            ev.getLocation() === aula
          );

          let stato = "⏭️ Già presente";

          if (eventoEsistente) {
            counters.global.invariati++;
            counters.perCorso.get(chiaveCalendario).invariati++;
          } else {
            const eventoSimile = eventiInCache.find(ev =>
              ev.getTitle() === titoloModulo &&
              formatDate(ev.getStartTime()) === formatDate(inizio)
            );

            if (eventoSimile) {
              eventoSimile.setTime(inizio, fine);
              eventoSimile.setDescription(docente);
              eventoSimile.setLocation(aula);
              stato = "🔄 Aggiornato";
              counters.global.aggiornati++;
              counters.perCorso.get(chiaveCalendario).aggiornati++;
            } else {
              // Aggiungi note se presenti
              let descrizione = docente;
              if (lezione.tb_note && lezione.tb_note.trim()) {
                descrizione += `\n\nNote: ${lezione.tb_note.trim()}`;
              }

              calendario.createEvent(titoloModulo, inizio, fine, {
                description: descrizione,
                location: aula
              });
              stato = "✅ Creato";
              counters.global.creati++;
              counters.perCorso.get(chiaveCalendario).creati++;
            }
          }

          rowsToInsert.push({
            data: [corso, annoAccademico, semestre, titoloModulo, docente, aula, inizio, fine, nomeCalendario, stato],
            stato: stato
          });

        } catch (e) {
          counters.global.errori++;
          Logger.log(`❌ Errore elaborazione lezione: ${e}`);
          rowsToInsert.push({
            data: [corso, annoAccademico, "-", "Errore nel recupero dati", e.toString(), "-", new Date(), new Date(), nomeCalendario, "⚠️ Errore"],
            stato: "⚠️ Errore"
          });
        }
      });
    });

    // CALENDARIO UNIFICATO - Gestione dopo tutti i calendari per corso
    updateProgressBar(progressSheet, 78, "Creazione calendario unificato...", "Aggiornamento calendario con tutte le lezioni");

    const nomeCalendarioUnificato = CONFIG.unifiedCalendarName; // ✅ Da configurazione
    let calendarioUnificato = CalendarApp.getCalendarsByName(nomeCalendarioUnificato)[0];

    if (!calendarioUnificato) {
      calendarioUnificato = CalendarApp.createCalendar(nomeCalendarioUnificato);
      calendarioUnificato.setColor("#1a73e8"); // Blu Google
    }

    // Cache eventi esistenti calendario unificato
    const eventiUnificatiEsistenti = calendarioUnificato.getEvents(oggi, tra180giorni);
    const cacheUnificato = new Map();
    eventiUnificatiEsistenti.forEach(evento => {
      const key = `${evento.getTitle()}___${formatDate(evento.getStartTime())}`;
      if (!cacheUnificato.has(key)) cacheUnificato.set(key, []);
      cacheUnificato.get(key).push(evento);
    });

    const eventiAPIUnificato = new Set();
    let unificatoCreati = 0, unificatoAggiornati = 0, unificatoInvariati = 0;

    // Processa tutte le lezioni per il calendario unificato
    allData.forEach(lezione => {
      const inizio = new Date(lezione.tb_dallehh);
      const inizioSoloData = new Date(inizio);
      inizioSoloData.setHours(0, 0, 0, 0);

      if (inizioSoloData < oggi) return;

      const corso = lezione.tb_decorso?.trim() || "Corso Sconosciuto";
      const titoloModulo = lezione.tb_desmodulo?.trim() || "Lezione SPD";
      const docente = `${lezione.tb_nomdoc?.trim() || ""} ${lezione.tb_cogdoc?.trim() || ""}`.trim() || "Docente da definire";
      const aula = lezione.aula?.trim() || lezione.sede?.trim() || "Online";
      const fine = new Date(lezione.tb_allehh);
      const codiceAnno = lezione.co_codfamcorso || "";
      const annoAccademico = convertYearCodeLocal(codiceAnno, yearMapping);

      // Titolo evento con corso e anno accademico alla fine
      const titoloCompleto = `[${corso}] ${titoloModulo} (${annoAccademico})`;

      const apiKey = `${titoloCompleto}___${inizio.toISOString()}___${fine.toISOString()}`;
      eventiAPIUnificato.add(apiKey);

      const cacheKey = `${titoloCompleto}___${formatDate(inizio)}`;
      const eventiInCache = cacheUnificato.get(cacheKey) || [];

      const newStartSec = toSeconds(inizio);
      const newEndSec = toSeconds(fine);

      const eventoEsistente = eventiInCache.find(ev =>
        toSeconds(ev.getStartTime()) === newStartSec &&
        toSeconds(ev.getEndTime()) === newEndSec &&
        ev.getDescription() === docente &&
        ev.getLocation() === aula
      );

      if (eventoEsistente) {
        unificatoInvariati++;
      } else {
        const eventoSimile = eventiInCache.find(ev =>
          ev.getTitle() === titoloCompleto &&
          formatDate(ev.getStartTime()) === formatDate(inizio)
        );

        if (eventoSimile) {
          eventoSimile.setTime(inizio, fine);
          eventoSimile.setDescription(docente);
          eventoSimile.setLocation(aula);
          unificatoAggiornati++;
        } else {
          let descrizione = docente;
          if (lezione.tb_note && lezione.tb_note.trim()) {
            descrizione += `\n\nNote: ${lezione.tb_note.trim()}`;
          }

          calendarioUnificato.createEvent(titoloCompleto, inizio, fine, {
            description: descrizione,
            location: aula
          });
          unificatoCreati++;
        }
      }
    });

    // Elimina eventi obsoleti dal calendario unificato
    let unificatoEliminati = 0;
    eventiUnificatiEsistenti.forEach(evento => {
      const key = `${evento.getTitle()}___${evento.getStartTime().toISOString()}___${evento.getEndTime().toISOString()}`;
      if (!eventiAPIUnificato.has(key)) {
        evento.deleteEvent();
        unificatoEliminati++;
      }
    });

    // Inserimento batch righe
    updateProgressBar(progressSheet, 80, "Aggiornamento foglio dati...", "Scrittura risultati nel foglio");

    if (rowsToInsert.length > 0) {
      const values = rowsToInsert.map(row => row.data);
      const batchSize = CONFIG.sync.batchSize; // ✅ Da configurazione

      for (let i = 0; i < values.length; i += batchSize) {
        const batch = values.slice(i, i + batchSize);
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, batch.length, headers.length).setValues(batch);

        // Applica colori
        batch.forEach((_, index) => {
          const rowData = rowsToInsert[i + index];
          if (rowData.stato in CONFIG.calendarColors) {
            sheet.getRange(startRow + index, 1, 1, headers.length)
              .setBackground(CONFIG.calendarColors[rowData.stato]);
          }
        });
      }
    }

    updateProgressBar(progressSheet, 90, "Aggiornamento dashboard...", "Generazione statistiche");

    let dashboardRow = 4;
    counters.perCorso.forEach((stats, chiaveCalendario) => {
      dashboard.getRange(dashboardRow, 1, 1, 7).setValues([[
        stats.codiceCorso,
        stats.corso,
        stats.annoAccademico,
        stats.creati,
        stats.aggiornati,
        stats.eliminati,
        stats.invariati
      ]]);

      dashboard.getRange(dashboardRow, 2).setBackground(corsoColorMap.get(stats.codiceCorso))
        .setFontColor("#ffffff")
        .setFontWeight("bold");

      dashboardRow++;
    });

    // Riga calendario unificato
    dashboard.getRange(dashboardRow + 1, 1, 1, 7).setValues([[
      "-",
      "📅 CALENDARIO UNIFICATO",
      "Tutte le lezioni",
      unificatoCreati,
      unificatoAggiornati,
      unificatoEliminati,
      unificatoInvariati
    ]]);
    dashboard.getRange(dashboardRow + 1, 2).setBackground("#1a73e8")
      .setFontColor("#ffffff")
      .setFontWeight("bold");

    // Totali
    dashboard.getRange(dashboardRow + 2, 1).setValue("TOTALE")
      .setFontWeight("bold")
      .setBackground("#5f6368")
      .setFontColor("#ffffff");

    dashboard.getRange(dashboardRow + 2, 2, 1, 2).merge()
      .setValue("-")
      .setBackground("#e8eaed");

    dashboard.getRange(dashboardRow + 2, 4, 1, 4).setValues([[
      counters.global.creati + unificatoCreati,
      counters.global.aggiornati + unificatoAggiornati,
      counters.global.eliminati + unificatoEliminati,
      counters.global.invariati + unificatoInvariati
    ]])
    .setFontWeight("bold")
    .setBackground("#e8eaed");

    // Log
    updateProgressBar(progressSheet, 95, "Salvataggio log...", "Registrazione attività");

    const logSheet = ss.getSheetByName(CONFIG.sheets.log) || ss.insertSheet(CONFIG.sheets.log);
    if (logSheet.getLastRow() === 0) {
      logSheet.getRange(1, 1, 1, 6).setValues([["Data/Ora", "✅ Creati", "🔄 Aggiornati", "❌ Eliminati", "⏭️ Invariati", "⚠️ Errori"]])
        .setFontWeight("bold")
        .setBackground("#f8f9fa");
    }

    logSheet.appendRow([
      new Date(),
      counters.global.creati,
      counters.global.aggiornati,
      counters.global.eliminati,
      counters.global.invariati,
      counters.global.errori
    ]);

    // Completamento
    updateProgressBar(progressSheet, 100, "✅ Sincronizzazione completata!",
      `Processati ${allData.length} record in ${calendariCreati.size} calendari`);

    // Mostra riepilogo
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);

    let riepilogo = `
📊 RIEPILOGO SINCRONIZZAZIONE

⏱️ Tempo impiegato: ${Math.floor(duration / 60)}m ${duration % 60}s
📅 Calendari gestiti: ${calendariCreati.size + 1}
📚 Record processati: ${allData.length}

✅ Eventi creati: ${counters.global.creati + unificatoCreati}
🔄 Eventi aggiornati: ${counters.global.aggiornati + unificatoAggiornati}
⏭️ Eventi invariati: ${counters.global.invariati + unificatoInvariati}
⚠️ Errori/Corsi saltati: ${counters.global.errori}

Consulta il foglio "Dashboard" per i dettagli per corso e anno accademico.
`;

    // Avviso corsi sconosciuti
    if (corsiSconosciuti.size > 0) {
      riepilogo += `\n\n⚠️ ATTENZIONE: ${corsiSconosciuti.size} codici corso non hanno una descrizione valida.\n`;
      riepilogo += `Compila il foglio "Mappatura Corsi" per includerli nella prossima sincronizzazione.\n`;
      riepilogo += `Codici saltati: ${Array.from(corsiSconosciuti).join(", ")}`;
    }

    // Mostra riepilogo solo se esecuzione manuale
    if (!isAutomatic) {
      ui.alert('Sincronizzazione Completata', riepilogo, ui.ButtonSet.OK);
    } else {
      // In esecuzione automatica, registra nel log
      Logger.log("SINCRONIZZAZIONE COMPLETATA");
      Logger.log(riepilogo);

      // Opzionale: invia email di riepilogo
      const email = Session.getActiveUser().getEmail();
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: "SPD - Sincronizzazione Calendari Completata",
          body: riepilogo
        });
      }
    }

  } catch (error) {
    const errorMsg = `❌ ERRORE: ${error.toString()}`;
    if (progressSheet) {
      updateProgressBar(progressSheet, 0, errorMsg, "Sincronizzazione interrotta");
    }
    Logger.log(`❌ Errore critico: ${error}`);

    if (!isAutomatic) {
      ui.alert('Errore', `Si è verificato un errore critico:\n\n${error.toString()}`, ui.ButtonSet.OK);
    } else {
      // In esecuzione automatica, invia email di errore
      const email = Session.getActiveUser().getEmail();
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: "SPD - ERRORE Sincronizzazione Calendari",
          body: `Si è verificato un errore durante la sincronizzazione automatica:\n\n${error.toString()}\n\nControlla il log per maggiori dettagli.`
        });
      }
    }
  }
}

// NOTA: Le altre funzioni (sincronizzaCalendariDocenti, aggiornaAnniCalendari, ecc.)
// saranno migliorate nelle prossime fasi. Per ora manteniamo solo la funzione principale aggiornata.
