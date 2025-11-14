/**
 * ====================================================================
 * CALENDARI SPD - Sistema di Sincronizzazione Calendari Google
 * ====================================================================
 *
 * VERSIONE MIGLIORATA - FASE 1: SICUREZZA
 * ✅ Credenziali protette (non più hardcoded)
 * ✅ Configurazione centralizzata
 * ✅ Sistema di setup guidato
 *
 * PRIMA CONFIGURAZIONE:
 * 1. Menu: 🔄 Sincronizzazione SPD → ⚙️ Configura Credenziali
 * 2. Inserisci username e password quando richiesto
 * 3. Le credenziali vengono salvate in modo sicuro
 * 4. Procedi con la sincronizzazione normale
 *
 * ====================================================================
 */


// ====================================================================
// SEZIONE 1: CONFIGURAZIONE E SICUREZZA
// ====================================================================

/**
 * Ottiene la configurazione globale dell'applicazione
 * @returns {Object} Configurazione completa
 */
function getConfig() {
  return {
    // Fogli di lavoro
    sheets: {
      main: "Foglio1",
      log: "Log",
      progress: "Progressi",
      dashboard: "Dashboard",
      mappingYears: "Mappatura Anni",
      mappingCourses: "Mappatura Corsi",
      teachersLessons: "Lezioni Docenti",
      teachersDashboard: "Dashboard Docenti",
      teachersProgress: "Progressi Docenti",
      updateProgress: "Progressi Aggiornamento",
      updateLog: "Log Aggiornamento Anni"
    },

    // Colori per stati calendario
    calendarColors: {
      "✅ Creato": "#d4edda",
      "🔄 Aggiornato": "#cce5ff",
      "⏭️ Già presente": "#f8f9fa",
      "❌ Eliminato": "#f8d7da",
      "⚠️ Errore": "#fff3cd"
    },

    // Colori per corsi (ciclici)
    courseColors: [
      "#4285F4", "#DB4437", "#F4B400", "#0F9D58", "#AB47BC",
      "#00ACC1", "#FF7043", "#9E9D24", "#5C6BC0", "#F06292"
    ],

    // Colori per docenti (ciclici)
    teacherColors: [
      "#4285F4", "#DB4437", "#F4B400", "#0F9D58", "#AB47BC",
      "#00ACC1", "#FF7043", "#9E9D24", "#5C6BC0", "#F06292",
      "#8E24AA", "#D81B60", "#43A047", "#FB8C00", "#039BE5"
    ],

    // Impostazioni API
    api: {
      loginUrl: "https://restauth.mtsinformatica.com/api/Account/GetPartecipante",
      dataUrl: "https://restpartecipanti.mtsinformatica.com/api/Partecipanti/getallLezioniCorso/SPD/0",
      timeout: 30000,
      maxRetries: 3,
      retryDelay: 2000
    },

    // Impostazioni sincronizzazione
    sync: {
      daysInFuture: 180,
      batchSize: 100,
      autoSyncHours: 4
    },

    // Prefissi calendari
    calendarPrefix: "SPD - ",
    unifiedCalendarName: "SPD - Tutte le Lezioni",
    teacherCalendarPrefix: "SPD - Docente - "
  };
}

/**
 * Ottiene le credenziali dal Properties Service
 * SICURO: Non espone mai le credenziali nel codice
 */
function getCredentials() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const user = scriptProperties.getProperty('SPD_API_USER');
  const pass = scriptProperties.getProperty('SPD_API_PASS');

  if (!user || !pass) {
    throw new Error(
      '❌ CREDENZIALI NON CONFIGURATE!\n\n' +
      'Esegui la funzione setupCredentials() dal menu:\n' +
      '🔄 Sincronizzazione SPD → ⚙️ Configura Credenziali\n\n' +
      'Oppure esegui manualmente la funzione "setupCredentials"'
    );
  }

  return { user: user, pass: pass };
}

/**
 * Salva le credenziali in modo sicuro
 */
function saveCredentials(username, password) {
  if (!username || !password) {
    throw new Error('Username e password sono obbligatori');
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SPD_API_USER', username);
  scriptProperties.setProperty('SPD_API_PASS', password);

  Logger.log('✅ Credenziali salvate con successo');
}

/**
 * Rimuove le credenziali salvate
 */
function deleteCredentials() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('SPD_API_USER');
  scriptProperties.deleteProperty('SPD_API_PASS');
  Logger.log('🗑️ Credenziali rimosse');
}

/**
 * Verifica se le credenziali sono configurate
 */
function areCredentialsConfigured() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const user = scriptProperties.getProperty('SPD_API_USER');
  const pass = scriptProperties.getProperty('SPD_API_PASS');
  return !!(user && pass);
}

/**
 * Funzione interattiva per configurare le credenziali
 */
function setupCredentials() {
  const ui = SpreadsheetApp.getUi();

  if (areCredentialsConfigured()) {
    const response = ui.alert(
      '⚠️ Credenziali già configurate',
      'Le credenziali API sono già state configurate.\n\nVuoi aggiornarle?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  const userResponse = ui.prompt(
    '🔐 Configurazione - Step 1/2',
    'Inserisci lo USERNAME per l\'API SPD:',
    ui.ButtonSet.OK_CANCEL
  );

  if (userResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('❌ Configurazione annullata');
    return;
  }

  const username = userResponse.getResponseText().trim();
  if (!username) {
    ui.alert('❌ Errore', 'Username non può essere vuoto', ui.ButtonSet.OK);
    return;
  }

  const passResponse = ui.prompt(
    '🔐 Configurazione - Step 2/2',
    'Inserisci la PASSWORD per l\'API SPD:\n\n⚠️ Verrà salvata in modo sicuro.',
    ui.ButtonSet.OK_CANCEL
  );

  if (passResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('❌ Configurazione annullata');
    return;
  }

  const password = passResponse.getResponseText().trim();
  if (!password) {
    ui.alert('❌ Errore', 'Password non può essere vuota', ui.ButtonSet.OK);
    return;
  }

  try {
    saveCredentials(username, password);
    ui.alert(
      '✅ Configurazione Completata',
      'Le credenziali sono state salvate con successo!\n\n' +
      '🔒 Archiviate in modo sicuro\n' +
      'Ora puoi procedere con la sincronizzazione.',
      ui.ButtonSet.OK
    );

    const testResult = ui.alert(
      '🧪 Test Credenziali',
      'Vuoi testare le credenziali?',
      ui.ButtonSet.YES_NO
    );
    if (testResult === ui.Button.YES) {
      testCredentials();
    }
  } catch (error) {
    ui.alert('❌ Errore', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Testa le credenziali
 */
function testCredentials() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('🔄 Test in corso...', 'Test API', -1);

    const credentials = getCredentials();
    const config = getConfig();

    const response = UrlFetchApp.fetch(config.api.loginUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        userName: credentials.user,
        password: credentials.pass
      }),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseData = JSON.parse(response.getContentText());

    ss.toast('', '', 1);

    if (responseCode === 200 && responseData.token) {
      ui.alert(
        '✅ Test SUCCESSO',
        'Le credenziali sono corrette!\n\n' +
        '🔑 Token ricevuto\n' +
        '👤 Utente: ' + credentials.user + '\n\n' +
        'Sistema pronto.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '❌ Test FALLITO',
        'Credenziali non corrette.\n\n' +
        'Codice: ' + responseCode + '\n' +
        'Verifica e riprova.',
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    ui.alert('❌ Errore Test', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Mostra stato configurazione
 */
function showConfigStatus() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();

  let message = '📊 STATO CONFIGURAZIONE\n\n';

  message += '🔐 Credenziali API:\n';
  if (areCredentialsConfigured()) {
    const cred = getCredentials();
    message += `   ✅ Configurate\n`;
    message += `   👤 Username: ${cred.user}\n`;
    message += `   🔑 Password: ${'*'.repeat(12)}\n\n`;
  } else {
    message += `   ❌ NON configurate\n\n`;
  }

  message += '⚙️ Impostazioni:\n';
  message += `   📅 Giorni futuri: ${config.sync.daysInFuture}\n`;
  message += `   ⏰ Auto-sync: ogni ${config.sync.autoSyncHours} ore\n\n`;

  message += areCredentialsConfigured()
    ? '✅ Sistema pronto'
    : '⚠️ Configura le credenziali';

  ui.alert('Stato Configurazione', message, ui.ButtonSet.OK);
}


// ====================================================================
// SEZIONE 2: FUNZIONI HELPER GLOBALI
// ====================================================================

function getYearMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mappingSheet = setupMappingSheet();
  const data = mappingSheet.getDataRange().getValues();
  const mapping = new Map();

  for (let i = 1; i < data.length; i++) {
    const [codice, descrizione] = data[i];
    if (codice && descrizione) {
      mapping.set(codice.toString(), descrizione.toString());
    }
  }

  return mapping;
}

function setupMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const CONFIG = getConfig();

  let mappingSheet = ss.getSheetByName(CONFIG.sheets.mappingYears);
  if (!mappingSheet) {
    mappingSheet = ss.insertSheet(CONFIG.sheets.mappingYears);

    mappingSheet.getRange(1, 1, 1, 3).setValues([["Codice", "Anno Accademico", "Note"]])
      .setFontWeight("bold")
      .setBackground("#1a73e8")
      .setFontColor("#ffffff");

    const currentYear = new Date().getFullYear();
    const defaultMappings = [];

    for (let i = -5; i <= 2; i++) {
      const year = currentYear + i;
      defaultMappings.push([
        year.toString(),
        `${year}/${year + 1}`,
        i === 0 ? "Anno corrente" : ""
      ]);
    }

    mappingSheet.getRange(2, 1, defaultMappings.length, 3).setValues(defaultMappings);
    mappingSheet.getRange(2, 1, defaultMappings.length, 3)
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment("center");

    mappingSheet.setColumnWidth(1, 100);
    mappingSheet.setColumnWidth(2, 150);
    mappingSheet.setColumnWidth(3, 200);
  }

  return mappingSheet;
}

function convertYearCode(codice, yearMapping) {
  if (!codice) return "Anno Sconosciuto";

  const codiceStr = codice.toString();

  if (yearMapping.has(codiceStr)) {
    return yearMapping.get(codiceStr);
  }

  const year = parseInt(codiceStr);
  if (!isNaN(year) && year > 2000 && year < 2100) {
    return `${year}/${year + 1}`;
  }

  return `Anno ${codiceStr}`;
}

function setupCorsiMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const CONFIG = getConfig();

  let mappingSheet = ss.getSheetByName(CONFIG.sheets.mappingCourses);
  if (!mappingSheet) {
    mappingSheet = ss.insertSheet(CONFIG.sheets.mappingCourses);

    mappingSheet.getRange(1, 1, 1, 4).setValues([["Codice Corso", "Descrizione Breve", "Descrizione Completa", "Note"]])
      .setFontWeight("bold")
      .setBackground("#0f9d58")
      .setFontColor("#ffffff");

    mappingSheet.getRange(2, 1, 1, 4).merge()
      .setValue("⚠️ Compila questa tabella per evitare calendari 'Corso Sconosciuto'.")
      .setFontSize(10)
      .setFontStyle("italic")
      .setBackground("#fff3cd");

    mappingSheet.setColumnWidth(1, 120);
    mappingSheet.setColumnWidth(2, 250);
    mappingSheet.setColumnWidth(3, 350);
    mappingSheet.setColumnWidth(4, 200);

    mappingSheet.getRange(1, 1, 2, 4).setBorder(true, true, true, true, true, true);
  }

  return mappingSheet;
}

function getCorsiMapping() {
  const mappingSheet = setupCorsiMappingSheet();
  const data = mappingSheet.getDataRange().getValues();
  const mapping = new Map();

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


// ====================================================================
// SEZIONE 3: SINCRONIZZAZIONE CALENDARI PER CORSO
// ====================================================================

function sincronizzaCalendariPerCorso() {
  const CONFIG = getConfig();
  const credentials = getCredentials(); // ✅ CREDENZIALI SICURE

  const counters = {
    global: { creati: 0, aggiornati: 0, eliminati: 0, invariati: 0, errori: 0 },
    perCorso: new Map()
  };

  const calendariCreati = new Map();
  const eventiAPIperCalendario = new Map();
  const corsoColorMap = new Map();

  const toSeconds = date => Math.floor(date.getTime() / 1000);
  const formatDate = date => Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let ui = null;
  let isAutomatic = false;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    isAutomatic = true;
    Logger.log("Esecuzione automatica rilevata");
  }

  function setupProgressSheet() {
    let progressSheet = ss.getSheetByName(CONFIG.sheets.progress);
    if (!progressSheet) {
      progressSheet = ss.insertSheet(CONFIG.sheets.progress);
    }

    progressSheet.clear();
    progressSheet.setColumnWidths(1, 4, 150);

    progressSheet.getRange(1, 1, 1, 4).merge()
      .setValue("🔄 SINCRONIZZAZIONE IN CORSO")
      .setFontSize(16)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#1a73e8")
      .setFontColor("#ffffff");

    progressSheet.getRange(3, 1).setValue("Progresso:");
    progressSheet.getRange(3, 2, 1, 3).merge();
    progressSheet.getRange(5, 1).setValue("Stato:");
    progressSheet.getRange(5, 2, 1, 3).merge();
    progressSheet.getRange(7, 1).setValue("Dettagli:");
    progressSheet.getRange(8, 1, 1, 4).merge();
    progressSheet.getRange(10, 1).setValue("Tempo trascorso:");
    progressSheet.getRange(10, 2).setValue("0:00");
    progressSheet.getRange(11, 1).setValue("Tempo stimato:");
    progressSheet.getRange(11, 2).setValue("Calcolo...");

    progressSheet.getRange(3, 1).setFontWeight("bold");
    progressSheet.getRange(5, 1).setFontWeight("bold");
    progressSheet.getRange(7, 1).setFontWeight("bold");
    progressSheet.getRange(10, 1).setFontWeight("bold");
    progressSheet.getRange(11, 1).setFontWeight("bold");

    return progressSheet;
  }

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

    let barColor = "#4285F4";
    if (percentage >= 75) barColor = "#0F9D58";
    else if (percentage >= 50) barColor = "#F4B400";
    else if (percentage >= 25) barColor = "#FF7043";

    progressSheet.getRange(3, 2).setFontColor(barColor);
    SpreadsheetApp.flush();
  }

  const startTime = new Date();
  function updateTimer(progressSheet) {
    const elapsed = new Date() - startTime;
    const minutes = Math.floor(elapsed / 60000);
    const seconds = Math.floor((elapsed % 60000) / 1000);
    progressSheet.getRange(10, 2).setValue(`${minutes}:${seconds.toString().padStart(2, '0')}`);
  }

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

  try {
    const progressSheet = setupProgressSheet();
    const dashboard = setupDashboard();
    const yearMapping = getYearMapping();
    const corsiMapping = getCorsiMapping();

    updateProgressBar(progressSheet, 0, "Inizializzazione...", "Preparazione ambiente");

    const sheet = ss.getSheetByName(CONFIG.sheets.main);
    sheet.clear();
    const headers = ["Corso", "Anno Accademico", "Semestre", "Titolo", "Docente", "Aula", "Data Inizio", "Data Fine", "Calendario", "Stato"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setBackground("#1a73e8")
      .setFontColor("#ffffff");

    updateProgressBar(progressSheet, 5, "Autenticazione...", "Connessione SPD");

    const loginResponse = UrlFetchApp.fetch(CONFIG.api.loginUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ userName: credentials.user, password: credentials.pass }),
      muteHttpExceptions: true
    });

    const loginData = JSON.parse(loginResponse.getContentText());
    const token = loginData.token;

    if (!token) {
      throw new Error("Autenticazione fallita");
    }

    updateProgressBar(progressSheet, 10, "Autenticazione completata", "Token ottenuto");

    updateProgressBar(progressSheet, 15, "Recupero dati...", "Download dati SPD");

    const response = UrlFetchApp.fetch(CONFIG.api.dataUrl, {
      method: "get",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Errore recupero dati: HTTP ${response.getResponseCode()}`);
    }

    const allData = JSON.parse(response.getContentText());
    updateProgressBar(progressSheet, 25, "Dati recuperati", `${allData.length} record`);

    const oggi = new Date();
    oggi.setHours(0, 0, 0, 0);
    const tra180giorni = new Date();
    tra180giorni.setDate(oggi.getDate() + CONFIG.sync.daysInFuture);

    const eventiCachePerCalendario = new Map();
    const rowsToInsert = [];
    const dataByCorsoAnno = new Map();
    const codiceToCorsoNome = new Map();
    const corsoToCode = new Map();
    const corsiSconosciuti = new Set();

    allData.forEach(lezione => {
      const corso = lezione.tb_decorso?.trim() || "";
      const codiceCorso = lezione.tb_codcorso?.toString() || "NOCODE";

      if (corsiMapping.has(codiceCorso)) {
        const mappedNome = corsiMapping.get(codiceCorso).breve;
        codiceToCorsoNome.set(codiceCorso, mappedNome);
      } else if (corso && corso !== "Corso Sconosciuto") {
        if (!codiceToCorsoNome.has(codiceCorso)) {
          codiceToCorsoNome.set(codiceCorso, corso);
        } else {
          const nomeEsistente = codiceToCorsoNome.get(codiceCorso);
          if (corso.length < nomeEsistente.length) {
            codiceToCorsoNome.set(codiceCorso, corso);
          }
        }
      } else {
        corsiSconosciuti.add(codiceCorso);
      }

      if (corso) {
        corsoToCode.set(corso, codiceCorso);
      }
    });

    allData.forEach(lezione => {
      const inizio = new Date(lezione.tb_dallehh);
      const inizioSoloData = new Date(inizio);
      inizioSoloData.setHours(0, 0, 0, 0);

      if (inizioSoloData < oggi) return;

      const codiceCorso = lezione.tb_codcorso?.toString() || "NOCODE";
      const codiceAnno = lezione.co_codfamcorso || "";
      const annoAccademico = convertYearCode(codiceAnno, yearMapping);
      const key = `${codiceCorso}___${annoAccademico}`;

      if (!dataByCorsoAnno.has(key)) {
        dataByCorsoAnno.set(key, []);
      }
      dataByCorsoAnno.get(key).push(lezione);
    });

    let processedGroups = 0;
    const totalGroups = dataByCorsoAnno.size;

    dataByCorsoAnno.forEach((lezioni, corsoAnnoKey) => {
      processedGroups++;
      const progressPercentage = 30 + (processedGroups / totalGroups * 50);

      const [codiceCorso, annoAccademico] = corsoAnnoKey.split("___");
      const corso = codiceToCorsoNome.get(codiceCorso) || `Corso ${codiceCorso}`;

      if (!codiceToCorsoNome.has(codiceCorso)) {
        Logger.log(`⚠️ Corso sconosciuto: ${codiceCorso}`);
        counters.global.errori++;
        return;
      }

      updateProgressBar(
        progressSheet,
        Math.round(progressPercentage),
        `Elaborazione ${corso}`,
        `${processedGroups}/${totalGroups}`
      );

      updateTimer(progressSheet);

      const nomeCalendario = `${CONFIG.calendarPrefix}${codiceCorso} - ${corso} - ${annoAccademico}`;
      const chiaveCalendario = corsoAnnoKey;

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

      if (!corsoColorMap.has(codiceCorso)) {
        const colorIndex = corsoColorMap.size % CONFIG.courseColors.length;
        corsoColorMap.set(codiceCorso, CONFIG.courseColors[colorIndex]);
      }

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

        const eventiEsistenti = calendario.getEvents(oggi, tra180giorni);
        const cache = new Map();
        eventiEsistenti.forEach(evento => {
          const key = `${evento.getTitle()}___${formatDate(evento.getStartTime())}`;
          if (!cache.has(key)) cache.set(key, []);
          cache.get(key).push(evento);
        });
        eventiCachePerCalendario.set(chiaveCalendario, cache);
      }

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
          Logger.log(`❌ Errore: ${e}`);
          rowsToInsert.push({
            data: [corso, annoAccademico, "-", "Errore", e.toString(), "-", new Date(), new Date(), nomeCalendario, "⚠️ Errore"],
            stato: "⚠️ Errore"
          });
        }
      });
    });

    updateProgressBar(progressSheet, 78, "Calendario unificato...", "Aggiornamento");

    const nomeCalendarioUnificato = CONFIG.unifiedCalendarName;
    let calendarioUnificato = CalendarApp.getCalendarsByName(nomeCalendarioUnificato)[0];

    if (!calendarioUnificato) {
      calendarioUnificato = CalendarApp.createCalendar(nomeCalendarioUnificato);
      calendarioUnificato.setColor("#1a73e8");
    }

    const eventiUnificatiEsistenti = calendarioUnificato.getEvents(oggi, tra180giorni);
    const cacheUnificato = new Map();
    eventiUnificatiEsistenti.forEach(evento => {
      const key = `${evento.getTitle()}___${formatDate(evento.getStartTime())}`;
      if (!cacheUnificato.has(key)) cacheUnificato.set(key, []);
      cacheUnificato.get(key).push(evento);
    });

    const eventiAPIUnificato = new Set();
    let unificatoCreati = 0, unificatoAggiornati = 0, unificatoInvariati = 0;

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
      const annoAccademico = convertYearCode(codiceAnno, yearMapping);

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

    let unificatoEliminati = 0;
    eventiUnificatiEsistenti.forEach(evento => {
      const key = `${evento.getTitle()}___${evento.getStartTime().toISOString()}___${evento.getEndTime().toISOString()}`;
      if (!eventiAPIUnificato.has(key)) {
        evento.deleteEvent();
        unificatoEliminati++;
      }
    });

    updateProgressBar(progressSheet, 80, "Aggiornamento foglio...", "Scrittura risultati");

    if (rowsToInsert.length > 0) {
      const values = rowsToInsert.map(row => row.data);
      const batchSize = CONFIG.sync.batchSize;

      for (let i = 0; i < values.length; i += batchSize) {
        const batch = values.slice(i, i + batchSize);
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, batch.length, headers.length).setValues(batch);

        batch.forEach((_, index) => {
          const rowData = rowsToInsert[i + index];
          if (rowData.stato in CONFIG.calendarColors) {
            sheet.getRange(startRow + index, 1, 1, headers.length)
              .setBackground(CONFIG.calendarColors[rowData.stato]);
          }
        });
      }
    }

    updateProgressBar(progressSheet, 90, "Dashboard...", "Statistiche");

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

    updateProgressBar(progressSheet, 95, "Salvataggio log...", "Registrazione");

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

    updateProgressBar(progressSheet, 100, "✅ Completato!",
      `${allData.length} record in ${calendariCreati.size} calendari`);

    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);

    let riepilogo = `
📊 RIEPILOGO

⏱️ Tempo: ${Math.floor(duration / 60)}m ${duration % 60}s
📅 Calendari: ${calendariCreati.size + 1}
📚 Record: ${allData.length}

✅ Creati: ${counters.global.creati + unificatoCreati}
🔄 Aggiornati: ${counters.global.aggiornati + unificatoAggiornati}
⏭️ Invariati: ${counters.global.invariati + unificatoInvariati}
⚠️ Errori: ${counters.global.errori}
`;

    if (corsiSconosciuti.size > 0) {
      riepilogo += `\n⚠️ ${corsiSconosciuti.size} codici corso sconosciuti.\n`;
      riepilogo += `Compila "Mappatura Corsi".\n`;
      riepilogo += `Codici: ${Array.from(corsiSconosciuti).join(", ")}`;
    }

    if (!isAutomatic) {
      ui.alert('Completato', riepilogo, ui.ButtonSet.OK);
    } else {
      Logger.log("COMPLETATO");
      Logger.log(riepilogo);

      const email = Session.getActiveUser().getEmail();
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: "SPD - Sincronizzazione Completata",
          body: riepilogo
        });
      }
    }

  } catch (error) {
    const errorMsg = `❌ ERRORE: ${error.toString()}`;
    if (progressSheet) {
      updateProgressBar(progressSheet, 0, errorMsg, "Interrotto");
    }
    Logger.log(`❌ Errore: ${error}`);

    if (!isAutomatic) {
      ui.alert('Errore', error.toString(), ui.ButtonSet.OK);
    } else {
      const email = Session.getActiveUser().getEmail();
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: "SPD - ERRORE Sincronizzazione",
          body: `Errore:\n\n${error.toString()}`
        });
      }
    }
  }
}


// ====================================================================
// SEZIONE 4: SINCRONIZZAZIONE CALENDARI DOCENTI
// ====================================================================

function sincronizzaCalendariDocenti() {
  const CONFIG = getConfig();
  const credentials = getCredentials(); // ✅ CREDENZIALI SICURE

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  function formatNomeDocente(nome, cognome) {
    const formatName = (str) => {
      if (!str) return "";
      return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
    };

    const nomeFormattato = formatName(nome);
    const cognomeFormattato = formatName(cognome);

    return `${nomeFormattato} ${cognomeFormattato}`.trim();
  }

  const toSeconds = date => Math.floor(date.getTime() / 1000);
  const formatDate = date => Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  function setupProgressSheet() {
    let progressSheet = ss.getSheetByName(CONFIG.sheets.teachersProgress);
    if (!progressSheet) {
      progressSheet = ss.insertSheet(CONFIG.sheets.teachersProgress);
    }

    progressSheet.clear();
    progressSheet.setColumnWidths(1, 4, 150);

    progressSheet.getRange(1, 1, 1, 4).merge()
      .setValue("👥 SINCRONIZZAZIONE DOCENTI")
      .setFontSize(16)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#34a853")
      .setFontColor("#ffffff");

    progressSheet.getRange(3, 1).setValue("Progresso:");
    progressSheet.getRange(3, 2, 1, 3).merge();
    progressSheet.getRange(5, 1).setValue("Stato:");
    progressSheet.getRange(5, 2, 1, 3).merge();
    progressSheet.getRange(7, 1).setValue("Dettagli:");
    progressSheet.getRange(8, 1, 1, 4).merge();

    return progressSheet;
  }

  function updateProgress(progressSheet, percentage, message, details = "") {
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

    SpreadsheetApp.flush();
  }

  try {
    const progressSheet = setupProgressSheet();
    const yearMapping = getYearMapping();

    updateProgress(progressSheet, 0, "Inizializzazione...", "Preparazione");

    const sheet = ss.getSheetByName(CONFIG.sheets.teachersLessons) || ss.insertSheet(CONFIG.sheets.teachersLessons);
    sheet.clear();
    const headers = ["Docente", "Corso", "Anno Accademico", "Titolo", "Aula", "Data Inizio", "Data Fine", "Calendario", "Stato"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setBackground("#34a853")
      .setFontColor("#ffffff");

    updateProgress(progressSheet, 10, "Autenticazione...", "Connessione SPD");

    const loginResponse = UrlFetchApp.fetch(CONFIG.api.loginUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ userName: credentials.user, password: credentials.pass }),
      muteHttpExceptions: true
    });

    const loginData = JSON.parse(loginResponse.getContentText());
    const token = loginData.token;

    if (!token) {
      throw new Error("Autenticazione fallita");
    }

    updateProgress(progressSheet, 20, "Recupero dati...", "Download lezioni");

    const response = UrlFetchApp.fetch(CONFIG.api.dataUrl, {
      method: "get",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Errore API: ${response.getResponseCode()}`);
    }

    const allData = JSON.parse(response.getContentText());

    const oggi = new Date();
    oggi.setHours(0, 0, 0, 0);
    const tra180giorni = new Date();
    tra180giorni.setDate(oggi.getDate() + CONFIG.sync.daysInFuture);

    const lezioniPerDocente = new Map();
    const docentiColorMap = new Map();

    allData.forEach(lezione => {
      const inizio = new Date(lezione.tb_dallehh);
      const inizioSoloData = new Date(inizio);
      inizioSoloData.setHours(0, 0, 0, 0);

      if (inizioSoloData < oggi) return;

      const nomeDocente = formatNomeDocente(lezione.tb_nomdoc, lezione.tb_cogdoc);
      if (!nomeDocente || nomeDocente === "Docente Da Definire") return;

      if (!lezioniPerDocente.has(nomeDocente)) {
        lezioniPerDocente.set(nomeDocente, []);
      }
      lezioniPerDocente.get(nomeDocente).push(lezione);
    });

    const totalDocenti = lezioniPerDocente.size;
    let processedDocenti = 0;
    const counters = { creati: 0, aggiornati: 0, eliminati: 0, invariati: 0 };
    const rowsToInsert = [];
    const calendariDocenti = new Map();

    lezioniPerDocente.forEach((lezioni, nomeDocente) => {
      processedDocenti++;
      updateProgress(
        progressSheet,
        30 + (processedDocenti / totalDocenti * 50),
        `Docente: ${nomeDocente}`,
        `${processedDocenti}/${totalDocenti}`
      );

      if (!docentiColorMap.has(nomeDocente)) {
        const colorIndex = docentiColorMap.size % CONFIG.teacherColors.length;
        docentiColorMap.set(nomeDocente, CONFIG.teacherColors[colorIndex]);
      }

      const nomeCalendario = `${CONFIG.teacherCalendarPrefix}${nomeDocente}`;
      let calendario = CalendarApp.getCalendarsByName(nomeCalendario)[0];

      if (!calendario) {
        calendario = CalendarApp.createCalendar(nomeCalendario);
        calendario.setColor(docentiColorMap.get(nomeDocente));
      }

      calendariDocenti.set(nomeDocente, calendario);

      const eventiEsistenti = calendario.getEvents(oggi, tra180giorni);
      const cache = new Map();
      eventiEsistenti.forEach(evento => {
        const key = `${evento.getTitle()}___${formatDate(evento.getStartTime())}`;
        if (!cache.has(key)) cache.set(key, []);
        cache.get(key).push(evento);
      });

      const eventiAPI = new Set();

      lezioni.forEach(lezione => {
        const corso = lezione.tb_decorso?.trim() || "Corso Sconosciuto";
        const annoAccademico = convertYearCode(lezione.co_codfamcorso, yearMapping);
        const titoloModulo = lezione.tb_desmodulo?.trim() || "Lezione SPD";
        const aula = lezione.aula?.trim() || lezione.sede?.trim() || "Online";
        const inizio = new Date(lezione.tb_dallehh);
        const fine = new Date(lezione.tb_allehh);

        const titoloEvento = `[${corso}] ${titoloModulo}`;
        const apiKey = `${titoloEvento}___${inizio.toISOString()}___${fine.toISOString()}`;
        eventiAPI.add(apiKey);

        const cacheKey = `${titoloEvento}___${formatDate(inizio)}`;
        const eventiInCache = cache.get(cacheKey) || [];

        const newStartSec = toSeconds(inizio);
        const newEndSec = toSeconds(fine);

        const eventoEsistente = eventiInCache.find(ev =>
          toSeconds(ev.getStartTime()) === newStartSec &&
          toSeconds(ev.getEndTime()) === newEndSec &&
          ev.getLocation() === aula
        );

        let stato = "⏭️ Già presente";

        if (eventoEsistente) {
          counters.invariati++;
        } else {
          const eventoSimile = eventiInCache.find(ev =>
            ev.getTitle() === titoloEvento &&
            formatDate(ev.getStartTime()) === formatDate(inizio)
          );

          if (eventoSimile) {
            eventoSimile.setTime(inizio, fine);
            eventoSimile.setLocation(aula);
            let descrizione = `Corso: ${corso}\nAnno: ${annoAccademico}`;
            if (lezione.tb_note && lezione.tb_note.trim()) {
              descrizione += `\n\nNote: ${lezione.tb_note.trim()}`;
            }
            eventoSimile.setDescription(descrizione);
            stato = "🔄 Aggiornato";
            counters.aggiornati++;
          } else {
            let descrizione = `Corso: ${corso}\nAnno: ${annoAccademico}`;
            if (lezione.tb_note && lezione.tb_note.trim()) {
              descrizione += `\n\nNote: ${lezione.tb_note.trim()}`;
            }

            calendario.createEvent(titoloEvento, inizio, fine, {
              description: descrizione,
              location: aula
            });
            stato = "✅ Creato";
            counters.creati++;
          }
        }

        rowsToInsert.push({
          data: [nomeDocente, corso, annoAccademico, titoloModulo, aula, inizio, fine, nomeCalendario, stato],
          stato: stato
        });
      });

      eventiEsistenti.forEach(evento => {
        const key = `${evento.getTitle()}___${evento.getStartTime().toISOString()}___${evento.getEndTime().toISOString()}`;
        if (!eventiAPI.has(key)) {
          evento.deleteEvent();
          counters.eliminati++;
        }
      });
    });

    updateProgress(progressSheet, 85, "Aggiornamento fogli...", "Scrittura");

    if (rowsToInsert.length > 0) {
      const values = rowsToInsert.map(row => row.data);
      const batchSize = CONFIG.sync.batchSize;

      for (let i = 0; i < values.length; i += batchSize) {
        const batch = values.slice(i, i + batchSize);
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, batch.length, headers.length).setValues(batch);
      }
    }

    const dashboard = ss.getSheetByName(CONFIG.sheets.teachersDashboard) || ss.insertSheet(CONFIG.sheets.teachersDashboard);
    dashboard.clear();

    dashboard.getRange(1, 1, 1, 6).merge()
      .setValue("📊 DASHBOARD DOCENTI")
      .setFontSize(18)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#f8f9fa");

    dashboard.getRange(3, 1).setValue("Docente");
    dashboard.getRange(3, 2).setValue("N° Lezioni");
    dashboard.getRange(3, 3).setValue("Codici Corsi");
    dashboard.getRange(3, 4).setValue("Corsi");
    dashboard.getRange(3, 5).setValue("Calendario");
    dashboard.getRange(3, 6).setValue("Colore");

    dashboard.getRange(3, 1, 1, 6)
      .setFontWeight("bold")
      .setBackground("#e8eaed")
      .setHorizontalAlignment("center");

    let dashboardRow = 4;
    lezioniPerDocente.forEach((lezioni, nomeDocente) => {
      const corsiSet = new Set(lezioni.map(l => l.tb_decorso));
      const corsiList = Array.from(corsiSet).join(", ");

      const codiciSet = new Set(lezioni.map(l => l.tb_codcorso).filter(c => c));
      const codiciList = Array.from(codiciSet).join(", ");

      dashboard.getRange(dashboardRow, 1, 1, 6).setValues([[
        nomeDocente,
        lezioni.length,
        codiciList || "-",
        corsiList,
        `${CONFIG.teacherCalendarPrefix}${nomeDocente}`,
        "■"
      ]]);

      dashboard.getRange(dashboardRow, 6)
        .setBackground(docentiColorMap.get(nomeDocente))
        .setHorizontalAlignment("center");

      dashboardRow++;
    });

    dashboard.getRange(dashboardRow + 1, 1, 1, 6).merge()
      .setValue(`Totale: ${lezioniPerDocente.size} | Creati: ${counters.creati} | Aggiornati: ${counters.aggiornati} | Eliminati: ${counters.eliminati}`)
      .setFontWeight("bold")
      .setBackground("#5f6368")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");

    updateProgress(progressSheet, 100, "✅ Completato!", `${lezioniPerDocente.size} docenti`);

    const riepilogo = `
📊 RIEPILOGO DOCENTI

👥 Docenti: ${lezioniPerDocente.size}
📅 Calendari: ${calendariDocenti.size}

✅ Creati: ${counters.creati}
🔄 Aggiornati: ${counters.aggiornati}
❌ Eliminati: ${counters.eliminati}
⏭️ Invariati: ${counters.invariati}
`;

    ui.alert('Completato', riepilogo, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log(`❌ Errore: ${error}`);
    ui.alert('Errore', error.toString(), ui.ButtonSet.OK);
  }
}


// ====================================================================
// SEZIONE 5: AGGIORNA ANNI NEI CALENDARI
// ====================================================================

function aggiornaAnniCalendari() {
  const CONFIG = getConfig();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '🔄 Aggiornamento Anni',
    'Questa operazione aggiornerà i nomi dei calendari.\n\nPuò richiedere alcuni minuti.\n\nProcedere?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const toSeconds = date => Math.floor(date.getTime() / 1000);
  const formatDate = date => Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  function setupProgressSheet() {
    let progressSheet = ss.getSheetByName(CONFIG.sheets.updateProgress);
    if (!progressSheet) {
      progressSheet = ss.insertSheet(CONFIG.sheets.updateProgress);
    }

    progressSheet.clear();
    progressSheet.setColumnWidths(1, 4, 150);

    progressSheet.getRange(1, 1, 1, 4).merge()
      .setValue("🔄 AGGIORNAMENTO ANNI")
      .setFontSize(16)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#f4b400")
      .setFontColor("#ffffff");

    progressSheet.getRange(3, 1).setValue("Progresso:");
    progressSheet.getRange(3, 2, 1, 3).merge();
    progressSheet.getRange(5, 1).setValue("Stato:");
    progressSheet.getRange(5, 2, 1, 3).merge();
    progressSheet.getRange(7, 1).setValue("Dettagli:");
    progressSheet.getRange(8, 1, 1, 4).merge();

    return progressSheet;
  }

  function updateProgress(progressSheet, percentage, message, details = "") {
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

    SpreadsheetApp.flush();
  }

  try {
    const progressSheet = setupProgressSheet();
    const yearMapping = getYearMapping();
    const startTime = new Date();

    updateProgress(progressSheet, 0, "Inizializzazione...", "Caricamento mappatura");

    let calendariRinominati = 0;
    let eventiAggiornati = 0;
    let errori = 0;

    const allCalendars = CalendarApp.getAllCalendars();
    const spdCalendars = allCalendars.filter(cal => cal.getName().startsWith(CONFIG.calendarPrefix));

    updateProgress(progressSheet, 10, "Calendari trovati", `${spdCalendars.length} calendari`);

    updateProgress(progressSheet, 15, "Verifica calendari...", "Analisi nomi");

    const calendarPattern = /^SPD - (.+) - (\d{4}\/\d{4})$/;
    const yearPattern = /(\d{4})\/\d{4}/;

    spdCalendars.forEach((calendar, index) => {
      const calendarName = calendar.getName();
      const match = calendarName.match(calendarPattern);

      if (match) {
        const corso = match[1];
        const annoVecchio = match[2];

        const yearMatch = annoVecchio.match(yearPattern);
        if (yearMatch) {
          const codiceAnno = yearMatch[1];
          const annoNuovo = convertYearCode(codiceAnno, yearMapping);

          if (annoVecchio !== annoNuovo) {
            const nuovoNome = `${CONFIG.calendarPrefix}${corso} - ${annoNuovo}`;

            try {
              calendariRinominati++;
              Logger.log(`Da rinominare: "${calendarName}" → "${nuovoNome}"`);
            } catch (e) {
              errori++;
              Logger.log(`Errore: ${e}`);
            }
          }
        }
      }

      const progress = 15 + (index / spdCalendars.length * 35);
      updateProgress(progressSheet, Math.round(progress),
        "Verifica calendari",
        `${index + 1}/${spdCalendars.length}`
      );
    });

    updateProgress(progressSheet, 50, "Calendario unificato...", "Ricerca eventi");

    const calendarioUnificato = spdCalendars.find(cal => cal.getName() === CONFIG.unifiedCalendarName);

    if (calendarioUnificato) {
      const oggi = new Date();
      oggi.setHours(0, 0, 0, 0);
      const tra180giorni = new Date();
      tra180giorni.setDate(oggi.getDate() + CONFIG.sync.daysInFuture);

      const eventi = calendarioUnificato.getEvents(oggi, tra180giorni);
      updateProgress(progressSheet, 55, "Eventi trovati", `${eventi.length} eventi`);

      const eventoPattern = /^(\[.+\]\s.+)\s\((\d{4}\/\d{4})\)$/;

      eventi.forEach((evento, index) => {
        const titoloVecchio = evento.getTitle();
        const match = titoloVecchio.match(eventoPattern);

        if (match) {
          const parteBase = match[1];
          const annoVecchio = match[2];

          const yearMatch = annoVecchio.match(yearPattern);
          if (yearMatch) {
            const codiceAnno = yearMatch[1];
            const annoNuovo = convertYearCode(codiceAnno, yearMapping);

            if (annoVecchio !== annoNuovo) {
              const titoloNuovo = `${parteBase} (${annoNuovo})`;

              try {
                evento.setTitle(titoloNuovo);
                eventiAggiornati++;
              } catch (e) {
                errori++;
                Logger.log(`Errore evento: ${e}`);
              }
            }
          }
        }

        if (index % 10 === 0) {
          const progress = 55 + (index / eventi.length * 40);
          updateProgress(progressSheet, Math.round(progress),
            "Aggiornamento eventi",
            `${index + 1}/${eventi.length}`
          );
        }
      });
    } else {
      updateProgress(progressSheet, 55, "Calendario unificato non trovato", "Skip");
    }

    updateProgress(progressSheet, 95, "Salvataggio log...", "Registrazione");

    const logSheet = ss.getSheetByName(CONFIG.sheets.updateLog) || ss.insertSheet(CONFIG.sheets.updateLog);
    if (logSheet.getLastRow() === 0) {
      logSheet.getRange(1, 1, 1, 5).setValues([["Data/Ora", "Calendari", "Eventi", "Errori", "Durata"]])
        .setFontWeight("bold")
        .setBackground("#f8f9fa");
    }

    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);

    logSheet.appendRow([
      new Date(),
      calendariRinominati,
      eventiAggiornati,
      errori,
      duration
    ]);

    updateProgress(progressSheet, 100, "✅ Completato!",
      `Eventi: ${eventiAggiornati}`);

    let riepilogo = `
📊 RIEPILOGO AGGIORNAMENTO

⏱️ Tempo: ${duration}s
📅 Eventi aggiornati: ${eventiAggiornati}
⚠️ Errori: ${errori}
`;

    if (calendariRinominati > 0) {
      riepilogo += `\n⚠️ ${calendariRinominati} calendari da rinominare manualmente.\n`;
      riepilogo += `Controlla il log.`;
    }

    ui.alert('Completato', riepilogo, ui.ButtonSet.OK);

  } catch (error) {
    Logger.log(`❌ Errore: ${error}`);
    ui.alert('Errore', error.toString(), ui.ButtonSet.OK);
  }
}


// ====================================================================
// SEZIONE 6: TRIGGER E MENU
// ====================================================================

function setupCalendarTrigger() {
  const CONFIG = getConfig();

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sincronizzaCalendariPerCorso') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('sincronizzaCalendariPerCorso')
    .timeBased()
    .everyHours(CONFIG.sync.autoSyncHours)
    .create();

  SpreadsheetApp.getUi().alert(
    'Trigger Configurato',
    `Sincronizzazione automatica ogni ${CONFIG.sync.autoSyncHours} ore.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Sincronizzazione SPD')
    .addItem('⚙️ Configura Credenziali', 'setupCredentials')
    .addItem('🧪 Test Credenziali', 'testCredentials')
    .addItem('📊 Stato Configurazione', 'showConfigStatus')
    .addSeparator()
    .addItem('▶️ Avvia Sincronizzazione', 'sincronizzaCalendariPerCorso')
    .addItem('👥 Sincronizza Docenti', 'sincronizzaCalendariDocenti')
    .addItem('🔄 Aggiorna Anni Calendari', 'aggiornaAnniCalendari')
    .addItem('⏰ Configura Auto-Sync', 'setupCalendarTrigger')
    .addSeparator()
    .addItem('📊 Dashboard', 'showDashboard')
    .addItem('📋 Log', 'showLog')
    .addItem('📅 Mappatura Anni', 'showMappingSheet')
    .addItem('📚 Mappatura Corsi', 'showCorsiMappingSheet')
    .addToUi();
}

function showDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('Dashboard');
  if (dashboard) {
    ss.setActiveSheet(dashboard);
  } else {
    SpreadsheetApp.getUi().alert('Dashboard non trovato', 'Esegui prima una sincronizzazione.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function showLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('Log');
  if (log) {
    ss.setActiveSheet(log);
  } else {
    SpreadsheetApp.getUi().alert('Log non trovato', 'Esegui prima una sincronizzazione.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function showMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let mapping = ss.getSheetByName('Mappatura Anni');
  if (!mapping) {
    mapping = setupMappingSheet();
  }
  ss.setActiveSheet(mapping);
}

function showCorsiMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let mapping = ss.getSheetByName('Mappatura Corsi');
  if (!mapping) {
    mapping = setupCorsiMappingSheet();
  }
  ss.setActiveSheet(mapping);
}


// ====================================================================
// SEZIONE 7: WEB APP (OPZIONALE - PER USO FUTURO)
// ====================================================================

// NOTA: Le funzioni Web App sono commentate per sicurezza.
// Decommentale solo se necessario e aggiungi autenticazione appropriata.

/*
function doGet(e) {
  // ⚠️ ATTENZIONE: Implementare autenticazione prima di usare!
  return ContentService
    .createTextOutput(JSON.stringify({error: 'Web App disabilitata per sicurezza'}))
    .setMimeType(ContentService.MimeType.JSON);
}
*/
