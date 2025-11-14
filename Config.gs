/**
 * Config.gs - Gestione Configurazione Centralizzata e Sicura
 *
 * IMPORTANTE: Le credenziali NON sono più hardcoded.
 * Usare setupCredentials() per configurarle la prima volta.
 */

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
      timeout: 30000, // 30 secondi
      maxRetries: 3,
      retryDelay: 2000 // 2 secondi
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
 *
 * @returns {Object} {user: string, pass: string}
 * @throws {Error} Se le credenziali non sono configurate
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
      'Oppure esegui manualmente:\n' +
      '1. Vai su Estensioni → Apps Script\n' +
      '2. Esegui la funzione "setupCredentials"\n' +
      '3. Inserisci username e password quando richiesto'
    );
  }

  return {
    user: user,
    pass: pass
  };
}

/**
 * Salva le credenziali in modo sicuro
 * NOTA: Questa funzione va eseguita UNA SOLA VOLTA per configurare il sistema
 *
 * @param {string} username - Username API
 * @param {string} password - Password API
 */
function saveCredentials(username, password) {
  if (!username || !password) {
    throw new Error('Username e password sono obbligatori');
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SPD_API_USER', username);
  scriptProperties.setProperty('SPD_API_PASS', password);

  Logger.log('✅ Credenziali salvate con successo in modo sicuro');
}

/**
 * Rimuove le credenziali salvate
 * ATTENZIONE: Dopo questa operazione sarà necessario riconfigurare
 */
function deleteCredentials() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('SPD_API_USER');
  scriptProperties.deleteProperty('SPD_API_PASS');

  Logger.log('🗑️ Credenziali rimosse');
}

/**
 * Verifica se le credenziali sono configurate
 * @returns {boolean}
 */
function areCredentialsConfigured() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const user = scriptProperties.getProperty('SPD_API_USER');
  const pass = scriptProperties.getProperty('SPD_API_PASS');

  return !!(user && pass);
}

/**
 * Ottiene informazioni sullo stato della configurazione
 * @returns {Object} Stato configurazione
 */
function getConfigStatus() {
  const credentialsConfigured = areCredentialsConfigured();

  return {
    credentialsConfigured: credentialsConfigured,
    message: credentialsConfigured
      ? '✅ Credenziali configurate correttamente'
      : '⚠️ Credenziali non ancora configurate',
    ready: credentialsConfigured
  };
}
