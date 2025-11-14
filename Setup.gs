/**
 * Setup.gs - Funzioni di Setup e Configurazione Iniziale
 *
 * Queste funzioni permettono di configurare il sistema in modo sicuro
 * attraverso un'interfaccia utente guidata.
 */

/**
 * Funzione interattiva per configurare le credenziali
 * Mostra un dialog per inserire username e password in modo sicuro
 *
 * COME USARE:
 * 1. Menu: 🔄 Sincronizzazione SPD → ⚙️ Configura Credenziali
 * 2. Inserisci username e password quando richiesto
 * 3. Le credenziali vengono salvate in modo sicuro e criptato
 */
function setupCredentials() {
  const ui = SpreadsheetApp.getUi();

  // Verifica se già configurate
  const status = getConfigStatus();
  if (status.credentialsConfigured) {
    const response = ui.alert(
      '⚠️ Credenziali già configurate',
      'Le credenziali API sono già state configurate.\n\n' +
      'Vuoi aggiornarle con nuove credenziali?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }
  }

  // Richiedi username
  const userResponse = ui.prompt(
    '🔐 Configurazione Credenziali API - Step 1/2',
    'Inserisci lo USERNAME per l\'API SPD:\n\n' +
    '(Esempio: superuser@login.scuoladesign.spd)',
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

  // Richiedi password
  const passResponse = ui.prompt(
    '🔐 Configurazione Credenziali API - Step 2/2',
    'Inserisci la PASSWORD per l\'API SPD:\n\n' +
    '⚠️ La password verrà salvata in modo sicuro e NON sarà visibile nel codice.',
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

  // Salva credenziali
  try {
    saveCredentials(username, password);

    ui.alert(
      '✅ Configurazione Completata',
      'Le credenziali sono state salvate con successo!\n\n' +
      '🔒 Le credenziali sono archiviate in modo sicuro e criptato.\n' +
      '📝 NON sono visibili nel codice.\n\n' +
      'Ora puoi procedere con la sincronizzazione.',
      ui.ButtonSet.OK
    );

    // Test immediato delle credenziali
    const testResult = ui.alert(
      '🧪 Test Credenziali',
      'Vuoi testare subito le credenziali configurate?',
      ui.ButtonSet.YES_NO
    );

    if (testResult === ui.Button.YES) {
      testCredentials();
    }

  } catch (error) {
    ui.alert(
      '❌ Errore',
      'Errore durante il salvataggio delle credenziali:\n\n' + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * Testa le credenziali configurate facendo una chiamata API reale
 * Mostra il risultato in un alert
 */
function testCredentials() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Mostra progress
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const toast = ss.toast('🔄 Test credenziali in corso...', 'Test API', -1);

    // Ottieni credenziali
    const credentials = getCredentials();
    const config = getConfig();

    // Chiama API di login
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

    // Nascondi toast
    ss.toast('', '', 1);

    if (responseCode === 200 && responseData.token) {
      ui.alert(
        '✅ Test Credenziali - SUCCESSO',
        'Le credenziali sono corrette!\n\n' +
        '🔑 Token ricevuto: ' + responseData.token.substring(0, 20) + '...\n' +
        '👤 Utente: ' + credentials.user + '\n\n' +
        'Il sistema è pronto per la sincronizzazione.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '❌ Test Credenziali - FALLITO',
        'Le credenziali sembrano non essere corrette.\n\n' +
        '📊 Codice risposta: ' + responseCode + '\n' +
        '📝 Messaggio: ' + (responseData.message || 'Errore sconosciuto') + '\n\n' +
        'Verifica username e password e riprova.',
        ui.ButtonSet.OK
      );
    }

  } catch (error) {
    ui.alert(
      '❌ Errore Test',
      'Errore durante il test delle credenziali:\n\n' + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * Mostra lo stato della configurazione corrente
 */
function showConfigStatus() {
  const ui = SpreadsheetApp.getUi();
  const status = getConfigStatus();
  const config = getConfig();

  let message = '📊 STATO CONFIGURAZIONE\n\n';

  // Credenziali
  message += '🔐 Credenziali API:\n';
  if (status.credentialsConfigured) {
    const cred = getCredentials();
    message += `   ✅ Configurate\n`;
    message += `   👤 Username: ${cred.user}\n`;
    message += `   🔑 Password: ${'*'.repeat(12)}\n\n`;
  } else {
    message += `   ❌ NON configurate\n\n`;
  }

  // Endpoint API
  message += '🌐 Endpoint API:\n';
  message += `   Login: ${config.api.loginUrl}\n`;
  message += `   Dati: ${config.api.dataUrl}\n\n`;

  // Impostazioni sincronizzazione
  message += '⚙️ Impostazioni:\n';
  message += `   📅 Giorni futuri: ${config.sync.daysInFuture}\n`;
  message += `   📦 Batch size: ${config.sync.batchSize}\n`;
  message += `   ⏰ Auto-sync: ogni ${config.sync.autoSyncHours} ore\n\n`;

  // Stato finale
  message += status.ready
    ? '✅ Sistema pronto per la sincronizzazione'
    : '⚠️ Configura le credenziali per iniziare';

  ui.alert('Stato Configurazione', message, ui.ButtonSet.OK);
}

/**
 * Reset completo della configurazione
 * ATTENZIONE: Rimuove tutte le credenziali salvate
 */
function resetConfiguration() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '⚠️ ATTENZIONE - Reset Configurazione',
    'Questa operazione rimuoverà TUTTE le credenziali salvate.\n\n' +
    'Sarà necessario riconfigurarle prima della prossima sincronizzazione.\n\n' +
    'Sei sicuro di voler procedere?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    try {
      deleteCredentials();
      ui.alert(
        '✅ Reset Completato',
        'Tutte le credenziali sono state rimosse.\n\n' +
        'Esegui "Configura Credenziali" per configurare nuovamente il sistema.',
        ui.ButtonSet.OK
      );
    } catch (error) {
      ui.alert(
        '❌ Errore',
        'Errore durante il reset:\n\n' + error.toString(),
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * Setup iniziale completo del sistema
 * Crea tutti i fogli necessari e configura le credenziali
 */
function firstTimeSetup() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '👋 Benvenuto - Setup Iniziale',
    'Questa procedura guidata configurerà il sistema per la prima volta.\n\n' +
    'Verranno:\n' +
    '• Configurate le credenziali API in modo sicuro\n' +
    '• Creati i fogli di lavoro necessari\n' +
    '• Testate le credenziali\n\n' +
    'Tempo stimato: 2-3 minuti\n\n' +
    'Vuoi procedere?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // 1. Configura credenziali
    ui.alert(
      'Step 1/3 - Credenziali',
      'Procediamo con la configurazione delle credenziali API.',
      ui.ButtonSet.OK
    );
    setupCredentials();

    // 2. Crea fogli di mappatura
    ui.alert(
      'Step 2/3 - Fogli di Lavoro',
      'Creazione fogli di mappatura...',
      ui.ButtonSet.OK
    );
    setupMappingSheet();
    setupCorsiMappingSheet();

    // 3. Completa
    ui.alert(
      '✅ Setup Completato!',
      'Il sistema è stato configurato con successo.\n\n' +
      '📋 Prossimi passi:\n' +
      '1. Compila il foglio "Mappatura Corsi" (opzionale)\n' +
      '2. Verifica il foglio "Mappatura Anni" (opzionale)\n' +
      '3. Esegui la prima sincronizzazione\n\n' +
      'Usa il menu "🔄 Sincronizzazione SPD" per tutte le operazioni.',
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert(
      '❌ Errore Setup',
      'Si è verificato un errore durante il setup:\n\n' + error.toString(),
      ui.ButtonSet.OK
    );
  }
}
