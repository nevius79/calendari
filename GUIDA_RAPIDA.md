# 📘 Guida Rapida - Calendari SPD

## 🎯 FILE DA USARE

**CalendariSPD.gs** - Un unico file con tutto il sistema

---

## 🚀 INSTALLAZIONE (5 MINUTI)

### 1️⃣ Copia il file in Google Apps Script

1. Apri il tuo Google Spreadsheet
2. Menu: **Estensioni** → **Apps Script**
3. Elimina il codice esistente (se presente)
4. Copia **tutto** il contenuto di `CalendariSPD.gs`
5. Incolla nell'editor
6. Clicca **Salva** (icona dischetto)

### 2️⃣ Ricarica il foglio

1. Torna al Google Spreadsheet
2. Ricarica la pagina (F5 o Cmd+R)
3. Vedrai il menu **🔄 Sincronizzazione SPD**

### 3️⃣ Configura le credenziali (PRIMA VOLTA)

1. Menu: **🔄 Sincronizzazione SPD** → **⚙️ Configura Credenziali**
2. Inserisci **username** quando richiesto
3. Inserisci **password** quando richiesto
4. ✅ Le credenziali vengono salvate in modo sicuro

**IMPORTANTE:** Le credenziali sono salvate in modo criptato e NON sono visibili nel codice!

### 4️⃣ Test (opzionale ma consigliato)

1. Menu: **🔄 Sincronizzazione SPD** → **🧪 Test Credenziali**
2. Se vedi "✅ Test SUCCESSO" → tutto OK!
3. Se vedi "❌ Test FALLITO" → verifica username/password

---

## 📖 USO QUOTIDIANO

### Sincronizzazione Calendari per Corso

Menu: **🔄 Sincronizzazione SPD** → **▶️ Avvia Sincronizzazione**

Cosa fa:
- ✅ Crea/aggiorna calendari per ogni corso
- ✅ Crea calendario unificato "SPD - Tutte le Lezioni"
- ✅ Aggiorna eventi esistenti
- ✅ Genera dashboard e log

Tempo: ~2-5 minuti (dipende dal numero di lezioni)

### Sincronizzazione Calendari Docenti

Menu: **🔄 Sincronizzazione SPD** → **👥 Sincronizza Docenti**

Cosa fa:
- ✅ Crea un calendario per ogni docente
- ✅ Include tutte le lezioni del docente
- ✅ Genera dashboard docenti

### Sincronizzazione Automatica

Menu: **🔄 Sincronizzazione SPD** → **⏰ Configura Auto-Sync**

Configura la sincronizzazione automatica ogni 4 ore.

---

## 📊 FOGLI GENERATI

Dopo la prima sincronizzazione vedrai questi fogli:

| Foglio | Descrizione |
|--------|-------------|
| **Dashboard** | Statistiche per corso |
| **Dashboard Docenti** | Statistiche per docente |
| **Foglio1** | Elenco dettagliato lezioni |
| **Lezioni Docenti** | Elenco lezioni per docente |
| **Log** | Storico sincronizzazioni |
| **Mappatura Anni** | Conversione codici anno → Anno Accademico |
| **Mappatura Corsi** | Conversione codici corso → Nome corso |

---

## ⚙️ CONFIGURAZIONI AVANZATE

### Mappatura Anni Accademici

Menu: **🔄 Sincronizzazione SPD** → **📅 Mappatura Anni**

Personalizza la conversione dei codici anno (es. 2024 → 2024/2025)

**Esempio:**

| Codice | Anno Accademico | Note |
|--------|-----------------|------|
| 2024 | 2024/2025 | Anno corrente |
| 2023 | 2023/2024 | |
| 2025 | 2025/2026 | |

### Mappatura Corsi

Menu: **🔄 Sincronizzazione SPD** → **📚 Mappatura Corsi**

Assegna nomi ai corsi sconosciuti.

**Esempio:**

| Codice Corso | Descrizione Breve | Descrizione Completa | Note |
|--------------|-------------------|----------------------|------|
| 101 | Design Grafico | Design Grafico Base | Primo anno |
| 202 | UX Design | User Experience Design | Secondo anno |

**⚠️ IMPORTANTE:** Se vedi corsi "sconosciuti", compila questo foglio!

---

## 🔐 SICUREZZA

### Dove sono salvate le credenziali?

Le credenziali sono salvate in **Properties Service** di Google:
- ✅ Criptate
- ✅ Non visibili nel codice
- ✅ Accessibili solo dal tuo account Google
- ✅ Non condivise se condividi il foglio

### Come cambio le credenziali?

Menu: **🔄 Sincronizzazione SPD** → **⚙️ Configura Credenziali**

Ti chiederà se vuoi sovrascrivere, rispondi **Sì**.

### Come verifico lo stato?

Menu: **🔄 Sincronizzazione SPD** → **📊 Stato Configurazione**

Mostra:
- 🔐 Credenziali configurate (sì/no)
- 👤 Username (visibile)
- 🔑 Password (nascosta)
- ⚙️ Impostazioni sistema

---

## 🆘 RISOLUZIONE PROBLEMI

### Errore: "Credenziali non configurate"

**Soluzione:**
1. Menu: **⚙️ Configura Credenziali**
2. Inserisci username e password
3. Riprova

### Errore: "Autenticazione fallita"

**Soluzione:**
1. Verifica username e password
2. Menu: **🧪 Test Credenziali**
3. Se fallisce, riconfigura con **⚙️ Configura Credenziali**

### Calendari "Corso Sconosciuto"

**Soluzione:**
1. Menu: **📚 Mappatura Corsi**
2. Compila con i codici corso mancanti
3. Riesegui sincronizzazione

### La sincronizzazione è lenta

**Normale!**
- Prima volta: 5-10 minuti (crea tutti i calendari)
- Successive: 2-3 minuti (aggiorna solo modifiche)

### Non vedo il menu

**Soluzione:**
1. Ricarica la pagina (F5)
2. Attendi 5 secondi
3. Se non appare, vai in Estensioni → Apps Script
4. Esegui manualmente la funzione `onOpen`

---

## 📈 PROSSIMI MIGLIORAMENTI (FASI FUTURE)

Se vuoi continuare con i miglioramenti graduali:

### ✅ FASE 1 - COMPLETATA
- Sicurezza credenziali

### 🔜 FASE 2 - PERFORMANCE
- Riduzione chiamate API
- Batch operations ottimizzate
- Cache intelligente
- Velocità +40%

### 🔜 FASE 3 - MODULARITÀ
- Codice organizzato in moduli
- Funzioni riutilizzabili
- Più facile manutenzione
- Testing semplificato

### 🔜 FASE 4 - FEATURES
- Notifiche email personalizzate
- Export calendario iCal
- Filtri avanzati
- Statistiche estese

---

## 📞 SUPPORTO

### Log e Debug

Tutti i log sono disponibili in:
1. **Foglio "Log"** - Storico sincronizzazioni
2. **Apps Script** → **Esecuzioni** - Log dettagliati

### Documenti Generati

Ogni sincronizzazione genera:
- ✅ Dashboard con statistiche
- ✅ Log con timestamp
- ✅ Fogli progressi (durante esecuzione)

---

## 🎉 FATTO!

Ora hai un sistema completo e sicuro per sincronizzare i calendari SPD!

**Ricorda:**
- ✅ Credenziali sicure e criptate
- ✅ Un solo file da gestire
- ✅ Menu facile da usare
- ✅ Tutto il codice originale funziona

**Prossimo step:** Vuoi continuare con le altre fasi di miglioramento?
