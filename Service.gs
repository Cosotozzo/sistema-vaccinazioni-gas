/**
 * Service.gs - Servizi di Business Logic e Gestione Dati
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistema Vaccinale - Studio Medico')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// Restituisce SOLO i vaccini disponibili con almeno 1 dose per la compilazione del modulo
function getVaccineData() {
  const sheet = getDb().getSheetByName(SHEET_VACCINI);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const headers = data.shift().map(h => String(h || '').toLowerCase().trim());
  const denominazioneIdx = headers.indexOf('denominazionevaccino');
  const lottoIdx = headers.indexOf('numerolotto');
  const statoIdx = headers.indexOf('stato');
  const dosiIdx = headers.indexOf('dosidisponibili');

  return data
    .filter(row => {
      const stato = (row[statoIdx] || '').toString().toLowerCase().trim();
      const dosi = parseInt(row[dosiIdx], 10);
      return stato !== 'completato' && stato !== 'esaurito' && (!isNaN(dosi) && dosi > 0);
    })
    .map(row => ({
      denominazione: row[denominazioneIdx] || '',
      lotto: row[lottoIdx] || '',
      dosi: row[dosiIdx] !== undefined ? row[dosiIdx] : 0
    }));
}

// Restituisce TUTTI i vaccini per la tabella di gestione
function getAllVaccinesForManagement() {
  const sheet = getDb().getSheetByName(SHEET_VACCINI);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data.shift().map(h => String(h || '').toLowerCase().trim());
  const denominazioneIdx = headers.indexOf('denominazionevaccino');
  const lottoIdx = headers.indexOf('numerolotto');
  const statoIdx = headers.indexOf('stato');
  const dosiIdx = headers.indexOf('dosidisponibili');

  return data.map((row, index) => ({
    rowIndex: index + 2,
    denominazione: row[denominazioneIdx] || '',
    lotto: row[lottoIdx] || '',
    stato: row[statoIdx] || 'Attivo',
    dosi: row[dosiIdx] !== undefined ? parseInt(row[dosiIdx], 10) || 0 : 0
  }));
}

// Inserimento nuovo lotto vaccino
function addVaccineBatch(denominazione, lotto, dosi) {
  if (!denominazione || !lotto) {
    return { success: false, message: 'Denominazione e Numero Lotto sono obbligatori.' };
  }

  const numDosi = parseInt(dosi, 10) || 0;
  const initialStatus = numDosi > 0 ? 'Attivo' : 'Completato';
  const sheet = getDb().getSheetByName(SHEET_VACCINI);
  sheet.appendRow([denominazione, lotto, initialStatus, new Date(), numDosi]);
  
  logAction('Medico', `Inserito nuovo lotto: ${denominazione} - Lotto ${lotto} (${numDosi} dosi)`);
  return { success: true, message: 'Lotto vaccino inserito con successo!' };
}

// Aggiornamento puntuale delle dosi con ripristino o chiusura automatica dello stato
function updateVaccineQuantity(rowIndex, newQuantity) {
  const sheet = getDb().getSheetByName(SHEET_VACCINI);
  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
    return { success: false, message: 'Indice riga non valido.' };
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h || '').toLowerCase().trim());
  const dosiIdx = headers.indexOf('dosidisponibili');
  const statoIdx = headers.indexOf('stato');

  if (dosiIdx === -1) return { success: false, message: 'Colonna DosiDisponibili non trovata.' };

  const qty = parseInt(newQuantity, 10);
  if (isNaN(qty) || qty < 0) {
    return { success: false, message: 'Inserisci un numero di dosi valido.' };
  }

  sheet.getRange(rowIndex, dosiIdx + 1).setValue(qty);

  const newStatus = qty > 0 ? 'Attivo' : 'Completato';
  if (statoIdx !== -1) {
    sheet.getRange(rowIndex, statoIdx + 1).setValue(newStatus);
  }

  logAction('Medico', `Aggiornate dosi riga ${rowIndex} a: ${qty} (Stato: ${newStatus})`);
  return { 
    success: true, 
    newStatus: newStatus, 
    newDoses: qty, 
    message: `Dosi aggiornate a ${qty}. Stato: ${qty > 0 ? 'Disponibile' : 'Esaurito'}.` 
  };
}

// Decremento automatico di 1 dose post-somministrazione
function decrementVaccineDose(denominazione, lotto) {
  const sheet = getDb().getSheetByName(SHEET_VACCINI);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0].map(h => String(h || '').toLowerCase().trim());
  const denIdx = headers.indexOf('denominazionevaccino');
  const lottoIdx = headers.indexOf('numerolotto');
  const dosiIdx = headers.indexOf('dosidisponibili');
  const statoIdx = headers.indexOf('stato');

  if (denIdx === -1 || lottoIdx === -1 || dosiIdx === -1) return;

  for (let i = 1; i < data.length; i++) {
    const dVal = String(data[i][denIdx] || '').trim();
    const lVal = String(data[i][lottoIdx] || '').trim();

    if (dVal === String(denominazione).trim() && lVal === String(lotto).trim()) {
      let currentDoses = parseInt(data[i][dosiIdx], 10);
      if (isNaN(currentDoses)) currentDoses = 0;

      const newDoses = Math.max(0, currentDoses - 1);
      sheet.getRange(i + 1, dosiIdx + 1).setValue(newDoses);

      if (newDoses === 0 && statoIdx !== -1) {
        sheet.getRange(i + 1, statoIdx + 1).setValue('Completato');
        logAction('Sistema', `Lotto ${lotto} esaurito automaticamente (Dosi: 0)`);
      } else {
        logAction('Sistema', `Scalata 1 dose da lotto ${lotto} (Dosi rimanenti: ${newDoses})`);
      }
      break;
    }
  }
}

function getPatientSuggestions(searchTerm) {
  if (!searchTerm || searchTerm.trim().length < 3) return [];

  const cache = CacheService.getScriptCache();
  const CACHE_KEY_HEADER = 'patient_data_header';
  const CACHE_KEY_PREFIX = 'patient_data_chunk_';
  const CACHE_KEY_COUNT = 'patient_data_chunk_count';
  let data = [];

  const chunkCountStr = cache.get(CACHE_KEY_COUNT);

  if (chunkCountStr) {
    const headerRow = JSON.parse(cache.get(CACHE_KEY_HEADER));
    if (headerRow) {
      const chunkCount = parseInt(chunkCountStr, 10);
      const keys = [];
      for (let i = 0; i < chunkCount; i++) keys.push(CACHE_KEY_PREFIX + i);
      const cachedChunks = cache.getAll(keys);
      let patientRows = [];
      for (let i = 0; i < chunkCount; i++) {
        if (cachedChunks[CACHE_KEY_PREFIX + i]) {
          patientRows = patientRows.concat(JSON.parse(cachedChunks[CACHE_KEY_PREFIX + i]));
        }
      }
      data = [headerRow].concat(patientRows);
    }
  } 
  
  if (data.length === 0) {
    const sheet = getDb().getSheetByName(SHEET_PAZIENTI);
    const allSheetData = sheet.getDataRange().getValues();
    if (allSheetData.length < 1) return [];

    const headerRow = allSheetData.shift();
    const patientRows = allSheetData;
    const CHUNK_SIZE = 250;
    const chunks = {};
    let count = 0;

    for (let i = 0; i < patientRows.length; i += CHUNK_SIZE) {
      chunks[CACHE_KEY_PREFIX + count] = JSON.stringify(patientRows.slice(i, i + CHUNK_SIZE));
      count++;
    }
    chunks[CACHE_KEY_COUNT] = count.toString();
    chunks[CACHE_KEY_HEADER] = JSON.stringify(headerRow);
    cache.putAll(chunks, 21600);
    data = [headerRow].concat(patientRows);
  }

  if (data.length === 0) return [];
  
  const headers = data[0].map(h => String(h || '').toLowerCase().trim());
  const searchWords = searchTerm.toLowerCase().trim().split(' ').filter(w => w.length > 0);

  const cfIdx = headers.indexOf('codicefiscale');
  const cognomeIdx = headers.indexOf('cognome');
  const nomeIdx = headers.indexOf('nome');

  if (cfIdx === -1 || cognomeIdx === -1 || nomeIdx === -1) return [];

  const suggestions = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const cf = (row[cfIdx] || '').toString().toLowerCase().trim();
    const cognome = (row[cognomeIdx] || '').toString().toLowerCase().trim();
    const nome = (row[nomeIdx] || '').toString().toLowerCase().trim();

    const fullStr = `${nome} ${cognome} ${cf}`;
    if (searchWords.every(word => fullStr.includes(word))) {
      let patientObj = {};
      headers.forEach((hKey, idx) => { patientObj[hKey] = row[idx]; });
      suggestions.push(patientObj);
      if (suggestions.length >= 10) break;
    }
  }
  return suggestions;
}

function logAction(user, action) {
  let sheet = getDb().getSheetByName(SHEET_LOGS);
  if (!sheet) {
    sheet = getDb().insertSheet(SHEET_LOGS);
    sheet.appendRow(['Data', 'Ora', 'Utente', 'Azione']);
  }
  const now = new Date();
  sheet.appendRow([now.toLocaleDateString('it-IT'), now.toLocaleTimeString('it-IT'), user, action]);
}