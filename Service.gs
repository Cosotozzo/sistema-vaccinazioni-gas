/**
 * Services.gs - Servizi di Business Logic e Gestione Dati
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistema Vaccinale - Studio Medico')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
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

function getVaccineData() {
  const sheet = getDb().getSheetByName(SHEET_VACCINI);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const headers = data.shift().map(h => String(h || '').toLowerCase().trim());
  const denominazioneIdx = headers.indexOf('denominazionevaccino');
  const lottoIdx = headers.indexOf('numerolotto');
  const statoIdx = headers.indexOf('stato');

  return data
    .filter(row => (row[statoIdx] || '').toString().toLowerCase().trim() !== 'completato')
    .map(row => ({
      denominazione: row[denominazioneIdx] || '',
      lotto: row[lottoIdx] || ''
    }));
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