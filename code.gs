// ID del tuo foglio Google. Puoi trovarlo nell'URL del foglio.
const SPREADSHEET_ID = '1PC_2uIFExMog5TZvEQPElTj3Ogrs8dEHh-3yQVFiTEU';

// Nomi dei fogli di lavoro
const SHEET_PAZIENTI = 'DB_pazienti';
const SHEET_VACCINI = 'Vaccini';
const SHEET_CONSENSI = 'Consensi';

// Nome della cartella in Google Drive dove verranno salvati i PDF
const PDF_FOLDER_NAME = 'Consensi Vaccinazioni Firmati';

/**
 * Funzione principale che serve l'applicazione web.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Modulo Consenso Vaccinazione')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Esegue una ricerca autocomplete OTTIMIZZATA CON CACHE A BLOCCHI e GESTIONE HEADER CORRETTA.
 * @param {string} searchTerm - Il termine da cercare.
 * @returns {Array<object>} Un array di oggetti paziente.
 */
function getPatientSuggestions(searchTerm) {
  if (!searchTerm || searchTerm.trim().length < 3) {
    return [];
  }

  const cache = CacheService.getScriptCache();
  const CACHE_KEY_HEADER = 'patient_data_header';
  const CACHE_KEY_PREFIX = 'patient_data_chunk_';
  const CACHE_KEY_COUNT = 'patient_data_chunk_count';
  let data = [];

  const chunkCountStr = cache.get(CACHE_KEY_COUNT);

  if (chunkCountStr) {
    const headerRow = JSON.parse(cache.get(CACHE_KEY_HEADER));
    if (!headerRow) {
      cache.removeAll([CACHE_KEY_HEADER, CACHE_KEY_COUNT]);
    } else {
        const chunkCount = parseInt(chunkCountStr, 10);
        const keys = [];
        for (let i = 0; i < chunkCount; i++) {
            keys.push(CACHE_KEY_PREFIX + i);
        }
        
        const cachedChunks = cache.getAll(keys);
        let patientRows = [];
        for (let i = 0; i < chunkCount; i++) {
            const chunkKey = CACHE_KEY_PREFIX + i;
            if (cachedChunks[chunkKey]) {
                const rows = JSON.parse(cachedChunks[chunkKey]);
                patientRows = patientRows.concat(rows);
            }
        }
        data = [headerRow].concat(patientRows);
    }
  } 
  
  if (data.length === 0) {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PAZIENTI);
    const allSheetData = sheet.getDataRange().getValues();
    
    if (allSheetData.length < 1) return [];

    const headerRow = allSheetData.shift();
    const patientRows = allSheetData;
    
    const CHUNK_SIZE = 250;
    const chunks = {};
    let count = 0;

    for (let i = 0; i < patientRows.length; i += CHUNK_SIZE) {
      const chunk = patientRows.slice(i, i + CHUNK_SIZE);
      const chunkKey = CACHE_KEY_PREFIX + count;
      chunks[chunkKey] = JSON.stringify(chunk);
      count++;
    }
    
    chunks[CACHE_KEY_COUNT] = count.toString();
    chunks[CACHE_KEY_HEADER] = JSON.stringify(headerRow);
    
    cache.putAll(chunks, 21600);
    
    data = [headerRow].concat(patientRows);
  }

  if (data.length === 0) {
      return [];
  }
  
  const headers = data[0].map(h => String(h || '').toLowerCase().trim());
  const searchTermLower = searchTerm.toLowerCase().trim();
  
  const suggestions = [];
  
  const codiceFiscaleIndex = headers.indexOf('codicefiscale');
  const cognomeIndex = headers.indexOf('cognome');
  const nomeIndex = headers.indexOf('nome');

  if (codiceFiscaleIndex === -1 || cognomeIndex === -1 || nomeIndex === -1) {
    Logger.log("ERRORE: Colonne 'codicefiscale', 'cognome', o 'nome' non trovate.");
    return [];
  }

  const searchWords = searchTermLower.split(' ').filter(word => word.length > 0);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const codiceFiscale = (row[codiceFiscaleIndex] || '').toString().toLowerCase().trim();
    const cognome = (row[cognomeIndex] || '').toString().toLowerCase().trim();
    const nome = (row[nomeIndex] || '').toString().toLowerCase().trim();

    const searchablePatientString = `${nome} ${cognome} ${codiceFiscale}`;
    
    const isMatch = searchWords.every(word => searchablePatientString.includes(word));
    
    if (isMatch) {
      let patientData = {};
      headers.forEach((headerKey, index) => {
        patientData[headerKey] = row[index];
      });
      suggestions.push(patientData);
      
      if (suggestions.length >= 10) {
        break;
      }
    }
  }
  
  return suggestions;
}


/**
 * Recupera l'elenco dei vaccini disponibili.
 */
function getVaccineData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_VACCINI);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const denominazioneIndex = headers.indexOf('denominazionevaccino');
  const lottoIndex = headers.indexOf('numerolotto');
  const statoIndex = headers.indexOf('stato');

  const availableVaccines = data.filter(row => {
    const status = (row[statoIndex] || '').toString().toLowerCase().trim();
    return status !== 'completato';
  });

  const vaccines = availableVaccines.map(row => {
    return {
      denominazione: row[denominazioneIndex],
      lotto: row[lottoIndex]
    };
  });

  return vaccines;
}

/**
 * Salva i dati del consenso, genera il PDF e lo salva in Drive.
 */
function submitConsentForm(formData) {
  try {
    const folder = DriveApp.getFoldersByName(PDF_FOLDER_NAME).hasNext()
      ? DriveApp.getFoldersByName(PDF_FOLDER_NAME).next()
      : DriveApp.createFolder(PDF_FOLDER_NAME);

    // MODIFICATO: La funzione createPdfHtml è stata completamente riscritta
    const htmlForPdf = createPdfHtml(formData);
    const pdfBlob = Utilities.newBlob(htmlForPdf, MimeType.HTML).getAs(MimeType.PDF);
    pdfBlob.setName('Consenso_' + formData.cognome + '_' + formData.nome + '_' + new Date().toISOString().slice(0,10) + '.pdf');
    
    const pdfFile = folder.createFile(pdfBlob);
    const pdfUrl = pdfFile.getUrl();

    const consensiSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CONSENSI);
    const headers = consensiSheet.getRange(1, 1, 1, consensiSheet.getLastColumn()).getValues()[0];
    
    const newRow = headers.map(header => {
      const key = header.toLowerCase().trim();
      switch(key) {
        case 'timestamp': return new Date();
        case 'cognome': return formData.cognome || '';
        case 'nome': return formData.nome || '';
        case 'codicefiscale': return formData.codicefiscale || '';
        case 'datanascita': return formData.datanascita || '';
        case 'denominazionevaccino': return formData.vaccinoDenominazione || '';
        case 'numerolotto': return formData.vaccinoLotto || '';
        case 'luogovaccinazione': return formData.luogoVaccinazione || '';
        case 'esito': return 'Completato';
        case 'consensosomministrazione': return formData.consenso === 'Acconsente' ? 'Sì' : 'No';
        // NUOVO: Aggiunta colonna per consenso privacy
        case 'consensoprivacy': return formData.privacy === 'Acconsente' ? 'Sì' : 'No';
        case 'pdfurl': return pdfUrl;
        case 'hashpaziente': return formData.firmaPaziente ? 'Presente' : 'Mancante';
        case 'hashmedico': return formData.firmaMedico ? 'Presente' : 'Mancante';
        default: return formData[key] || '';
      }
    });

    consensiSheet.appendRow(newRow);

    // MODIFICATO: La risposta di successo ora non contiene più l'URL del PDF.
    // Il frontend mostrerà un toast generico di successo.
    return { status: 'success' };

  } catch (error) {
    Logger.log(error.toString());
    return { status: 'error', message: error.toString() };
  }
}

/**
 * MODIFICATO: Funzione completamente riscritta per un PDF più pulito e completo.
 * Crea l'HTML per il PDF utilizzando un layout migliorato.
 */
function createPdfHtml(data) {
  var consensoVaccinoText = data.consenso === 'Acconsente' 
    ? 'ACCONSENTE AD ESSERE SOTTOPOSTO/A ALLA VACCINAZIONE' 
    : 'NON ACCONSENTE AD ESSERE SOTTOPOSTO/A ALLA VACCINAZIONE';
    
  var consensoPrivacyText = data.privacy === 'Acconsente'
    ? 'ACCONSENTE al trattamento dei dati personali e biometrici'
    : 'NON ACCONSENTE al trattamento dei dati personali e biometrici';

  var html = "<html><head><style>" +
      "body { font-family: 'Helvetica', 'Arial', sans-serif; margin: 40px; font-size: 10pt; color: #333; line-height: 1.4; }" +
      "h1 { color: #2563EB; text-align: center; border-bottom: 2px solid #2563EB; padding-bottom: 10px; font-size: 18pt; }" +
      ".section { margin-top: 20px; page-break-inside: avoid; }" +
      ".section h2 { font-size: 13pt; color: #333; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-bottom: 12px;}" +
      ".row { display: flex; justify-content: space-between; margin-bottom: 8px; }" +
      ".row .item { width: 48%; }" +
      ".row .item-full { width: 100%; }" +
      ".label { font-weight: bold; }" +
      ".value { color: #555; }" +
      ".centered-text { text-align: center; }" +
      ".signature-container { margin-top: 40px; display: flex; justify-content: space-around; }" +
      ".signature-box { width: 45%; text-align: center; }" +
      ".signature-box p { margin-bottom: 5px; font-weight: bold;}" +
      ".signature-box img { max-width: 250px; height: 80px; border-bottom: 1px solid #888; margin-bottom: 5px; }" +
      ".signature-box .doctor-title { font-size: 9pt; font-style: italic; color: #666; }" +
      "ul { padding-left: 20px; text-align: justify; }" +
      "li { margin-bottom: 5px; }" +
      ".final-statement { margin-top: 15px; font-size: 11pt; text-align: center; font-weight: bold; }" +
      "</style></head><body>" +
      "<h1>Modulo di Consenso alla Vaccinazione</h1>" +
      
      "<div class='section'><h2>1. Dati Anagrafici Paziente</h2>" +
        "<div class='row'><div class='item'><span class='label'>Cognome:</span> <span class='value'>" + (data.cognome || '') + "</span></div><div class='item'><span class='label'>Nome:</span> <span class='value'>" + (data.nome || '') + "</span></div></div>" +
        "<div class='row'><div class='item'><span class='label'>Sesso:</span> <span class='value'>" + (data.sesso || '') + "</span></div><div class='item'><span class='label'>Data di Nascita:</span> <span class='value'>" + (data.datanascita || '') + "</span></div></div>" +
        "<div class='row'><div class='item-full'><span class='label'>Codice Fiscale:</span> <span class='value'>" + (data.codicefiscale || '') + "</span></div></div>" +
      "</div>" +

      "<div class='section'><h2>2. Residenza e Contatti</h2>" +
        "<div class='row'><div class='item'><span class='label'>Iscritto al SSR:</span> <span class='value'>" + (data.ssr || '') + "</span></div><div class='item'><span class='label'>Residente:</span> <span class='value'>" + (data.residenza || '') + "</span></div></div>" +
        "<div class='row'><div class='item-full'><span class='label'>Indirizzo:</span> <span class='value'>" + (data.indirizzo || '') + "</span></div></div>" +
        "<div class='row'><div class='item'><span class='label'>Comune:</span> <span class='value'>" + (data.comuneresidenza || '') + "</span></div><div class='item'><span class='label'>Telefono:</span> <span class='value'>" + (data.telefono || '') + "</span></div></div>" +
      "</div>" +

      "<div class='section'><h2>3. Dati a cura dell'Operatore Sanitario</h2>" +
        "<div class='row'><div class='item'><span class='label'>Nome vaccino:</span> <span class='value'>" + (data.vaccinoDenominazione || '') + "</span></div><div class='item'><span class='label'>Lotto N:</span> <span class='value'>" + (data.vaccinoLotto || '') + "</span></div></div>" +
        "<div class='row'><div class='item-full'><span class='label'>Luogo vaccinazione:</span> <span class='value'>" + (data.luogoVaccinazione || '') + "</span></div></div>" +
      "</div>" +
      
      "<div class='section'><h2>4. Dichiarazione e Consenso alla Vaccinazione</h2>" +
        "<p>Il/La sottoscritto/a dichiara di:</p>" +
        "<ul>" +
          "<li>aver ricevuto e letto la scheda informativa sintetica relativa alla vaccinazione antinfluenzale;</li>" +
          "<li>essere stato/a informato/a in modo chiaro e comprensibile sui benefici e sui potenziali rischi della vaccinazione;</li>" +
          "<li>aver avuto la possibilità di porre domande e di ricevere risposte adeguate ai propri quesiti;</li>" +
          "<li>aver compreso le informazioni ricevute e di prestare il proprio consenso alla somministrazione del vaccino.</li>" +
        "</ul>" +
        "<h3 class='final-statement'>" + consensoVaccinoText + "</h3>" +
      "</div>" +

      // NUOVA SEZIONE: Consenso Privacy
      "<div class='section'><h2>5. Consenso al Trattamento dei Dati Personali (GDPR)</h2>" +
        "<p>Il/La sottoscritto/a, ai sensi del Regolamento UE 2016/679, dichiara di essere stato/a informato/a che:</p>" +
        "<ul>" +
          "<li>I dati personali e sanitari saranno trattati esclusivamente per finalità connesse alla prestazione sanitaria e agli obblighi di legge.</li>" +
          "<li>Per la sottoscrizione digitale del presente modulo verranno raccolti dati biometrici al solo scopo di garantire l’autenticità, l’integrità e la validità legale della firma elettronica.</li>" +
        "</ul>" +
        "<h3 class='final-statement'>" + consensoPrivacyText + "</h3>" +
      "</div>" +

      "<div class='section'><h2>6. Firme</h2>" +
        "<div class='signature-container'>" +
          "<div class='signature-box'><p>Firma del Paziente</p><img src='" + (data.firmaPaziente || 'https://via.placeholder.com/250x80.png?text=Firma+Mancante') + "' /></div>" +
          "<div class='signature-box'>" +
            "<p>Firma dell'Operatore Sanitario</p>" +
            "<img src='" + (data.firmaMedico || 'https://via.placeholder.com/250x80.png?text=Firma+Mancante') + "' />" +
            // NUOVO: Aggiunta dicitura completa del medico
            "<div class='doctor-title'>Dott.ssa Arianna Baroni<br>Medico Chirurgo<br>062778 G/RM2</div>" +
          "</div>" +
        "</div>" +
      "</div>" +

      // MODIFICATO: Data allineata a sinistra
      "<p style='text-align: left; margin-top: 40px;'><span class='label'>Data sottoscrizione:</span> " + new Date().toLocaleDateString('it-IT') + "</p>" +

    "</body></html>";
  
  return html;
}