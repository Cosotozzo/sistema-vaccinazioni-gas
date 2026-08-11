/**
 * PdfEngine.gs - Motore di rendering e generazione documenti PDF
 */

function submitConsentForm(formData) {
  try {
    const isVaccineAccepted = formData.consensoVaccino === 'Acconsente';
    const isGdprAccepted = formData.consensoPrivacy === 'Acconsente';
    
    let pdfUrl = '';

    const folderIter = DriveApp.getFoldersByName(PDF_FOLDER_NAME);
const folder = folderIter.hasNext() ? folderIter.next() : DriveApp.createFolder(PDF_FOLDER_NAME);

    const localDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const htmlForPdf = createPdfHtml(formData);
    const pdfBlob = Utilities.newBlob(htmlForPdf, MimeType.HTML).getAs(MimeType.PDF);
    pdfBlob.setName(`Consenso_${formData.cognome}_${formData.nome}_${localDateStr}.pdf`);
    
    const pdfFile = folder.createFile(pdfBlob);
    pdfUrl = pdfFile.getUrl();

    // Registrazione sul Foglio Consensi
    const consensiSheet = getDb().getSheetByName(SHEET_CONSENSI);
    const headers = consensiSheet.getRange(1, 1, 1, consensiSheet.getLastColumn()).getValues()[0];
    
    const newRow = headers.map(header => {
      const key = header.toLowerCase().trim().replace(/\s+/g, '');
      switch(key) {
        case 'timestamp': return new Date();
        case 'cognome': return formData.cognome || '';
        case 'nome': return formData.nome || '';
        case 'codicefiscale': return formData.codicefiscale || '';
        case 'datanascita': return formData.datanascita || '';
        case 'denominazionevaccino': return formData.vaccinoDenominazione || '';
        case 'numerolotto': return isVaccineAccepted ? (formData.vaccinoLotto || '') : 'NON SOMMINISTRATO';
        case 'luogovaccinazione': return formData.luogoVaccinazione || '';
        case 'esito': return isVaccineAccepted ? (isGdprAccepted ? 'Completato Digitale' : 'Completato Cartaceo') : 'Diniego/Rifiuto';
        case 'consensoprivacy': return isGdprAccepted ? 'Sì' : 'No';
        case 'consensosomministrazione': return isVaccineAccepted ? 'Sì' : 'No';
        case 'pdfurl': return pdfUrl;
        case 'hashpaziente': return formData.firmaPazienteVaccino ? 'Presente' : 'Mancante';
        case 'hashmedico': return formData.firmaMedico ? 'Presente' : 'Mancante';
        default: return formData[key] || '';
      }
    });

    consensiSheet.appendRow(newRow);
    
    // REGOLA LOGICA DOSI:
    // Se il vaccino è ACCONSENTITO (sia con GDPR che senza GDPR), scala 1 dose per allineare il magazzino.
    // Se il vaccino è NEGATO, NON scalare dosi.
    let isPaperRequired = false;

    if (isVaccineAccepted) {
      decrementVaccineDose(formData.vaccinoDenominazione, formData.vaccinoLotto);
      if (!isGdprAccepted) {
        isPaperRequired = true;
      }
    }

    logAction('Sistema', `Modulo registrato per ${formData.cognome} ${formData.nome} (Vaccino: ${formData.consensoVaccino}, GDPR: ${formData.consensoPrivacy})`);
    
    return { 
      status: 'success', 
      pdfUrl: pdfUrl,
      isPaperRequired: isPaperRequired,
      message: isPaperRequired 
        ? "Dose scalata dal magazzino. NECESSARIA FIRMA SU MODULO CARTACEO." 
        : "Operazione completata e registrata con successo."
    };

  } catch (error) {
    Logger.log(error.toString());
    return { status: 'error', message: error.toString() };
  }
}

function createPdfHtml(data) {
  const consensoVaccinoText = data.consensoVaccino === 'Acconsente' 
    ? 'ACCONSENTE AD ESSERE SOTTOPOSTO/A ALLA VACCINAZIONE' 
    : 'NON ACCONSENTE AD ESSERE SOTTOPOSTO/A ALLA VACCINAZIONE';
    
  const consensoPrivacyText = data.consensoPrivacy === 'Acconsente'
    ? 'ACCONSENTE al trattamento dei dati personali e biometrici'
    : 'NON ACCONSENTE al trattamento dei dati personali e biometrici';

  const imgPrivacy = data.firmaPazienteBiometrico 
    ? `<img src="${data.firmaPazienteBiometrico}" style="height: 42px; max-width: 100%; object-fit: contain;" />` 
    : `<div style="height: 42px; line-height: 42px; font-size: 8pt; color: #94a3b8; font-style: italic;">[Firma Non Presente]</div>`;

  const imgVaccino = data.firmaPazienteVaccino 
    ? `<img src="${data.firmaPazienteVaccino}" style="height: 42px; max-width: 100%; object-fit: contain;" />` 
    : `<div style="height: 42px; line-height: 42px; font-size: 8pt; color: #94a3b8; font-style: italic;">[Firma Non Presente]</div>`;

  const imgMedico = data.firmaMedico 
    ? `<img src="${data.firmaMedico}" style="height: 42px; max-width: 100%; object-fit: contain;" />` 
    : `<div style="height: 42px; line-height: 42px; font-size: 8pt; color: #94a3b8; font-style: italic;">[Firma Non Presente]</div>`;

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @page { size: A4; margin: 10mm 12mm; }
        body { font-family: 'Helvetica Neue', Arial, sans-serif; font-size: 9.5pt; color: #1e293b; line-height: 1.3; margin: 0; }
        .header-title { text-align: center; color: #1d4ed8; font-size: 15pt; font-weight: bold; border-bottom: 2px solid #1d4ed8; padding-bottom: 6px; margin-bottom: 12px; }
        .section-block { margin-bottom: 10px; background: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px 12px; }
        .section-title { font-size: 10.5pt; font-weight: bold; color: #0f172a; border-bottom: 1px solid #e2e8f0; padding-bottom: 3px; margin-bottom: 6px; text-transform: uppercase; }
        .label { font-weight: bold; color: #475569; }
        .value { font-weight: 600; color: #0f172a; }
        ul { margin: 3px 0 6px 16px; padding: 0; }
        li { margin-bottom: 2px; text-align: justify; }
        .statement-box { text-align: center; font-size: 9.5pt; font-weight: bold; color: #0f172a; background: #f1f5f9; padding: 5px; border-radius: 4px; margin-top: 4px; }
        .signature-box { width: 210px; text-align: center; border: 1px solid #cbd5e1; padding: 4px; border-radius: 6px; background: #fafafa; }
        .signature-box p { font-size: 8pt; font-weight: bold; margin: 0 0 2px 0; color: #334155; border-bottom: 1px solid #cbd5e1; padding-bottom: 2px; }
        .doctor-title { font-size: 7.5pt; font-style: italic; color: #64748b; margin-top: 2px; }
      </style>
    </head>
    <body>
      <div class="header-title">Modulo di Consenso alla Vaccinazione</div>

      <!-- SEZIONE 1: ANAGRAFICA -->
      <div class="section-block">
        <div class="section-title">1. Dati Anagrafici Paziente</div>
        <table width="100%" border="0" cellspacing="0" cellpadding="2">
          <tr>
            <td width="50%"><span class="label">Cognome:</span> <span class="value">${data.cognome || ''}</span></td>
            <td width="50%"><span class="label">Nome:</span> <span class="value">${data.nome || ''}</span></td>
          </tr>
          <tr>
            <td width="50%"><span class="label">Codice Fiscale:</span> <span class="value">${data.codicefiscale || ''}</span></td>
            <td width="50%"><span class="label">Data di Nascita:</span> <span class="value">${data.datanascita || ''}</span></td>
          </tr>
        </table>
      </div>

      <!-- SEZIONE 2: GDPR PRIVACY -->
      <div class="section-block">
        <div class="section-title">2. Consenso al Trattamento Dati Personali e Biometrici (GDPR)</div>
        <p style="margin: 0 0 4px 0;">Il/La sottoscritto/a, ai sensi del Regolamento UE 2016/679, dichiara che:</p>
        <ul>
          <li>I dati personali e sanitari saranno trattati esclusivamente per finalità connesse alla prestazione sanitaria.</li>
          <li>Verranno raccolti dati biometrici al solo scopo di garantire l'autenticità e la validità legale della firma elettronica.</li>
        </ul>
        <div class="statement-box">${consensoPrivacyText}</div>
        
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top: 8px;">
          <tr>
            <td align="right">
              <div class="signature-box">
                <p>Firma Paziente (GDPR / Privacy)</p>
                ${imgPrivacy}
              </div>
            </td>
          </tr>
        </table>
      </div>

      <!-- SEZIONE 3: SOMMINISTRAZIONE E FIRME -->
      <div class="section-block">
        <div class="section-title">3. Dati e Consenso alla Vaccinazione</div>
        <table width="100%" border="0" cellspacing="0" cellpadding="2" style="margin-bottom: 4px;">
          <tr>
            <td width="50%"><span class="label">Nome vaccino:</span> <span class="value">${data.vaccinoDenominazione || ''}</span></td>
            <td width="50%"><span class="label">Lotto N:</span> <span class="value">${data.vaccinoLotto || ''}</span></td>
          </tr>
          <tr>
            <td colspan="2"><span class="label">Luogo vaccinazione:</span> <span class="value">${data.luogoVaccinazione || ''}</span></td>
          </tr>
        </table>
        
        <p style="margin: 0 0 4px 0;">Il/La sottoscritto/a dichiara di:</p>
        <ul>
          <li>aver ricevuto e letto la scheda informativa sintetica relativa alla vaccinazione;</li>
          <li>essere stato/a informato/a sui benefici e sui potenziali rischi della vaccinazione;</li>
          <li>aver avuto la possibilità di porre domande e di ricevere risposte adeguate;</li>
          <li>aver compreso le informazioni e di prestare il proprio consenso alla somministrazione.</li>
        </ul>
        <div class="statement-box">${consensoVaccinoText}</div>
        
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top: 8px;">
          <tr>
            <td width="50%" align="left" valign="top">
              <div class="signature-box">
                <p>Firma Operatore Sanitario</p>
                ${imgMedico}
                <div class="doctor-title">Dott.ssa Arianna Baroni<br>Medico Chirurgo</div>
              </div>
            </td>
            <td width="50%" align="right" valign="top">
              <div class="signature-box">
                <p>Firma Paziente (Consenso Vaccino)</p>
                ${imgVaccino}
              </div>
            </td>
          </tr>
        </table>
      </div>

      <div style="font-size: 8pt; color: #64748b; text-align: right; margin-top: 4px;">
        Data e ora sottoscrizione: <strong>${new Date().toLocaleDateString('it-IT')} ${new Date().toLocaleTimeString('it-IT')}</strong>
      </div>
    </body>
    </html>
  `;
}