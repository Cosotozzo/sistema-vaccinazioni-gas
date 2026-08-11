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

  // Unifichiamo la firma visiva del paziente come nel cartaceo
  const imgPaziente = data.firmaPazienteVaccino 
    ? `<img src="${data.firmaPazienteVaccino}" style="height: 50px; max-width: 100%; object-fit: contain;" />` 
    : `<div style="height: 50px; line-height: 50px; font-size: 8pt; color: #94a3b8; font-style: italic;">[Firma Non Presente]</div>`;

  const imgMedico = data.firmaMedico 
    ? `<img src="${data.firmaMedico}" style="height: 50px; max-width: 100%; object-fit: contain;" />` 
    : `<div style="height: 50px; line-height: 50px; font-size: 8pt; color: #94a3b8; font-style: italic;">[Firma Non Presente]</div>`;

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @page { size: A4; margin: 12mm 15mm; }
        body { font-family: 'Helvetica', 'Arial', sans-serif; font-size: 10pt; color: #000; line-height: 1.4; margin: 0; }
        .header-title { text-align: center; font-size: 14pt; font-weight: bold; margin-bottom: 15px; }
        .section-title { font-size: 11pt; font-weight: bold; margin-top: 15px; margin-bottom: 5px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 5px; }
        td { padding: 3px 0; vertical-align: top; }
        .label { font-weight: bold; }
        .value { font-weight: normal; }
        ul { margin: 5px 0 10px 20px; padding: 0; }
        li { margin-bottom: 5px; text-align: justify; }
        .statement-box { text-align: center; font-size: 10pt; font-weight: bold; padding: 10px; border: 1px solid #000; margin: 10px 0; }
        .signature-title { font-weight: bold; margin-bottom: 5px; }
      </style>
    </head>
    <body>
      <div class="header-title">Modulo di Consenso alla Vaccinazione</div>

      <div class="section-title">1. Dati Anagrafici Paziente</div>
      <table>
        <tr>
          <td width="50%"><span class="label">Cognome:</span> <span class="value">${data.cognome || ''}</span></td>
          <td width="50%"><span class="label">Nome:</span> <span class="value">${data.nome || ''}</span></td>
        </tr>
        <tr>
          <td width="50%"><span class="label">Sesso:</span> <span class="value">${data.sesso || ''}</span></td>
          <td width="50%"><span class="label">Codice Fiscale:</span> <span class="value">${data.codicefiscale || ''}</span></td>
        </tr>
      </table>

      <div class="section-title">2. Residenza e Contatti</div>
      <table>
        <tr>
          <td width="50%"><span class="label">Data di Nascita:</span> <span class="value">${data.datanascita || ''}</span></td>
          <td width="50%"><span class="label">Iscritto al SSR:</span> <span class="value">${data.iscrittossr || 'S.S.R. LAZIO'}</span></td>
        </tr>
        <tr>
          <td width="50%"><span class="label">Indirizzo:</span> <span class="value">${data.indirizzo || ''}</span></td>
          <td width="50%"><span class="label">Comune:</span> <span class="value">${data.comune || ''}</span></td>
        </tr>
        <tr>
          <td width="50%"><span class="label">Residente:</span> <span class="value">${data.residente || 'NELLA REGIONE LAZIO'}</span></td>
          <td width="50%"><span class="label">Telefono:</span> <span class="value">${data.telefono || ''}</span></td>
        </tr>
      </table>

      <div class="section-title">3. Dati a cura dell'Operatore Sanitario</div>
      <table>
        <tr>
          <td width="50%"><span class="label">Nome vaccino:</span> <span class="value">${data.vaccinoDenominazione || ''}</span></td>
          <td width="50%"><span class="label">Luogo vaccinazione:</span> <span class="value">${data.luogoVaccinazione || ''}</span></td>
        </tr>
        <tr>
          <td colspan="2"><span class="label">Lotto N:</span> <span class="value">${data.vaccinoLotto || ''}</span></td>
        </tr>
      </table>

      <div class="section-title">4. Dichiarazione e Consenso alla Vaccinazione</div>
      <p style="margin: 0 0 5px 0;">Il/La sottoscritto/a dichiara di:</p>
      <ul>
        <li>aver ricevuto e letto la scheda informativa sintetica relativa alla vaccinazione antinfluenzale;</li>
        <li>essere stato/a informato/a in modo chiaro e comprensibile sui benefici e sui potenziali rischi della vaccinazione;</li>
        <li>aver avuto la possibilità di porre domande e di ricevere risposte adeguate ai propri quesiti;</li>
        <li>aver compreso le informazioni ricevute e di prestare il proprio consenso alla somministrazione del vaccino.</li>
      </ul>
      <div class="statement-box">${consensoVaccinoText}</div>

      <div class="section-title">5. Consenso al Trattamento dei Dati Personali (GDPR)</div>
      <p style="margin: 0 0 5px 0;">Il/La sottoscritto/a, ai sensi del Regolamento UE 2016/679, dichiara di essere stato/a informato/a che:</p>
      <ul>
        <li>I dati personali e sanitari saranno trattati esclusivamente per finalità connesse alla prestazione sanitaria e agli obblighi di legge.</li>
        <li>Per la sottoscrizione digitale del presente modulo verranno raccolti dati biometrici al solo scopo di garantire l'autenticità, l'integrità e la validità legale della firma elettronica.</li>
      </ul>
      <div class="statement-box">${consensoPrivacyText}</div>

      <div class="section-title">6. Firme</div>
      <table style="margin-top: 15px;">
        <tr>
          <td width="50%" align="center">
            <div class="signature-title">Firma del Paziente</div>
            ${imgPaziente}
          </td>
          <td width="50%" align="center">
            <div class="signature-title">Firma dell'Operatore Sanitario</div>
            ${imgMedico}
          </td>
        </tr>
        <tr>
          <td width="50%" align="left" style="padding-top: 15px;">
            <span class="label">Data sottoscrizione:</span> <span class="value">${new Date().toLocaleDateString('it-IT')}</span>
          </td>
          <td width="50%" align="center" style="padding-top: 15px;">
            <span class="label">Dott.ssa Arianna Baroni</span><br>
            Medico Chirurgo<br>
            062778 G/RM2
          </td>
        </tr>
      </table>
    </body>
    </html>
  `;
}
