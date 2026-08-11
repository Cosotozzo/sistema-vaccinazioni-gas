/**
 * Config.gs - Configurazione globale dell'applicazione
 */

// ID del Foglio di calcolo Google (Database)
// ATTENZIONE: Questo dato rimane lato server e NON è mai visibile nel frontend per ragioni di sicurezza.
const SPREADSHEET_ID = '1PC_2uIFExMog5TZvEQPElTj3Ogrs8dEHh-3yQVFiTEU';

// Nomi dei Fogli di lavoro (Tab) presenti all'interno del Google Sheets
const SHEET_ACCOUNTS = 'Account';
const SHEET_LOGS = 'Log';
const SHEET_PAZIENTI = 'DB_pazienti';
const SHEET_VACCINI = 'Vaccini';
const SHEET_CONSENSI = 'Consensi';

// Nome Cartella Google Drive per il salvataggio dei PDF
// Nota: Questa è la cartella che verrà sincronizzata automaticamente con il NAS Synology
const PDF_FOLDER_NAME = 'Consensi Vaccinazioni Firmati';

// Utility globale per ottenere rapidamente l'istanza del Foglio Database in tutto il codice backend
function getDb() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}