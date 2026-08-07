/**
 * Config.gs - Configurazione globale dell'applicazione
 */

// ID del Foglio di calcolo Google (Database)
const SPREADSHEET_ID = '1PC_2uIFExMog5TZvEQPElTj3Ogrs8dEHh-3yQVFiTEU';

// Nomi dei Fogli di lavoro
const SHEET_ACCOUNTS = 'Account';
const SHEET_LOGS = 'Log';
const SHEET_PAZIENTI = 'DB_pazienti';
const SHEET_VACCINI = 'Vaccini';
const SHEET_CONSENSI = 'Consensi';

// Nome Cartella Google Drive per i PDF
const PDF_FOLDER_NAME = 'Consensi Vaccinazioni Firmati';

// Utility per ottenere l'istanza del Foglio
function getDb() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}