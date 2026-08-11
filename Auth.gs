/**
 * Auth.gs - Modulo di Autenticazione e Sicurezza
 */

function checkLogin(email, password) {
  // Username CASE-INSENSITIVE
  const cleanEmail = (email || '').toString().toLowerCase().trim();
  // Password CASE-SENSITIVE
  const inputHash = hashPassword(password); 
  
  const sheet = getDb().getSheetByName(SHEET_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const dbEmail = (data[i][2] || '').toString().toLowerCase().trim();
    const dbHash = data[i][3];
    
    if (dbEmail === cleanEmail && dbHash === inputHash) { 
      const fullName = `${data[i][0]} ${data[i][1]}`;
      const role = (data[i][4] || '').toString().toLowerCase().trim();
      
      // Controllo Ruolo: Se non è assegnato 'medico' o 'segretario'
      if (role !== 'medico' && role !== 'segretario') {
        logAction(fullName, 'Login bloccato: Ruolo non autorizzato');
        return { 
          success: false, 
          isUnauthorized: true, // Questo flag attiva la schermata specifica lato Frontend
          message: 'Utenza non autorizzata.' 
        };
      }
      
      logAction(fullName, `Login effettuato (Ruolo: ${role})`);
      return { success: true, name: fullName, role: role };
    }
  }
  return { success: false, isUnauthorized: false, message: 'Credenziali non valide o account inesistente.' };
}

function registerUser(nome, cognome, email, password, confirmPassword) {
  if (password !== confirmPassword) return { success: false, message: 'Le password non coincidono.' };
  if (!validatePasswordRules(password)) return { success: false, message: 'La password non rispetta i requisiti di sicurezza.' };

  const cleanEmail = (email || '').toString().toLowerCase().trim();
  const sheet = getDb().getSheetByName(SHEET_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toString().toLowerCase().trim() === cleanEmail) {
      return { success: false, message: 'Indirizzo email già registrato.' };
    }
  }
  
  const hashedPass = hashPassword(password);
  sheet.appendRow([nome, cognome, cleanEmail, hashedPass, '']);
  logAction(`${nome} ${cognome}`, 'Registrazione nuova utenza');
  
  return { success: true, message: 'Registrazione completata! In attesa di approvazione.' };
}

function recoverPassword(email) {
  const cleanEmail = (email || '').toString().toLowerCase().trim();
  const sheet = getDb().getSheetByName(SHEET_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toString().toLowerCase().trim() === cleanEmail) {
      const tempPass = Math.random().toString(36).slice(-8) + "12!"; 
      sheet.getRange(i + 1, 4).setValue(hashPassword(tempPass));
      GmailApp.sendEmail(cleanEmail, "Reset Accesso", `La tua password temporanea è: ${tempPass}`);
      return { success: true, message: 'Password temporanea inviata via email.' };
    }
  }
  return { success: false, message: 'Email non trovata.' };
}

function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length === 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function validatePasswordRules(password) {
  const numbersCount = (password.match(/\d/g) || []).length;
  const hasSpecial = /[!@#$%^&*()_+\-=\[\]{};':"\\|,.<>\/?]+/.test(password);
  return numbersCount >= 2 && hasSpecial && password.length >= 8;
}