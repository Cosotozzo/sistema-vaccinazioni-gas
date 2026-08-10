/**
 * Auth.gs - Modulo di Autenticazione e Sicurezza
 */

function checkLogin(email, password) {
  const sheet = getDb().getSheetByName(SHEET_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  const inputHash = hashPassword(password);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email && data[i][3] === inputHash) { 
      const fullName = `${data[i][0]} ${data[i][1]}`;
      // Recupero ruolo da Colonna E (indice 4), default 'medico' se vuoto
      const role = (data[i][4] || 'medico').toString().toLowerCase().trim();
      
      logAction(fullName, `Login effettuato (Ruolo: ${role})`);
      return { 
        success: true, 
        name: fullName, 
        role: role 
      };
    }
  }
  return { success: false, message: 'Credenziali non valide o account inesistente.' };
}

function registerUser(nome, cognome, email, password, confirmPassword) {
  if (password !== confirmPassword) {
    return { success: false, message: 'Le due password non coincidono.' };
  }

  if (!validatePasswordRules(password)) {
    return { success: false, message: 'La password deve contenere almeno 8 caratteri, 1 carattere speciale e almeno 2 numeri.' };
  }

  const sheet = getDb().getSheetByName(SHEET_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email) {
      return { success: false, message: 'Indirizzo email già registrato.' };
    }
  }
  
  const hashedPass = hashPassword(password);
  // Di default le nuove registrazioni vengono create come 'segretario' (modificabile da DB)
  sheet.appendRow([nome, cognome, email, hashedPass, 'segretario']);
  logAction(`${nome} ${cognome}`, 'Registrazione nuova utenza (segretario)');
  
  return { success: true, message: 'Account registrato con successo!' };
}

function recoverPassword(email) {
  const sheet = getDb().getSheetByName(SHEET_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email) {
      const fullName = `${data[i][0]} ${data[i][1]}`;
      const tempPass = Math.random().toString(36).slice(-8) + "12!"; 
      const tempHash = hashPassword(tempPass);
      
      sheet.getRange(i + 1, 4).setValue(tempHash);
      
      const subject = "Reset Accesso - Studio Medico";
      const body = `Gentile ${fullName},\n\nLa tua nuova password temporanea è: ${tempPass}\n\nAmministrazione Studio Medico`;
      
      GmailApp.sendEmail(email, subject, body);
      logAction(fullName, 'Richiesta reset password');
      return { success: true, message: 'Password temporanea inviata via email.' };
    }
  }
  return { success: false, message: 'Email non trovata nel database.' };
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
  const isLongEnough = password.length >= 8;
  return numbersCount >= 2 && hasSpecial && isLongEnough;
}