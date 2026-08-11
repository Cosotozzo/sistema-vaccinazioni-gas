/**
 * Auth.gs - Modulo di Autenticazione e Sicurezza (Email Case-Insensitive, Password Case-Sensitive)
 */

function checkLogin(email, password) {
  const cleanEmail = (email || '').toString().toLowerCase().trim();
  const inputHash = hashPassword(password); // La password rimane case-sensitive
  
  const sheet = getDb().getSheetByName(SHEET_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const dbEmail = (data[i][2] || '').toString().toLowerCase().trim();
    const dbHash = data[i][3];
    
    if (dbEmail === cleanEmail && dbHash === inputHash) { 
      const fullName = `${data[i][0]} ${data[i][1]}`;
      const role = (data[i][4] || '').toString().toLowerCase().trim();
      
      // BLOCCO ACCESSO: Se il ruolo non è assegnato dall'Amministratore
      if (!role) {
        logAction(fullName, 'Tentativo di login fallito: Ruolo non ancora assegnato');
        return { 
          success: false, 
          message: 'Account registrato ma non ancora abilitato. In attesa di assegnazione del ruolo da parte dell\'Amministratore.' 
        };
      }
      
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

  const cleanEmail = (email || '').toString().toLowerCase().trim();
  const sheet = getDb().getSheetByName(SHEET_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  
  // Controllo duplicati case-insensitive
  for (let i = 1; i < data.length; i++) {
    const dbEmail = (data[i][2] || '').toString().toLowerCase().trim();
    if (dbEmail === cleanEmail) {
      return { success: false, message: 'Indirizzo email già registrato.' };
    }
  }
  
  const hashedPass = hashPassword(password);
  // Salviamo l'email già normalizzata in minuscolo sul DB
  sheet.appendRow([nome, cognome, cleanEmail, hashedPass, '']);
  logAction(`${nome} ${cognome}`, 'Registrazione nuova utenza (In attesa di attivazione)');
  
  return { 
    success: true, 
    message: 'Registrazione completata! Il tuo account sarà attivo non appena l\'Amministratore assegnerà il ruolo dal database.' 
  };
}

function recoverPassword(email) {
  const cleanEmail = (email || '').toString().toLowerCase().trim();
  const sheet = getDb().getSheetByName(SHEET_ACCOUNTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const dbEmail = (data[i][2] || '').toString().toLowerCase().trim();
    if (dbEmail === cleanEmail) {
      const fullName = `${data[i][0]} ${data[i][1]}`;
      const tempPass = Math.random().toString(36).slice(-8) + "12!"; 
      const tempHash = hashPassword(tempPass);
      
      sheet.getRange(i + 1, 4).setValue(tempHash);
      
      const subject = "Reset Accesso - Studio Medico";
      const body = `Gentile ${fullName},\n\nLa tua nuova password temporanea è: ${tempPass}\n\nAmministrazione Studio Medico`;
      
      GmailApp.sendEmail(cleanEmail, subject, body);
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