// ==========================================
// ENHANCED CONFIGURATION & SETUP
// ==========================================
const TEMPLATE_ID = '1lv30v34OWkInWpn6L0g3aevHgIGV-g0lNtAwHETXLvs'; 
const FOLDER_ID = '1IMbUPOQoJXqDFRxbaWIiWsoMpwdTA7cl';   
const SHEET_LOANS = "Loans";
const SHEET_PENDING = "Pending";
const SHEET_CONFIG = "Config";
const SHEET_AUDIT = "Audit Log";

// ==========================================
// FIXED AUTH SYSTEM - NO OAUTH REQUIRED
// ==========================================

function initializeSystem() {
  const scriptProps = PropertiesService.getScriptProperties();
  if (!scriptProps.getProperty('adminPassword')) {
    scriptProps.setProperty('adminPassword', 'admin123');
  }
  if (!scriptProps.getProperty('managers')) {
    scriptProps.setProperty('managers', 'admin@pawnme.com');
  }
  if (!scriptProps.getProperty('sessionStore')) {
    scriptProps.setProperty('sessionStore', '{}');
  }
}

function loginUser(password, email) {
  initializeSystem();
  const scriptProps = PropertiesService.getScriptProperties();
  const correctPassword = scriptProps.getProperty('adminPassword');
  
  if (password !== correctPassword) {
    return { success: false, message: 'Invalid password' };
  }
  
  // Generate session token
  const sessionToken = Utilities.getUuid();
  
  // Check if user is manager (bootstrap: if managers is still the default placeholder,
  // promote the first successfully authenticated user to Manager)
  const managersRaw = (scriptProps.getProperty('managers') || '').trim();
  const managers = managersRaw
    .split(',')
    .map(e => e.trim().toLowerCase())
    .filter(Boolean);
  let isManagerUser = managers.includes(String(email).toLowerCase());
  if (!isManagerUser && (managers.length === 0 || (managers.length === 1 && managers[0] === 'admin@pawnme.com'))) {
    scriptProps.setProperty('managers', String(email).toLowerCase());
    isManagerUser = true;
  }
  
  // Store session with 2-hour expiry
  const sessions = JSON.parse(scriptProps.getProperty('sessionStore') || '{}');
  sessions[sessionToken] = {
    email: email,
    isManager: isManagerUser,
    timestamp: Date.now()
  };
  
  // Clean up old sessions (older than 2 hours)
  Object.keys(sessions).forEach(token => {
    if (Date.now() - sessions[token].timestamp > 7200000) {
      delete sessions[token];
    }
  });
  
  scriptProps.setProperty('sessionStore', JSON.stringify(sessions));
  
  return {
    success: true,
    sessionToken: sessionToken,
    role: isManagerUser ? 'Manager' : 'Clerk',
    email: email
  };
}

function getUserRole(sessionToken) {
  if (!sessionToken) return 'Clerk';
  
  const scriptProps = PropertiesService.getScriptProperties();
  const sessions = JSON.parse(scriptProps.getProperty('sessionStore') || '{}');
  const session = sessions[sessionToken];
  
  if (session && Date.now() - session.timestamp < 7200000) {
    return session.isManager ? 'Manager' : 'Clerk';
  }
  return 'Clerk';
}

function validateSession(sessionToken) {
  if (!sessionToken) return false;
  
  const scriptProps = PropertiesService.getScriptProperties();
  const sessions = JSON.parse(scriptProps.getProperty('sessionStore') || '{}');
  const session = sessions[sessionToken];
  
  return session && Date.now() - session.timestamp < 7200000;
}

function updateManagerList(emails, sessionToken) {
  if (!validateSession(sessionToken)) {
    return { success: false, message: 'Session expired' };
  }

  const scriptProps = PropertiesService.getScriptProperties();
  const currentManagersRaw = (scriptProps.getProperty('managers') || '').trim();
  const currentManagers = currentManagersRaw
    .split(',')
    .map(e => e.trim().toLowerCase())
    .filter(Boolean);

  const isBootstrap = currentManagers.length === 0 || (currentManagers.length === 1 && currentManagers[0] === 'admin@pawnme.com');
  if (!isBootstrap && getUserRole(sessionToken) !== 'Manager') {
    return { success: false, message: 'Unauthorized' };
  }

  scriptProps.setProperty('managers', emails);

  const sessions = JSON.parse(scriptProps.getProperty('sessionStore') || '{}');
  const session = sessions[sessionToken];
  if (session) {
    const nextManagers = String(emails || '')
      .split(',')
      .map(e => e.trim().toLowerCase())
      .filter(Boolean);
    session.isManager = nextManagers.includes(String(session.email).toLowerCase());
    sessions[sessionToken] = session;
    scriptProps.setProperty('sessionStore', JSON.stringify(sessions));
  }

  return { success: true, message: "Managers updated" };
}

function getManagerList(sessionToken) {
  if (!validateSession(sessionToken)) return '';
  
  const scriptProps = PropertiesService.getScriptProperties();
  return scriptProps.getProperty('managers') || '';
}

function changePassword(newPassword, sessionToken) {
  if (!validateSession(sessionToken)) {
    return { success: false, message: 'Session expired' };
  }
  
  if (newPassword.length < 4) {
    return { success: false, message: 'Password too short' };
  }
  
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty('adminPassword', newPassword);
  return { success: true, message: 'Password updated' };
}

// ==========================================
// ROUTING
// ==========================================
function doGet(e) {
  if (e.parameter.page === 'form') {
    return HtmlService.createTemplateFromFile('Form')
      .evaluate().setTitle('Apply Now | PawnMe')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createTemplateFromFile('index')
      .evaluate().setTitle('PawnMe Vault Admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// AUTO-REPAIR SYSTEM
// ==========================================
function debugSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Loans Sheet
  let ws = ss.getSheetByName(SHEET_LOANS);
  if (!ws) {
    ws = ss.insertSheet(SHEET_LOANS);
    ws.appendRow([
      "ID","Name","Principal","Rate","Months","Total Due","Status",
      "Item Type","Brand","Serial","Condition","Est Value",
      "National ID","Phone","Email","Address","Country",
      "Contract URL","Photo URL","Due Date","Paid Amount",
      "Extension Fee","Notes","Sale Price","Sale Date"
    ]);
  }
  
  // Pending Sheet
  let pws = ss.getSheetByName(SHEET_PENDING);
  if (!pws) {
    pws = ss.insertSheet(SHEET_PENDING);
    pws.appendRow(["Timestamp","Full Name","Email","Phone","National ID","Address","Country","Item Type","Brand","Serial","Condition","Requested Amount","Requested Duration","Photo"]);
  }
  
  // Config Sheet
  let cfg = ss.getSheetByName(SHEET_CONFIG);
  if (!cfg) {
    cfg = ss.insertSheet(SHEET_CONFIG);
    cfg.appendRow(["Setting", "Value"]);
  }
  
  // Audit Log Sheet
  let audit = ss.getSheetByName(SHEET_AUDIT);
  if (!audit) {
    audit = ss.insertSheet(SHEET_AUDIT);
    audit.appendRow([
      "Timestamp","User Email","User Role","Action","Target ID","Target Type","Details","IP Address"
    ]);
  }
  
  // Initialize auth system
  initializeSystem();
  
  return "System Verified. All sheets ready.";
}

// ==========================================
// ENHANCED DASHBOARD & FINANCIALS
// ==========================================
function getStats(sessionToken) {
  if (!validateSession(sessionToken)) {
    return { error: 'Session expired', loans: [] };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_LOANS);
  if (!ws) return { 
    totalLent: 0, realizedGain: 0, bookedLoss: 0, 
    weeklyGain: 0, weeklyLoss: 0, inventory: 0, 
    vaultValue: 0, loans: [], userRole: getUserRole(sessionToken)
  };

  const data = ws.getDataRange().getValues();
  let metrics = { 
    totalLent: 0, realizedGain: 0, bookedLoss: 0, 
    weeklyGain: 0, weeklyLoss: 0, inventory: 0, vaultValue: 0 
  };
  let recentLoans = [];
  
  const now = new Date();
  const oneWeekAgo = new Date();
  oneWeekAgo.setDate(now.getDate() - 7);

  if (data.length > 1) {
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (!row[0]) continue;

      const principal = Number(row[2]) || 0;
      const totalDue = Number(row[5]) || 0;
      const status = String(row[6]);
      const estValue = Number(row[11]) || 0;
      const paidAmount = Number(row[19]) || 0;
      const extFee = Number(row[20]) || 0;
      const salePrice = Number(row[22]) || 0;
      
      let dateStr = "N/A";
      let rowDate = new Date();
      try {
        rowDate = new Date(row[18]);
        dateStr = rowDate.toLocaleDateString();
      } catch(e) {}

      const isRecent = rowDate >= oneWeekAgo;

      if (status === 'Active' || status === 'Defaulted') {
        metrics.totalLent += principal;
      }

      // Inventory & Vault Value
      if (status === 'Active' || status === 'Defaulted') {
        metrics.inventory++;
        metrics.vaultValue += estValue;
      }

      // Financial Calculations
      if (status === "Active") {
        metrics.realizedGain += extFee;
        if (isRecent && extFee > 0) metrics.weeklyGain += extFee;
      }
      else if (status === "Paid") {
        const totalCollected = (paidAmount > 0 ? paidAmount : totalDue) + extFee;
        const netResult = totalCollected - principal;
        if (netResult >= 0) {
          metrics.realizedGain += netResult;
          if (isRecent) metrics.weeklyGain += netResult;
        } else {
          metrics.bookedLoss += Math.abs(netResult);
          if (isRecent) metrics.weeklyLoss += Math.abs(netResult);
        }
      }
      else if (status === "Defaulted") {
        const recovered = paidAmount + extFee;
        const netResult = recovered - principal;
        if (netResult >= 0) {
          metrics.realizedGain += netResult;
          if (isRecent) metrics.weeklyGain += netResult;
        } else {
          metrics.bookedLoss += Math.abs(netResult);
          if (isRecent) metrics.weeklyLoss += Math.abs(netResult);
        }
      }
      else if (status === "Sold") {
        const recovered = paidAmount + extFee + salePrice;
        const netResult = recovered - principal;
        if (netResult >= 0) {
          metrics.realizedGain += netResult;
          if (isRecent) metrics.weeklyGain += netResult;
        } else {
          metrics.bookedLoss += Math.abs(netResult);
          if (isRecent) metrics.weeklyLoss += Math.abs(netResult);
        }
      }

      recentLoans.push({
        id: row[0], name: row[1], amount: principal, total: totalDue,
        status: status, itemType: row[7], brand: row[8], serial: row[9],
        condition: row[10], estValue: estValue, date: dateStr,
        contractUrl: row[16], photoUrl: row[17], email: row[14],
        paidSoFar: paidAmount, extFees: extFee, salePrice: salePrice
      });
    }
  }
  return { ...metrics, loans: recentLoans, userRole: getUserRole(sessionToken) };
}

// ==========================================
// ENHANCED LOAN MANAGEMENT
// ==========================================
function createLoan(formData, sessionToken) {
  if (!validateSession(sessionToken)) {
    return { success: false, message: 'Session expired' };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let ws = ss.getSheetByName(SHEET_LOANS);
    if(!ws) { debugSystem(); ws = ss.getSheetByName(SHEET_LOANS); }
    
    const principal = parseFloat(formData.amount);
    const rate = parseFloat(formData.rate);
    const months = parseInt(formData.duration);
    const interest = principal * (rate / 100) * months;
    const total = principal + interest;
    const estValue = parseFloat(formData.estValue) || (principal * 1.5);
    
    const dueDate = new Date();
    dueDate.setMonth(dueDate.getMonth() + months);
    const id = "LN-" + Date.now().toString().slice(-8);

    // Write immediately
    ws.appendRow([
      id, formData.name, principal, rate + "%", months, total.toFixed(2), 
      "Active", formData.itemType, formData.brand, formData.serial || "", 
      formData.condition, estValue, formData.nationalId, formData.phone, 
      formData.email, formData.address, formData.country || "", "Generating...", "Uploading...", 
      dueDate, 0, 0, "", 0, ""
    ]);
    
    const rowIndex = ws.getLastRow();
    
    // Upload photo
    if (formData.photoUrl) {
      ws.getRange(rowIndex, 18).setValue(formData.photoUrl);
    } else if (formData.imageFile && formData.imageName) {
      const photoUrl = uploadFile(formData.imageFile, "Evidence_" + id + "_" + formData.imageName);
      ws.getRange(rowIndex, 18).setValue(photoUrl);
    }
    
    // Generate contract
    const contractUrl = generatePDF(id, formData, total, dueDate);
    ws.getRange(rowIndex, 17).setValue(contractUrl);
    
    // Send email
    try {
      MailApp.sendEmail({
        to: formData.email,
        subject: `PawnMe - Loan Confirmation ${id}`,
        htmlBody: `<p>Dear ${formData.name},</p>
          <p>Your loan has been approved. Here are the details:</p>
          <ul>
            <li>Loan ID: <b>${id}</b></li>
            <li>Principal: <b>${formatTry_(principal)}</b></li>
            <li>Total Due: <b>${formatTry_(total)}</b></li>
            <li>Due Date: <b>${dueDate.toLocaleDateString()}</b></li>
          </ul>
          <p>Contract: <a href="${contractUrl}">View PDF</a></p>
          <p style="color:#64748b;font-size:12px;margin-top:12px;">PawnMe Vault Admin</p>`
      });
    } catch(e) {}
    
    return { success: true, message: "Loan Created!", id: id };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}

function addPayment(id, amount, sessionToken) {
  if (!validateSession(sessionToken)) {
    return { success: false, message: 'Session expired' };
  }
  
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const current = parseFloat(data[i][19]) || 0;
      const total = parseFloat(data[i][5]);
      const newPaid = current + parseFloat(amount);
      ws.getRange(i + 1, 20).setValue(newPaid);
      if (newPaid >= total) {
        ws.getRange(i + 1, 7).setValue("Paid");
      }
      return { success: true };
    }
  }
  return { success: false };
}

function extendLoan(id, months, fee, sessionToken) {
  if (!validateSession(sessionToken)) {
    return { success: false, message: 'Session expired' };
  }
  
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const d = new Date(data[i][18]);
      const f = parseFloat(data[i][20]) || 0;
      d.setMonth(d.getMonth() + parseInt(months));
      ws.getRange(i + 1, 19).setValue(d);
      ws.getRange(i + 1, 21).setValue(f + parseFloat(fee));
      return { success: true };
    }
  }
  return { success: false };
}

function updateLoanStatus(id, st, sessionToken) {
  const role = getUserRole(sessionToken);
  if (st === 'Defaulted' && role !== 'Manager') {
    return { success: false, message: "Only managers can default loans" };
  }
  
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      ws.getRange(i + 1, 7).setValue(st);
      return { success: true };
    }
  }
  return { success: false };
}

function liquidateLoan(id, salePrice, sessionToken) {
  const role = getUserRole(sessionToken);
  if (role !== 'Manager') {
    return { success: false, message: "Only managers can liquidate assets" };
  }
  
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  const data = ws.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id && data[i][6] === 'Defaulted') {
      const principal = Number(data[i][2]);
      const paidAmount = Number(data[i][19]) || 0;
      const sale = parseFloat(salePrice);
      
      ws.getRange(i + 1, 23).setValue(sale);
      ws.getRange(i + 1, 24).setValue(new Date());
      ws.getRange(i + 1, 7).setValue("Sold");
      
      const recovered = paidAmount + sale;
      const result = recovered - principal;
      
      return { 
        success: true, 
        profit: result >= 0,
        amount: Math.abs(result)
      };
    }
  }
  return { success: false, message: "Loan not found or not defaulted" };
}

function getVaultInventory(sessionToken) {
  if (!validateSession(sessionToken)) return [];
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_LOANS);
  if (!ws) return [];
  
  const data = ws.getDataRange().getValues();
  let inventory = [];
  
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][6]);
    if (status === 'Active' || status === 'Defaulted') {
      inventory.push({
        id: data[i][0],
        itemType: data[i][7],
        brand: data[i][8],
        serial: data[i][9],
        condition: data[i][10],
        estValue: Number(data[i][11]) || 0,
        principal: Number(data[i][2]),
        status: status,
        dueDate: data[i][18],
        photoUrl: data[i][17]
      });
    }
  }
  return inventory;
}

// ==========================================
// PENDING & CRM
// ==========================================
function submitApplication(fd) {
  let ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PENDING);
  if(!ws) { debugSystem(); ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PENDING); }

  let photoUrl = "";
  if (fd && fd.imageFile && fd.imageName) {
    try {
      photoUrl = uploadFile(fd.imageFile, "Collateral_" + String(fd.name || '').replace(/\s+/g, '_') + "_" + fd.imageName);
    } catch (e) {
      photoUrl = "";
    }
  }

  ws.appendRow([new Date(), fd.name, fd.email, fd.phone, fd.nationalId, fd.address, fd.country || "", fd.itemType, fd.brand, fd.serial, fd.condition, fd.amount, fd.duration, photoUrl]);
  return { success: true };
}

function getPendingApps(sessionToken) {
  if (!validateSession(sessionToken)) return [];
  
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PENDING);
  if (!ws || ws.getLastRow() < 2) return [];
  return ws.getRange(2, 1, ws.getLastRow()-1, ws.getLastColumn()).getValues()
    .map((r, i) => ({ 
      rowIndex: i+2, name: r[1], email: r[2], phone: r[3],
      itemType: r[7], brand: r[8], requestedAmount: r[11], 
      requestedDuration: r[12],
      photoUrl: r[13]
    }));
}

function getPendingApp(rowIndex, sessionToken) {
  if (!validateSession(sessionToken)) return null;

  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PENDING);
  if (!ws || ws.getLastRow() < rowIndex) return null;

  const r = ws.getRange(rowIndex, 1, 1, ws.getLastColumn()).getValues()[0];
  return {
    rowIndex: rowIndex,
    requestedAmount: r[10],
    requestedDuration: r[11],
    photoUrl: r[12]
  };
}

function approveApplication(idx, fd, sessionToken) {
  if (!validateSession(sessionToken)) {
    return { success: false, message: 'Session expired' };
  }
  
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PENDING);
  const r = ws.getRange(idx, 1, 1, ws.getLastColumn()).getValues()[0];
  
  // Use the requested amount from the application, not from approval form
  const principal = parseFloat(r[10]); // requestedAmount column
  const rate = parseFloat(fd.rate) || 15; // default 15% if not specified
  const months = parseInt(r[11]); // requestedDuration column
  const interest = principal * (rate / 100) * months;
  const total = principal + interest;
  
  const dueDate = new Date();
  dueDate.setMonth(dueDate.getMonth() + months);
  const id = "LN-" + Date.now().toString().slice(-8);
  
  const loanData = {
    name: r[1], 
    email: r[2], 
    phone: r[3], 
    nationalId: r[4], 
    address: r[5], 
    country: r[6], 
    itemType: r[7], 
    brand: r[8], 
    serial: r[9] || "", 
    condition: r[10], 
    amount: principal, // Use the requested amount
    rate: rate, 
    duration: months, // Use the requested duration
    estValue: principal * 1.5, // Standard 1.5x multiplier
    photoUrl: r[13],
    // Additional data needed for createLoan
    imageFile: r[14], // imageFile if present
    imageName: r[15]  // imageName if present
  };
  
  createLoan(loanData, sessionToken);
  ws.deleteRow(idx);
  return { success: true };
}

function getCustomerHistory(email, sessionToken) {
  if (!validateSession(sessionToken)) {
    return {history:[],stats:{}};
  }
  
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  if(!ws) return {history:[],stats:{}};
  const data = ws.getDataRange().getValues();
  let h=[], s={totalLoans:0,paid:0,defaulted:0,activeCount:0,currentExposure:0};
  for(let i=1;i<data.length;i++){
    if(String(data[i][14]).toLowerCase()===String(email).toLowerCase()){
      let st=data[i][6], a=Number(data[i][2]);
      s.totalLoans++; 
      if(st==='Paid')s.paid++; 
      else if(st==='Defaulted' || st==='Sold')s.defaulted++; 
      else if(st==='Active'){s.activeCount++;s.currentExposure+=a;}
      h.push({date:data[i][18], item:data[i][7], status:st, amount:a});
    }
  }
  return {history:h, stats:s};
}

// ==========================================
// AUTOMATION & ALERTS
// ==========================================
function checkReminders() {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  const data = ws.getDataRange().getValues();
  const today = new Date();
  
  for(let i=1; i<data.length; i++){
    if(data[i][6] === 'Active') {
      const due = new Date(data[i][18]);
      const diff = Math.ceil((due - today) / (1000 * 60 * 60 * 24));
      
      if(diff === 3 && data[i][14]) {
        try {
          MailApp.sendEmail({
            to: data[i][14],
            subject: "⚠️ PawnMe - Loan Due in 3 Days",
            htmlBody: `<p>Dear ${data[i][1]},</p>
              <p>Your loan <b>${data[i][0]}</b> for <b>${data[i][7]}</b> is due in 3 days.</p>
              <p><b>Amount Due:</b> ${formatTry_(data[i][5])}</p>
              <p>If you’d like to settle or extend your loan, please visit us or reply to this email.</p>
              <p style="color:#64748b;font-size:12px;margin-top:12px;">PawnMe Vault Admin</p>`
          });
        } catch(e) {}
      }
    }
  }
}

function formatTry_(n) {
  try {
    return new Intl.NumberFormat('tr-TR', { style: 'currency', currency: 'TRY' }).format(Number(n) || 0);
  } catch (e) {
    const v = Number(n) || 0;
    return '₺' + v.toFixed(2);
  }
}

// ==========================================
// UTILS
// ==========================================
function getScriptUrl() { return ScriptApp.getService().getUrl(); }

function uploadFile(b64, name) {
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(b64.split(',')[1]),
      'image/jpeg',
      name
    );
    const file = DriveApp.getFolderById(FOLDER_ID)
      .createFile(blob)
      .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(e) {
    return "Upload Failed";
  }
}

function generatePDF(id, data, total, due) {
  try {
    const copy = DriveApp.getFileById(TEMPLATE_ID)
      .makeCopy('Contract_' + id, DriveApp.getFolderById(FOLDER_ID));
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();
    
    // Extract values safely
    const principal = Number(data.amount) || 0;
    const rateNum = data && data.rate !== undefined ? parseFloat(String(data.rate).replace('%', '')) : 15;
    const monthsVal = data && (data.duration !== undefined || data.months !== undefined)
      ? (data.duration !== undefined ? data.duration : data.months)
      : 3;
    const estValue = data && data.estValue !== undefined ? Number(data.estValue) : (principal * 1.5);
    
    // Calculate interest
    const interestAmount = principal * (rateNum / 100) * monthsVal;
    const totalRepayment = principal + interestAmount;
    
    // Format dates
    const today = new Date();
    const formattedToday = today.toLocaleDateString('en-NG', { 
      day: '2-digit', 
      month: '2-digit', 
      year: 'numeric' 
    });
    const formattedDue = due.toLocaleDateString('en-NG', { 
      day: '2-digit', 
      month: '2-digit', 
      year: 'numeric' 
    });
    
    // COMPLETE REPLACEMENTS - matching the DOCX template exactly
    const replacements = {
      // Loan identifiers
      "{{ID}}": id,
      "{{LOAN_ID}}": id,
      
      // Dates
      "{{DATE}}": formattedToday,
      "{{DUE_DATE}}": formattedDue,
      
      // Customer information
      "{{NAME}}": data && data.name !== undefined ? data.name : '',
      "{{EMAIL}}": data && data.email !== undefined ? data.email : '',
      "{{PHONE}}": data && data.phone !== undefined ? data.phone : '',
      "{{ADDRESS}}": data && data.address !== undefined ? data.address : '',
      "{{NATIONAL_ID}}": data && data.nationalId !== undefined ? data.nationalId : '',
      "{{Country}}": data && data.country !== undefined ? data.country : '',
      
      // Item information
      "{{ITEM_TYPE}}": data && data.itemType !== undefined ? data.itemType : '',
      "{{ITEM}}": data && data.itemType !== undefined ? data.itemType : '',
      "{{BRAND}}": data && data.brand !== undefined ? data.brand : '',
      "{{MODEL}}": data && data.brand !== undefined ? data.brand : '',
      "{{SERIAL}}": data && data.serial !== undefined ? data.serial : 'N/A',
      "{{SERIAL_NUMBER}}": data && data.serial !== undefined ? data.serial : 'N/A',
      "{{CONDITION}}": data && data.condition !== undefined ? data.condition : '',
      "{{EST_VALUE}}": estValue.toFixed(2),
      
      // Financial details
      "{{AMOUNT}}": principal.toFixed(2),
      "{{PRINCIPAL}}": principal.toFixed(2),
      "{{RATE}}": String(rateNum),
      "{{MONTHS}}": String(monthsVal),
      "{{TERM}}": String(monthsVal),
      "{{INTEREST_AMOUNT}}": interestAmount.toFixed(2),
      "{{TOTAL}}": totalRepayment.toFixed(2),
      "{{TOTAL_REPAYMENT}}": totalRepayment.toFixed(2)
    };
    
    // Replace all placeholders
    Object.keys(replacements).forEach(placeholder => {
      const value = replacements[placeholder];
      // Escape special regex characters in placeholder
      const escapedPlaceholder = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      body.replaceText(escapedPlaceholder, value === undefined || value === null ? '' : String(value));
    });
    
    doc.saveAndClose();
    
    // Convert to PDF
    const pdf = DriveApp.getFolderById(FOLDER_ID)
      .createFile(copy.getAs('application/pdf'))
      .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Delete the temporary Doc copy
    copy.setTrashed(true);
    
    return pdf.getUrl();
  } catch(e) {
    Logger.log('PDF Generation Error: ' + e.toString());
    return "PDF Generation Failed: " + e.message;
  }
}