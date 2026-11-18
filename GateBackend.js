/**
 * ==============================================================================
 * LIVE GATE SCANNER BACKEND & REST API
 * ==============================================================================
 * Backend service powering both the HTML5 mobile gate scanner interface and
 * external web apps (e.g. React check-in dashboard):
 * - Sanitized Guest Directory (Excludes private emails & phone numbers from gate staff)
 * - Real-time Ticket Verification with Strikethrough / Cancellation Guard
 * - Flexible Headcount Stepper Entry (Adults, Children, Infants)
 * - Complete Audit Trail logging in dedicated CheckIns sheet
 * - REST API endpoints for GET (directory/stats) and POST (check-in)
 */

/**
 * Serves the mobile gate scanner application or returns JSON API data.
 */
function doGet(e) {
  const config = getGateConfig();

  // JSON REST API: Fetch Attendee Directory
  if (e && e.parameter && (e.parameter.format === 'json' || e.parameter.action === 'directory')) {
    const directoryResult = apiGetAttendeeDirectory();
    return ContentService.createTextOutput(JSON.stringify(directoryResult.data || []))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // JSON REST API: Fetch Live Gate Statistics
  if (e && e.parameter && e.parameter.action === 'stats') {
    const statsResult = apiGetGateStats();
    return ContentService.createTextOutput(JSON.stringify(statsResult))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Default: Render HTML5 Mobile Scanner Web App
  return HtmlService.createHtmlOutputFromFile('ScannerApp')
    .setTitle(`Gate Entry - ${config.EVENT_NAME}`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handles incoming POST requests from external clients (e.g. React check-in app).
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Empty POST body.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const payload = JSON.parse(e.postData.contents);

    // 1. External React App toggle check-in format: { id: rowId, checkedIn: boolean }
    if (payload.id !== undefined && payload.checkedIn !== undefined) {
      const result = updateCheckInStatusById(payload.id, payload.checkedIn);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Direct ticket check-in format: { ticketCode, adults, children, infants, note }
    const result = apiProcessCheckIn(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Server error: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Updates check-in status by row index (used by external dashboard).
 */
function updateCheckInStatusById(rowId, isCheckedIn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { success: false, message: 'Database lock timeout. Please retry.' };
  }

  try {
    const config = getGateConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = ss.getSheetByName(config.SHEET_RESPONSES);
    if (!sheet1) return { success: false, message: 'Responses sheet not found.' };

    const rowIndex = Number(rowId);
    if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > sheet1.getLastRow()) {
      return { success: false, message: `Invalid row index: ${rowId}` };
    }

    const totalAdults = Number(sheet1.getRange(rowIndex, config.COLS.ADULTS).getValue()) || 0;
    const totalChildren = Number(sheet1.getRange(rowIndex, config.COLS.CHILDREN).getValue()) || 0;
    const totalInfants = Number(sheet1.getRange(rowIndex, config.COLS.INFANTS).getValue()) || 0;

    const newAdults = isCheckedIn ? totalAdults : 0;
    const newChildren = isCheckedIn ? totalChildren : 0;
    const newInfants = isCheckedIn ? totalInfants : 0;
    const statusText = isCheckedIn ? 'Fully Checked In' : 'Not Checked In';
    const statusBg = isCheckedIn ? '#d4edda' : '#f8f9fa';

    sheet1.getRange(rowIndex, config.COLS.ADULTS_CHECKED).setValue(newAdults);
    sheet1.getRange(rowIndex, config.COLS.CHILDREN_CHECKED).setValue(newChildren);
    sheet1.getRange(rowIndex, config.COLS.INFANTS_CHECKED).setValue(newInfants);
    sheet1.getRange(rowIndex, config.COLS.GATE_STATUS).setValue(statusText).setBackground(statusBg);

    return {
      success: true,
      message: `Row ${rowIndex} updated to: ${statusText}`,
      checkedIn: isCheckedIn
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Checks if a spreadsheet row has been marked with strikethrough (canceled).
 */
function isRowStrikethrough(fontLineRow) {
  if (!fontLineRow || !Array.isArray(fontLineRow)) return false;
  const config = getGateConfig();
  return fontLineRow[config.COLS.TIMESTAMP - 1] === 'line-through' ||
         fontLineRow[config.COLS.NAME - 1] === 'line-through' ||
         fontLineRow[config.COLS.EMAIL - 1] === 'line-through' ||
         fontLineRow[config.COLS.TICKET_CODE - 1] === 'line-through';
}

/**
 * Fetches attendee directory with sensitive personal info (email & phone) omitted for privacy.
 */
function apiGetAttendeeDirectory() {
  const config = getGateConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName(config.SHEET_RESPONSES);
  if (!sheet1) return { success: false, data: [] };

  const dataRange = sheet1.getDataRange();
  const data = dataRange.getValues();
  const fontLines = dataRange.getFontLines();

  const directory = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ticketCode = String(row[config.COLS.TICKET_CODE - 1] || '').trim();
    const name = String(row[config.COLS.NAME - 1] || '').trim();

    if (!ticketCode && !name) continue;

    const isCanceled = isRowStrikethrough(fontLines[i]);
    const totalAdults = Number(row[config.COLS.ADULTS - 1]) || 0;
    const totalChildren = Number(row[config.COLS.CHILDREN - 1]) || 0;
    const totalInfants = Number(row[config.COLS.INFANTS - 1]) || 0;

    const checkedAdults = Number(row[config.COLS.ADULTS_CHECKED - 1]) || 0;
    const checkedChildren = Number(row[config.COLS.CHILDREN_CHECKED - 1]) || 0;
    const checkedInfants = Number(row[config.COLS.INFANTS_CHECKED - 1]) || 0;

    const paymentStatus = String(row[config.COLS.PAYMENT_STATUS - 1] || 'NotPaid').trim();
    let gateStatus = String(row[config.COLS.GATE_STATUS - 1] || 'Not Checked In').trim();
    if (isCanceled) gateStatus = 'Canceled';

    const numAttendees = totalAdults + totalChildren + totalInfants;
    const totalChecked = checkedAdults + checkedChildren + checkedInfants;
    const isCheckedIn = (numAttendees > 0 && totalChecked >= numAttendees);

    const nameParts = name.split(' ');
    const firstName = nameParts[0] || 'Guest';
    const lastName = nameParts.slice(1).join(' ') || '';

    directory.push({
      id: i + 1,
      ticketCode: ticketCode,
      name: name,
      firstName: firstName,
      lastName: lastName,
      numAttendees: numAttendees,
      paymentStatus: paymentStatus,
      isPaid: paymentStatus.toLowerCase() === 'paid',
      checkedIn: isCheckedIn,
      isCanceled: isCanceled,
      gateStatus: gateStatus,
      totalAdults: totalAdults,
      totalChildren: totalChildren,
      totalInfants: totalInfants,
      checkedAdults: checkedAdults,
      checkedChildren: checkedChildren,
      checkedInfants: checkedInfants,
      remTotal: numAttendees - totalChecked
    });
  }

  return { success: true, data: directory };
}

/**
 * Verifies a scanned or typed ticket code against the responses sheet.
 */
function apiVerifyTicket(rawCode) {
  if (!rawCode) return { success: false, message: 'Empty ticket code provided.' };
  
  const ticketCode = String(rawCode).trim();
  const config = getGateConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName(config.SHEET_RESPONSES);
  
  if (!sheet1) return { success: false, message: 'Responses sheet not found in spreadsheet.' };

  const dataRange = sheet1.getDataRange();
  const data = dataRange.getValues();
  const fontLines = dataRange.getFontLines();

  let foundRow = -1;
  let attendee = null;
  let isCanceled = false;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const codeInSheet = String(row[config.COLS.TICKET_CODE - 1] || '').trim();
    
    if (codeInSheet.toLowerCase() === ticketCode.toLowerCase()) {
      foundRow = i + 1;
      isCanceled = isRowStrikethrough(fontLines[i]);
      
      const totalAdults = Number(row[config.COLS.ADULTS - 1]) || 0;
      const totalChildren = Number(row[config.COLS.CHILDREN - 1]) || 0;
      const totalInfants = Number(row[config.COLS.INFANTS - 1]) || 0;

      const checkedAdults = Number(row[config.COLS.ADULTS_CHECKED - 1]) || 0;
      const checkedChildren = Number(row[config.COLS.CHILDREN_CHECKED - 1]) || 0;
      const checkedInfants = Number(row[config.COLS.INFANTS_CHECKED - 1]) || 0;

      const paymentStatus = String(row[config.COLS.PAYMENT_STATUS - 1] || 'NotPaid').trim();
      
      attendee = {
        rowIndex: foundRow,
        ticketCode: codeInSheet,
        name: String(row[config.COLS.NAME - 1] || 'Guest').trim(),
        paymentStatus: paymentStatus,
        isPaid: paymentStatus.toLowerCase() === 'paid',
        isCanceled: isCanceled,
        
        totalAdults: totalAdults,
        totalChildren: totalChildren,
        totalInfants: totalInfants,
        
        checkedAdults: checkedAdults,
        checkedChildren: checkedChildren,
        checkedInfants: checkedInfants,
        
        remAdults: Math.max(0, totalAdults - checkedAdults),
        remChildren: Math.max(0, totalChildren - checkedChildren),
        remInfants: Math.max(0, totalInfants - checkedInfants),
        
        gateStatus: isCanceled ? 'Canceled / Void' : String(row[config.COLS.GATE_STATUS - 1] || 'Not Checked In')
      };
      break;
    }
  }

  if (!attendee) {
    return { success: false, message: `Ticket "${ticketCode}" was not found.` };
  }

  const isFullyCheckedIn = (!isCanceled && attendee.remAdults === 0 && attendee.remChildren === 0 && attendee.remInfants === 0);

  return {
    success: true,
    isCanceled: isCanceled,
    isFullyCheckedIn: isFullyCheckedIn,
    data: attendee
  };
}

/**
 * Records attendee check-in and logs the entry to the CheckIns audit sheet.
 */
function apiProcessCheckIn(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { success: false, message: 'Server busy. Please tap confirm again in a moment.' };
  }

  try {
    const config = getGateConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = ss.getSheetByName(config.SHEET_RESPONSES);
    let checkInSheet = ss.getSheetByName(config.SHEET_CHECKINS);

    if (!checkInSheet) {
      checkInSheet = ss.insertSheet(config.SHEET_CHECKINS);
      checkInSheet.appendRow(['Timestamp', 'TicketCode', 'Name', 'AdultsCheckedIn', 'ChildrenCheckedIn', 'InfantsCheckedIn', 'GateNote']);
      checkInSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f0f0f0');
    }

    const verification = apiVerifyTicket(payload.ticketCode);
    if (!verification.success) return verification;
    if (verification.isCanceled) {
      return { success: false, message: 'ENTRY REJECTED: This reservation is marked as CANCELED.' };
    }

    const current = verification.data;
    const rowIndex = current.rowIndex;

    const inAdults = Math.min(Number(payload.adults) || 0, current.remAdults);
    const inChildren = Math.min(Number(payload.children) || 0, current.remChildren);
    const inInfants = Math.min(Number(payload.infants) || 0, current.remInfants);

    if (inAdults === 0 && inChildren === 0 && inInfants === 0) {
      return { success: false, message: 'No remaining attendees selected for check-in.' };
    }

    const updatedAdults = current.checkedAdults + inAdults;
    const updatedChildren = current.checkedChildren + inChildren;
    const updatedInfants = current.checkedInfants + inInfants;

    const totalAllowed = current.totalAdults + current.totalChildren + current.totalInfants;
    const totalNowChecked = updatedAdults + updatedChildren + updatedInfants;

    let newStatus = 'Partially Checked In';
    let statusBg = '#fff3cd';

    if (totalNowChecked >= totalAllowed) {
      newStatus = 'Fully Checked In';
      statusBg = '#d4edda';
    }

    sheet1.getRange(rowIndex, config.COLS.ADULTS_CHECKED).setValue(updatedAdults);
    sheet1.getRange(rowIndex, config.COLS.CHILDREN_CHECKED).setValue(updatedChildren);
    sheet1.getRange(rowIndex, config.COLS.INFANTS_CHECKED).setValue(updatedInfants);
    sheet1.getRange(rowIndex, config.COLS.GATE_STATUS).setValue(newStatus).setBackground(statusBg);

    const timestamp = new Date();
    checkInSheet.appendRow([
      timestamp,
      current.ticketCode,
      current.name,
      inAdults,
      inChildren,
      inInfants,
      payload.note || 'Gate Entry'
    ]);

    return {
      success: true,
      message: `Checked in: ${inAdults} Adults, ${inChildren} Children, ${inInfants} Infants`,
      remaining: {
        adults: current.totalAdults - updatedAdults,
        children: current.totalChildren - updatedChildren,
        infants: current.totalInfants - updatedInfants
      },
      gateStatus: newStatus
    };
  } catch (err) {
    return { success: false, message: 'Error processing check-in: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Calculates current event statistics (Total Paid vs Total Checked In).
 */
function apiGetGateStats() {
  const config = getGateConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName(config.SHEET_RESPONSES);
  if (!sheet1) return { totalPaid: 0, totalChecked: 0 };

  const dataRange = sheet1.getDataRange();
  const data = dataRange.getValues();
  const fontLines = dataRange.getFontLines();

  let totalPaidAttendees = 0;
  let totalCheckedInAttendees = 0;

  for (let i = 1; i < data.length; i++) {
    if (isRowStrikethrough(fontLines[i])) continue;

    const row = data[i];
    const isPaid = String(row[config.COLS.PAYMENT_STATUS - 1] || '').trim().toLowerCase() === 'paid';
    
    if (isPaid) {
      const adults = Number(row[config.COLS.ADULTS - 1]) || 0;
      const children = Number(row[config.COLS.CHILDREN - 1]) || 0;
      const infants = Number(row[config.COLS.INFANTS - 1]) || 0;
      totalPaidAttendees += (adults + children + infants);

      const chkAdults = Number(row[config.COLS.ADULTS_CHECKED - 1]) || 0;
      const chkKids = Number(row[config.COLS.CHILDREN_CHECKED - 1]) || 0;
      const chkInfants = Number(row[config.COLS.INFANTS_CHECKED - 1]) || 0;
      totalCheckedInAttendees += (chkAdults + chkKids + chkInfants);
    }
  }

  return {
    totalPaid: totalPaidAttendees,
    totalChecked: totalCheckedInAttendees
  };
}
