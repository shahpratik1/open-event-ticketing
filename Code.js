/**
 * ==============================================================================
 * EVENT MANAGEMENT & TICKETING SYSTEM - AUTOMATION ENGINE
 * ==============================================================================
 * Production-grade automation engine featuring:
 * - Dynamic Attendee & Capacity Quota Tracking
 * - Automated Registration & Payment Emails with token interpolation
 * - QR Code Ticket Generation & Email Delivery for Paid Guests
 * - Strikethrough Guard: Automatically skips canceled / voided spreadsheet rows
 * - In-sheet Payment Status Dropdowns & Dynamic Highlighting
 * - Automated Daily Reminders & Form Submission Triggers
 */

/**
 * Fallback stub: Intercepts legacy trigger calls targeting "myFunction"
 */
function myFunction() {
  console.warn('A legacy trigger invoked "myFunction". No operation performed.');
}

/**
 * Creates custom navigation menu in Google Sheets on document open.
 */
function onOpen(e) {
  try {
    const config = getAppConfig();
    const ui = SpreadsheetApp.getUi();
    ui.createMenu(`🎪 ${config.EVENT_NAME}`)
      .addItem('1. Update Attendee Counts', 'updateAttendeeCounts')
      .addItem('2. Send Registration / Payment Emails', 'sendRegistrationEmails')
      .addItem('3. Send Confirmation Emails (Paid + QR Code)', 'sendConfirmationEmails')
      .addItem('4. Send Payment Reminders', 'sendPaymentReminders')
      .addItem('5. Send Feedback Emails', 'sendFeedbackEmails')
      .addSeparator()
      .addItem('⚙️ Initialize Script Properties', 'setupProperties')
      .addItem('📋 Create/Refresh Payment Dropdowns (Col M)', 'applyPaymentDropdowns')
      .addItem('⏰ Install Automated Daily & Submit Triggers', 'installAutomatedTriggers')
      .addItem('🧹 Clean & Remove All Triggers', 'removeAllProjectTriggers')
      .addToUi();
  } catch (err) {
    console.warn('UI not accessible in current execution context:', err.message);
  }
}

/**
 * Real-time listener: Colors cell automatically when user selects Paid/NotPaid
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const config = getAppConfig();
  const sheet = e.range.getSheet();
  if (sheet.getName() !== config.SHEETS.RESPONSES) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (col === config.COLS.PAYMENT_STATUS && row > 1) {
    const val = String(e.value || e.range.getValue()).trim().toLowerCase();
    if (val === 'paid') {
      e.range.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    } else if (val === 'notpaid') {
      e.range.setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('normal');
    }
  }
}

/**
 * Applies "Paid" / "NotPaid" dropdown validation across Column M
 */
function applyPaymentDropdowns() {
  const config = getAppConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName(config.SHEETS.RESPONSES);
  if (!sheet1) return;

  const maxRows = Math.max(sheet1.getMaxRows() - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['NotPaid', 'Paid'], true)
    .setAllowInvalid(false)
    .build();

  const paymentRange = sheet1.getRange(2, config.COLS.PAYMENT_STATUS, maxRows, 1);
  paymentRange.setDataValidation(rule);

  safeToast('Dropdowns added to Column M (Paid / NotPaid).', 'Dropdowns Ready');
}

/**
 * Trigger target for real-time form submission processing
 */
function onFormSubmitTrigger(e) {
  applyPaymentDropdowns();
  sendRegistrationEmails();
  updateAttendeeCounts();
}

/**
 * Recalculates total reserved and confirmed attendees (ignores strikethrough rows).
 */
function updateAttendeeCounts() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;

  try {
    const config = getAppConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = ss.getSheetByName(config.SHEETS.RESPONSES);
    const sheet2 = ss.getSheetByName(config.SHEETS.CONFIG_TEMPLATES);
    if (!sheet1 || !sheet2) return;

    const dataRange = sheet1.getDataRange();
    const data = dataRange.getValues();
    const fontLines = dataRange.getFontLines();
    let reservedCount = 0;
    let confirmedCount = 0;

    for (let i = 1; i < data.length; i++) {
      if (isRowStrikethrough(fontLines[i], config)) continue;

      const row = data[i];
      const adults = Number(row[config.COLS.ADULTS - 1]) || 0;
      const children = Number(row[config.COLS.CHILDREN - 1]) || 0;
      const infants = Number(row[config.COLS.INFANTS - 1]) || 0;
      const partySize = adults + children + infants;
      const paymentStatus = String(row[config.COLS.PAYMENT_STATUS - 1] || '').trim().toLowerCase();

      reservedCount += partySize;
      if (paymentStatus === 'paid') {
        confirmedCount += partySize;
      }
    }

    sheet2.getRange(config.SHEET2.RESERVED_COUNT.row, config.SHEET2.RESERVED_COUNT.col).setValue(reservedCount);
    sheet2.getRange(config.SHEET2.CONFIRMED_COUNT.row, config.SHEET2.CONFIRMED_COUNT.col).setValue(confirmedCount);

    safeToast(`Reserved: ${reservedCount} | Confirmed: ${confirmedCount}`, 'Attendee Count Updated');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Sends registration/payment request emails (skips strikethrough rows).
 */
function sendRegistrationEmails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const config = getAppConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = ss.getSheetByName(config.SHEETS.RESPONSES);
    const sheet2 = ss.getSheetByName(config.SHEETS.CONFIG_TEMPLATES);

    const dataRange = sheet1.getDataRange();
    const data = dataRange.getValues();
    const fontLines = dataRange.getFontLines();
    if (data.length <= 1) return;

    const adultCost = Number(sheet2.getRange(config.SHEET2.COST_ADULT.row, config.SHEET2.COST_ADULT.col).getValue()) || 0;
    const childCost = Number(sheet2.getRange(config.SHEET2.COST_CHILD.row, config.SHEET2.COST_CHILD.col).getValue()) || 0;
    const maxCapacity = Number(sheet2.getRange(config.SHEET2.MAX_ATTENDEES.row, config.SHEET2.MAX_ATTENDEES.col).getValue()) || config.DEFAULT_MAX_ATTENDEES;

    const subjectTemplate = sheet2.getRange(config.SHEET2.TEMPLATES.REGISTRATION, 1).getValue();
    const bodyTemplate = sheet2.getRange(config.SHEET2.TEMPLATES.REGISTRATION, 2).getValue();

    const dropdownRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['NotPaid', 'Paid'], true)
      .setAllowInvalid(false)
      .build();

    let currentReserved = Number(sheet2.getRange(config.SHEET2.RESERVED_COUNT.row, config.SHEET2.RESERVED_COUNT.col).getValue()) || 0;

    for (let i = 1; i < data.length; i++) {
      if (isRowStrikethrough(fontLines[i], config)) continue;

      const row = data[i];
      const rowIndex = i + 1;
      const email = String(row[config.COLS.EMAIL - 1] || '').trim();
      const regEmailStatus = String(row[config.COLS.REG_EMAIL - 1] || '').trim();

      if (!isValidEmail(email) || regEmailStatus !== '') continue;

      const adults = Number(row[config.COLS.ADULTS - 1]) || 0;
      const children = Number(row[config.COLS.CHILDREN - 1]) || 0;
      const infants = Number(row[config.COLS.INFANTS - 1]) || 0;
      const partySize = adults + children + infants;

      if (currentReserved + partySize > maxCapacity) {
        notifyCapacityLimitReached(maxCapacity);
        break;
      }

      const name = String(row[config.COLS.NAME - 1] || 'Guest').trim();
      const cleanFirstName = (name.split(' ')[0] || 'Guest').replace(/[^a-zA-Z0-9]/g, '');
      const ticketCode = `${config.TICKET_CODE_PREFIX}${rowIndex}-${cleanFirstName}-${adults}-${children}-${infants}`;
      const totalCost = (adults * adultCost) + (children * childCost);

      const tokens = {
        '<name>': name,
        '<TicketCode>': ticketCode,
        '<AdultTicketCost>': adultCost,
        '<ChildTicketCost>': childCost,
        '<TotalCost>': totalCost,
        '<EventName>': config.EVENT_NAME
      };

      const finalSubject = replaceTemplateTokens(subjectTemplate, tokens);
      const finalBody = replaceTemplateTokens(bodyTemplate, tokens);

      try {
        MailApp.sendEmail({
          to: email,
          subject: finalSubject,
          body: finalBody,
          name: config.SENDER_NAME
        });

        sheet1.getRange(rowIndex, config.COLS.TICKET_CODE).setValue(ticketCode);

        const paymentCell = sheet1.getRange(rowIndex, config.COLS.PAYMENT_STATUS);
        paymentCell.setDataValidation(dropdownRule);
        paymentCell.setValue('NotPaid').setBackground('#f8d7da').setFontColor('#721c24');

        sheet1.getRange(rowIndex, config.COLS.REG_EMAIL).setValue('Sent').setFontWeight('bold').setBackground('#d4edda');
        sheet1.getRange(rowIndex, config.COLS.CONF_EMAIL).setValue('Pending').setBackground('#fff3cd');

        currentReserved += partySize;
      } catch (err) {
        console.error(`Error sending registration email at row ${rowIndex}: ${err.message}`);
        sheet1.getRange(rowIndex, config.COLS.REG_EMAIL).setValue('Failed').setBackground('#f8d7da');
      }
    }

    sheet2.getRange(config.SHEET2.RESERVED_COUNT.row, config.SHEET2.RESERVED_COUNT.col).setValue(currentReserved);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Sends confirmation emails with QR code to verified "Paid" rows (skips strikethrough rows).
 */
function sendConfirmationEmails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const config = getAppConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = ss.getSheetByName(config.SHEETS.RESPONSES);
    const sheet2 = ss.getSheetByName(config.SHEETS.CONFIG_TEMPLATES);

    const dataRange = sheet1.getDataRange();
    const data = dataRange.getValues();
    const fontLines = dataRange.getFontLines();
    if (data.length <= 1) return;

    const subjectTemplate = sheet2.getRange(config.SHEET2.TEMPLATES.CONFIRMATION, 1).getValue();
    const bodyTemplate = sheet2.getRange(config.SHEET2.TEMPLATES.CONFIRMATION, 2).getValue();

    for (let i = 1; i < data.length; i++) {
      if (isRowStrikethrough(fontLines[i], config)) continue;

      const row = data[i];
      const rowIndex = i + 1;
      const email = String(row[config.COLS.EMAIL - 1] || '').trim();
      const paymentStatus = String(row[config.COLS.PAYMENT_STATUS - 1] || '').trim().toLowerCase();
      const confStatus = String(row[config.COLS.CONF_EMAIL - 1] || '').trim();
      const ticketCode = String(row[config.COLS.TICKET_CODE - 1] || '').trim();
      const name = String(row[config.COLS.NAME - 1] || 'Guest').trim();

      if (isValidEmail(email) && paymentStatus === 'paid' && confStatus !== 'Sent') {
        const tokens = {
          '<name>': name,
          '<TicketCode>': ticketCode,
          '<EventName>': config.EVENT_NAME
        };

        const finalSubject = replaceTemplateTokens(subjectTemplate, tokens);
        const finalBody = replaceTemplateTokens(bodyTemplate, tokens);
        const qrBlob = fetchQrBlob(ticketCode, config);

        const formattedText = finalBody.replace(/\n/g, '<br>');
        const htmlBodyContent = `
          <div style="font-family: Arial, Helvetica, sans-serif; font-size: 15px; color: #222; line-height: 1.6; max-width: 620px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;">
            ${formattedText}
            <div style="margin: 25px 0; padding: 20px; border: 2px dashed #0d6efd; border-radius: 12px; background-color: #f8faff; text-align: center;">
              <h3 style="margin: 0 0 10px 0; color: #0d6efd; font-size: 18px; text-transform: uppercase; letter-spacing: 0.5px;">Official Entry Pass</h3>
              <img src="cid:ticketQrCode" alt="Ticket QR Code" width="180" height="180" style="display: block; margin: 10px auto; border-radius: 6px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);" />
              <div style="font-family: monospace; font-size: 18px; font-weight: bold; color: #111; margin-top: 8px; letter-spacing: 1px;">${ticketCode}</div>
              <p style="margin: 6px 0 0 0; color: #666; font-size: 13px;">Please present this QR code at the check-in desk for entry.</p>
            </div>
            <p style="font-size: 12px; color: #888; text-align: center; margin-top: 20px;">${config.ORGANIZATION_NAME} &bull; Entry Ticket</p>
          </div>
        `;

        try {
          const emailOptions = {
            to: email,
            subject: finalSubject,
            body: finalBody,
            htmlBody: htmlBodyContent,
            name: config.SENDER_NAME
          };

          if (qrBlob) {
            emailOptions.inlineImages = { ticketQrCode: qrBlob };
          }

          MailApp.sendEmail(emailOptions);

          sheet1.getRange(rowIndex, config.COLS.CONF_EMAIL).setValue('Sent').setFontWeight('bold').setBackground('#d4edda');
          sheet1.getRange(rowIndex, config.COLS.PAYMENT_STATUS).setFontWeight('bold').setBackground('#d4edda').setFontColor('#155724');
        } catch (err) {
          console.error(`Error sending confirmation email at row ${rowIndex}: ${err.message}`);
        }
      }
    }
    updateAttendeeCounts();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Sends payment reminders to unpaid registrants (skips strikethrough rows).
 */
function sendPaymentReminders() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const config = getAppConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = ss.getSheetByName(config.SHEETS.RESPONSES);
    const sheet2 = ss.getSheetByName(config.SHEETS.CONFIG_TEMPLATES);

    const dataRange = sheet1.getDataRange();
    const data = dataRange.getValues();
    const fontLines = dataRange.getFontLines();
    if (data.length <= 1) return;

    const adultCost = Number(sheet2.getRange(config.SHEET2.COST_ADULT.row, config.SHEET2.COST_ADULT.col).getValue()) || 0;
    const childCost = Number(sheet2.getRange(config.SHEET2.COST_CHILD.row, config.SHEET2.COST_CHILD.col).getValue()) || 0;
    const subjectTemplate = sheet2.getRange(config.SHEET2.TEMPLATES.REMINDER, 1).getValue();
    const bodyTemplate = sheet2.getRange(config.SHEET2.TEMPLATES.REMINDER, 2).getValue();

    const now = new Date().getTime();
    const allowedTimeMs = config.PAYMENT_DUE_DAYS * 24 * 60 * 60 * 1000;

    for (let i = 1; i < data.length; i++) {
      if (isRowStrikethrough(fontLines[i], config)) continue;

      const row = data[i];
      const rowIndex = i + 1;
      const regTimestamp = row[config.COLS.TIMESTAMP - 1];
      const paymentStatus = String(row[config.COLS.PAYMENT_STATUS - 1] || '').trim().toLowerCase();
      const reminderStatus = String(row[config.COLS.REMINDER_EMAIL - 1] || '').trim();
      const email = String(row[config.COLS.EMAIL - 1] || '').trim();

      if (!regTimestamp || !isValidEmail(email)) continue;

      const regTime = new Date(regTimestamp).getTime();

      if (now - regTime > allowedTimeMs && paymentStatus === 'notpaid' && reminderStatus === '') {
        const name = String(row[config.COLS.NAME - 1] || 'Guest').trim();
        const ticketCode = String(row[config.COLS.TICKET_CODE - 1] || '').trim();
        const adults = Number(row[config.COLS.ADULTS - 1]) || 0;
        const children = Number(row[config.COLS.CHILDREN - 1]) || 0;
        const totalCost = (adults * adultCost) + (children * childCost);

        const tokens = {
          '<name>': name,
          '<TicketCode>': ticketCode,
          '<AdultTicketCost>': adultCost,
          '<ChildTicketCost>': childCost,
          '<TotalCost>': totalCost,
          '<EventName>': config.EVENT_NAME
        };

        const finalSubject = replaceTemplateTokens(subjectTemplate, tokens);
        const finalBody = replaceTemplateTokens(bodyTemplate, tokens);

        try {
          MailApp.sendEmail({
            to: email,
            subject: finalSubject,
            body: finalBody,
            name: config.SENDER_NAME
          });

          sheet1.getRange(rowIndex, config.COLS.REMINDER_EMAIL).setValue('Sent').setFontWeight('bold').setBackground('#fff3cd');
        } catch (err) {
          console.error(`Error sending reminder to row ${rowIndex}: ${err.message}`);
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Sends event feedback requests to all confirmed attendees (skips strikethrough rows).
 */
function sendFeedbackEmails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const config = getAppConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = ss.getSheetByName(config.SHEETS.RESPONSES);
    const sheet2 = ss.getSheetByName(config.SHEETS.CONFIG_TEMPLATES);

    const dataRange = sheet1.getDataRange();
    const data = dataRange.getValues();
    const fontLines = dataRange.getFontLines();
    if (data.length <= 1) return;

    const subjectTemplate = sheet2.getRange(config.SHEET2.TEMPLATES.FEEDBACK, 1).getValue();
    const bodyTemplate = sheet2.getRange(config.SHEET2.TEMPLATES.FEEDBACK, 2).getValue();

    for (let i = 1; i < data.length; i++) {
      if (isRowStrikethrough(fontLines[i], config)) continue;

      const row = data[i];
      const paymentStatus = String(row[config.COLS.PAYMENT_STATUS - 1] || '').trim().toLowerCase();
      const confStatus = String(row[config.COLS.CONF_EMAIL - 1] || '').trim();
      const email = String(row[config.COLS.EMAIL - 1] || '').trim();
      const name = String(row[config.COLS.NAME - 1] || 'Guest').trim();

      if (paymentStatus === 'paid' && confStatus === 'Sent' && isValidEmail(email)) {
        const tokens = { 
          '<name>': name,
          '<EventName>': config.EVENT_NAME
        };
        const finalSubject = replaceTemplateTokens(subjectTemplate, tokens);
        const finalBody = replaceTemplateTokens(bodyTemplate, tokens);

        try {
          MailApp.sendEmail({
            to: email,
            subject: finalSubject,
            body: finalBody,
            name: config.SENDER_NAME
          });
        } catch (err) {
          console.error(`Error sending feedback email to ${email}: ${err.message}`);
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------------------------
// TRIGGER SETUP & CLEANUP UTILITIES
// ------------------------------------------------------------------------------

function removeAllProjectTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  console.log(`Purged ${triggers.length} triggers.`);
  safeToast('All project triggers have been completely purged.', 'Triggers Cleaned');
}

function installAutomatedTriggers() {
  removeAllProjectTriggers();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.newTrigger('onFormSubmitTrigger')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('sendPaymentReminders')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();

  applyPaymentDropdowns();
  safeToast('Automated triggers and Column M dropdowns installed.', 'Triggers Active');
}

// ------------------------------------------------------------------------------
// HELPER FUNCTIONS
// ------------------------------------------------------------------------------

/**
 * Detects if a row has been marked as strikethrough/canceled.
 * Checks Timestamp (Col A), Name (Col B), and Email (Col C).
 */
function isRowStrikethrough(fontLineRow, config) {
  if (!fontLineRow || !Array.isArray(fontLineRow)) return false;
  const cfg = config || getAppConfig();
  return fontLineRow[cfg.COLS.TIMESTAMP - 1] === 'line-through' ||
         fontLineRow[cfg.COLS.NAME - 1] === 'line-through' ||
         fontLineRow[cfg.COLS.EMAIL - 1] === 'line-through';
}

function fetchQrBlob(textData, config) {
  const cfg = config || getAppConfig();
  try {
    const encoded = encodeURIComponent(textData);
    const url = `${cfg.QR_API_BASE_URL}?text=${encoded}&size=220&margin=1`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      return response.getBlob().setName('ticket_qrcode.png');
    }
  } catch (e) {
    console.error(`QR Code fetch failed for "${textData}": ${e.message}`);
  }
  return null;
}

function replaceTemplateTokens(templateStr, replacements) {
  let output = String(templateStr || '');
  for (const [key, val] of Object.entries(replacements)) {
    const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    output = output.replace(new RegExp(escapedKey, 'g'), val !== undefined ? val : '');
  }
  return output;
}

function isValidEmail(email) {
  return typeof email === 'string' && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
}

function safeToast(message, title) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      ss.toast(message, title, 5);
    }
  } catch (e) {
    console.log(`[Toast] ${title}: ${message}`);
  }
}

function notifyCapacityLimitReached(maxCapacity) {
  const config = getAppConfig();
  config.ORGANIZER_EMAILS.forEach(email => {
    MailApp.sendEmail(
      email,
      `Maximum Capacity Reached - ${config.EVENT_NAME}`,
      `Attention: The reservation ceiling of ${maxCapacity} attendees has been reached. Please turn off accepting responses in Google Forms.`
    );
  });
}
