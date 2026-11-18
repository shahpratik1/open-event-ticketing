/**
 * ==============================================================================
 * EVENT MANAGEMENT & CHECK-IN SYSTEM - CONFIGURATION & SCRIPT PROPERTIES
 * ==============================================================================
 * This module manages environment-specific variables using Google Apps Script's
 * PropertiesService, ensuring no sensitive emails, organization names, or IDs
 * are hardcoded into the source files.
 */

/**
 * Initializes default Script Properties in Google Apps Script.
 * Run this function once from the Apps Script editor (Select 'setupProperties' -> Run)
 * or via the custom Google Sheets menu: '🎪 Event Management' -> '⚙️ Initialize Script Properties'.
 *
 * @param {Object} [customProps] Optional key-value object to override defaults.
 */
function setupProperties(customProps = {}) {
  const scriptProperties = PropertiesService.getScriptProperties();

  const defaultProperties = {
    ORGANIZER_EMAILS: 'organizer@example.com',
    EVENT_SENDER_NAME: 'Event Management Team',
    EVENT_NAME: 'Community Event 2026',
    ORGANIZATION_NAME: 'Event Organizers Association',
    DEFAULT_MAX_ATTENDEES: '250',
    PAYMENT_DUE_DAYS: '3',
    SHEET_NAME_RESPONSES: 'Sheet1',
    SHEET_NAME_TEMPLATES: 'Sheet2',
    SHEET_NAME_CHECKINS: 'CheckIns',
    TICKET_CODE_PREFIX: 'T',
    QR_API_BASE_URL: 'https://quickchart.io/qr'
  };

  const propertiesToApply = Object.assign({}, defaultProperties, customProps);
  const currentProperties = scriptProperties.getProperties();
  const updatedProperties = {};

  for (const [key, value] of Object.entries(propertiesToApply)) {
    // Only set if not already defined, or if explicitly provided in customProps
    if (!currentProperties[key] || customProps[key] !== undefined) {
      updatedProperties[key] = String(value);
    }
  }

  if (Object.keys(updatedProperties).length > 0) {
    scriptProperties.setProperties(updatedProperties);
    console.log('Script properties successfully initialized/updated:', updatedProperties);
    safeToast('Script Properties initialized successfully.', 'Properties Ready');
  } else {
    console.log('Script properties are already configured.');
    safeToast('Script Properties already exist. No overwrite needed.', 'Properties Unchanged');
  }

  return scriptProperties.getProperties();
}

/**
 * Helper to retrieve a script property with an optional fallback.
 *
 * @param {string} key - Property key name.
 * @param {string} [fallback=''] - Default fallback value if property is unset.
 * @returns {string} The property value or fallback.
 */
function getScriptProp(key, fallback = '') {
  try {
    const val = PropertiesService.getScriptProperties().getProperty(key);
    return (val !== null && val !== undefined && val.trim() !== '') ? val.trim() : fallback;
  } catch (err) {
    console.warn(`Could not read property "${key}":`, err.message);
    return fallback;
  }
}

/**
 * Returns organizer notification emails as an array.
 * Reads comma-separated list from Script Properties (ORGANIZER_EMAILS).
 *
 * @returns {string[]} List of valid organizer emails.
 */
function getOrganizerEmails() {
  const raw = getScriptProp('ORGANIZER_EMAILS', 'organizer@example.com');
  return raw
    .split(',')
    .map(email => email.trim())
    .filter(email => email.length > 0);
}

/**
 * Constructs the core runtime configuration object for registration and emails.
 *
 * @returns {Object} Application configuration object.
 */
function getAppConfig() {
  return {
    SHEETS: {
      RESPONSES: getScriptProp('SHEET_NAME_RESPONSES', 'Sheet1'),
      CONFIG_TEMPLATES: getScriptProp('SHEET_NAME_TEMPLATES', 'Sheet2'),
    },
    SENDER_NAME: getScriptProp('EVENT_SENDER_NAME', 'Event Management Team'),
    ORGANIZER_EMAILS: getOrganizerEmails(),
    EVENT_NAME: getScriptProp('EVENT_NAME', 'Community Event 2026'),
    ORGANIZATION_NAME: getScriptProp('ORGANIZATION_NAME', 'Event Organizers Association'),
    DEFAULT_MAX_ATTENDEES: Number(getScriptProp('DEFAULT_MAX_ATTENDEES', '250')) || 250,
    PAYMENT_DUE_DAYS: Number(getScriptProp('PAYMENT_DUE_DAYS', '3')) || 3,
    TICKET_CODE_PREFIX: getScriptProp('TICKET_CODE_PREFIX', 'T'),
    QR_API_BASE_URL: getScriptProp('QR_API_BASE_URL', 'https://quickchart.io/qr'),

    // 1-based column positions in Sheet 1 (Google Form Responses)
    COLS: {
      TIMESTAMP: 1,      // Col A: Timestamp
      NAME: 2,           // Col B: Guest Full Name
      EMAIL: 3,          // Col C: Guest Email
      PHONE: 4,          // Col D: Guest Phone
      ADULTS: 5,         // Col E: Number of Adults
      CHILDREN: 6,       // Col F: Number of Children
      INFANTS: 7,        // Col G: Number of Infants
      CONSENT: 8,        // Col H: Consent / Terms
      TICKET_CODE: 12,   // Col L: Generated Ticket Code
      PAYMENT_STATUS: 13,// Col M: Payment Status (Dropdown: Paid / NotPaid)
      REG_EMAIL: 14,     // Col N: Registration Email Status
      CONF_EMAIL: 15,    // Col O: Confirmation Email Status
      REMINDER_EMAIL: 16,// Col P: Payment Reminder Status
    },

    // Coordinate references in Sheet 2 (Configuration & Email Templates)
    SHEET2: {
      COST_ADULT: { row: 11, col: 2 },
      COST_CHILD: { row: 12, col: 2 },
      CONFIRMED_COUNT: { row: 13, col: 2 },
      RESERVED_COUNT: { row: 14, col: 2 },
      MAX_ATTENDEES: { row: 15, col: 2 },

      TEMPLATES: {
        CONFIRMATION: 2, // Row 2: Subject (Col A), Body (Col B)
        REGISTRATION: 4, // Row 4: Subject (Col A), Body (Col B)
        FEEDBACK: 7,     // Row 7: Subject (Col A), Body (Col B)
        REMINDER: 17,    // Row 17: Subject (Col A), Body (Col B)
      }
    }
  };
}

/**
 * Constructs the runtime configuration object for the live gate entry & scanner.
 *
 * @returns {Object} Gate configuration object.
 */
function getGateConfig() {
  return {
    SHEET_RESPONSES: getScriptProp('SHEET_NAME_RESPONSES', 'Sheet1'),
    SHEET_CHECKINS: getScriptProp('SHEET_NAME_CHECKINS', 'CheckIns'),
    EVENT_NAME: getScriptProp('EVENT_NAME', 'Community Event 2026'),
    COLS: {
      TIMESTAMP: 1,         // Col A: Timestamp
      NAME: 2,              // Col B: Guest Full Name
      EMAIL: 3,             // Col C: Email (Excluded from scanner directory for privacy)
      PHONE: 4,             // Col D: Phone (Excluded from scanner directory for privacy)
      ADULTS: 5,            // Col E: Total Adults on Ticket
      CHILDREN: 6,          // Col F: Total Children on Ticket
      INFANTS: 7,           // Col G: Total Infants on Ticket
      TICKET_CODE: 12,      // Col L: Ticket Code
      PAYMENT_STATUS: 13,   // Col M: Payment Status
      
      // Live check-in tracking columns in Responses Sheet
      ADULTS_CHECKED: 19,   // Col S: Checked-in Adults Count
      CHILDREN_CHECKED: 20, // Col T: Checked-in Children Count
      INFANTS_CHECKED: 21,  // Col U: Checked-in Infants Count
      GATE_STATUS: 22       // Col V: Gate Status ('Not Checked In' | 'Partially Checked In' | 'Fully Checked In' | 'Canceled')
    }
  };
}
