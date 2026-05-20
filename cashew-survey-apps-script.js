// ============================================================
//  CONFIG — edit only this section before deploying
// ============================================================
const SHEET_ID    = '1e4JSfAw4BG7Ru8FvsGoqHHz4nE88cP1jz53UhLYEDlc'; // ← your sheet
const SHEET_NAME  = 'Sheet1';                // tab name shown at the bottom of the sheet
const SECRET_KEY  = 'easier';               // must match the "secret" field in your HTML
const NOTIFY_EMAIL = 'lakshya.tiwari1201@gmail.com'; // notification recipient

// Rate-limiting: max submissions accepted from one IP per hour
const MAX_PER_HOUR = 5;

// ============================================================
//  MAIN ENTRY POINT
// ============================================================
function doGet() {
  return ContentService
    .createTextOutput('The Cashew Project survey endpoint is live. Submissions accepted via POST only.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    // 1. Parse the body ----------------------------------------
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, error: 'No data received' });
    }

    const data = JSON.parse(e.postData.contents);

    // 2. Validate shared secret --------------------------------
    if (data.secret !== SECRET_KEY) {
      return jsonResponse({ success: false, error: 'Unauthorized' });
    }

    // 3. Sanitize inputs ---------------------------------------
    const email = sanitize(data.email || '');
    const phone = sanitize(data.phone || '');
    const q1    = sanitize(data.q1    || '');
    const q2    = sanitize(data.q2    || '');
    const q3    = sanitize(data.q3    || '');
    const q4    = sanitize(data.q4    || '');

    if (!email || !isValidEmail(email)) {
      return jsonResponse({ success: false, error: 'Invalid email' });
    }

    // 4. Rate limiting -----------------------------------------
    const ip = (e.parameter && e.parameter.userIp) ? e.parameter.userIp : 'unknown';
    if (!checkRateLimit(ip)) {
      return jsonResponse({ success: false, error: 'Too many submissions. Try again later.' });
    }

    // 5. Spam filter -------------------------------------------
    if (isSpam(email, q1, q2, q3, q4)) {
      // Silently accept without writing — spammers shouldn't know they're blocked
      return jsonResponse({ success: true });
    }

    // 6. Write to Sheet ----------------------------------------
    const sheet = SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName(SHEET_NAME);

    const timestamp = new Date().toISOString();

    sheet.appendRow([timestamp, email, phone, q1, q2, q3, q4]);

    // 7. Email notification ------------------------------------
    sendNotification(timestamp, email, phone, q1, q2, q3, q4);

    return jsonResponse({ success: true });

  } catch (err) {
    console.error('doPost error:', err.toString());
    return jsonResponse({ success: false, error: 'Server error. Please try again.' });
  }
}

// ============================================================
//  HELPERS
// ============================================================

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sanitize(str) {
  return String(str).replace(/<[^>]*>/g, '').trim().substring(0, 500);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function isSpam(email, q1, q2, q3, q4) {
  const combined = [email, q1, q2, q3, q4].join(' ').toLowerCase();
  const spamKeywords = ['http://', 'https://', 'viagra', 'casino', 'click here', 'free money'];
  return spamKeywords.some(kw => combined.includes(kw));
}

// ============================================================
//  RATE LIMITING
//  Uses Script Properties as a lightweight per-IP counter.
//  Each IP gets a JSON array of timestamps; entries older than
//  1 hour are pruned on every check.
// ============================================================

function checkRateLimit(ip) {
  const props  = PropertiesService.getScriptProperties();
  const key    = 'rl_' + ip;
  const now    = Date.now();
  const window = 60 * 60 * 1000; // 1 hour in ms

  let log = [];
  try { log = JSON.parse(props.getProperty(key) || '[]'); } catch (_) {}

  log = log.filter(t => now - t < window);

  if (log.length >= MAX_PER_HOUR) return false;

  log.push(now);
  props.setProperty(key, JSON.stringify(log));
  return true;
}

// ============================================================
//  EMAIL NOTIFICATION
// ============================================================

function sendNotification(timestamp, email, phone, q1, q2, q3, q4) {
  try {
    const subject = '🥜 New Cashew Project Survey Response';
    const body = [
      'New survey response received!\n',
      'Timestamp:           ' + timestamp,
      'Email:               ' + (email || '—'),
      'Phone:               ' + (phone || '—'),
      'Purchase Channel:    ' + q1,
      'Biggest Frustration: ' + q2,
      'Purchase Driver:     ' + q3,
      'Delivery Format:     ' + q4,
      '\nView all responses:',
      'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit'
    ].join('\n');

    GmailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (err) {
    // Don't let a notification failure break the submission
    console.error('Email notification failed:', err.toString());
  }
}
