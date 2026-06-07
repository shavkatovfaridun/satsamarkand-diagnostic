// ═══════════════════════════════════════════════════════════════
// SAT SAMARKAND — GOOGLE APPS SCRIPT (unified)
// Handles: registrations · diagnostics · token management · admin queries
// Deploy: New deployment → Web app → Execute as: Me → Anyone
// ═══════════════════════════════════════════════════════════════

// ─── CONFIGURATION ───────────────────────────────────────────
const SHEET_ID            = '1nneDzo_Uzj5sNvh5r5DHQHJT73CRPrvOfWUev9e5PXw';
const TELEGRAM_BOT_TOKEN  = '8790225726:AAGcx7Izzl2fcyZVAoFq0ZEgurS9ZbmfScE';
const TELEGRAM_CHAT_ID    = '1632587141';
const ADMIN_PASSWORD      = '@SatSam2026';

// Sheet tab names
const REGISTRATIONS_TAB = 'Registrations';
const DIAGNOSTICS_TAB   = 'Diagnostics';
const TOKENS_TAB        = 'Tokens';


// ═══════════════════════════════════════════════════════════════
// ENTRY POINTS
// ═══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    let data = (e && e.parameter && Object.keys(e.parameter).length > 0)
      ? e.parameter
      : {};
    if (Object.keys(data).length === 0 && e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch (_) { data = {}; }
    }

    const type = (data.type || 'registration').toLowerCase();

    if (type === 'diagnostic') handleDiagnostic(data);
    else if (type === 'mock_score') handleMockScore(data);
    else handleRegistration(data);

    return json({ success: true });
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return json({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    const params   = (e && e.parameter) || {};
    const action   = (params.action || '').toLowerCase();
    const callback = params.callback || '';

    let result;

    if (action === 'validate') {
      // validateToken returns a ContentService output directly
      return validateToken(params.token || '');
    } else if (action === 'getsubmissions') {
      if (params.password !== ADMIN_PASSWORD) {
        result = { success: false, error: 'Unauthorized' };
      } else {
        result = { success: true, submissions: getDiagnosticSubmissions() };
      }
    } else if (action === 'get_scores') {
      result = getMockScores();
    } else if (action === 'get_registrations') {
      result = getRegistrations();
    } else if (action === 'ping') {
      result = { success: true, status: 'alive', time: new Date().toISOString() };
    } else {
      result = { success: false, error: 'Unknown action' };
    }

    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return json(result);
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return json({ success: false, error: err.toString() });
  }
}


// ═══════════════════════════════════════════════════════════════
// TOKEN MANAGEMENT
// ═══════════════════════════════════════════════════════════════

function validateToken(token) {
  if (!token) return json({ valid: false, reason: 'missing' });

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TOKENS_TAB);
  if (!sheet) return json({ valid: false, reason: 'sheet_missing' });

  // Columns: Token | Name | Status | DateIssued | DateStarted | DateCompleted
  const rows = sheet.getDataRange().getValues();
  const tok  = token.trim().toUpperCase();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toUpperCase() === tok) {
      const status = String(rows[i][2]).toLowerCase().trim();
      if (status === 'used') return json({ valid: false, reason: 'already_used' });

      sheet.getRange(i + 1, 3).setValue('in_progress');
      sheet.getRange(i + 1, 5).setValue(new Date().toISOString());

      const studentName = rows[i][1] ? String(rows[i][1]).trim() : '';
      return json({ valid: true, name: studentName });
    }
  }

  return json({ valid: false, reason: 'not_found' });
}

/**
 * generateTokens() — run from Apps Script editor to mint tokens.
 * Change COUNT before running.
 */
function generateTokens() {
  const COUNT = 30;
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, TOKENS_TAB,
    ['Token', 'Name', 'Status', 'DateIssued', 'DateStarted', 'DateCompleted']);

  const tokens = [];
  for (let i = 0; i < COUNT; i++) {
    const token = 'SAT-' + Math.random().toString(36).substr(2, 6).toUpperCase();
    sheet.appendRow([token, '', 'unused', new Date().toISOString(), '', '']);
    tokens.push(token);
  }
  Logger.log('Generated ' + COUNT + ' tokens:\n' + tokens.join('\n'));
  SpreadsheetApp.getUi().alert('Done! Generated ' + COUNT + ' tokens.\nSee the "Tokens" sheet.');
}

/**
 * resetToken('SAT-XXXXXX') — run from editor console
 */
function resetToken(token) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TOKENS_TAB);
  if (!sheet) { Logger.log('Tokens sheet not found.'); return; }
  const rows = sheet.getDataRange().getValues();
  const tok  = String(token).trim().toUpperCase();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toUpperCase() === tok) {
      sheet.getRange(i + 1, 3).setValue('unused');
      sheet.getRange(i + 1, 5).setValue('');
      sheet.getRange(i + 1, 6).setValue('');
      Logger.log('Token reset: ' + token);
      return;
    }
  }
  Logger.log('Token not found: ' + token);
}

/**
 * assignTokenToStudent('SAT-XXXXXX', 'Student Name') — run from editor console
 */
function assignTokenToStudent(token, name) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TOKENS_TAB);
  if (!sheet) { Logger.log('Tokens sheet not found.'); return; }
  const rows = sheet.getDataRange().getValues();
  const tok  = String(token).trim().toUpperCase();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toUpperCase() === tok) {
      sheet.getRange(i + 1, 2).setValue(name);
      Logger.log('Assigned "' + name + '" to ' + token);
      return;
    }
  }
  Logger.log('Token not found: ' + token);
}


// ═══════════════════════════════════════════════════════════════
// REGISTRATION HANDLER
// ═══════════════════════════════════════════════════════════════

function handleRegistration(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, REGISTRATIONS_TAB,
    ['Timestamp', 'Name', 'Phone', 'Age', 'Grade', 'School', 'Program', 'Heard', 'Message']);

  const ts = new Date();
  sheet.appendRow([
    ts,
    data.name    || '',
    data.phone   || '',
    data.age     || '',
    data.grade   || '',
    data.school  || '',
    data.program || '',
    data.heard   || '',
    data.message || ''
  ]);

  const msg =
    '🔔 *New Lead — SAT Samarkand*\n\n' +
    '👤 *Name:* ' + (data.name  || '—') + '\n' +
    '📞 *Phone:* `' + (data.phone || '—') + '`\n' +
    (data.age     ? '🎂 *Age:* ' + data.age + '\n' : '') +
    (data.grade   ? '🎓 *Grade:* ' + data.grade + '\n' : '') +
    (data.school  ? '🏫 *School:* ' + data.school + '\n' : '') +
    (data.program ? '📚 *Program:* ' + data.program + '\n' : '') +
    (data.heard   ? '📡 *Source:* ' + data.heard + '\n' : '') +
    (data.message ? '\n💬 *Message:* ' + data.message + '\n' : '') +
    '\n📅 ' + Utilities.formatDate(ts, 'Asia/Tashkent', 'dd MMM yyyy HH:mm') +
    '\n\n⚡ _Contact within 24 hours_';

  sendTelegram(msg);
}


// ═══════════════════════════════════════════════════════════════
// DIAGNOSTIC HANDLER
// ═══════════════════════════════════════════════════════════════

function handleDiagnostic(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, DIAGNOSTICS_TAB, [
    'Timestamp', 'Name', 'Phone', 'Age', 'Grade',
    'Goal SAT', 'Target Date', 'University',
    'Math Score', 'Math Total', 'Math %',
    'Eng Score',  'Eng Total',  'Eng %',
    'Total Score', 'Total Max', 'Total %',
    'CEFR Level', 'Status', 'Recommendation',
    'Tab Switches', 'Fullscreen Exits', 'Duration', 'Token', 'Timestamp'
  ]);

  const ts = new Date();

  // Use dynamic totals — works for any question count, forever
  const mathScore   = parseInt(data.mathScore)   || 0;
  const mathTotal   = parseInt(data.mathTotal)   || 20;
  const mathPct     = parseInt(data.mathPct)     || Math.round(mathScore / mathTotal * 100);
  const engScore    = parseInt(data.engScore)    || 0;
  const engTotal    = parseInt(data.engTotal)    || 20;
  const engPct      = parseInt(data.engPct)      || Math.round(engScore / engTotal * 100);
  const totalScore  = mathScore + engScore;
  const totalMax    = mathTotal + engTotal;
  const totalPct    = parseInt(data.totalPct)    || Math.round(totalScore / totalMax * 100);

  const cefr            = data.cefr         || determineCEFR(engPct);
  const passed          = data.passed === true || data.passed === 'true' || data.passed === 'YES';
  const status          = passed ? 'Passed' : 'Below threshold';
  const recommendation  = generateRecommendation(totalPct, mathPct, engPct);
  const tabSwitches     = parseInt(data.tabSwitches)     || 0;
  const fullscreenExits = parseInt(data.fullscreenExits) || 0;

  sheet.appendRow([
    ts,
    data.name       || '',
    data.phone      || '',
    data.age        || '',
    data.grade      || '',
    data.goal       || '',
    data.targetDate || '',
    data.university || '',
    mathScore, mathTotal, mathPct,
    engScore,  engTotal,  engPct,
    totalScore, totalMax, totalPct,
    cefr, status, recommendation,
    tabSwitches, fullscreenExits,
    data.duration   || '',
    data.token      || '',
    ts.toISOString()
  ]);

  // Row colour: green = passed, amber = needs work
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 25)
       .setBackground(passed ? '#dcfce7' : '#fef3c7');

  // ── Telegram notification ──────────────────────────────────
  const passEmoji  = passed ? '✅' : '⚠️';
  const statusText = passed ? 'PASSED — SAT-Ready' : 'Foundation Needed';
  const cheatNote  = (tabSwitches > 0 || fullscreenExits > 0)
    ? '\n⚠️ *Anti-cheat:* ' + tabSwitches + ' tab switches · ' + fullscreenExits + ' fullscreen exits'
    : '';
  const durationNote = data.duration ? '\n⏱ *Duration:* ' + data.duration : '';

  const msg =
    '🎯 *New Diagnostic — SAT Samarkand*\n\n' +
    '👤 *Name:* ' + (data.name  || '—') + '\n' +
    '📞 *Phone:* `' + (data.phone || '—') + '`\n' +
    '🎂 *Age:* ' + (data.age || '—') + ' · *Grade:* ' + (data.grade || '—') + '\n' +
    (data.goal       ? '🎯 *Goal SAT:* ' + data.goal + '+\n'        : '') +
    (data.university ? '🏛 *University:* ' + data.university + '\n' : '') +
    '\n📊 *Score:* ' + totalScore + '/' + totalMax + ' (' + totalPct + '%)\n' +
    '   📐 Math:    ' + mathScore + '/' + mathTotal + ' (' + mathPct + '%)\n' +
    '   📖 English: ' + engScore  + '/' + engTotal  + ' (' + engPct  + '%) — ' + cefr + '\n\n' +
    passEmoji + ' *STATUS:* ' + statusText + '\n' +
    '🎓 *RECOMMENDED:* ' + recommendation +
    cheatNote + durationNote + '\n\n' +
    '📅 ' + Utilities.formatDate(ts, 'Asia/Tashkent', 'dd MMM yyyy HH:mm') + '\n\n' +
    (passed ? '🔥 *Hot lead — call within 24h!*' : '💼 Lead needs nurturing.');

  sendTelegram(msg);

  // Mark token as used (if provided)
  if (data.token) {
    const tokenSheet = ss.getSheetByName(TOKENS_TAB);
    if (tokenSheet) {
      const rows = tokenSheet.getDataRange().getValues();
      const tok  = String(data.token).trim().toUpperCase();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim().toUpperCase() === tok) {
          tokenSheet.getRange(i + 1, 3).setValue('used');
          tokenSheet.getRange(i + 1, 6).setValue(ts.toISOString());
          break;
        }
      }
    }
  }
}


// ═══════════════════════════════════════════════════════════════
// MOCK SCORES
// ═══════════════════════════════════════════════════════════════

function handleMockScore(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, 'MockScores',
    ['Timestamp', 'Name', 'Score', 'Date', 'Added By']);
  sheet.appendRow([new Date(), data.name || '', data.score || '', data.date || '', data.added_by || 'Admin']);
}

function getMockScores() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MockScores');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const scores = rows.slice(1).filter(r => r[1] && r[2]).map(r => ({ name: r[1], score: Number(r[2]), date: r[3] }));
  scores.sort((a, b) => b.score - a.score);
  return scores.slice(0, 10);
}


// ═══════════════════════════════════════════════════════════════
// ADMIN — fetch submissions
// ═══════════════════════════════════════════════════════════════

function getDiagnosticSubmissions() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DIAGNOSTICS_TAB);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 25).getValues();
  return values.map((row, i) => ({
    id:              i + 1,
    date:            row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
    name:            row[1],
    phone:           row[2],
    age:             row[3],
    grade:           row[4],
    goal:            row[5],
    when:            row[6],
    uni:             row[7],
    mathScore:       row[8],
    mathTotal:       row[9],
    mathPct:         row[10],
    engScore:        row[11],
    engTotal:        row[12],
    engPct:          row[13],
    totalScore:      row[14],
    totalMax:        row[15],
    totalPct:        row[16],
    cefr:            row[17],
    status:          row[18] === 'Passed' ? 'passed' : 'failed',
    recommendation:  row[19],
    tabSwitches:     row[20],
    fullscreenExits: row[21],
    duration:        row[22],
    token:           row[23]
  })).reverse();
}

function getRegistrations() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(REGISTRATIONS_TAB);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1).filter(r => r[1]).map(r => ({
    time: r[0], name: r[1], phone: r[2], age: r[3], grade: r[4], school: r[5], program: r[6]
  })).reverse();
}


// ═══════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#0B1F3A')
         .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function determineCEFR(engPct) {
  if (engPct >= 86) return 'C1+';
  if (engPct >= 66) return 'B2';
  if (engPct >= 46) return 'B1';
  if (engPct >= 26) return 'A2';
  return 'A1';
}

function generateRecommendation(totalPct, mathPct, engPct) {
  if (totalPct >= 95) return 'Private 1-on-1';
  if (totalPct >= 80) return 'SAT MAX (1500+ Guaranteed)';
  if (totalPct >= 60) {
    if (mathPct < engPct - 15) return 'Intensive Program (Math Focus)';
    if (engPct < mathPct - 15) return 'Intensive Program (English Focus)';
    return 'Intensive Program';
  }
  if (totalPct >= 40) return '1 Subject Program (Foundation)';
  return 'Foundation Program (Pre-SAT)';
}

function sendTelegram(text) {
  try {
    UrlFetchApp.fetch(
      'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage',
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          chat_id: TELEGRAM_CHAT_ID,
          text: text,
          parse_mode: 'Markdown',
          disable_web_page_preview: true
        }),
        muteHttpExceptions: true
      }
    );
  } catch (err) {
    Logger.log('Telegram error: ' + err);
  }
}


// ═══════════════════════════════════════════════════════════════
// TESTS — run manually from Apps Script editor
// ═══════════════════════════════════════════════════════════════

function testRegistration() {
  handleRegistration({ name: 'Test User', phone: '+998-95-113-16-00', program: 'Full SAT' });
}

function testDiagnostic() {
  handleDiagnostic({
    name: 'Test Student', phone: '+998-95-113-16-00',
    age: '17', grade: '11th grade', goal: '1500',
    targetDate: '6months', university: 'METU',
    mathScore: '16', mathTotal: '20', mathPct: '80',
    engScore:  '14', engTotal:  '20', engPct:  '70',
    totalPct: '75', cefr: 'B2', passed: 'true',
    tabSwitches: '0', fullscreenExits: '0', duration: '28m 12s'
  });
}
