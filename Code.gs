// ═══════════════════════════════════════════════════════════════
// SAT SAMARKAND — GOOGLE APPS SCRIPT
// Handles: diagnostics · abandoned tests · token management
// Deploy: New deployment → Web app → Execute as: Me → Anyone
// ═══════════════════════════════════════════════════════════════

// ─── CONFIGURATION ───────────────────────────────────────────
const SHEET_ID           = '1nneDzo_Uzj5sNvh5r5DHQHJT73CRPrvOfWUev9e5PXw';
const TELEGRAM_BOT_TOKEN = '8790225726:AAGcx7Izzl2fcyZVAoFq0ZEgurS9ZbmfScE';
const TELEGRAM_CHAT_ID   = '1632587141';
const ADMIN_PASSWORD     = '@SatSam2026';

const DIAGNOSTICS_TAB = 'Diagnostics';
const TOKENS_TAB      = 'Tokens';


// ═══════════════════════════════════════════════════════════════
// ENTRY POINTS
// ═══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    let data = (e && e.parameter && Object.keys(e.parameter).length > 0)
      ? e.parameter : {};
    if (Object.keys(data).length === 0 && e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch (_) { data = {}; }
    }

    const action = (data.action || '').toLowerCase();

    if      (action === 'diagnostic') handleDiagnostic(data);
    else if (action === 'abandon')    handleAbandoned(data);

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
      return validateToken(params.token || '');
    } else if (action === 'getsubmissions') {
      if (params.password !== ADMIN_PASSWORD) {
        result = { success: false, error: 'Unauthorized' };
      } else {
        result = { success: true, submissions: getDiagnosticSubmissions() };
      }
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

// Run from Apps Script editor to mint tokens. Change COUNT first.
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

// Run from editor console: resetToken('SAT-XXXXXX')
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

// Run from editor console: assignTokenToStudent('SAT-XXXXXX', 'Student Name')
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
// ABANDONED HANDLER — fires when student leaves mid-test
// ═══════════════════════════════════════════════════════════════

function handleAbandoned(data) {
  const ts  = new Date();
  const msg =
    '🚪 *Left Mid-Test — SAT Samarkand*\n\n' +
    '👤 *Name:* '  + (data.name  || '—') + '\n' +
    '📞 *Phone:* ' + (data.phone || '—') + '\n' +
    (data.grade ? '🎓 *Grade:* ' + data.grade + '\n' : '') +
    '\n📍 *Left at:* ' + (data.stage || '—') + '\n' +
    (data.duration ? '⏱ *Time spent:* ' + data.duration + '\n' : '') +
    (parseInt(data.tabSwitches) > 0 || parseInt(data.fullscreenExits) > 0
      ? '⚠️ *Anti-cheat:* ' + (data.tabSwitches||0) + ' tab switches · ' + (data.fullscreenExits||0) + ' fullscreen exits\n'
      : '') +
    '\n📅 ' + Utilities.formatDate(ts, 'Asia/Tashkent', 'dd MMM yyyy HH:mm');
  sendTelegram(msg);
}


// ═══════════════════════════════════════════════════════════════
// DIAGNOSTIC HANDLER — fires when student completes the test
// ═══════════════════════════════════════════════════════════════

function handleDiagnostic(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, DIAGNOSTICS_TAB, [
    'Timestamp', 'Name', 'Phone', 'Grade', 'Goal SAT',
    'Math Score', 'Math Total', 'Math %',
    'Eng Score',  'Eng Total',  'Eng %',
    'Total Score', 'Total Max', 'Total %',
    'CEFR', 'Eng Group', 'Math Group', 'Recommendation',
    'Tab Switches', 'Fullscreen Exits', 'Duration', 'Token'
  ]);

  const ts = new Date();

  const mathScore   = parseInt(data.mathScore) || 0;
  const mathTotal   = parseInt(data.mathTotal) || 20;
  const mathPct     = parseInt(data.mathPct)   || Math.round(mathScore / mathTotal * 100);
  const engScore    = parseInt(data.engScore)  || 0;
  const engTotal    = parseInt(data.engTotal)  || 20;
  const engPct      = parseInt(data.engPct)    || Math.round(engScore / engTotal * 100);
  const totalScore  = mathScore + engScore;
  const totalMax    = mathTotal + engTotal;
  const totalPct    = parseInt(data.totalPct)  || Math.round(totalScore / totalMax * 100);

  const cefr            = data.cefr || determineCEFR(engPct);
  const engGroup        = engPct  >= 70 ? 'SAT Ready' : 'Foundation';
  const mathGroup       = mathPct >= 70 ? 'SAT Ready' : 'Foundation';
  const recommendation  = generateRecommendation(totalPct, mathPct, engPct);
  const tabSwitches     = parseInt(data.tabSwitches)     || 0;
  const fullscreenExits = parseInt(data.fullscreenExits) || 0;

  sheet.appendRow([
    ts,
    data.name  || '', data.phone || '', data.grade || '', data.goal || '',
    mathScore, mathTotal, mathPct,
    engScore,  engTotal,  engPct,
    totalScore, totalMax, totalPct,
    cefr, engGroup, mathGroup, recommendation,
    tabSwitches, fullscreenExits,
    data.duration || '', data.token || ''
  ]);

  // Green row = both sections SAT Ready; amber = at least one Foundation
  const allReady = engGroup === 'SAT Ready' && mathGroup === 'SAT Ready';
  sheet.getRange(sheet.getLastRow(), 1, 1, 22)
       .setBackground(allReady ? '#dcfce7' : '#fef3c7');

  // Telegram
  const engStatus  = engPct  >= 70 ? '✅ SAT Ready' : '🟡 Foundation';
  const mathStatus = mathPct >= 70 ? '✅ SAT Ready' : '🟡 Foundation';
  const msg =
    '🎯 *Diagnostic Complete — SAT Samarkand*\n\n' +
    '👤 *Name:* '  + (data.name  || '—') + '\n' +
    '📞 *Phone:* ' + (data.phone || '—') + '\n' +
    (data.grade ? '🎓 *Grade:* ' + data.grade + '\n' : '') +
    (data.goal  ? '🎯 *Goal:* '  + data.goal  + '\n' : '') +
    '\n📊 *Results:*\n' +
    '  English: ' + engScore  + '/' + engTotal  + ' (' + engPct  + '%) — ' + engStatus  + '\n' +
    '  Math:    ' + mathScore + '/' + mathTotal + ' (' + mathPct + '%) — ' + mathStatus + '\n' +
    '  Total:   ' + totalScore + '/' + totalMax + ' (' + totalPct + '%)\n\n' +
    '🎓 *Recommended:* ' + recommendation + '\n' +
    (tabSwitches > 0 || fullscreenExits > 0
      ? '⚠️ *Anti-cheat:* ' + tabSwitches + ' tab switches · ' + fullscreenExits + ' fullscreen exits\n'
      : '') +
    (data.duration ? '⏱ *Duration:* ' + data.duration + '\n' : '') +
    '\n📅 ' + Utilities.formatDate(ts, 'Asia/Tashkent', 'dd MMM yyyy HH:mm') + '\n\n' +
    (allReady ? '🔥 *Both sections ready — hot lead!*' : '💼 *Needs foundation work first.*');

  sendTelegram(msg);

  // Mark token as used
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
// ADMIN
// ═══════════════════════════════════════════════════════════════

function getDiagnosticSubmissions() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DIAGNOSTICS_TAB);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 22).getValues();
  return values.map((row, i) => ({
    id: i + 1,
    date:            row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
    name:            row[1],  phone:          row[2],
    grade:           row[3],  goal:           row[4],
    mathScore:       row[5],  mathTotal:      row[6],  mathPct:   row[7],
    engScore:        row[8],  engTotal:       row[9],  engPct:    row[10],
    totalScore:      row[11], totalMax:       row[12], totalPct:  row[13],
    cefr:            row[14], engGroup:       row[15], mathGroup: row[16],
    recommendation:  row[17],
    tabSwitches:     row[18], fullscreenExits:row[19],
    duration:        row[20], token:          row[21]
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
  if (totalPct >= 85) return 'Advanced Group';
  if (totalPct >= 65) return 'Upper-Intermediate Group';
  if (totalPct >= 45) {
    if (mathPct < engPct - 15) return 'Intermediate Group (Math Focus)';
    if (engPct < mathPct - 15) return 'Intermediate Group (English Focus)';
    return 'Intermediate Group';
  }
  if (totalPct >= 25) return 'Elementary Group';
  return 'Beginner Group';
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
// TEST — run manually from Apps Script editor
// ═══════════════════════════════════════════════════════════════

function testDiagnostic() {
  handleDiagnostic({
    name: 'Test Student', phone: '+998-95-113-16-00',
    grade: '11th grade', goal: '1500',
    mathScore: '16', mathTotal: '20', mathPct: '80',
    engScore:  '12', engTotal:  '20', engPct:  '60',
    totalPct: '70', cefr: 'B1',
    tabSwitches: '0', fullscreenExits: '0', duration: '28m 12s'
  });
}

function testAbandoned() {
  handleAbandoned({
    name: 'Test Student', phone: '+998-95-113-16-00',
    grade: '11th grade', stage: 'English Section — Question 8/20',
    duration: '12m 5s', tabSwitches: '1', fullscreenExits: '0'
  });
}
