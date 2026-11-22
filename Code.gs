/**
 * SafeHangar - Google Sheets backend (Google Apps Script Web App)
 *
 * Deployment: Extensions > Apps Script > Deploy > New deployment >
 *   Type: Web app > Execute as: Me > Who has access: Anyone
 *   Copy the /exec URL and paste it into API_URL in script.js.
 *
 * Sheet tabs required (headers in row 1, created by setupTab()):
 *   Aircraft, Inspections, Hazards, Facilities, Users
 *
 * Run from the editor only (never exposed over HTTP):
 *   seedDemoData()   - write demo rows into empty tabs
 *   clearAllData()   - empty Aircraft/Inspections/Hazards (keep Facilities/Users)
 *   clearEverything()- empty all five tabs
 *   resetDemoData()  - clearEverything() + seedDemoData()
 */

var SHEETS = {
  AIRCRAFT: 'Aircraft',
  INSPECTIONS: 'Inspections',
  HAZARDS: 'Hazards',
  FACILITIES: 'Facilities',
  USERS: 'Users'
};

/** Column headers per tab. setupTab() creates them if missing. */
var HEADERS = {
  Aircraft: ['ID', 'Registration', 'Model', 'Serial Number', 'Operator', 'Status',
             'Registered At'],
  Inspections: ['ID', 'Aircraft', 'Inspection Type', 'Inspector', 'Date',
                'Fuselage', 'Wings', 'Engine', 'Gear', 'Overall Result', 'Remarks',
                'Created At'],
  Hazards: ['ID', 'Category', 'Severity', 'Facility', 'Location', 'Reported By',
            'Date Observed', 'Description', 'Status', 'Created At'],
  Facilities: ['ID', 'Name', 'Address', 'Region'],
  Users: ['ID', 'Name', 'Role', 'Email', 'Status']
};

/* ------------------------------------------------------------------ */
/* Web app entry points                                                */
/* ------------------------------------------------------------------ */

function doGet(e) {
  return handle(e.parameter || {}, 'GET');
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { body = {}; }
  return handle(body, 'POST');
}

function handle(params, method) {
  try {
    ensureSetup();
    var action = String(params.action || '');
    var result;

    switch (action) {
      case 'list':        result = apiList(params.entity); break;
      case 'create':      result = apiCreate(params.entity, params.record || {}); break;
      case 'update':      result = apiUpdate(params.entity, params.id, params.record || {}); break;
      case 'delete':      result = apiDelete(params.entity, params.id); break;
      case 'bootstrap':   result = apiBootstrap(); break;
      case 'ping':        result = { ok: true, time: new Date().toISOString() }; break;
      default:            throw new Error('Unknown action: ' + action);
    }
    return json_({ ok: true, data: result });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Generic CRUD                                                         */
/* ------------------------------------------------------------------ */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function setupTab(name) {
  var book = ss_();
  var sheet = book.getSheetByName(name);
  if (!sheet) sheet = book.insertSheet(name);
  var expected = HEADERS[name];
  if (!expected) throw new Error('Unknown sheet: ' + name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected])
      .setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureSetup() {
  Object.keys(HEADERS).forEach(function (name) { setupTab(name); });
}

function readAll_(sheetName) {
  var sheet = setupTab(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = HEADERS[sheetName].length;
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var headers = HEADERS[sheetName];
  return values.map(function (row, i) {
    var obj = { _row: i + 2 };
    headers.forEach(function (h, c) { obj[h] = row[c]; });
    return obj;
  }).filter(function (r) {
    // skip fully blank rows
    return headers.some(function (h) { return r[h] !== '' && r[h] !== null; });
  });
}

function nextId_(prefix) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty('SEQ_' + prefix) || 0) + 1;
    props.setProperty('SEQ_' + prefix, String(n));
    return prefix + '-' + ('000' + n).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function cleanRecord_(sheetName, record) {
  var headers = HEADERS[sheetName];
  var out = [];
  headers.forEach(function (h) {
    var v = record[h];
    if (v === undefined || v === null) v = '';
    out.push(v);
  });
  return out;
}

function apiList(entity) {
  var name = SHEETS[String(entity).toUpperCase()] || entity;
  return readAll_(name).map(function (r) { delete r._row; return r; });
}

function apiCreate(entity, record) {
  var name = SHEETS[String(entity).toUpperCase()] || entity;
  var sheet = setupTab(name);
  var idKey = HEADERS[name][0]; // 'ID'
  if (!record[idKey]) record[idKey] = nextId_(name.slice(0, 2).toUpperCase());
  if (HEADERS[name].indexOf('Created At') >= 0 && !record['Created At']) {
    record['Created At'] = new Date().toISOString();
  }
  sheet.appendRow(cleanRecord_(name, record));
  return record;
}

function apiUpdate(entity, id, record) {
  var name = SHEETS[String(entity).toUpperCase()] || entity;
  var sheet = setupTab(name);
  var rows = readAll_(name);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][HEADERS[name][0]]) === String(id)) {
      var merged = rows[i];
      HEADERS[name].forEach(function (h) {
        if (record[h] !== undefined) merged[h] = record[h];
      });
      sheet.getRange(rows[i]._row, 1, 1, HEADERS[name].length)
        .setValues([cleanRecord_(name, merged)]);
      delete merged._row;
      return merged;
    }
  }
  throw new Error('Record not found: ' + id);
}

function apiDelete(entity, id) {
  var name = SHEETS[String(entity).toUpperCase()] || entity;
  var sheet = setupTab(name);
  var rows = readAll_(name);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][HEADERS[name][0]]) === String(id)) {
      sheet.deleteRow(rows[i]._row);
      return { deleted: id };
    }
  }
  throw new Error('Record not found: ' + id);
}

/* ------------------------------------------------------------------ */
/* One-call bootstrap for the dashboard                                */
/* ------------------------------------------------------------------ */

function apiBootstrap() {
  return {
    aircraft: apiList('aircraft'),
    inspections: apiList('inspections'),
    hazards: apiList('hazards'),
    facilities: apiList('facilities'),
    users: apiList('users')
  };
}

/* ------------------------------------------------------------------ */
/* Clear / reset utilities.                                            */
/* Run these FROM THE APPS SCRIPT EDITOR ONLY (function dropdown ->    */
/* Run). They are deliberately NOT reachable through the web app:     */
/* handle() only routes list/create/update/delete/bootstrap/ping.     */
/*                                                                     */
/*   clearAllData()    - clears the 3 transactional tabs (Aircraft,    */
/*                       Inspections, Hazards); keeps the Facilities   */
/*                       directory and Users, which the forms need.    */
/*   clearEverything() - clears ALL five tabs (fresh empty database).  */
/*   resetDemoData()   - clearEverything() + seedDemoData() again.     */
/* ------------------------------------------------------------------ */

function clearRows_(name) {
  var sheet = setupTab(name);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.deleteRows(2, lastRow - 1);
}

function clearAllData() {
  ['Aircraft', 'Inspections', 'Hazards'].forEach(clearRows_);
  Logger.log('Cleared Aircraft, Inspections and Hazards. Facilities and Users were kept.');
}

function clearEverything() {
  Object.keys(HEADERS).forEach(clearRows_);
  Logger.log('Cleared all tabs. The database is now empty (headers only).');
}

function resetDemoData() {
  clearEverything();
  seedDemoData();
  Logger.log('Database cleared and reseeded with the demo rows.');
}


function seedDemoData() {
  ensureSetup();
  var fac = setupTab(SHEETS.FACILITIES);
  if (fac.getLastRow() < 2) {
    [
      ['Lufthansa Technik Philippines', 'MacroAsia Special Economic Zone, Villamor Air Base, Pasay City, Metro Manila 1309', 'Metro Manila'],
      ['Aviation Partnership (Philippines) Corporation (Aplus)', 'Building 8006, Domestic Road, Pasay City, Metro Manila 1301', 'Metro Manila'],
      ['INAEC Aviation Corporation - Hangars A & B', 'General Aviation Area, Manila Domestic Airport, Pasay City, Metro Manila', 'Metro Manila'],
      ['INAEC Aviation Corporation - Hangar D', 'Andrews Avenue, Manila Domestic Airport, Pasay City, Metro Manila', 'Metro Manila'],
      ['PhilJets Aero Services Inc.', 'AAOP Hangar 2, Manila Domestic Airport, Domestic Road corner MIAA Road, Pasay City, Metro Manila 1300', 'Metro Manila'],
      ['Asian Aerospace Corporation', 'Lima Gate, Andrews Avenue, NAIA Complex, Pasay City, Metro Manila 1300', 'Metro Manila'],
      ['SIA Engineering (Philippines) Corporation', 'Building 7592, A. Bonifacio Avenue, Civil Aviation Complex, Clark Freeport Zone, Pampanga 2023', 'Clark, Pampanga'],
      ['Metrojet Engineering (Clark) Ltd.', 'Building N7727, M.A. Roxas Highway, Clark Civil Aviation Complex, Clark Freeport Zone, Pampanga 2023', 'Clark, Pampanga'],
      ['Dornier Technology, Inc.', 'Hangar 7224, Clark International Airport, Clark Freeport Zone, Pampanga', 'Clark, Pampanga'],
      ['Asian Aerospace Corporation - Clark Facility', 'Clark Freeport Zone, Pampanga', 'Clark, Pampanga'],
      ['INAEC North', 'Building 7216, Civil Aviation Complex, Clark International Airport, Pampanga', 'Clark, Pampanga'],
      ['First Aviation Academy, Inc. (FAA)', 'Building 8303, Southwest Apron B, Argonaut Highway, Subic Bay International Airport, Subic Bay Freeport Zone, Zambales 2200', 'Subic, Zambales']
    ].forEach(function (row) {
      fac.appendRow(['', row[0], row[1], row[2]]); // ID blank -> assign below
    });
    // assign IDs
    var last = fac.getLastRow();
    for (var r = 2; r <= last; r++) {
      if (!fac.getRange(r, 1).getValue()) {
        fac.getRange(r, 1).setValue(nextId_('FA'));
      }
    }
  }

  var air = setupTab(SHEETS.AIRCRAFT);
  if (air.getLastRow() < 2) {
    [
      ['RP-C1234', 'Cessna 172', '172-8841', 'Training Fleet', 'Active'],
      ['RP-C5678', 'Piper PA-28', 'PA28-2291', 'Private', 'Active'],
      ['RP-C9012', 'Cessna 152', 'C152-3377', 'Training Fleet', 'Active'],
      ['RP-C3456', 'Cessna 208 Caravan', '208B-0102', 'Private', 'Active'],
      ['RP-C7890', 'Beechcraft Baron 58', 'BE-58-4410', 'Private', 'Maintenance']
    ].forEach(function (row) {
      apiCreate('aircraft', {
        Registration: row[0], Model: row[1], 'Serial Number': row[2],
        Operator: row[3], Status: row[4], 'Registered At': new Date().toISOString()
      });
    });
  }

  var insp = setupTab(SHEETS.INSPECTIONS);
  if (insp.getLastRow() < 2) {
    apiCreate('inspections', {
      Aircraft: 'RP-C1234', 'Inspection Type': 'Pre-flight inspection',
      Inspector: 'J. Mendoza', Date: '2026-11-15',
      Fuselage: 'Pass', Wings: 'Pass', Engine: 'Pass', Gear: 'Pass',
      'Overall Result': 'Passed', Remarks: 'No defects observed.'
    });
    apiCreate('inspections', {
      Aircraft: 'RP-C5678', 'Inspection Type': 'Engine inspection',
      Inspector: 'R. Santos', Date: '2026-11-14',
      Fuselage: 'Pass', Wings: 'Pass', Engine: 'Needs review', Gear: 'Pass',
      'Overall Result': 'Needs review', Remarks: 'Oil residue near cylinder No. 3.'
    });
    apiCreate('inspections', {
      Aircraft: 'RP-C9012', 'Inspection Type': 'Structural inspection',
      Inspector: 'A. Cruz', Date: '2026-11-13',
      Fuselage: 'Fail', Wings: 'Fail', Engine: 'Pass', Gear: 'Pass',
      'Overall Result': 'Failed', Remarks: 'Crack found on left wing spar fitting.'
    });
  }

  var haz = setupTab(SHEETS.HAZARDS);
  if (haz.getLastRow() < 2) {
    apiCreate('hazards', {
      Category: 'Fuel or Chemical Spillage', Severity: 'High',
      Facility: 'Lufthansa Technik Philippines', Location: 'Bay 2, near the wing stand',
      'Reported By': 'M. Reyes', 'Date Observed': '2026-11-15',
      Description: 'Hydraulic fluid puddle left unmarked on the hangar floor.',
      Status: 'Open'
    });
    apiCreate('hazards', {
      Category: 'Foreign Object Debris (FOD)', Severity: 'Medium',
      Facility: 'SIA Engineering (Philippines) Corporation', Location: 'Apron walkway B',
      'Reported By': 'K. Lim', 'Date Observed': '2026-11-12',
      Description: 'Loose bolts and rivets found near the tow path.',
      Status: 'Closed'
    });
  }

  var usr = setupTab(SHEETS.USERS);
  if (usr.getLastRow() < 2) {
    apiCreate('users', {
      Name: 'System Admin', Role: 'Administrator',
      Email: 'admin@safehangar.example', Status: 'Active'
    });
    apiCreate('users', {
      Name: 'J. Mendoza', Role: 'Inspector',
      Email: 'jmendoza@safehangar.example', Status: 'Active'
    });
    apiCreate('users', {
      Name: 'R. Santos', Role: 'Inspector',
      Email: 'rsantos@safehangar.example', Status: 'Active'
    });
  }
}
