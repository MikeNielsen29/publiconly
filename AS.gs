/**
 * ============================================================
 * MIPORO ICT Solution — Google Sheets Database Companion
 * google-apps-script.gs
 *
 * GINAGAWA:
 *  - Tumatanggap ng full data snapshot mula sa MIPORO ICT Solution web app
 *    (Settings → Push to Sheets) at nagsusulat sa mga sheets.
 *  - Nagbabalik ng buong data kapag may GET request
 *    (Settings → Pull from Sheets).
 *
 * SETUP (isang beses lang):
 *  1. Gumawa ng bagong Google Sheet (hal. "MIPORO ICT DB").
 *  2. Extensions → Apps Script — i-paste ang BUONG file na ito.
 *  3. Palitan ang SECRET sa ibaba ng sarili mong passphrase.
 *  4. Deploy → New deployment → Type: Web app
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     Kopyahin ang Web App URL (https://script.google.com/macros/s/…/exec)
 *  5. Sa MIPORO ICT Solution → Settings → i-paste ang URL at Secret.
 * ============================================================
 */

var SECRET = 'mike201422'; // ⚠️ PALITAN ITO!

/* Sheet name ↔ data key mapping (naka-align sa AS.allData()) */
var KEYS = [
  'business', 'accounts', 'customers', 'suppliers', 'items',
  'employees', 'sales', 'receipts', 'purchases', 'expenses',
  'gj', 'payrollRuns', 'invAdjustments', 'opening', 'taxOverrides', 'settings'
];

function doPost(e){
  try{
    var raw = (e && e.postData && e.postData.contents) || '';
    if(!raw) return _json({ ok: false, error: 'empty body: nag-close agad ang receiving — siguraduhing POST ang gamit at may laman ang body' });
    var body;
    try{ body = JSON.parse(raw); }
    catch(pe){ return _json({ ok: false, error: 'invalid JSON: ' + pe.message }); }
    if(String(body.secret) !== String(SECRET)) return _json({ ok: false, error: 'bad secret' });
    saveAll_(body.data || {});
    return _json({ ok: true, msg: 'saved ' + new Date().toISOString() });
  }catch(err){
    return _json({ ok: false, error: err.message });
  }
}

function doGet(e){
  try{
    if(String(e.parameter.secret) !== String(SECRET)) return _json({ ok: false, error: 'bad secret' });
    return _json(loadAll_());
  }catch(err){
    return _json({ ok: false, error: err.message });
  }
}

/* ---------------- internals ---------------- */
function saveAll_(data){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  KEYS.forEach(function(k){
    var rows = data[k];
    if(rows === undefined || rows === null) return;
    if(!Array.isArray(rows)) rows = [rows];
    // Alisin ang mga empty objects (hal. [{}]) para hindi magka-0 columns
    rows = rows.filter(function(r){
      if(r === null || r === undefined) return false;
      if(typeof r !== 'object') return true;
      return Object.keys(r).some(function(c){
        var v = r[c];
        return v !== null && v !== undefined && String(v) !== '';
      });
    });
    var sh = ss.getSheetByName(k) || ss.insertSheet(k);
    sh.clearContents();
    if(!rows.length){ sh.getRange(1, 1).setValue('(empty)'); return; }
    /* normalize rows: objects → arrays of columns */
    var cols = [];
    rows.forEach(function(r){
      if(r !== null && typeof r === 'object'){
        Object.keys(r).forEach(function(c){ if(cols.indexOf(c) < 0) cols.push(c); });
      } else if(cols.indexOf('value') < 0){
        cols.push('value'); // primitive rows (hal. opening lines)
      }
    });
    if(!cols.length){ sh.getRange(1, 1).setValue('(empty)'); return; }
    var matrix = rows.map(function(r){
      return cols.map(function(c){
        var v = (r !== null && typeof r === 'object') ? (c === 'value' && typeof r !== 'object' ? r : r[c]) : (c === 'value' ? r : '');
        if(typeof r !== 'object' && cols.length === 1) v = r;
        if(v === null || v === undefined) return '';
        if(typeof v === 'object') return JSON.stringify(v);
        return v;
      });
    });
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    if(matrix.length) sh.getRange(2, 1, matrix.length, cols.length).setValues(matrix);
  });
}

function loadAll_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = { ok: true };
  KEYS.forEach(function(k){
    var sh = ss.getSheetByName(k);
    if(!sh){ out[k] = []; return; }
    var vals = sh.getDataRange().getValues();
    if(!vals.length || String(vals[0][0]) === '(empty)'){ out[k] = []; return; }
    var head = vals[0];
    var rows = [];
    for(var i = 1; i < vals.length; i++){
      var o = {};
      var hasAny = false;
      for(var j = 0; j < head.length; j++){
        var v = vals[i][j];
        if(v === '' || v === null) continue;
        hasAny = true;
        var s = String(v);
        /* restore objects/arrays saved as JSON strings */
        if((s[0] === '{' && s[s.length-1] === '}') || (s[0] === '[' && s[s.length-1] === ']')){
          try{ v = JSON.parse(s); }catch(e2){ v = s; }
        }
        o[head[j]] = v;
      }
      if(hasAny) rows.push(o);
    }
    out[k] = rows;
  });
  return out;
}

function _txt(s){ return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); }
function _json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }