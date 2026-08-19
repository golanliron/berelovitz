/**
 * קרן ברלוביץ׳ — API לניהול מאגר העמותות ומדידת האימפקט
 * Google Apps Script Web App מחובר לגיליון:
 *   "קרן ברלוביץ׳ — מאגר עמותות ומדידת אימפקט"
 *   ID: 1kGjCardDU9b5vcJjCJ6s8mGkkniXM32-Xxk_UDFrrWE
 *
 * נותן API מלא (קריאה + כתיבה) לאתר app.html:
 *   GET  ?action=list           → כל העמותות (JSON)
 *   GET  ?action=meta           → סטטיסטיקות לדשבורד
 *   POST {action:'add', row}    → הוספת עמותה
 *   POST {action:'update', id, row} → עריכת עמותה
 *   POST {action:'delete', id}  → מחיקת עמותה
 *   POST {action:'log_read}     → (פנימי)
 *
 * פריסה: Extensions → Apps Script → Deploy → New deployment →
 *   Web app → Execute as: Me, Who has access: Anyone → Deploy.
 */

var SHEET_ID = '1kGjCardDU9b5vcJjCJ6s8mGkkniXM32-Xxk_UDFrrWE';
var DATA_TAB = 'Sheet1';   // שם לשונית הנתונים (ברירת מחדל של גיליון חדש)
var LOG_TAB  = 'activity_log';

// עמודות המאגר (חייב להתאים לשורת הכותרת בגיליון)
var HEADERS = [
  'id','track','org','project','contact','email','phone','amount',
  'beneficiaries','what_for','problem','results','measurement','budget_use','notes',
  'source_timestamp',
  'approval_status','donation_sent','donation_date','email_sent','email_date',
  'team_status','team_notes'
];

/* ---------- helpers ---------- */

function _ss()   { return SpreadsheetApp.openById(SHEET_ID); }
function _sheet(){
  var ss = _ss();
  var sh = ss.getSheetByName(DATA_TAB);
  if (!sh) sh = ss.getSheets()[0]; // fallback ללשונית הראשונה
  return sh;
}
function _json(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _readAll(){
  var sh = _sheet();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return { headers: HEADERS, rows: [] };
  var head = values[0].map(function(h){ return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < values.length; i++){
    var r = values[i];
    // דלג על שורות ריקות לגמרי
    if (r.join('').trim() === '') continue;
    var obj = {};
    for (var c = 0; c < head.length; c++){
      obj[head[c]] = r[c];
    }
    obj._rowIndex = i + 1; // מספר השורה בגיליון (1-based)
    rows.push(obj);
  }
  return { headers: head, rows: rows };
}

function _log(action, entity, detail){
  try {
    var ss = _ss();
    var lg = ss.getSheetByName(LOG_TAB);
    if (!lg){
      lg = ss.insertSheet(LOG_TAB);
      lg.appendRow(['timestamp','action','entity','detail']);
    }
    lg.appendRow([new Date(), action, entity || '', detail || '']);
  } catch(e) { /* לוג לא קריטי */ }
}

function _nextId(rows){
  var max = 0;
  rows.forEach(function(r){
    var n = parseInt(r.id, 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

function _stats(rows){
  var total = 0, benef = 0, byTrack = {}, statusCounts = {approved:0, sent:0, pending:0};
  rows.forEach(function(r){
    var a = parseFloat(String(r.amount).replace(/[^\d.]/g,''));
    if (!isNaN(a)) total += a;
    var b = parseFloat(String(r.beneficiaries).replace(/[^\d.]/g,''));
    if (!isNaN(b)) benef += b;
    var t = r.track || 'ללא מסלול';
    byTrack[t] = (byTrack[t] || 0) + 1;
    var st = String(r.approval_status || '').trim();
    if (st === 'אושר') statusCounts.approved++;
    var ds = String(r.donation_sent || '').trim();
    if (ds === 'כן') statusCounts.sent++;
    if (!st || st === 'ממתין לאישור') statusCounts.pending++;
  });
  return {
    count: rows.length,
    totalAmount: total,
    totalBeneficiaries: benef,
    tracks: byTrack,
    trackCount: Object.keys(byTrack).length,
    status: statusCounts
  };
}

/* ---------- GET ---------- */

function doGet(e){
  var action = (e && e.parameter && e.parameter.action) || 'list';
  var data = _readAll();
  if (action === 'meta'){
    return _json({ ok:true, stats: _stats(data.rows) });
  }
  // ברירת מחדל: list — מחזיר גם את הנתונים וגם סטטיסטיקות
  return _json({
    ok: true,
    headers: data.headers,
    rows: data.rows,
    stats: _stats(data.rows)
  });
}

/* ---------- POST ---------- */

function doPost(e){
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err){
    return _json({ ok:false, error:'bad json' });
  }
  var action = body.action;
  var sh = _sheet();
  var data = _readAll();

  if (action === 'add'){
    var id = _nextId(data.rows);
    var row = HEADERS.map(function(h){
      if (h === 'id') return id;
      if (h === 'source_timestamp' && !body.row[h]) return new Date();
      return (body.row && body.row[h] != null) ? body.row[h] : '';
    });
    sh.appendRow(row);
    _log('add', 'org#' + id, (body.row && body.row.org) || '');
    return _json({ ok:true, id: id });
  }

  if (action === 'update'){
    var target = null;
    data.rows.forEach(function(r){ if (String(r.id) === String(body.id)) target = r; });
    if (!target) return _json({ ok:false, error:'not found' });
    var head = data.headers;
    for (var c = 0; c < head.length; c++){
      var key = head[c];
      if (key === 'id') continue;
      if (body.row && body.row.hasOwnProperty(key)){
        sh.getRange(target._rowIndex, c + 1).setValue(body.row[key]);
      }
    }
    _log('update', 'org#' + body.id, (body.row && body.row.org) || '');
    return _json({ ok:true });
  }

  if (action === 'delete'){
    var tgt = null;
    data.rows.forEach(function(r){ if (String(r.id) === String(body.id)) tgt = r; });
    if (!tgt) return _json({ ok:false, error:'not found' });
    sh.deleteRow(tgt._rowIndex);
    _log('delete', 'org#' + body.id, (tgt.org) || '');
    return _json({ ok:true });
  }

  return _json({ ok:false, error:'unknown action' });
}
