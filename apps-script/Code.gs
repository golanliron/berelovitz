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

function _amountRange(a){
  if (a <= 0) return 'לא ידוע';
  if (a < 10000) return 'עד ₪10,000';
  if (a < 25000) return '₪10,000–25,000';
  if (a < 50000) return '₪25,000–50,000';
  return '₪50,000 ומעלה';
}
function _stats(rows){
  var total = 0, benef = 0, benefCount = 0;
  var byTrack = {}, byApproval = {}, byAmount = {}, byTeam = {};
  var statusCounts = {approved:0, sent:0, pending:0};
  rows.forEach(function(r){
    var a = parseFloat(String(r.amount).replace(/[^\d.]/g,''));
    if (isNaN(a)) a = 0;
    total += a;
    // beneficiaries is free-text ("כ-1000 צעירים") — extract the FIRST number when present.
    var braw = String(r.beneficiaries || '').replace(/,/g,'');
    var bm = braw.match(/\d+/);
    if (bm) { benef += parseInt(bm[0],10); benefCount++; }
    // breakdowns
    var t = r.track || 'ללא מסלול';
    byTrack[t] = (byTrack[t] || 0) + 1;
    var ap = String(r.approval_status || '').trim() || 'ממתין לאישור';
    byApproval[ap] = (byApproval[ap] || 0) + 1;
    var ar = _amountRange(a);
    byAmount[ar] = (byAmount[ar] || 0) + 1;
    var tm = String(r.team_status || '').trim() || 'ללא סטטוס';
    byTeam[tm] = (byTeam[tm] || 0) + 1;
    // status tiles
    var st = String(r.approval_status || '').trim();
    if (st === 'אושר') statusCounts.approved++;
    var ds = String(r.donation_sent || '').trim();
    if (ds === 'כן') statusCounts.sent++;
    if (!st || st === 'ממתין לאישור') statusCounts.pending++;
  });
  return {
    count: rows.length,
    totalAmount: total,
    totalBeneficiaries: benef,       // sum of first-number-in-each-field
    beneficiariesReported: benefCount, // how many of the orgs reported a countable number
    tracks: byTrack,
    trackCount: Object.keys(byTrack).length,
    byApproval: byApproval,
    byAmount: byAmount,
    byTeam: byTeam,
    status: statusCounts
  };
}

/* ---------- GET ---------- */

function doGet(e){
  var action = (e && e.parameter && e.parameter.action) || 'list';

  // ---- auth: בדיקה אם ת.ז מורשית להיכנס ----
  if (action === 'auth'){
    var id = String((e.parameter.id || '')).replace(/\D/g,'');
    var allowed = _isAuthorized(id);
    return _json({ ok:true, allowed: allowed });
  }
  // ---- authorized list (לניהול מהאדמין) ----
  if (action === 'authorized'){
    return _json({ ok:true, authorized: _readAuthorized() });
  }

  var data = _readAll();
  if (action === 'meta'){
    return _json({ ok:true, stats: _stats(data.rows) });
  }
  // ברירת מחדל: list — מחזיר גם את הנתונים, גם סטטיסטיקות, גם רשימת מסלולים
  return _json({
    ok: true,
    headers: data.headers,
    rows: data.rows,
    stats: _stats(data.rows),
    tracks: _readTracks(data.rows)
  });
}

/* ---------- authorized IDs (ת.ז מורשות) ---------- */
var AUTH_TAB = 'authorized';
function _authSheet(create){
  var ss = _ss();
  var a = ss.getSheetByName(AUTH_TAB);
  if (!a && create){
    a = ss.insertSheet(AUTH_TAB);
    a.appendRow(['id','name','role']);
  }
  return a;
}
function _readAuthorized(){
  var a = _authSheet(false);
  var out = [];
  if (a){
    var v = a.getDataRange().getValues();
    for (var i=1;i<v.length;i++){
      var id = String(v[i][0]).replace(/\D/g,'');
      if (!id) continue;
      out.push({ id:id, name:v[i][1]||'', role:v[i][2]||'', _rowIndex:i+1 });
    }
  }
  return out;
}
function _isAuthorized(id){
  id = String(id).replace(/\D/g,'');
  if (!id) return false;
  var list = _readAuthorized();
  for (var i=0;i<list.length;i++){ if (list[i].id === id) return true; }
  return false;
}

/* ---------- tracks (מסלולים) ---------- */
var TRACKS_TAB = 'tracks';
function _tracksSheet(create){
  var ss = _ss();
  var t = ss.getSheetByName(TRACKS_TAB);
  if (!t && create){
    t = ss.insertSheet(TRACKS_TAB);
    t.appendRow(['name','date','budget','description']);
    // seed with the 5 existing tracks so the list isn't empty
    ['עם כלביא','שאגת הארי','קול קורא 25','דוח מענק 25','השלמת פרטים'].forEach(function(n){
      t.appendRow([n,'','','']);
    });
  }
  return t;
}
// returns [{name,date,budget,description,orgCount}]
function _readTracks(orgRows){
  var counts = {};
  (orgRows||[]).forEach(function(r){ var t=r.track||''; if(t) counts[t]=(counts[t]||0)+1; });
  var t = _tracksSheet(false);
  var defined = [];
  if (t){
    var v = t.getDataRange().getValues();
    for (var i=1;i<v.length;i++){
      if (String(v[i][0]).trim()==='') continue;
      defined.push({ name:String(v[i][0]).trim(), date:v[i][1]||'', budget:v[i][2]||'', description:v[i][3]||'', _rowIndex:i+1 });
    }
  }
  var seen = {};
  var out = defined.map(function(d){ seen[d.name]=true; return {name:d.name,date:d.date,budget:d.budget,description:d.description,orgCount:counts[d.name]||0}; });
  // include any track that exists on orgs but isn't defined yet
  Object.keys(counts).forEach(function(n){ if(!seen[n]) out.push({name:n,date:'',budget:'',description:'',orgCount:counts[n]}); });
  return out;
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

  // ---- bulk import: מוסיף מספר עמותות בבת אחת (מאקסל/CSV) ----
  if (action === 'import'){
    var items = body.rows || [];
    var nextId = _nextId(data.rows);
    var added = 0;
    items.forEach(function(item){
      if (!item || !String(item.org||'').trim()) return; // דלג על שורות בלי שם עמותה
      var rid = nextId++;
      var newRow = HEADERS.map(function(h){
        if (h === 'id') return rid;
        if (h === 'source_timestamp') return item[h] || new Date();
        return (item[h] != null) ? item[h] : '';
      });
      sh.appendRow(newRow);
      added++;
    });
    _log('import', 'orgs', added + ' עמותות');
    return _json({ ok:true, added: added });
  }

  // ---- tracks: הוספה/עריכה/מחיקה של מסלול ----
  if (action === 'track_add'){
    var t = _tracksSheet(true);
    t.appendRow([body.name||'', body.date||'', body.budget||'', body.description||'']);
    _log('track_add', body.name||'', '');
    return _json({ ok:true });
  }
  if (action === 'track_update'){
    var t2 = _tracksSheet(true);
    var vv = t2.getDataRange().getValues();
    for (var i=1;i<vv.length;i++){
      if (String(vv[i][0]).trim() === String(body.origName||body.name).trim()){
        t2.getRange(i+1,1,1,4).setValues([[body.name||'', body.date||'', body.budget||'', body.description||'']]);
        _log('track_update', body.name||'', '');
        return _json({ ok:true });
      }
    }
    // לא נמצא -> הוסף
    t2.appendRow([body.name||'', body.date||'', body.budget||'', body.description||'']);
    return _json({ ok:true });
  }
  if (action === 'track_delete'){
    var t3 = _tracksSheet(true);
    var v3 = t3.getDataRange().getValues();
    for (var j=v3.length-1;j>=1;j--){
      if (String(v3[j][0]).trim() === String(body.name).trim()){ t3.deleteRow(j+1); }
    }
    _log('track_delete', body.name||'', '');
    return _json({ ok:true });
  }

  // ---- authorized IDs: הוספה/מחיקה של ת.ז מורשית ----
  if (action === 'auth_add'){
    var aid = String(body.id||'').replace(/\D/g,'');
    if (!aid) return _json({ ok:false, error:'no id' });
    var as = _authSheet(true);
    // מנע כפילות
    var existing = _readAuthorized();
    for (var k=0;k<existing.length;k++){ if (existing[k].id===aid) return _json({ ok:true, dup:true }); }
    as.appendRow([aid, body.name||'', body.role||'']);
    _log('auth_add', aid, body.name||'');
    return _json({ ok:true });
  }
  if (action === 'auth_delete'){
    var did = String(body.id||'').replace(/\D/g,'');
    var as2 = _authSheet(true);
    var v2 = as2.getDataRange().getValues();
    for (var m=v2.length-1;m>=1;m--){
      if (String(v2[m][0]).replace(/\D/g,'') === did){ as2.deleteRow(m+1); }
    }
    _log('auth_delete', did, '');
    return _json({ ok:true });
  }

  return _json({ ok:false, error:'unknown action' });
}
