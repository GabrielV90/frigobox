/*************************************************************************
 * CELLE FRIGO · CONTROLLO ORDINI — Backend Google Apps Script
 * (Krio · con supporto cartelle organizzative appFolderId)
 *
 * COSA FA
 *  - Verifica il PIN di accesso (Proprietà script "PIN", default 1234)
 *  - Salva gli ordini in un foglio Google (creato in automatico)
 *  - Salva foto / documenti / video su Drive, una cartella per ordine
 *  - Cartelle organizzative dell'app (appFolderId + settings.folders)
 *  - Restituisce sempre JSON, tutto resta online e accessibile ovunque
 *
 * COME PUBBLICARE (una volta sola)
 *  1) script.google.com → Nuovo progetto → incolla questo file
 *  2) In alto a sinistra: nome progetto (es. "Celle Frigo API")
 *  3) Menu ⚙ Impostazioni progetto → NON serve nulla di particolare
 *  4) (facoltativo) Imposta il PIN:
 *     Impostazioni progetto → Proprietà script → Aggiungi proprietà
 *     Nome: PIN   Valore: il tuo codice (es. 7391)
 *  5) Distribuisci → Nuova distribuzione → tipo "App web"
 *       Esegui come: Me stesso
 *       Chi ha accesso: Chiunque
 *     → Distribuisci → autorizza → COPIA l'URL che finisce con /exec
 *  6) Incolla quell'URL nelle Impostazioni (⚙︎) dell'app.
 *
 * Ad ogni modifica del codice: Distribuisci → Gestisci distribuzioni →
 * matita ✎ → Versione: Nuova → Distribuisci (l'URL resta lo stesso).
 *************************************************************************/

var ROOT_FOLDER_NAME = 'CELLE FRIGO — ORDINI';   // cartella Drive radice
var SHEET_NAME       = 'Ordini';
// folderId     = cartella Drive dei media (foto/documenti)
// appFolderId  = cartella organizzativa dell'app (es. "2026")
var HEADERS = ['id','codice','cliente','targa','tipoCella','dataOrdine',
               'scadenza','stato','optional','note','folderId','appFolderId','updatedAt',
               'budgetOre','oreEffettive','readyAt','fasi','operatore','updatedBy','clientUpdatedAt'];

/* ---------- entry points ---------- */
function doPost(e){
  try{
    var req = JSON.parse(e.postData.contents || '{}');
    return handle(req);
  }catch(err){
    return out({ok:false, error:String(err)});
  }
}
function doGet(e){
  // supporto GET per test rapido dal browser
  var p = (e && e.parameter) ? e.parameter : {};
  if(!p.action) return out({ok:true, service:'Celle Frigo API', status:'online'});
  return handle(p);
}

var WRITE_ACTIONS = {save:1, 'delete':1, upload:1, deleteFile:1, setSettings:1};

function handle(req){
  // Ruolo in base al codice: 'edit' (modifica) o 'view' (sola lettura)
  var role = roleFor(req.pin);
  if(!role) return out({ok:false, code:'AUTH', error:'Codice errato'});
  if(WRITE_ACTIONS[req.action] && role !== 'edit')
    return out({ok:false, code:'READONLY', error:'Accesso in sola lettura'});

  switch(req.action){
    case 'auth':        return out({ok:true, role:role});
    case 'list':        return out({ok:true, orders:listOrders(), settings:getSettings()});
    case 'getSettings': return out({ok:true, settings:getSettings()});
    case 'setSettings': return out({ok:true, settings:putSettings(req.settings)});
    case 'save':        return out({ok:true, order:saveOrder(req.order)});
    case 'delete':      return out({ok:true}, deleteOrder(req.id));
    case 'files':       return out({ok:true, files:listFiles(req)});
    case 'upload':      return out({ok:true, file:uploadFile(req)});
    case 'deleteFile':  return out({ok:true}, deleteFile(req.fileId));
    default:            return out({ok:false, error:'Azione sconosciuta: '+req.action});
  }
}

/* ---------- impostazioni condivise (organico + cartelle app) ---------- */
function getSettings(){
  var raw = PropertiesService.getScriptProperties().getProperty('SETTINGS');
  try{ return raw ? JSON.parse(raw) : {}; }catch(_){ return {}; }
}
function putSettings(s){
  // merge con settings esistenti per non perdere chiavi non inviate
  var prev = getSettings();
  var obj = s || {};
  Object.keys(obj).forEach(function(k){ prev[k] = obj[k]; });
  PropertiesService.getScriptProperties().setProperty('SETTINGS', JSON.stringify(prev));
  return prev;
}

/* ---------- auth ---------- */
// Proprietà script:
//   PIN       = codice di chi può MODIFICARE (default 1234)
//   PIN_VIEW  = codice per la SOLA LETTURA (lascia vuoto per disattivarlo)
function roleFor(pin){
  var props = PropertiesService.getScriptProperties();
  var edit = props.getProperty('PIN') || '1990';
  var view = props.getProperty('PIN_VIEW') || '0000';
  pin = String(pin||'');
  if(pin && pin === String(edit)) return 'edit';
  if(view && pin === String(view)) return 'view';
  return null;
}

/* ---------- sheet helpers ---------- */
function getSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if(!ss){
    // se lo script è standalone crea/riusa un foglio dedicato
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('SHEET_ID');
    if(id){ ss = SpreadsheetApp.openById(id); }
    else{
      ss = SpreadsheetApp.create('Celle Frigo — Database ordini');
      props.setProperty('SHEET_ID', ss.getId());
    }
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  if(!sh){
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  if(sh.getLastRow()===0) sh.appendRow(HEADERS);
  ensureHeaders(sh);
  return sh;
}
// aggiunge in coda eventuali colonne mancanti (migrazione fogli già esistenti)
function ensureHeaders(sh){
  var lastCol = sh.getLastColumn();
  var existing = lastCol>0 ? sh.getRange(1,1,1,lastCol).getValues()[0] : [];
  var changed=false;
  HEADERS.forEach(function(h){ if(existing.indexOf(h)===-1){ existing.push(h); changed=true; } });
  if(changed) sh.getRange(1,1,1,existing.length).setValues([existing]);
}
function rowsToObjects(sh){
  var data = sh.getDataRange().getValues();
  var head = data.shift();
  return data.map(function(r){
    var o={}; head.forEach(function(h,i){ o[h]=r[i]; });
    try{ o.optional = o.optional ? JSON.parse(o.optional) : []; }catch(_){ o.optional=[]; }
    try{ o.budget = o.budgetOre ? JSON.parse(o.budgetOre) : {}; }catch(_){ o.budget={}; }
    try{ o.actual = o.oreEffettive ? JSON.parse(o.oreEffettive) : {}; }catch(_){ o.actual={}; }
    try{ o.fasi = o.fasi ? JSON.parse(o.fasi) : []; }catch(_){ o.fasi=[]; }
    o.operatore = o.operatore ? String(o.operatore) : '';
    o.updatedBy = o.updatedBy ? String(o.updatedBy) : '';
    o.clientUpdatedAt = o.clientUpdatedAt ? String(o.clientUpdatedAt) : '';
    // normalizza date (Sheet può restituire Date)
    ['dataOrdine','scadenza','readyAt'].forEach(function(k){
      if(o[k] instanceof Date) o[k]=Utilities.formatDate(o[k], Session.getScriptTimeZone(),'yyyy-MM-dd');
      else o[k]=o[k]?String(o[k]):'';
    });
    // normalizza appFolderId (cartella organizzativa app)
    o.appFolderId = o.appFolderId ? String(o.appFolderId) : '';
    if(!o.appFolderId) o.appFolderId = null;
    o.mediaCount = o.folderId ? folderCount(o.folderId) : 0;
    return o;
  });
}
function findRow(sh, id){
  var last = sh.getLastRow();
  if(last < 2) return -1;                 // solo intestazione: nessun ordine ancora
  var ids = sh.getRange(2,1,last-1,1).getValues();
  for(var i=0;i<ids.length;i++){ if(String(ids[i][0])===String(id)) return i+2; }
  return -1;
}

/* ---------- orders CRUD ---------- */
function listOrders(){ return rowsToObjects(getSheet()); }

function saveOrder(o){
  var sh = getSheet();
  o.updatedAt = new Date().toISOString();
  if((o.stato==='pronto' || o.stato==='cons') && !o.readyAt)
    o.readyAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  // assicura la cartella Drive dell'ordine, annidata sotto il cliente
  var existing = findOrderById(o.id);
  var known = existing ? existing.folderId : '';
  var driveFolderId = ensureOrderFolder(o.codice || o.id, o.id, o.cliente, known).getId();
  // cartella organizzativa app (NON confondere con Drive)
  // accetta sia appFolderId sia folderId dal client se non è un id Drive
  var appFolderId = '';
  if(o.appFolderId != null && o.appFolderId !== '') appFolderId = String(o.appFolderId);
  else if(o.folderId != null && o.folderId !== '' && String(o.folderId).charAt(0) === 'F')
    appFolderId = String(o.folderId); // id cartelle app generati dal frontend (prefisso F)
  else if(existing && existing.appFolderId) appFolderId = String(existing.appFolderId);
  var row = HEADERS.map(function(h){
    if(h==='optional')     return JSON.stringify(o.optional||[]);
    if(h==='budgetOre')    return JSON.stringify(o.budget||{});
    if(h==='oreEffettive') return JSON.stringify(o.actual||{});
    if(h==='fasi')         return JSON.stringify(o.fasi||[]);
    if(h==='folderId')     return driveFolderId;
    if(h==='appFolderId')  return appFolderId;
    return o[h]!=null ? o[h] : '';
  });
  var r = findRow(sh, o.id);
  if(r>0) sh.getRange(r,1,1,HEADERS.length).setValues([row]);
  else    sh.appendRow(row);
  o.folderId = driveFolderId;
  o.appFolderId = appFolderId || null;
  o.mediaCount = folderCount(driveFolderId);
  return o;
}
function deleteOrder(id){
  var sh = getSheet();
  var r = findRow(sh, id);
  if(r>0) sh.deleteRow(r);
  // NB: la cartella Drive con i file NON viene eliminata (sicurezza)
}

/* ---------- Drive helpers ---------- */
function getRoot(){
  var it = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
}
function ensureChildFolder(parent, name){
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
// cerca la cartella di un ordine (marcatore [id]) sia nella vecchia posizione piatta
// sia dentro le cartelle cliente
function findOrderFolder(root, id){
  var marker = '['+id+']';
  var it = root.getFolders();
  while(it.hasNext()){
    var f = it.next();
    if(f.getName().indexOf(marker) > -1) return f;   // vecchia struttura: ordine sotto la radice
    var it2 = f.getFolders();                         // dentro una cartella cliente
    while(it2.hasNext()){ var g = it2.next(); if(g.getName().indexOf(marker) > -1) return g; }
  }
  return null;
}
// Struttura Drive: RADICE / <Cliente> / <codice [id]> / (foto, documenti, video)
function ensureOrderFolder(codice, id, cliente, knownFolderId){
  var root = getRoot();
  var clientFolder = ensureChildFolder(root, sanitize(cliente || 'Senza cliente'));
  var wanted = sanitize(codice || id) + ' [' + id + ']';
  var folder = null;
  if(knownFolderId){ folder = folderById(knownFolderId); }
  if(!folder){ folder = findOrderFolder(root, id); }
  if(folder){
    // sposta sotto il cliente corretto se necessario
    try{
      var pit = folder.getParents();
      var pid = pit.hasNext() ? pit.next().getId() : null;
      if(pid !== clientFolder.getId()) folder.moveTo(clientFolder);
    }catch(_){}
    if(folder.getName() !== wanted){ try{ folder.setName(wanted); }catch(_){} }
    return folder;
  }
  return clientFolder.createFolder(wanted);
}
function folderById(id){
  try{ return DriveApp.getFolderById(id); }catch(_){ return null; }
}
function folderCount(id){
  var f = folderById(id); if(!f) return 0;
  var n=0, it=f.getFiles(); while(it.hasNext()){ it.next(); n++; }
  return n;
}
function listFiles(req){
  var f = null;
  // usa folderId se lo conosciamo dall'ordine, altrimenti risali da id/codice
  var order = findOrderById(req.id);
  if(order && order.folderId) f = folderById(order.folderId);
  if(!f) f = ensureOrderFolder(req.codice, req.id, req.cliente, order && order.folderId);
  var out=[], it=f.getFiles();
  while(it.hasNext()){
    var file=it.next();
    out.push({
      id:file.getId(),
      name:file.getName(),
      mimeType:file.getMimeType(),
      url:'https://drive.google.com/uc?export=view&id='+file.getId(),
      thumb:'https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w400',
      updated:file.getLastUpdated().toISOString()
    });
  }
  // più recenti prima
  out.sort(function(a,b){ return a.updated<b.updated?1:-1; });
  return out;
}
function findOrderById(id){
  var sh=getSheet(), r=findRow(sh,id);
  if(r<0) return null;
  // usa header reali del foglio (può avere colonne in più dopo migrazione)
  var lastCol = sh.getLastColumn();
  var head = sh.getRange(1,1,1,lastCol).getValues()[0];
  var vals=sh.getRange(r,1,1,lastCol).getValues()[0];
  var o={}; head.forEach(function(h,i){o[h]=vals[i];}); return o;
}
function uploadFile(req){
  var order = findOrderById(req.id);
  var folder = (order && order.folderId) ? (folderById(order.folderId)||ensureOrderFolder(req.codice,req.id,req.cliente,order.folderId))
                                         : ensureOrderFolder(req.codice, req.id, req.cliente);
  var bytes = Utilities.base64Decode(req.data);
  var blob = Utilities.newBlob(bytes, req.mimeType, req.name);
  var file = folder.createFile(blob);
  try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(_){}
  return {
    id:file.getId(), name:file.getName(), mimeType:file.getMimeType(),
    url:'https://drive.google.com/uc?export=view&id='+file.getId(),
    thumb:'https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w400'
  };
}
function deleteFile(fileId){
  try{ DriveApp.getFileById(fileId).setTrashed(true); }catch(_){}
}

/* ---------- utils ---------- */
function sanitize(s){ return String(s||'').replace(/[\\/:*?"<>|\[\]]/g,'-').trim() || 'ordine'; }
function out(obj, _){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
