/***********************************************************
 * RMS Google Sheets Database - Apps Script backend
 * ======================================================
 * ALL-IN-ONE: This single file contains the code AND the
 * appsscript.json manifest (see the MANIFEST section below).
 *
 * HOW TO USE:
 * 1. Open your Google Sheet (the one you want to use as DB)
 * 2. Menu: Extensions > Apps Script
 * 3. Delete the default "function myFunction() {}"
 * 4. Paste THIS whole file into the editor
 * 5. Click Save (Ctrl+S), name it: "RMS Database"
 * 6. Apply the manifest (ONLY ONCE):
 *    Menu: Project Settings -> scroll to "Show appsscript.json
 *    manifest file in editor" -> turn it ON. The appsscript.json
 *    file appears in the left file list. Delete its contents and
 *    paste the MANIFEST JSON block from the section below.
 *    (It must be a real .json file - remove the outer comment
 *    slash-asterisk markers before pasting)
 * 7. At the top, click "Run > setup()" and authorize the app
 *    (choose your account -> Advanced -> Go to project -> Allow)
 *    This creates/validates the "Recruitment" tab and headers.
 * 8. Click Deploy > New deployment > gear icon > "Web app"
 *      Description: RMS Database
 *      Execute as:  Me
 *      Who has access: Anyone
 *    Click Deploy (authorize again if asked)
 * 9. Copy the "Web app URL" (ends with /exec)
 * 10. Open index.html and paste it into:
 *      const _GSHEET_URL = 'YOUR_APPS_SCRIPT_WEB_APP_URL';
 * 11. Open the app, log in (admin / admin123), click the
 *     "Sync Google Sheets" button in the sidebar.
 *
 * PDF FILES:
 * - Uploaded CVs are stored TWO ways (fully automatic):
 *   1. Preferred: uploaded to a Google Drive folder "RMS CV Files"
 *      and a view link is stored in the sheet (cvDriveId /
 *      cvDriveUrl / cvDrivePreview). No base64 is written to the
 *      sheet in this case.
 *   2. Fallback: if Drive is not authorized/unavailable, the PDF
 *      base64 is stored DIRECTLY IN THE SHEET, split across the
 *      columns cvPdfData_1 .. cvPdfData_16 (a single cell can only
 *      hold 50,000 chars, so large PDFs are chunked; max ~570KB).
 *      These columns are auto-hidden so the sheet stays readable.
 *      The PDF is rebuilt automatically when records are loaded.
 * - Either way, the "cvDriveOpen" column shows a clickable link:
 *   "Open PDF" (Drive file) or "Review PDF" (built-in viewer served
 *   by this web app: ?action=viewpdf&id=<candidateId>). Clicking the
 *   link opens the PDF in the browser as an inline preview (iframe),
 *   with a "Download PDF" button for saving - no Drive authorization
 *   needed.
 * - After redeploying with this code, run once to fix existing rows:
 *   open the app and call   ?action=backfillopen   (or run
 *   backfillOpenLinks() in the editor).
 * - Deleting a candidate also trashes its PDF from Drive (if any).
 * - If you RE-DEPLOY after editing, re-run setup() first so the
 *   new columns are added to the sheet.
 * - If CV upload to Drive fails with a scope error, the record is
 *   still saved with the PDF stored in the sheet, so nothing is
 *   lost. To also enable Drive uploads, re-deploy the web app
 *   (Manage deployments -> Edit -> New version -> Deploy) so the
 *   manifest scopes (script.external_request, drive.file) are
 *   re-authorized.
 ***********************************************************/

/* ==========================================================
 * MANIFEST  (appsscript.json) - paste into the appsscript.json
 * file. Enable it under: Project Settings -> "Show appsscript.json
 * manifest file in editor".
 *
{
  "timeZone": "Asia/Phnom_Penh",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/script.external_request"
  ]
}
 * ========================================================== */

var SHEET_NAME = 'Recruitment';
var CV_FOLDER_NAME = 'RMS CV Files';

// PDF stored directly in the sheet (fallback when Drive upload is not authorized).
// A sheet cell can hold max 50,000 chars, so the base64 PDF is split into chunks,
// each chunk in its own column (cvPdfData_1 .. cvPdfData_N).
var PDF_CELL_LIMIT = 48000; // safe chars per chunk (under the 50k cell limit)
var PDF_CHUNKS = 16;        // 16 * 48k = 768k chars base64 (~570KB binary) per candidate

var HEADERS = [
  'id',
  'position',
  'department',
  'numHiring',
  'requestDate',
  'name',
  'phone',
  'email',
  'cvDate',
  'interviewDate',
  'interviewTime',
  'interviewType',
  'status',
  'stage',
  'interviewPanel',
  'interviewLocation',
  'manager',
  'cvLink',
  'remarks',
  'cvPdfName',
  'createdAt',
  'interviewHistory',
  '_roundResult',
  '_roundScore',
  '_roundFeedback',
  'cvDriveId',
  'cvDriveUrl',
  'cvDrivePreview',
  'cvDriveOpen'
];
for (var _ci = 1; _ci <= PDF_CHUNKS; _ci++) {
  HEADERS.push('cvPdfData_' + _ci);
}

/** Run this once from the Apps Script editor to prepare the sheet. */
function setup() {
  var sheet = getSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    for (var i = 0; i < HEADERS.length; i++) {
      sheet.getRange(1, i + 1).setValue(HEADERS[i]);
    }
    Logger.log('Header mismatch fixed. Columns: ' + HEADERS.length);
  }
  Logger.log('Sheet "' + SHEET_NAME + '" is ready. Rows: ' + (sheet.getLastRow() - 1));
  return 'ready';
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getCvFolder() {
  var it = DriveApp.getFoldersByName(CV_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(CV_FOLDER_NAME);
}

/** Make sure row 1 has the exact current HEADERS (self-healing after script updates). */
function ensureHeaders(sheet) {
  var current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  for (var i = 0; i < HEADERS.length; i++) {
    if (String(current[i] || '') !== HEADERS[i]) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      Logger.log('Header row updated to ' + HEADERS.length + ' columns.');
      break;
    }
  }
  try { hidePdfColumns(sheet); } catch (err) {}
}

/** Hide the raw base64 chunk columns so the sheet stays readable. */
function hidePdfColumns(sheet) {
  var ci = colIndex('cvPdfData_1');
  if (ci === -1) return;
  sheet.hideColumns(ci, PDF_CHUNKS);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** GET requests: ?action=ping  -> {ok:true}   (no params) -> array of candidates */
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'ping') {
      return json({ ok: true, service: 'RMS Google Sheets Database', time: new Date().toISOString() });
    }
    if (e && e.parameter && e.parameter.action === 'drivecheck') {
      var report = {};
      try { report.rootId = DriveApp.getRootFolder().getId(); } catch (err) { report.rootError = String(err); }
      try {
        var apiId = uploadPdfViaApi('rms-drive-check.txt', 'ok', 'text/plain');
        report.apiUpload = 'OK id=' + apiId;
        try { DriveApp.getFileById(apiId).setTrashed(true); } catch (e2) {}
      } catch (err) { report.apiError = String(err); }
      try {
        var it = DriveApp.getFoldersByName(CV_FOLDER_NAME);
        report.folderSearch = it.hasNext() ? 'found' : 'none';
      } catch (err) { report.folderError = String(err); }
      return json(report);
    }
    if (e && e.parameter && e.parameter.action === 'viewpdf') {
      return servePdfReview(Number(e.parameter.id));
    }
    if (e && e.parameter && e.parameter.action === 'pdfdata') {
      return json(fetchPdfData(Number(e.parameter.id)));
    }
    if (e && e.parameter && e.parameter.action === 'backfillopen') {
      return json({ ok: true, result: backfillOpenLinks() });
    }
    var sheet = getSheet();
    ensureHeaders(sheet);
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var rows = [];
    for (var r = 1; r < values.length; r++) {
      var obj = {};
      var hasId = false;
      var pdfParts = [];
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c]);
        var v = values[r][c];
        if (h === 'id') {
          if (v === '' || v === null || v === undefined) continue;
          v = Number(v);
          hasId = true;
        }
        if (h.indexOf('cvPdfData_') === 0) {
          if (v && v !== '') pdfParts.push(String(v));
          continue;
        }
        if ((h === 'interviewHistory') && v && v !== '') {
          try { v = JSON.parse(v); } catch (err) { v = []; }
        }
        obj[h] = v;
      }
      if (pdfParts.length) obj.cvPdfData = pdfParts.join('');
      if (hasId) rows.push(obj);
    }
    return json(rows);
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** Read the base64 PDF + file name for a candidate id straight from the sheet. */
function fetchPdfData(id) {
  var sheet = getSheet();
  ensureHeaders(sheet);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var b64 = '';
  var name = '';
  var r;
  for (r = 1; r < values.length; r++) {
    if (Number(values[r][0]) === Number(id)) break;
  }
  if (r >= values.length) return { ok: false, error: 'PDF not found (id ' + id + ')' };
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]);
    var v = values[r][c];
    if (h === 'cvPdfName' && v) name = String(v);
    if (h.indexOf('cvPdfData_') === 0 && v) b64 += String(v);
  }
  if (!b64) return { ok: false, error: 'No PDF stored for id ' + id };
  return { ok: true, data: b64, name: name };
}

/** Serve an HTML page that embeds a PDF (from the sheet's base64 chunks) for review. */
function servePdfReview(id) {
  try {
    var pdf = fetchPdfData(id);
    if (!pdf.ok) {
      return HtmlService.createHtmlOutput('<h3>' + pdf.error + '</h3>');
    }
    var b64 = pdf.data;
    var name = pdf.name;
    var title = (name || 'cv_' + id + '.pdf').replace(/"/g, '');
    var uri = 'data:application/pdf;base64,' + b64;
    var base = ScriptApp.getService().getUrl();
    var html =
      '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + title + '</title></head>' +
      '<style>' +
      'body{margin:0;font-family:Arial,sans-serif}' +
      '.bar{display:flex;align-items:center;justify-content:space-between;padding:10px 14px;background:#1a73e8;color:#fff;flex-wrap:wrap;gap:8px}' +
      '.bar strong{font-size:14px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}' +
      '.bar .actions{display:flex;gap:8px}' +
      '.bar a{color:#fff;text-decoration:none;border:1px solid #fff;border-radius:4px;padding:6px 14px;font-size:13px;white-space:nowrap}' +
      '.bar a:hover{background:rgba(255,255,255,.15)}' +
      'iframe{width:100%;height:calc(100vh - 56px);border:none}' +
      '.loading{position:absolute;top:56px;left:0;right:0;text-align:center;color:#666;padding:40px;font-size:14px}' +
      '</style>' +
      '<body>' +
      '<div class="bar"><strong>' + title + '</strong>' +
      '<div class="actions"><a id="dl" href="' + uri + '" download="' + title + '">Download PDF</a></div></div>' +
      '<div class="loading" id="loading">Loading preview...</div>' +
      '<iframe id="viewer" src=""></iframe>' +
      '<script>' +
      'var id=' + id + ';' +
      'var url="' + base + '?action=pdfdata&id="+id;' +
      'fetch(url).then(function(r){return r.json();}).then(function(d){' +
      '  if(!d || !d.ok || !d.data){document.getElementById("loading").textContent="Unable to load PDF.";return;}' +
      '  var bin=atob(d.data);' +
      '  var arr=new Uint8Array(bin.length);' +
      '  for(var i=0;i<bin.length;i++){arr[i]=bin.charCodeAt(i);}' +
      '  var blob=new Blob([arr],{type:"application/pdf"});' +
      '  var obj=URL.createObjectURL(blob);' +
      '  document.getElementById("viewer").src=obj;' +
      '  document.getElementById("dl").href=obj;' +
      '  document.getElementById("loading").style.display="none";' +
      '}).catch(function(e){document.getElementById("loading").textContent="Error: "+e.message;});' +
      '</script>' +
      '</body></html>';
    return HtmlService.createHtmlOutput(html);
  } catch (err) {
    return HtmlService.createHtmlOutput('<h3>Error: ' + String(err) + '</h3>');
  }
}

/** Fix the cvDriveOpen column for existing rows that stored their PDF as base64 in the sheet. */
function backfillOpenLinks() {
  var sheet = getSheet();
  ensureHeaders(sheet);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var ci = {};
  for (var c = 0; c < headers.length; c++) ci[String(headers[c])] = c;
  var base = ScriptApp.getService().getUrl();
  var writes = [];
  for (var r = 1; r < values.length; r++) {
    var hasB64 = false;
    for (var cc = 0; cc < headers.length; cc++) {
      if (String(headers[cc]).indexOf('cvPdfData_') === 0 && values[r][cc]) { hasB64 = true; break; }
    }
    var hasDrive = values[r][ci['cvDriveUrl']] ? true : false;
    var formula;
    if (hasDrive) {
      formula = '=HYPERLINK("' + values[r][ci['cvDriveUrl']] + '","Open PDF")';
    } else if (hasB64) {
      formula = '=HYPERLINK("' + base + '?action=viewpdf&id=' + values[r][ci['id']] + '","Review PDF")';
    } else {
      formula = '';
    }
    writes.push([r + 1, formula]);
  }
  if (writes.length) {
    var range = sheet.getRange(2, ci['cvDriveOpen'] + 1, writes.length, 1);
    var out = writes.map(function (w) { return [w[1]]; });
    range.setValues(out);
  }
  return 'fixed ' + writes.length + ' rows';
}

/** POST requests: body = { action:'add'|'update'|'delete', data:{...} } */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(20000);
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var data = body.data || {};
    var sheet = getSheet();
    ensureHeaders(sheet);
    if (action === 'add' || action === 'update') {
      data.id = Number(data.id);
      var pdf = processPdf(data);
      var cvError = '';
      if (pdf && pdf.error) {
        // Drive upload not possible -> store the PDF directly in the sheet
        var chunks = splitPdfData(data.cvPdfData || '');
        if (chunks.length <= PDF_CHUNKS) {
          cvError = '';
          Logger.log('Drive upload skipped for ' + data.id + ', PDF stored in sheet columns (' + chunks.length + ' chunks).');
        } else {
          cvError = 'PDF too large to store in the sheet. It stays in your browser storage only: ' + pdf.error;
          delete data.cvPdfData;
        }
        pdf = null;
      }
      if (pdf) {
        data.cvDriveId = pdf.id;
        data.cvDriveUrl = pdf.url;
        data.cvDrivePreview = pdf.preview;
        delete data.cvPdfData; // PDF is on Drive, do not bloat the sheet with base64
      }
      if (action === 'add') {
        appendRecord(sheet, data);
      } else {
        updateRecord(sheet, data);
      }
      return json({ ok: true, id: data.id, cvError: cvError, cvDriveId: data.cvDriveId || '', cvDriveUrl: data.cvDriveUrl || '', cvDrivePreview: data.cvDrivePreview || '' });
    }
    if (action === 'delete') {
      deleteRecord(sheet, data.id);
      return json({ ok: true });
    }
    return json({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * If the incoming record has cvPdfData (base64) and has NOT been
 * uploaded yet (no cvDriveId), upload it to Drive and return the
 * file links. Returns null when there is nothing to upload.
 *
 * Uploads via the Drive REST API (files.create) so it only needs the
 * narrow "drive.file" scope.
 */
function processPdf(data) {
  if (data.cvDriveId && data.cvDriveId !== '') return null;
  if (!data.cvPdfData || data.cvPdfData === '') return null;
  try {
    var name = data.cvPdfName || ('cv_' + data.id + '.pdf');
    if (!/\.pdf$/i.test(name)) name = name + '.pdf';
    var fileId = uploadPdfViaApi(name, data.cvPdfData, 'application/pdf');
    return {
      id: fileId,
      url: 'https://drive.google.com/file/d/' + fileId + '/view',
      preview: 'https://drive.google.com/file/d/' + fileId + '/preview'
    };
  } catch (err) {
    Logger.log('processPdf error: ' + err);
    return { error: String(err) };
  }
}

/** Upload bytes (base64 string) to Drive via the REST API files.create (works with drive.file scope). */
function uploadPdfViaApi(name, base64Data, mimeType) {
  var token = ScriptApp.getOAuthToken();
  var boundary = 'rms_' + Date.now();
  var meta = JSON.stringify({ name: name, mimeType: mimeType || 'application/pdf' });
  var head = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta + '\r\n--' + boundary + '\r\nContent-Type: ' + (mimeType || 'application/pdf') + '\r\n\r\n';
  var tail = '\r\n--' + boundary + '--';
  var headBytes = Utilities.newBlob(head).getBytes();
  var dataBytes = Utilities.base64Decode(base64Data);
  var tailBytes = Utilities.newBlob(tail).getBytes();
  var payload = headBytes.concat(dataBytes).concat(tailBytes);
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'multipart/related; boundary=' + boundary
    },
    payload: payload,
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code !== 200 && code !== 201) {
    throw new Error('Drive API upload failed (' + code + '): ' + body.slice(0, 300));
  }
  var data = JSON.parse(body);
  return data.id;
}

function colIndex(name) {
  var i = HEADERS.indexOf(name);
  return i === -1 ? -1 : i + 1;
}

function findRow(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (Number(ids[i][0]) === Number(id)) return i + 2; // 1-based sheet row
  }
  return -1;
}

/** Split a base64 string into chunks that fit inside a sheet cell (50k limit). */
function splitPdfData(b64) {
  var chunks = [];
  if (!b64) return chunks;
  for (var i = 0; i < b64.length; i += PDF_CELL_LIMIT) {
    chunks.push(b64.substring(i, i + PDF_CELL_LIMIT));
  }
  return chunks;
}

function serializeRow(data) {
  var chunks = splitPdfData(data.cvPdfData || '');
  var row = [];
  for (var i = 0; i < HEADERS.length; i++) {
    var h = HEADERS[i];
    var v;
    if (h.indexOf('cvPdfData_') === 0) {
      var idx = parseInt(h.split('_')[1], 10); // 1-based chunk number
      v = chunks[idx - 1] || '';
    } else {
      v = data[h];
      if (v === undefined || v === null) v = '';
      if (h === 'interviewHistory' && typeof v === 'object') v = JSON.stringify(v);
      if (h === 'cvDriveOpen') {
        if (data.cvDriveUrl && data.cvDriveUrl !== '') {
          v = '=HYPERLINK("' + data.cvDriveUrl + '","Open PDF")';
        } else if (data.cvPdfData && data.cvPdfData !== '') {
          v = '=HYPERLINK("' + ScriptApp.getService().getUrl() + '?action=viewpdf&id=' + data.id + '","Review PDF")';
        } else {
          v = '';
        }
      }
    }
    row.push(v);
  }
  return row;
}

function trashDriveFile(fileId) {
  if (!fileId) return;
  try {
    var token = ScriptApp.getOAuthToken();
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId), {
      method: 'patch',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ trashed: true }),
      muteHttpExceptions: true
    });
  } catch (err) {}
}

function appendRecord(sheet, data) {
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, HEADERS.length).setValues([serializeRow(data)]);
}

function updateRecord(sheet, data) {
  var r = findRow(sheet, data.id);
  if (r === -1) {
    appendRecord(sheet, data);
  } else {
    var ci = colIndex('cvDriveId');
    var oldId = (ci !== -1) ? sheet.getRange(r, ci).getValue() : '';
    if (oldId && data.cvDriveId && String(oldId) !== String(data.cvDriveId)) {
      trashDriveFile(oldId);
    }
    sheet.getRange(r, 1, 1, HEADERS.length).setValues([serializeRow(data)]);
  }
}

function deleteRecord(sheet, id) {
  var r = findRow(sheet, id);
  if (r === -1) return;
  var ci = colIndex('cvDriveId');
  if (ci !== -1) {
    var v = sheet.getRange(r, ci).getValue();
    if (v) trashDriveFile(v);
  }
  sheet.deleteRow(r);
}
