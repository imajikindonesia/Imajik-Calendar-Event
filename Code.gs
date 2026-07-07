const PARENT_FOLDER_ID = '1cdRJ6oB0jcmzIDgS5bygQfRmxj-yfkI4'; // folder 09_IMAJIK_Calender
const DB_FILE_NAME = 'IMAJIK Events Database';
const SHEET_NAME = 'Events';

// Kolom internal (mapping by index ke spreadsheet kolom A–S)
const COLUMNS = [
  'id','date','title','clientName','category','time',
  'location','pic','picPhone',
  'quotSent','invoiceSent','paymentChecked','notes',
  'feeTotal','dpAmount','dpDate','pelunasanAmount','pelunasanDate',
  'briefFolderUrl','dateEnd'
];

// Header tampilan untuk spreadsheet (urutan sama dengan COLUMNS)
const HEADERS = [
  'ID','DATE EVENT','NAMA EVENT','COMPANY','KATEGORI','JAM STANDBY',
  'LOKASI EVENT','PIC NAME','PIC Phone',
  'QUOT','INVOICE','PAYMENT','CATATAN',
  'JUMLAH NOMINAL','Downpayment','DP - Tanggal','Pelunasan','Lunas - Date',
  'Upload Brief','DATE END'
];

// ================= WEB APP HANDLERS =================

function doGet(e) {
  const action = e.parameter.action || 'listEvents';
  if (action === 'listEvents') {
    return jsonResponse({ success: true, events: listEvents() });
  }
  return jsonResponse({ success: false, error: 'Unknown action' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'uploadFiles') return jsonResponse(uploadFiles(body));
    if (body.action === 'saveEvent') return jsonResponse({ success: true, event: saveEvent(body.event) });
    if (body.action === 'deleteEvent') { deleteEvent(body.id, body.briefFolderUrl); return jsonResponse({ success: true }); }
    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ================= GOOGLE DRIVE =================

function getParentFolder() {
  return DriveApp.getFolderById(PARENT_FOLDER_ID);
}

function getOrCreateFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

// ================= UPLOAD FILES (Brief -> Google Drive) =================

function uploadFiles(body) {
  const parent = getParentFolder();
  const folderName = 'Brief - ' + (body.eventId || 'Tanpa-ID').trim();
  const briefFolder = getOrCreateFolder(parent, folderName);

  const links = (body.files || []).map(f => {
    const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mimeType, f.name);
    const file = briefFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { name: f.name, url: file.getUrl() };
  });

  return { success: true, folderUrl: briefFolder.getUrl(), files: links };
}

// ================= SPREADSHEET DATABASE =================

function getSheet() {
  const parent = getParentFolder();
  const files = parent.getFilesByName(DB_FILE_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(DB_FILE_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(parent);
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); sheet.appendRow(HEADERS); }
  return sheet;
}

// Helper: konversi Date object dari Sheets ke string format yyyy-MM-dd
function toDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

// Helper: konversi Date/time object dari Sheets ke string format HH:mm
function toTimeStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  }
  return String(val);
}

function rowToEvent(row) {
  const obj = {};
  COLUMNS.forEach((col, i) => { obj[col] = row[i]; });

  // Konversi tanggal & waktu dari Date object Sheets ke string
  obj.date = toDateStr(obj.date);
  obj.dateEnd = toDateStr(obj.dateEnd);
  obj.time = toTimeStr(obj.time);
  obj.dpDate = toDateStr(obj.dpDate);
  obj.pelunasanDate = toDateStr(obj.pelunasanDate);

  // Field yang tidak ada di spreadsheet — set default
  obj.locationUrl = '';
  obj.briefFiles = [];

  // String fields
  obj.category = obj.category || 'other';
  obj.location = obj.location || '';
  obj.briefFolderUrl = obj.briefFolderUrl || '';

  // Status: "Terkirim" → true, "Done" → true, lainnya → false
  obj.quotSent = obj.quotSent === true || obj.quotSent === 'TRUE' || obj.quotSent === 'Terkirim';
  obj.invoiceSent = obj.invoiceSent === true || obj.invoiceSent === 'TRUE' || obj.invoiceSent === 'Terkirim';
  obj.paymentChecked = obj.paymentChecked === true || obj.paymentChecked === 'TRUE' || obj.paymentChecked === 'Done';

  obj.feeTotal = Number(obj.feeTotal) || 0;
  obj.dpAmount = Number(obj.dpAmount) || 0;
  obj.pelunasanAmount = Number(obj.pelunasanAmount) || 0;

  return obj;
}

function eventToRow(ev) {
  return COLUMNS.map(col => {
    if (col === 'quotSent') return ev.quotSent ? 'Terkirim' : '-';
    if (col === 'invoiceSent') return ev.invoiceSent ? 'Terkirim' : '-';
    if (col === 'paymentChecked') return ev.paymentChecked ? 'Done' : '-';
    return (ev[col] !== undefined && ev[col] !== null) ? ev[col] : '';
  });
}

// ================= CRUD OPERATIONS =================

function listEvents() {
  const data = getSheet().getDataRange().getValues();
  return data.slice(1).filter(r => r[0]).map(rowToEvent);
}

function saveEvent(ev) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const idCol = COLUMNS.indexOf('id');
  if (!ev.id) {
    // Fallback: generate ID format DD/MM/YY_Client
    const parts = (ev.date || '').split('-');
    if (parts.length === 3) {
      const dd = parts[2], mm = parts[1], yy = parts[0].slice(-2);
      const client = (ev.clientName || 'Internal').trim();
      ev.id = dd + '/' + mm + '/' + yy + '_' + client;
    } else {
      ev.id = 'evt_' + new Date().getTime();
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === ev.id) {
      sheet.getRange(i + 1, 1, 1, COLUMNS.length).setValues([eventToRow(ev)]);
      return ev;
    }
  }
  sheet.appendRow(eventToRow(ev));
  return ev;
}

function deleteEvent(id, briefFolderUrl) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const idCol = COLUMNS.indexOf('id');
  const folderUrlCol = COLUMNS.indexOf('briefFolderUrl');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      // Hapus folder Brief di Google Drive kalau ada
      const folderUrl = briefFolderUrl || data[i][folderUrlCol];
      if (folderUrl) {
        try {
          const folderId = String(folderUrl).match(/folders\/([a-zA-Z0-9_-]+)/);
          if (folderId && folderId[1]) {
            DriveApp.getFolderById(folderId[1]).setTrashed(true);
          }
        } catch (e) { /* folder sudah tidak ada atau tidak bisa diakses, abaikan */ }
      }

      // Hapus row dari spreadsheet
      sheet.deleteRow(i + 1);
      return;
    }
  }
}
