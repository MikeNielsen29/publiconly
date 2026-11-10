/**
 * LCC Calendar Notebook — Google Apps Script Backend
 * Complete backend: Authentication, Entry Management, Photo Storage
 * 
 * DEPLOYMENT:
 * 1. Open script.google.com
 * 2. Create new project
 * 3. Paste this entire code into Code.gs
 * 4. Project Settings > Script Properties: add API_KEY (long random value),
 *    TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID.
 *    The same API_KEY value must be set as APPS_SCRIPT_SECRET in the app env.
 * 5. Deploy > New deployment > Web App
 * 6. Execute as: Me | Who has access: Anyone
 * 7. Copy Web App URL → set GOOGLE_APPS_SCRIPT_URL in the app env
 * 8. Run initialize() once from the editor to create database tables
 *
 * SECURITY: reads and writes (entries, users, audit) require the API key, which
 * only the Next.js server holds. Passwords are stored as salted SHA-256 digests;
 * legacy plaintext rows are upgraded automatically on the next successful login.
 *
 * PHOTOS: uploaded photos are written to the Drive folder named by the
 * DRIVE_FOLDER_NAME property (default "LCC Calendar Notebook Photos"). The
 * folder is created on first save and shared read-only by link so the app can
 * render the images. The resulting URLs are stored in the Entries row.
 *
 * TELEGRAM: with TELEGRAM_BOT_TOKEN + TELEGRAM_CHAT_ID set, every saved entry is
 * forwarded as one text message plus one message per photo.
 */

// API key lives in Script Properties so it is never committed to the repo.
const API_KEY = PropertiesService.getScriptProperties().getProperty('API_KEY') || '';

// Telegram config - set via Script Properties
const TELEGRAM_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN') || '';
const TELEGRAM_CHAT_ID = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID') || '';

// ==================== SHEET TABS & SCHEMAS ====================
const TABS = {
  USERS: 'Users',
  ENTRIES: 'Entries',
  PHOTOS: 'Photos',
  ADMIN_LOGS: 'AdminLogs'
};

const SCHEMAS = {
  Users: ['id', 'email', 'name', 'role', 'password_hash', 'created_at', 'is_active', 'updated_at'],
  Entries: ['id', 'control_number', 'date', 'photographer', 'contact_no', 'congregation_district', 'church_office', 'camera_used', 'other_details', 'caption', 'photo_urls', 'created_at', 'user_email', 'updated_at'],
  Photos: ['id', 'entry_id', 'photo_url', 'photo_name', 'uploaded_at'],
  AdminLogs: ['id', 'action', 'user_email', 'target_entry_id', 'timestamp']
};

// ==================== DO GET / DO POST / DO OPTIONS ====================
function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }
// ContentService has no setHeader()/CORS support. None is needed either: the
// browser never calls this web app directly — the Next.js server proxies every
// request — so there is no preflight to answer.
function doOptions(e) {
  return ContentService.createTextOutput('');
}

function handleRequest(e) {
  try {
    const action = e.parameter.action || e.parameters.action;
    const body = e.postData ? JSON.parse(e.postData.contents) : {};

    // Public actions (no auth required)
    if (action === 'test') return json({ ok: true, message: 'Backend connected' });
    if (action === 'validateLogin') return validateLogin(body);

    // Everything else requires the API key, which only the server holds
    const auth = checkAuth(e.parameter);
    if (!auth.ok) return json({ ok: false, error: auth.error });

    switch (action) {
      case 'saveEntry': return saveEntry(body);
      case 'getEntries': return getEntries(body.date || e.parameter.date);
      case 'getEntriesByMonth': return getEntriesByMonth(body.year, body.month);
      case 'getAllEntries': return getAllEntries();
      case 'getPhotosByEntry': return getPhotosByEntry(body.entryId);
      case 'getPhoto': return getPhoto(body);
      case 'deleteEntry': return deleteEntry(body);
      case 'updateEntry': return updateEntry(body);
      case 'listUsers': return listUsers();
      case 'createUser':
      case 'createUserAccount':
      case 'create-user': return createUser(body);
      case 'deleteUser':
      case 'delete-user': return deleteUser(body);
      case 'resetUser':
      case 'reset-account': return resetUser(body);
      case 'configureTelegram': return configureTelegram(body);
      default: return json({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json({ ok: false, error: err.toString() });
  }
}

// ==================== AUTHENTICATION ====================
function checkAuth(params) {
  if (!API_KEY) {
    return { ok: false, error: 'API key not configured' };
  }
  if (String((params && params.key) || '') === API_KEY) return { ok: true };
  return { ok: false, error: 'Unauthorized' };
}

// ==================== PASSWORD HASHING ====================
const PASSWORD_HASH_ROUNDS = 500;

function hashPassword(password, salt) {
  const useSalt = salt || Utilities.getUuid().replace(/-/g, '');
  let digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    useSalt + ':' + password,
    Utilities.Charset.UTF_8
  );
  for (let i = 0; i < PASSWORD_HASH_ROUNDS; i++) {
    digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, digest);
  }
  return 'sha256$' + useSalt + '$' + digestToHex(digest);
}

function digestToHex(bytes) {
  return bytes.map(function (b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

function isHashedPassword(stored) {
  return String(stored || '').indexOf('sha256$') === 0;
}

function verifyPassword(password, stored) {
  const value = String(stored || '');
  if (!value) return false;
  const parts = value.split('$');
  if (parts.length === 3 && parts[0] === 'sha256') {
    return hashPassword(password, parts[1]) === value;
  }
  // Legacy plaintext row — still accepted so existing users can migrate.
  return value === password;
}

// ==================== INITIALIZE DATABASE ====================
function initializeDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create all tabs first
  Object.values(TABS).forEach(tabName => {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
    }
    setupSheetHeaders(sheet, tabName);
  });

  // Create default admin user
  const usersSheet = ss.getSheetByName(TABS.USERS);
  const users = safeGetSheetData(usersSheet);
  const adminExists = users.some(u => (u.email || '').toLowerCase().trim() === 'admin@lcc.org' && (u.role || '').toLowerCase().trim() === 'admin');

  if (!adminExists) {
    const id = 'user_' + Utilities.getUuid().toString().slice(0, 8);
    const now = new Date().toISOString();
    usersSheet.appendRow([id, 'admin@lcc.org', 'Administrator', 'admin', hashPassword('admin123'), now, 'TRUE', now]);
  }

  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  PropertiesService.getScriptProperties().setProperty('SHEET_URL', ss.getUrl());

  return json({ ok: true, message: 'Database initialized', sheetId: ss.getId(), sheetUrl: ss.getUrl() });
}

function setupSheetHeaders(sheet, tabName) {
  const headers = SCHEMAS[tabName];
  if (!headers) return;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

// ==================== HELPER FUNCTIONS ====================
function safeGetSheetData(sheet) {
  if (!sheet) return [];
  try {
    return getSheetData(sheet);
  } catch (e) {
    return [];
  }
}

function getSheetData(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => (h || '').toString().trim());
  const timeZone = getTimeZone_(sheet);
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = normalizeCellValue_(row[i], h, timeZone);
    });
    return obj;
  });
}

// Sheets stores a written "2026-11-08" as a real Date, so reading it back gave
// "Sat Nov 08 2026 00:00:00 GMT+0800 (...)" and every date match failed.
function normalizeCellValue_(val, header, timeZone) {
  if (val === null || val === undefined || val === '') return '';

  if (val instanceof Date) {
    if (header === 'date') {
      return Utilities.formatDate(val, timeZone, 'yyyy-MM-dd');
    }
    return val.toISOString();
  }

  if (typeof val === 'object' && val.toString) return val.toString();

  if (header === 'date') return String(val).trim().slice(0, 10);
  return val;
}

function getTimeZone_(sheet) {
  try {
    return sheet.getParent().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  } catch (e) {
    return Session.getScriptTimeZone();
  }
}

// YYYY-MM-DD key for a stored date value, whatever shape it arrived in.
function dateKey_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value).trim().slice(0, 10);
}

function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setupSheetHeaders(sheet, sheetName);
  }
  return sheet;
}

// ==================== USER OPERATIONS ====================
function validateLogin(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName(TABS.USERS);
  const data = usersSheet.getDataRange().getValues();

  const email = (body.email || '').toString().toLowerCase().trim();
  const password = (body.password || '').toString().trim();

  if (!email || !password) return json({ ok: false, error: 'Invalid credentials' });

  const headers = data[0].map(h => (h || '').toString().trim());
  const emailIndex = headers.indexOf('email');
  const passwordIndex = headers.indexOf('password_hash');
  const activeIndex = headers.indexOf('is_active');
  const roleIndex = headers.indexOf('role');
  const idIndex = headers.indexOf('id');
  const nameIndex = headers.indexOf('name');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if ((row[emailIndex] || '').toString().toLowerCase().trim() !== email) continue;
    if ((row[activeIndex] || '').toString().toUpperCase().trim() !== 'TRUE') continue;

    const stored = row[passwordIndex];
    if (!verifyPassword(password, stored)) continue;

    // Upgrade legacy plaintext passwords on successful login.
    if (!isHashedPassword(stored)) {
      usersSheet.getRange(i + 1, passwordIndex + 1).setValue(hashPassword(password));
    }

    logAudit('LOGIN', email, 'SUCCESS');

    return json({
      ok: true,
      role: row[roleIndex],
      email: row[emailIndex],
      userId: row[idIndex],
      name: row[nameIndex]
    });
  }

  logAudit('LOGIN', email, 'FAILED');
  return json({ ok: false, error: 'Invalid credentials' });
}

function createUser(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName(TABS.USERS);
  const users = getSheetData(usersSheet);

  const email = (body.email || '').toLowerCase().trim();
  const name = (body.name || '').toString().trim();
  if (!email) return json({ ok: false, error: 'Email is required' });
  if (!name) return json({ ok: false, error: 'Name is required' });
  if (!body.password) return json({ ok: false, error: 'Password is required' });

  if (users.find(u => (u.email || '').toLowerCase().trim() === email)) {
    return json({ ok: false, error: 'User already exists' });
  }

  const id = 'user_' + Utilities.getUuid().toString().slice(0, 8);
  const now = new Date().toISOString();

  // Log audit trail
  logAudit('CREATE_USER', email, name);

  usersSheet.appendRow([
    id,
    email,
    name,
    body.role === 'admin' ? 'admin' : 'user',
    hashPassword(body.password),
    now,
    'TRUE',
    now
  ]);

  return json({ ok: true, id: id, message: 'User created successfully' });
}

function listUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName(TABS.USERS);
  // Never return password_hash to the client.
  const users = getSheetData(usersSheet).map(function (u) {
    return {
      id: u.id,
      email: u.email,
      name: u.name,
      role: u.role,
      is_active: u.is_active,
      created_at: u.created_at,
      updated_at: u.updated_at
    };
  });
  return json({ ok: true, users: users });
}

function deleteUser(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName(TABS.USERS);
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0].map(h => (h || '').toString().trim());
  const emailIndex = headers.indexOf('email');
  const idIndex = headers.indexOf('id');
  const roleIndex = headers.indexOf('role');

  const targetEmail = (body.email || '').toLowerCase().trim();
  if (!targetEmail) return json({ ok: false, error: 'Email is required' });

  // Deleting your own account would lock you out mid-session.
  const requesterEmail = (body.requesterEmail || '').toLowerCase().trim();
  if (requesterEmail && requesterEmail === targetEmail) {
    return json({ ok: false, error: 'You cannot delete your own account' });
  }

  // Prevent deleting the last admin
  const adminCount = data.slice(1).filter(row => (row[roleIndex] || '').toString().toLowerCase().trim() === 'admin').length;
  const targetUser = data.slice(1).find(row => (row[emailIndex] || '').toString().toLowerCase().trim() === targetEmail);
  if (targetUser && (targetUser[roleIndex] || '').toString().toLowerCase().trim() === 'admin' && adminCount <= 1) {
    return json({ ok: false, error: 'Cannot delete the last admin user' });
  }

  for (let i = 1; i < data.length; i++) {
    if ((data[i][emailIndex] || '').toString().toLowerCase().trim() === targetEmail) {
      usersSheet.deleteRow(i + 1);
      logAudit('DELETE_USER', targetEmail, targetEmail);
      return json({ ok: true, message: 'User deleted successfully' });
    }
  }
  return json({ ok: false, error: 'User not found' });
}

function resetUser(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName(TABS.USERS);
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0].map(h => (h || '').toString().trim());
  const emailIndex = headers.indexOf('email');
  const passwordIndex = headers.indexOf('password_hash');

  const targetEmail = (body.email || '').toLowerCase().trim();
  if (!targetEmail) return json({ ok: false, error: 'Email is required' });

  const newPassword = 'temp_' + Utilities.getUuid().toString().slice(0, 8);

  for (let i = 1; i < data.length; i++) {
    if ((data[i][emailIndex] || '').toString().toLowerCase().trim() === targetEmail) {
      usersSheet.getRange(i + 1, passwordIndex + 1).setValue(hashPassword(newPassword));
      logAudit('RESET_USER', targetEmail, targetEmail);
      return json({ ok: true, newPassword: newPassword, message: 'Password reset successfully' });
    }
  }
  return json({ ok: false, error: 'User not found' });
}

// ==================== ENTRY OPERATIONS ====================
// ---------------------- PHOTO STORAGE (Google Drive) ----------------------
function getPhotoFolder_() {
  const properties = PropertiesService.getScriptProperties();
  const folderId = properties.getProperty('DRIVE_FOLDER_ID');

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      Logger.log('Photo folder is gone, recreating: ' + e.toString());
    }
  }

  const folderName = properties.getProperty('DRIVE_FOLDER_NAME') || 'LCC Calendar Notebook Photos';
  const existing = DriveApp.getFoldersByName(folderName);
  const folder = existing.hasNext() ? existing.next() : DriveApp.createFolder(folderName);
  properties.setProperty('DRIVE_FOLDER_ID', folder.getId());
  return folder;
}

function savePhotoToDrive_(entryId, photo, index) {
  if (!photo || !photo.data) throw new Error('Photo data is missing');

  const mimeType = photo.mimeType || 'image/jpeg';
  const extension = mimeType === 'image/png' ? 'png' : (mimeType === 'image/webp' ? 'webp' : 'jpg');
  const fileName = entryId + '_' + (index + 1) + '.' + extension;

  const blob = Utilities.newBlob(Utilities.base64Decode(photo.data), mimeType, fileName);
  const file = getPhotoFolder_().createFile(blob);

  // Read-only for anyone with the link so the app and Telegram can display it.
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600';
  return {
    id: file.getId(),
    name: photo.name || fileName,
    url: url,
    ref: url
  };
}

// Telegram is the photo archive. Photos are sent as documents so the bytes the
// browser prepared are stored unchanged (sendPhoto would recompress and cap the
// long edge at 1280px). Only the file_id is kept, so Drive is never touched.
function telegramConfigured_() {
  return Boolean(TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID);
}

function savePhotoToTelegram_(photo, index, caption, photographer) {
  if (!photo || !photo.data) throw new Error('Photo data is missing');

  const mimeType = photo.mimeType || 'image/jpeg';
  const extension = mimeType === 'image/png' ? 'png' : (mimeType === 'image/webp' ? 'webp' : 'jpg');
  const fileName = photo.name || ('photo_' + (index + 1) + '.' + extension);

  const blob = Utilities.newBlob(Utilities.base64Decode(photo.data), mimeType, fileName);
  const payload = { chat_id: TELEGRAM_CHAT_ID, document: blob };

  // Tag the first photo so the album in the chat is searchable later.
  if (index === 0) {
    const parts = [];
    if (photographer) parts.push('Photographer: ' + photographer);
    if (caption) parts.push(String(caption));
    payload.caption = parts.join('\n').slice(0, 900);
  }

  const response = UrlFetchApp.fetch(telegramEndpoint_('sendDocument'), {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });

  const text = response.getContentText();
  if (response.getResponseCode() >= 300) {
    throw new Error('Telegram rejected the photo: ' + text.slice(0, 200));
  }

  const document = (JSON.parse(text).result || {}).document;
  if (!document || !document.file_id) {
    throw new Error('Telegram did not return a file_id');
  }

  return {
    name: document.file_name || fileName,
    fileId: document.file_id,
    ref: 'tg:' + document.file_id
  };
}

function saveEntry(body) {
  const entriesSheet = getSheet(TABS.ENTRIES);
  const photosSheet = getSheet(TABS.PHOTOS);
  const photos = Array.isArray(body.photos) ? body.photos : [];
  const storeOnTelegram = telegramConfigured_();

  const id = 'entry_' + Utilities.getUuid().toString().slice(0, 8);
  const controlNumber = 'LCC-' + Utilities.getUuid().toString().slice(0, 6).toUpperCase();
  const now = new Date().toISOString();

  // Store the photos first so the row keeps real references. Drive is only a
  // fallback for when Telegram is not configured - it is never written to
  // otherwise, which is what keeps the Drive quota untouched.
  const photoRefs = [];
  photos.forEach(function (photo, index) {
    try {
      const saved = storeOnTelegram
        ? savePhotoToTelegram_(photo, index, body.caption, body.photographer)
        : savePhotoToDrive_(id, photo, index);
      photoRefs.push(saved.ref);
      photosSheet.appendRow([
        'photo_' + Utilities.getUuid().toString().slice(0, 8),
        id, saved.ref, saved.name, now
      ]);
    } catch (photoError) {
      Logger.log('Photo upload failed: ' + photoError.toString());
    }
  });

  entriesSheet.appendRow([
    id, controlNumber, body.date || '', body.photographer || '',
    body.contact_no || '', body.congregation_district || '',
    body.church_office || '', body.camera_used || '',
    body.other_details || '', body.caption || '',
    JSON.stringify(photoRefs), now,
    body.user_email || '', now
  ]);

  logAudit('CREATE_ENTRY', body.user_email || 'unknown', id);

  // The details message is the human-readable half of the Telegram archive.
  // It never fails the save.
  try {
    sendToTelegram(body);
  } catch (telegramError) {
    Logger.log('Telegram forward failed: ' + telegramError.toString());
  }

  return json({
    ok: true,
    id: id,
    control_number: controlNumber,
    photo_urls: photoRefs,
    photo_count: photoRefs.length,
    storage: storeOnTelegram ? 'telegram' : 'drive',
    message: 'Entry saved successfully'
  });
}

function getEntries(date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.ENTRIES);
  const data = getSheetData(sheet);

  if (date) {
    const target = String(date).trim().slice(0, 10);
    return json({ ok: true, entries: data.filter(e => dateKey_(e.date) === target) });
  }
  return json({ ok: true, entries: data });
}

function getAllEntries() {
  return getEntries(null);
}

// Entries for one calendar month — used by the calendar's date markers.
function getEntriesByMonth(year, month) {
  const monthNumber = Number(month);
  if (!year || !monthNumber || monthNumber < 1 || monthNumber > 12) {
    return json({ ok: false, error: 'Valid year and month are required' });
  }

  const prefix = String(year) + '-' + ('0' + monthNumber).slice(-2);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = getSheetData(ss.getSheetByName(TABS.ENTRIES));

  return json({
    ok: true,
    entries: data.filter(function (entry) {
      return dateKey_(entry.date).indexOf(prefix) === 0;
    })
  });
}

function getPhotosByEntry(entryId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const photosSheet = ss.getSheetByName(TABS.PHOTOS);
  if (!photosSheet) return json({ ok: true, photos: [] });
  const data = getSheetData(photosSheet);
  const photos = data.filter(p => p.entry_id === entryId);
  return json({ ok: true, photos: photos });
}

// Deletes the entry from the app only. Photos already delivered to Telegram are
// the permanent archive and are deliberately left untouched there.
function deleteEntry(body) {
  const id = body.id;
  if (!id) return json({ ok: false, error: 'Entry id is required' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.ENTRIES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => (h || '').toString().trim());
  const idIndex = headers.indexOf('id');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === id) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) return json({ ok: false, error: 'Entry not found' });

  sheet.deleteRow(rowIndex + 1);
  const removedPhotos = deleteEntryPhotos_(ss, id);

  logAudit('DELETE_ENTRY', body.user_email || 'admin', id);
  return json({ ok: true, message: 'Entry deleted', removedPhotos: removedPhotos });
}

// Remove the Photos rows for an entry and trash any legacy Drive files.
// Telegram photos are intentionally kept: the chat is the archive.
function deleteEntryPhotos_(ss, entryId) {
  const photosSheet = ss.getSheetByName(TABS.PHOTOS);
  if (!photosSheet) return 0;

  const values = photosSheet.getDataRange().getValues();
  if (values.length < 2) return 0;

  const headers = values[0].map(function (h) { return (h || '').toString().trim(); });
  const entryIndex = headers.indexOf('entry_id');
  const urlIndex = headers.indexOf('photo_url');
  if (entryIndex === -1) return 0;

  const fileIds = [];

  // Walk backwards so deleting a row never shifts the ones still to check.
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][entryIndex] !== entryId) continue;

    if (urlIndex !== -1) {
      const match = String(values[i][urlIndex] || '').match(/[?&]id=([A-Za-z0-9_-]+)/);
      if (match) fileIds.push(match[1]);
    }

    photosSheet.deleteRow(i + 1);
  }

  fileIds.forEach(function (fileId) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (e) {
      // Already gone or moved - nothing left to clean up.
    }
  });

  return fileIds.length;
}

function updateEntry(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.ENTRIES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => (h || '').toString().trim());
  const idIndex = headers.indexOf('id');
  const updatedIndex = headers.indexOf('updated_at');

  if (!body.id) return json({ ok: false, error: 'Entry id is required' });

  const editable = ['date', 'photographer', 'contact_no', 'congregation_district',
    'church_office', 'camera_used', 'other_details', 'caption'];

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] !== body.id) continue;

    editable.forEach(function (field) {
      if (body[field] === undefined) return;
      const column = headers.indexOf(field);
      if (column !== -1) sheet.getRange(i + 1, column + 1).setValue(body[field]);
    });

    if (updatedIndex !== -1) {
      sheet.getRange(i + 1, updatedIndex + 1).setValue(new Date().toISOString());
    }

    logAudit('UPDATE_ENTRY', body.user_email || 'admin', body.id);
    return json({ ok: true, message: 'Entry updated successfully' });
  }

  return json({ ok: false, error: 'Entry not found' });
}

// ==================== AUDIT LOGGING ====================
function logAudit(action, userEmail, target) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logsSheet = ss.getSheetByName(TABS.ADMIN_LOGS);
    if (!logsSheet) {
      logsSheet = ss.insertSheet(TABS.ADMIN_LOGS);
      setupSheetHeaders(logsSheet, 'AdminLogs');
    }
    const id = 'log_' + Utilities.getUuid().toString().slice(0, 8);
    const now = new Date().toISOString();
    logsSheet.appendRow([id, action, userEmail || '', target || '', now]);
  } catch (e) {
    // Silently fail - audit logging should not break the main flow
  }
}

// ==================== TELEGRAM INTEGRATION ====================
function telegramEndpoint_(method) {
  return 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/' + method;
}

// Text only: the photos are already stored by savePhotoToTelegram_.
function sendToTelegram(entryData) {
  if (!telegramConfigured_()) {
    Logger.log('Telegram not configured - skipping');
    return;
  }

  // Plain text only: Markdown/HTML would break on names containing _ or *.
  const message = [
    '📸 NEW CALENDAR ENTRY',
    '',
    '📅 Date: ' + (entryData.date || '—'),
    '👤 Photographer: ' + (entryData.photographer || '—'),
    '📞 Contact No.: ' + (entryData.contact_no || '—'),
    '🏛 Congregation & District: ' + (entryData.congregation_district || '—'),
    '🏢 Church Office: ' + (entryData.church_office || '—'),
    '📷 Camera Used: ' + (entryData.camera_used || '—'),
    '⚙ Other Details: ' + (entryData.other_details || '—'),
    "📝 Caption (5W's & 1H): " + (entryData.caption || '—'),
    '👤 Submitted by: ' + (entryData.user_email || 'unknown')
  ].join('\n');

  try {
    const response = UrlFetchApp.fetch(telegramEndpoint_('sendMessage'), {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: TELEGRAM_CHAT_ID,
        text: message,
        disable_web_page_preview: true
      }),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() >= 300) {
      Logger.log('Telegram message error: ' + response.getContentText());
    }
  } catch (err) {
    Logger.log('Telegram message error: ' + err.toString());
  }

}

// Streams a stored photo back to the Next.js image proxy. The bot token never
// leaves the server, so this is the only way the browser can show a photo that
// lives in Telegram.
function getPhoto(body) {
  const fileId = String(body.file_id || '').trim();
  if (!fileId) return json({ ok: false, error: 'file_id is required' });
  if (!TELEGRAM_BOT_TOKEN) return json({ ok: false, error: 'Telegram is not configured' });

  try {
    const infoResponse = UrlFetchApp.fetch(
      telegramEndpoint_('getFile') + '?file_id=' + encodeURIComponent(fileId),
      { muteHttpExceptions: true }
    );
    const infoText = infoResponse.getContentText();

    if (infoResponse.getResponseCode() >= 300) {
      return json({ ok: false, error: 'Telegram getFile failed: ' + infoText.slice(0, 200) });
    }

    const filePath = (JSON.parse(infoText).result || {}).file_path;
    if (!filePath) return json({ ok: false, error: 'Telegram did not return a file path' });

    const fileResponse = UrlFetchApp.fetch(
      'https://api.telegram.org/file/bot' + TELEGRAM_BOT_TOKEN + '/' + filePath,
      { muteHttpExceptions: true }
    );

    if (fileResponse.getResponseCode() >= 300) {
      return json({ ok: false, error: 'Could not download the photo from Telegram' });
    }

    const blob = fileResponse.getBlob();
    return json({
      ok: true,
      mime_type: blob.getContentType() || 'image/jpeg',
      data: Utilities.base64Encode(blob.getBytes())
    });
  } catch (err) {
    return json({ ok: false, error: err.toString() });
  }
}

function configureTelegram(body) {
  const botToken = body.botToken || '';
  const chatId = body.chatId || '';
  if (!botToken || !chatId) {
    return json({ ok: false, error: 'Bot token and chat ID are required' });
  }
  PropertiesService.getScriptProperties().setProperty('TELEGRAM_BOT_TOKEN', botToken);
  PropertiesService.getScriptProperties().setProperty('TELEGRAM_CHAT_ID', chatId);
  return json({ ok: true, message: 'Telegram configured successfully' });
}

// ==================== JSON HELPER ====================
// ContentService has no setHeader(): setting CORS headers here throws
// "output.setHeader is not a function". Grep for it before adding any back.
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// ==================== INITIALIZE ====================
function main() {
  const result = initializeDatabase();
  Logger.log(result.getContent());
}