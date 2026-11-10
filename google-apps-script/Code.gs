/**
 * Dating Club leads and payment screenshot collection.
 * Deploy as a Web App: Execute as Me, access Anyone.
 * Images stay private and appear as images positioned over column C.
 */
const SPREADSHEET_ID = '1hO1RjQSyt9iRyBK9WkoSpow9yl7MVv0xtQ2Ly9wHzEU';
const SHEET_NAME = 'Sheet1';
const TIME_ZONE = 'Asia/Kolkata';
const MAX_IMAGE_BYTES = 1900000;

function doGet(e) {
  const p = (e && e.parameter) || {};
  const name = String(p.name || '').trim().replace(/\s+/g, ' ');
  const phone = String(p.phone || '').trim();
  if (!name || name.length > 120 || !/^[6-9]\d{9}$/.test(phone)) {
    return reply_('Invalid lead');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = sheet_();
    const dateTime = Utilities.formatDate(new Date(), TIME_ZONE, 'dd-MM hh:mma');
    // Keep the date as text: Sheets otherwise parses 10-11 as October 11.
    sheet.appendRow([safe_(name), '', '', phone]);
    const row = sheet.getLastRow();
    sheet.getRange(row, 2).setNumberFormat('@').setValue(dateTime);
    sheet.getRange(row, 1, 1, 3).setVerticalAlignment('middle');
    return reply_('OK');
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  let p;
  try {
    p = JSON.parse(e.postData.contents);
  } catch (_) {
    return reply_('Invalid JSON');
  }
  if (p.action !== 'screenshot') return reply_('Unknown action');
  const name = String(p.name || '').trim().replace(/\s+/g, ' ');
  const phone = String(p.phone || '').trim();
  if (!name || !/^[6-9]\d{9}$/.test(phone)) return reply_('Invalid member');
  if (p.mimeType !== 'image/jpeg') return reply_('Unsupported image');
  const base64 = String(p.image || '');
  if (!base64 || base64.length > 2600000) return reply_('Image too large');

  const bytes = Utilities.base64Decode(base64);
  if (!bytes.length || bytes.length > MAX_IMAGE_BYTES) {
    return reply_('Image exceeds 1.9 MB');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = sheet_();
    const last = sheet.getLastRow();
    if (last < 2) return reply_('No matching signup');
    const first = Math.max(2, last - 999);
    const data = sheet.getRange(first, 1, last - first + 1, 4).getDisplayValues();
    let row = -1;
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] === name && data[i][3] === phone) {
        row = first + i;
        break;
      }
    }
    if (row < 0) return reply_('No matching signup');

    const cell = sheet.getRange(row, 3);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', 'payment-proof.jpg');

    // Replace any old or misplaced image for this record.
    const oldImages = sheet.getImages().filter(img => {
      const anchor = img.getAnchorCell();
      return anchor.getRow() === row && anchor.getColumn() === 3;
    });
    const img = sheet.insertImage(blob, 3, row, 6, 6);
    img.setAnchorCell(cell)
       .setAnchorCellXOffset(6)
       .setAnchorCellYOffset(6)
       .setWidth(170)
       .setHeight(108)
       .setAltTextTitle('Payment screenshot')
       .setAltTextDescription('Screenshot attached to the corresponding Dating Club signup');
    oldImages.forEach(oldImage => oldImage.remove());
    sheet.setRowHeight(row, 122);
    sheet.getRange(row, 1, 1, 3).setVerticalAlignment('middle');

    // Keep a private Drive backup so screenshots can be recovered if Google
    // Sheets fails to render an over-grid image.
    let backupUrl = '';
    try {
      const backup = getProofFolder_().createFile(blob);
      backup.setName('datingclub_payment_' +
        Utilities.formatDate(new Date(), TIME_ZONE, 'yyyyMMdd_HHmmss') +
        '_row_' + row + '.jpg');
      backupUrl = backup.getUrl();
    } catch (err) {
      console.error('Optional backup failed: ' + err);
    }

    cell.setNote('Payment screenshot uploaded' +
      (backupUrl ? '\nPrivate backup: ' + backupUrl : ''));
    SpreadsheetApp.flush();
    return reply_('OK');
  } finally {
    lock.releaseLock();
  }
}

/** Run once in Apps Script editor to reposition previously uploaded images. */
function repairImagePositions() {
  const sheet = sheet_();
  let repaired = 0;
  for (const img of sheet.getImages()) {
    const anchor = img.getAnchorCell();
    const row = anchor.getRow();
    if (row < 2) continue;
    // Column D is hidden, so always position payment images over column C.
    const cell = sheet.getRange(row, 3);
    img.setAnchorCell(cell)
       .setAnchorCellXOffset(6)
       .setAnchorCellYOffset(6)
       .setWidth(170)
       .setHeight(108);
    sheet.setRowHeight(row, 122);
    sheet.getRange(row, 1, 1, 3).setVerticalAlignment('middle');
    repaired++;
  }
  SpreadsheetApp.flush();
  console.log('Repositioned ' + repaired + ' images');
  return repaired;
}

function getProofFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('proofFolderId');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (_) {}
  }
  const folder = DriveApp.createFolder('Dating Club Payment Screenshots');
  props.setProperty('proofFolderId', folder.getId());
  return folder;
}

function safe_(value) {
  return /^[=+@-]/.test(value) ? "'" + value : value;
}

function sheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Missing sheet tab');
  return sheet;
}

function reply_(message) {
  return ContentService.createTextOutput(message);
}
