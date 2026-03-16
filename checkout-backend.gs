// ============================================================
// GDLite GD-8017 Music - Google Apps Script Backend
// Получава поръчки от checkout страницата и ги записва в Google Sheets
// ============================================================
//
// ИНСТРУКЦИИ ЗА DEPLOY (следвай стъпките точно):
//
// СТЪПКА 1: Отвори Apps Script от Spreadsheet-а
//   Разширения > Apps Script > изтрий стария код > постави този файл
//
// СТЪПКА 2: Deploy
//   Deploy > Manage deployments > Edit (молив) > New version > Deploy
//
// ============================================================

var SPREADSHEET_ID = '1a4gIF-Ak3qBsMC2jgsSbMXo0j9hktKmb4fzBzE8mRT8';

function writeOrder(data) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Orders');

  if (!sheet) {
    sheet = ss.insertSheet('Orders');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Дата/час', 'Език', 'Продукт', 'Количество', 'Bump',
      'Обща сума', 'Доставчик', 'Цена доставка',
      'Ime и Фамилия', 'Телефон', 'Адрес', 'Град', 'Email', 'Бележки',
    ]);
    var headerRange = sheet.getRange(1, 1, 1, 14);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#E53E3E');
    headerRange.setFontColor('#FFFFFF');
  }

  var now = new Date();
  var timestamp = Utilities.formatDate(now, 'Europe/Sofia', 'dd.MM.yyyy HH:mm:ss');

  sheet.appendRow([
    timestamp,
    data.lang          || 'BG',
    data.product       || 'GDLite GD-8017 Music',
    data.qty           || 1,
    data.bump ? 'Да'  : 'Не',
    (data.totalBGN || data.totalEUR || '') + (data.lang === 'BG' ? 'лв.' : '€'),
    data.delivery      || 'ЕКОНТ',
    (data.deliveryBGN || data.deliveryEUR || '') + (data.lang === 'BG' ? 'лв.' : '€'),
    data.name          || '',
    data.phone         || '',
    data.address       || '',
    data.city          || '',
    data.email         || '',
    data.notes         || '',
  ]);
}

// GET handler - receives orders via ?payload= query param (avoids POST redirect issue)
function doGet(e) {
  if (e.parameter && e.parameter.payload) {
    try {
      var data = JSON.parse(e.parameter.payload);
      writeOrder(data);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'GDLite Checkout Backend is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST handler (fallback)
function doPost(e) {
  try {
    var raw = (e.parameter && e.parameter.payload)
      || (e.postData && e.postData.contents);
    var data = JSON.parse(raw);
    writeOrder(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Опционална имейл нотификация при нова поръчка
function sendNotificationEmail(data, timestamp) {
  var email   = Session.getActiveUser().getEmail();
  var subject = 'Нова поръчка: ' + (data.name || 'Непознат') + ' - ' + (data.totalBGN || data.totalEUR) + (data.lang === 'BG' ? 'лв.' : '€');
  var body =
    'Нова поръчка!\n\n' +
    'Дата: ' + timestamp + '\n' +
    'Клиент: ' + (data.name || '') + '\n' +
    'Телефон: ' + (data.phone || '') + '\n' +
    'Адрес: ' + (data.address || '') + ', ' + (data.city || '') + '\n' +
    'Продукт: ' + (data.product || '') + ' x' + (data.qty || 1) + '\n' +
    'Bump: ' + (data.bump ? 'Да' : 'Не') + '\n' +
    'Обща сума: ' + (data.totalBGN || data.totalEUR || '') + '\n' +
    'Доставчик: ' + (data.delivery || '') + '\n';
  MailApp.sendEmail(email, subject, body);
}
