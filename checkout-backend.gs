// ============================================================
// GDLite GD-8017 Music - Google Apps Script Backend
// Получава поръчки от checkout страницата и ги записва в Google Sheets
// ============================================================
//
// ИНСТРУКЦИИ ЗА DEPLOY (следвай стъпките точно):
//
// СТЪПКА 1: Създай Google Spreadsheet
//   1. Отиди на sheets.google.com
//   2. Създай нов spreadsheet
//   3. Преименувай го "GDLite Orders"
//   4. Копирай ID-то от URL-а (частта между /d/ и /edit)
//      Пример: https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXX/edit
//      ID = XXXXXXXXXXXXX
//   5. Постави ID-то на ред 33 долу: SPREADSHEET_ID = 'ТУК'
//
// СТЪПКА 2: Отвори Apps Script
//   1. В Spreadsheet: Разширения > Apps Script
//   2. Изтрий целия съществуващ код
//   3. Постави този файл
//
// СТЪПКА 3: Deploy
//   1. Натисни "Deploy" > "New deployment"
//   2. Type: Web app
//   3. Execute as: Me (your Google account)
//   4. Who has access: Anyone
//   5. Натисни "Deploy"
//   6. Копирай Web App URL (изглежда като https://script.google.com/macros/s/XXXXX/exec)
//
// СТЪПКА 4: Постави URL в checkout страниците
//   В gdlite-checkout-BG.html и gdlite-checkout-GR.html намери:
//     const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL';
//   И замени 'YOUR_APPS_SCRIPT_URL' с Web App URL-а от Стъпка 3.
//
// ============================================================

var SPREADSHEET_ID = '1a4gIF-Ak3qBsMC2jgsSbMXo0j9hktKmb4fzBzE8mRT8';

function doPost(e) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Orders');

    // Създай sheet ако не съществува
    if (!sheet) {
      sheet = ss.insertSheet('Orders');
    }

    // Добави headers ако sheet е празен
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Дата/час',
        'Език',
        'Продукт',
        'Количество',
        'Bump',
        'Обща сума',
        'Доставчик',
        'Цена доставка',
        'Име и Фамилия',
        'Телефон',
        'Адрес',
        'Град',
        'Email',
        'Бележки',
      ]);

      // Форматирай header ред
      var headerRange = sheet.getRange(1, 1, 1, 14);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E53E3E');
      headerRange.setFontColor('#FFFFFF');
    }

    // Парсни данните (URLSearchParams формат)
    var raw = e.parameter.payload || (e.postData && e.postData.contents);
    var data = JSON.parse(raw);

    // Запиши поръчката
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

    // Изпрати имейл нотификация (по желание - премахни коментара за да активираш)
    // sendNotificationEmail(data, timestamp);

    // Върни успех
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // Върни грешка
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Опционална имейл нотификация при нова поръчка
function sendNotificationEmail(data, timestamp) {
  var email   = Session.getActiveUser().getEmail(); // изпраща до теб
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

// GET handler (за тест - отвори URL-а директно в браузъра)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'GDLite Checkout Backend is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
