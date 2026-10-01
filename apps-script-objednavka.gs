/**
 * Příjem objednávek z objednavka.html -> zapisuje do tabulky s odpověďmi (stejné sloupce jako Google Formulář),
 * takže import v aplikaci Sklad je převezme beze změny.
 *
 * Nasazení: script.google.com -> Nový projekt -> vložit tento kód -> Nasadit -> Nové nasazení -> Webová aplikace
 *   Spustit jako: já · Přístup: kdokoli. Vzniklou URL vložit do CONFIG.ENDPOINT v objednavka.html.
 */
var SPREADSHEET_ID = '1Yi8u3c6WusM2LJYq9zHWzXys6DHKdhQCaBh9iqA2Nds'; // zdrojová tabulka z nastavení aplikace
var SHEET_NAME = 'Objednávky';
var TOKEN = '';                       // volitelné: stejné jako CONFIG.TOKEN ve formuláři
var NOTIFY = 'info@fkslovany.cz';     // upozornění na novou objednávku ('' = nevyplňovat)

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (TOKEN && body.token !== TOKEN) return out({ ok: false, error: 'neplatný token' });
    var o = body.order || {};
    if (!o['Jméno a příjmení'] || !o['E-mail']) return out({ ok: false, error: 'chybí povinné údaje' });

    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    var row = head.map(function () { return ''; });
    o['Časová značka'] = Utilities.formatDate(new Date(), 'Europe/Prague', 'd.M.yyyy H:mm:ss');
    Object.keys(o).forEach(function (key) {
      for (var i = 0; i < head.length; i++) {
        if (head[i].indexOf(key) === 0 && row[i] === '') { row[i] = o[key]; break; }  // stejné rozpoznávání jako v aplikaci (začátek názvu sloupce)
      }
    });
    sh.appendRow(row);
    lock.releaseLock();

    if (NOTIFY) MailApp.sendEmail(NOTIFY, 'Nová objednávka oblečení – ' + o['Jméno a příjmení'],
      Object.keys(o).filter(function (k) { return o[k]; }).map(function (k) { return k + ': ' + o[k]; }).join('\n'));
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}
function out(j) { return ContentService.createTextOutput(JSON.stringify(j)).setMimeType(ContentService.MimeType.JSON); }
