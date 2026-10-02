/**
 * Příjem objednávek z objednavka.html -> zapisuje do tabulky s odpověďmi (stejné sloupce jako Google Formulář),
 * takže import v aplikaci Sklad je převezme beze změny.
 *
 * Dále:
 *  - doGet (?action=stock) vrací formuláři, které kusy a velikosti jsou fyzicky skladem (jen ano/ne, čte list Pohyby databáze Skladu),
 *  - po zapsání objednávky pošle rodiči potvrzení e-mailem (včetně dostupnosti a informace o platbě přes KIS).
 *
 * Nasazení: script.google.com -> Nový projekt -> vložit tento kód -> Nasadit -> Nové nasazení -> Webová aplikace
 *   Spustit jako: já · Přístup: kdokoli. Vzniklou URL vložit do CONFIG.ENDPOINT v objednavka.html.
 */
var SPREADSHEET_ID = '1Yi8u3c6WusM2LJYq9zHWzXys6DHKdhQCaBh9iqA2Nds'; // zdrojová tabulka z nastavení aplikace
var SHEET_NAME = 'Objednávky';
var TOKEN = '';                       // volitelné: stejné jako CONFIG.TOKEN ve formuláři
var NOTIFY = 'info@fkslovany.cz';     // upozornění na novou objednávku ('' = nevyplňovat)
var DB_SPREADSHEET_ID = '1b5p7-rNaOBGqZklXdTXRm4j0QuujEgTTRZ_j1rnjXLc'; // databáze aplikace Sklad (list Pohyby)
var SENDER_NAME = 'FK Slovany Pardubice';
var SEND_FROM = 'hosek@fkslovany.cz';   // odesílací adresa; musí být v Gmailu nastavená jako alias (Odesílat e-maily jako)
var CONFIRM_PARENT = true;            // potvrzení objednávky rodiči (e-mail z formuláře)

/** Stav skladu pro formulář: { "Triko|M": true, ... } = fyzicky skladem (součet pohybů > 0). */
function doGet(e) {
  try {
    var q = (e && e.parameter) || {};
    if (TOKEN && q.token !== TOKEN) return out({ ok: false, error: 'neplatný token' });
    if (q.action !== 'stock') return out({ ok: false, error: 'neznámá akce' });
    var sh = SpreadsheetApp.openById(DB_SPREADSHEET_ID).getSheetByName('Pohyby');
    var rows = sh.getDataRange().getValues(), head = rows[0].map(String);
    var ci = head.indexOf('item'), cv = head.indexOf('velikost'), ck = head.indexOf('ks');
    if (ci < 0 || cv < 0 || ck < 0) return out({ ok: false, error: 'list Pohyby má jiné sloupce' });
    var sum = {};
    for (var i = 1; i < rows.length; i++) {
      var key = rows[i][ci] + '|' + rows[i][cv];
      sum[key] = (sum[key] || 0) + (Number(rows[i][ck]) || 0);
    }
    var stock = {};
    Object.keys(sum).forEach(function (k) { if (sum[k] > 0) stock[k] = true; });
    return out({ ok: true, stock: stock });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

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

    if (CONFIRM_PARENT && /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(o['E-mail'])) {
      try { sendMail(o['E-mail'], 'Objednávka klubového oblečení k nám dorazila!', confirmText(o), NOTIFY); } catch (mailErr) {}
    }
    if (NOTIFY) sendMail(NOTIFY, 'Nová objednávka oblečení – ' + o['Jméno a příjmení'],
      Object.keys(o).filter(function (k) { return o[k]; }).map(function (k) { return k + ': ' + o[k]; }).join('\n'), o['E-mail']);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}
/** Odešle e-mail z aliasu SEND_FROM (přes Gmail); když alias nebo oprávnění chybí, odešle z hlavní adresy účtu. */
function sendMail(to, subject, body, replyTo) {
  try {
    var opt = { name: SENDER_NAME };
    if (replyTo) opt.replyTo = replyTo;
    if (GmailApp.getAliases().indexOf(SEND_FROM) !== -1) opt.from = SEND_FROM;
    GmailApp.sendEmail(to, subject, body, opt);
  } catch (err) {
    MailApp.sendEmail(to, subject, body, replyTo ? { name: SENDER_NAME, replyTo: replyTo } : { name: SENDER_NAME });
  }
}
/** Jednorázově spusťte v editoru: vyvolá žádost o oprávnění k Gmailu a v protokolu ukáže dostupné aliasy. */
function autorizace() { Logger.log('Aliasy: ' + GmailApp.getAliases().join(', ')); }
function confirmText(o) {
  return 'Ahoj,\n\nobjednávka klubového oblečení pro ' + o['Jméno a příjmení'] + ' (' + o['Kategorie'] + ') k nám dorazila. Děkujeme!\n\n' +
    'Co jsi objednal(a):\n' + (o['Souhrn'] || (o['Typ setu'] + (o['Volitelné položky'] ? ', ' + o['Volitelné položky'] : ''))) + '\n\n' +
    'Celkem: ' + o['Celková cena'] + '\n\n' +
    'Jak to bude dál:\n' +
    '- Zboží, které máme skladem, bude připravené k vyzvednutí v rámci několika dní.\n' +
    '- Zboží, které není skladem, je skladem u dodavatele. Objednáme ho a termín vyzvednutí upřesníme.\n' +
    '- V následujících dnech bude platba zadána do KIS.\n\n' +
    'Až bude objednávka připravená, ozveme se e-mailem.\n\nFK Slovany Pardubice';
}
function out(j) { return ContentService.createTextOutput(JSON.stringify(j)).setMimeType(ContentService.MimeType.JSON); }
