/**
 * 地理課小遊戲 成績接收程式
 * 貼在試算表的「擴充功能 → Apps Script」裡，部署成網頁應用程式。
 * 學生完成遊戲後，網頁會把成績送到這裡，自動寫進「成績紀錄」工作表。
 */
const SHEET_LOG = '成績紀錄';
const SHEET_BEST = '最佳成績';
const HEADERS = ['上傳時間', '遊戲', '班級座號', '班級', '座號', '片數', '完成秒數', '完成時間', '紀錄編號'];
const SPREADSHEET_KEY = 'SCORE_SPREADSHEET_ID';

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);

    const sid = String(d.sid || '');
    if (!/^\d{5}$/.test(sid)) return reply({ ok: false, error: '班級座號必須是 5 個數字' });
    const sec = Number(d.seconds);
    if (!(sec > 0 && sec < 36000)) return reply({ ok: false, error: '完成時間不正確' });

    // 只留安全的文字，避免有人送公式進來
    const game = String(d.game || '').replace(/^[=+\-@\s]+/, '').slice(0, 40);
    const tenths = Math.round(sec * 10);
    const time = Math.floor(tenths / 600) + ':' + String(Math.floor(tenths / 10) % 60).padStart(2, '0') + '.' + tenths % 10;
    const rid = String(d.rid || '');
    if (!/^[0-9a-z-]{8,64}$/i.test(rid)) return reply({ ok: false, error: '紀錄編號不正確' });
    const count = Number(d.pieces);
    if (!game || !Number.isInteger(count) || count < 1 || count > 1000 || (game === '臺灣地形拼圖' && count !== 24)) {
      return reply({ ok: false, error: '遊戲或片數不正確' });
    }

    const sh = getLogSheet();

    // 同一筆成績重傳時不重複寫入
    const last = sh.getLastRow();
    if (last > 1) {
      const previous = sh.getRange(2, 9, last - 1, 1).createTextFinder(rid).matchEntireCell(true).useRegularExpression(false).findNext();
      if (previous) return reply({ ok: true, rid: rid, duplicate: true });
    }

    sh.appendRow([
      new Date(), game, Number(sid), Number(sid.slice(0, 3)), Number(sid.slice(3)),
      count, Math.round(sec * 10) / 10, time, rid
    ]);
    return reply({ ok: true, rid: rid });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// 用瀏覽器直接打開部署網址時，會看到這段，代表服務有在運作
function doGet() {
  return reply({ ok: true, message: '成績接收服務運作中' });
}

function getLogSheet() {
  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_KEY);
  if (!id) throw new Error('請先在試算表的 Apps Script 編輯器執行 setup，再部署網頁應用程式');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sh.getRange('H:I').setNumberFormat('@');
  }
  if (!ss.getSheetByName(SHEET_BEST)) {
    const best = ss.insertSheet(SHEET_BEST);
    best.getRange('A1').setFormula(
      "=QUERY('" + SHEET_LOG + "'!A:I," +
      "\"select B,D,E,C,min(G),count(G) where C is not null group by B,D,E,C order by B,D,E " +
      "label B '遊戲', D '班級', E '座號', C '班級座號', min(G) '最佳秒數', count(G) '挑戰次數'\",1)"
    );
    best.setFrozenRows(1);
  }
  return sh;
}

// 第一次使用前，在編輯器上方選這個函式按「執行」，會建立工作表並完成授權
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('請從 Google 試算表的「擴充功能 → Apps Script」執行 setup');
  PropertiesService.getScriptProperties().setProperty(SPREADSHEET_KEY, ss.getId());
  getLogSheet();
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
