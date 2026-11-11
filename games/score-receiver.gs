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
    let game = String(d.game || '').replace(/^[=+\-@\s]+/, '').slice(0, 40);
    if (d.mapId) game = publishedMapMeta_(String(d.mapId)).game;
    const tenths = Math.round(sec * 10);
    const time = Math.floor(tenths / 600) + ':' + String(Math.floor(tenths / 10) % 60).padStart(2, '0') + '.' + tenths % 10;
    const rid = String(d.rid || '');
    if (!/^[0-9a-z-]{8,64}$/i.test(rid)) return reply({ ok: false, error: '紀錄編號不正確' });
    const count = Number(d.pieces);
    if (!game || !Number.isInteger(count) || count < 1 || count > 1000 || ((d.mapId || game === '臺灣地形拼圖') && count !== 24)) {
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
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'map') return reply({ ok: true, protocol: 1, map: activeJigsawMap_(true) });
    return reply({ ok: true, message: '成績接收服務運作中' });
  } catch (err) { return reply({ ok: false, error: String(err) }); }
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
      "=IFNA(QUERY('" + SHEET_LOG + "'!A:I," +
      "\"select B,D,E,C,min(G),count(G) where C is not null group by B,D,E,C order by B,D,E " +
      "label B '遊戲', D '班級', E '座號', C '班級座號', min(G) '最佳秒數', count(G) '挑戰次數'\",1),{\"遊戲\",\"班級\",\"座號\",\"班級座號\",\"最佳秒數\",\"挑戰次數\"})"
    );
    best.setFrozenRows(1);
  }
  return sh;
}

// 第一次使用前，在編輯器上方選這個函式按「執行」，會建立工作表並完成授權
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('請從 Google 試算表的「擴充功能 → Apps Script」執行 setup');
  ss.setSpreadsheetTimeZone('Asia/Taipei');
  PropertiesService.getScriptProperties().setProperty(SPREADSHEET_KEY, ss.getId());
  getLogSheet();
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 管理介面只由私人試算表開啟，公開網頁只提供讀圖與交成績。
const MAP_SHEET = '拼圖地圖資料';
const ACTIVE_MAP_KEY = 'JIGSAW_ACTIVE_MAP';
const DEFAULT_MAP_ID = 'default-v1';
const MAP_CHUNK = 30000;
const MAX_MAP_LENGTH = 1200000;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('拼圖管理').addItem('更換全班地圖', 'openJigsawMapManager').addToUi();
}

function mapEditorSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_KEY);
  if (!ss || ss.getId() !== id) throw new Error('請從老師的成績試算表開啟「拼圖管理」');
  return ss;
}

function openJigsawMapManager() {
  mapEditorSpreadsheet_();
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('map-manager').setWidth(680).setHeight(660), '老師更換地圖');
}

function defaultJigsawMap_() {
  return { id: DEFAULT_MAP_ID, name: '臺灣地形拼圖', game: '臺灣地形拼圖', updatedAt: '', isDefault: true };
}

function mapDataSheet_(ss, create) {
  let sh = ss.getSheetByName(MAP_SHEET);
  if (!sh && create) {
    sh = ss.insertSheet(MAP_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([['地圖版本', '地圖名稱', '發布時間', '圖片段號', '總段數', '圖片資料']]);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function publishedMapMeta_(id) {
  if (id === DEFAULT_MAP_ID) return defaultJigsawMap_();
  if (!/^m[0-9a-f]{32}$/.test(id)) throw new Error('地圖版本不正確');
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty(SPREADSHEET_KEY));
  const sh = mapDataSheet_(ss, false);
  if (!sh || sh.getLastRow() < 2) throw new Error('找不到這個地圖版本');
  const match = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).useRegularExpression(false).findNext();
  if (!match) throw new Error('找不到這個地圖版本');
  const row = match.getRow(), v = sh.getRange(row, 1, 1, 5).getValues()[0];
  return { id: id, name: String(v[1]), updatedAt: String(v[2]), row: row, chunks: Number(v[4]), game: String(v[1]) + ' #' + id.slice(-8), isDefault: false };
}

function activeJigsawMap_(includeImage) {
  const id = PropertiesService.getScriptProperties().getProperty(ACTIVE_MAP_KEY) || DEFAULT_MAP_ID;
  const meta = publishedMapMeta_(id);
  if (!includeImage || meta.isDefault) return meta;
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty(SPREADSHEET_KEY));
  const sh = mapDataSheet_(ss, false);
  if (!Number.isInteger(meta.chunks) || meta.chunks < 1 || meta.chunks > 40) throw new Error('地圖資料不完整');
  const values = sh.getRange(meta.row, 1, meta.chunks, 6).getValues();
  let image = '';
  values.forEach(function(v, i) {
    if (v[0] !== id || Number(v[3]) !== i || Number(v[4]) !== meta.chunks || String(v[5]).slice(0, 6) !== 'chunk:') throw new Error('地圖資料不完整');
    image += String(v[5]).slice(6);
  });
  if (!/^data:image\/(jpeg|png);base64,[A-Za-z0-9+/=]+$/.test(image) || image.length > MAX_MAP_LENGTH) throw new Error('地圖圖片格式不正確');
  meta.image = image;
  delete meta.row;
  delete meta.chunks;
  return meta;
}

function getJigsawMapManagerState() {
  const ss = mapEditorSpreadsheet_(), sh = mapDataSheet_(ss, false);
  const history = [defaultJigsawMap_()];
  if (sh && sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    rows.forEach(function(v) {
      if (Number(v[3]) === 0) history.push({ id: String(v[0]), name: String(v[1]), updatedAt: String(v[2]) });
    });
  }
  return { current: activeJigsawMap_(false), history: [history[0]].concat(history.slice(1).reverse().slice(0, 30)) };
}

function saveJigsawMap(data) {
  const ss = mapEditorSpreadsheet_();
  const name = String(data && data.name || '').replace(/[\u0000-\u001f]/g, '').replace(/^[=+\-@\s]+/, '').trim().slice(0, 24);
  const image = String(data && data.image || '');
  if (!name) throw new Error('請輸入地圖名稱');
  const match = /^data:image\/(jpeg|png);base64,([A-Za-z0-9+/]+={0,2})$/.exec(image);
  if (!match || image.length > MAX_MAP_LENGTH) throw new Error('圖片過大或格式不正確，請重新選擇 JPG／PNG');
  const bytes = Utilities.base64Decode(match[2]);
  const b = bytes.slice(0, 8).map(function(v) { return (v + 256) % 256; });
  if (!(match[1] === 'jpeg' && b[0] === 255 && b[1] === 216 && b[2] === 255) && !(match[1] === 'png' && b.join(',') === '137,80,78,71,13,10,26,10')) throw new Error('圖片內容不是 JPG／PNG');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = mapDataSheet_(ss, true), id = 'm' + Utilities.getUuid().replace(/-/g, '').toLowerCase();
    const updatedAt = new Date().toISOString(), chunks = Math.ceil(image.length / MAP_CHUNK), start = sh.getLastRow() + 1;
    const rows = [];
    for (let i = 0; i < chunks; i++) rows.push([id, name, updatedAt, i, chunks, 'chunk:' + image.slice(i * MAP_CHUNK, (i + 1) * MAP_CHUNK)]);
    if (start + chunks - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), start + chunks - 1 - sh.getMaxRows());
    sh.getRange(start, 1, chunks, 6).setValues(rows);
    SpreadsheetApp.flush();
    // 圖片完整寫入後才切換；失敗不影響全班原本的地圖。
    PropertiesService.getScriptProperties().setProperty(ACTIVE_MAP_KEY, id);
    return { ok: true, id: id, name: name, updatedAt: updatedAt };
  } finally { lock.releaseLock(); }
}

function activateJigsawMap(id) {
  mapEditorSpreadsheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const meta = publishedMapMeta_(String(id));
    PropertiesService.getScriptProperties().setProperty(ACTIVE_MAP_KEY, meta.id);
    return { ok: true, id: meta.id, name: meta.name };
  } finally { lock.releaseLock(); }
}
