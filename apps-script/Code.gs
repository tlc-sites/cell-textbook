/**
 * 주님의교회 셀 교재 제작 — 구글 시트 저장용 Apps Script
 * 이 코드는 구글 시트의 [확장 프로그램 → Apps Script]에 붙여 넣어 사용합니다.
 * 시트의 '내용(자동)' 칸은 앱이 쓰는 칸이니 직접 고치지 마세요.
 * (완성본 PDF를 드라이브 폴더에 보관하므로 @OnlyCurrentDoc 는 쓰지 않습니다)
 */

var SHEET_BOOKS = '교재';
var SHEET_SETTINGS = '설정';
var SHEET_IMAGES = '사진';
var CHUNK = 45000;          // 시트 한 칸에 넣을 글자 수 (한 칸 최대 5만 자)
var DATA_COL_BOOK = 9;      // 교재 내용이 시작되는 칸
var DATA_COL_SET = 4;       // 설정 내용이 시작되는 칸
var DATA_COL_IMG = 3;       // 사진 내용이 시작되는 칸

/* ---------- 웹 앱 화면 ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('주님의교회 셀 교재 제작')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 저장 창구 (바깥 화면에서 부르는 곳) ----------
 * 화면은 구글 바깥 주소에 두고, 여기로 {fn, args} 를 보내 저장·불러오기를 합니다. */
var API_ = {
  checkPassword: checkPassword, init: init, listBooks: listBooks, loadBook: loadBook,
  getVersion: getVersion, saveBook: saveBook, createBook: createBook, saveSettings: saveSettings,
  savePdf: savePdf, deleteBook: deleteBook, presence: presence
};
function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = API_[req.fn];
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    out = { ok: true, data: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 시트 준비 ---------- */
function ss_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('이 코드가 구글 시트에 연결되어 있지 않습니다. "셀 교재 저장소" 시트의 [확장 프로그램 → Apps Script]에서 만든 코드인지 확인해 주세요.');
  return ss;
}

function sheet_(name, headers) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function books_() { return sheet_(SHEET_BOOKS, ['번호', '날짜', '한글 제목', '영어 제목', '저장한 사람', '저장 시각', '버전', '사진 번호', '내용(자동)']); }
function settings_() { return sheet_(SHEET_SETTINGS, ['이름', '저장한 사람', '저장 시각', '내용(자동)']); }
function images_() { return sheet_(SHEET_IMAGES, ['사진 번호', '저장 시각', '내용(자동)']); }

function me_() { try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; } }
function now_() { return new Date().toISOString(); }

function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function setText_(sh, row, col, vals) {
  var r = sh.getRange(row, col, 1, vals.length);
  r.setNumberFormat('@');
  r.setValues([vals.map(function (v) { return String(v); })]);
}
/* 긴 글을 여러 칸에 나눠 저장 */
function writeChunks_(sh, row, col, str) {
  str = str || '';
  var parts = [];
  for (var i = 0; i < str.length; i += CHUNK) parts.push(str.substring(i, i + CHUNK));
  if (!parts.length) parts = [''];
  var need = col + parts.length - 1;
  if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
  var lastCol = Math.max(sh.getLastColumn(), need);
  sh.getRange(row, col, 1, lastCol - col + 1).clearContent();
  setText_(sh, row, col, parts);
}
function readChunks_(sh, row, col) {
  var lastCol = sh.getLastColumn();
  if (lastCol < col) return '';
  return sh.getRange(row, col, 1, lastCol - col + 1).getValues()[0].map(function (v) { return String(v); }).join('');
}
function firstLine_(obj, lang) {
  var d = obj && obj.docs && obj.docs[lang];
  return d && d.title ? String(d.title).split('\n')[0] : '';
}

/* ---------- 사진 ---------- */
function saveImage_(dataUrl) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, dataUrl);
  var id = bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  var sh = images_();
  if (findRow_(sh, id)) return id;            // 같은 사진은 한 번만 저장
  var row = sh.getLastRow() + 1;
  setText_(sh, row, 1, [id, now_()]);
  writeChunks_(sh, row, DATA_COL_IMG, dataUrl);
  return id;
}
function loadImage_(id) {
  if (!id) return '';
  var sh = images_(), row = findRow_(sh, id);
  return row ? readChunks_(sh, row, DATA_COL_IMG) : '';
}

/* ---------- 편집 비밀번호 ----------
 * [프로젝트 설정 → 스크립트 속성]에 PASSWORD 가 없으면 비밀번호 없이 누구나 씁니다.
 * 나중에 다시 잠그고 싶으면 이름 PASSWORD 로 값을 넣기만 하면 됩니다. (다시 배포할 필요 없음) */
function pass_(pw) {
  var p = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  if (!p) return;
  if (String(pw || '') !== String(p)) { Utilities.sleep(800); throw new Error('BADPASS'); }
}
function checkPassword(pw) {
  try { pass_(pw); return 'ok'; } catch (e) { return e.message; }
}

/* ---------- 화면에서 부르는 기능 (모두 비밀번호 확인) ---------- */
function init(pw) {
  pass_(pw);
  var s = settings_(), row = findRow_(s, 'settings');
  return { email: me_(), books: listBooks_(), settings: row ? readChunks_(s, row, DATA_COL_SET) : '' };
}

function listBooks(pw) { pass_(pw); return listBooks_(); }

function listBooks_() {
  var sh = books_(), last = sh.getLastRow();
  if (last < 2) return [];
  var out = sh.getRange(2, 1, last - 1, 8).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      return { id: String(r[0]), date: String(r[1]), titleKo: String(r[2]), titleEn: String(r[3]),
               savedBy: String(r[4]), savedAt: String(r[5]), version: Number(r[6]) || 0 };
    });
  out.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  return out;
}

function loadBook(pw, id) {
  pass_(pw);
  var sh = books_(), row = findRow_(sh, id);
  if (!row) return null;
  var v = sh.getRange(row, 1, 1, 8).getValues()[0];
  return { id: String(v[0]), version: Number(v[6]) || 0, savedBy: String(v[4]), savedAt: String(v[5]),
           json: readChunks_(sh, row, DATA_COL_BOOK), banner: loadImage_(String(v[7])) };
}

function getVersion(pw, id) {
  pass_(pw);
  var sh = books_(), row = findRow_(sh, id);
  if (!row) return null;
  var v = sh.getRange(row, 1, 1, 8).getValues()[0];
  return { version: Number(v[6]) || 0, savedBy: String(v[4]), savedAt: String(v[5]) };
}

/* banner: null = 그대로, '' = 사진 없앰, 'data:...' = 새 사진 */
function saveBook(pw, id, json, baseVersion, banner, name, force) {
  pass_(pw);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = books_(), row = findRow_(sh, id);
    if (!row) throw new Error('교재를 찾지 못했습니다.');
    var cur = sh.getRange(row, 1, 1, 8).getValues()[0];
    var curVer = Number(cur[6]) || 0;
    if (!force && curVer !== Number(baseVersion)) {
      return { conflict: true, version: curVer, savedBy: String(cur[4]), savedAt: String(cur[5]) };
    }
    var obj = JSON.parse(json);
    var img = String(cur[7]);
    if (banner !== null && banner !== undefined) img = banner ? saveImage_(banner) : '';
    var who = name || me_(), at = now_(), ver = curVer + 1;
    var tKo = firstLine_(obj, 'ko'), tEn = firstLine_(obj, 'en');
    setText_(sh, row, 2, [obj.date || '', tKo, tEn, who, at, ver, img]);
    writeChunks_(sh, row, DATA_COL_BOOK, json);
    return { version: ver, savedBy: who, savedAt: at, date: obj.date || '', titleKo: tKo, titleEn: tEn };
  } finally {
    lock.releaseLock();
  }
}

/* 새 교재: 지금 교재(fromId)의 표지 사진과 색상을 이어받음 */
function createBook(pw, date, fromId) {
  pass_(pw);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = books_(), id = Utilities.getUuid(), colors = null, img = '';
    var fromRow = fromId ? findRow_(sh, fromId) : 0;
    if (fromRow) {
      img = String(sh.getRange(fromRow, 8).getValue());
      try { colors = JSON.parse(readChunks_(sh, fromRow, DATA_COL_BOOK)).colors || null; } catch (e) {}
    }
    var blank = { title: '', ref: '', body: '' };
    var obj = { date: date, lines: true, docs: { ko: blank, en: blank }, check: { ko: [], en: [] } };
    if (colors) obj.colors = colors;
    var row = sh.getLastRow() + 1;
    setText_(sh, row, 1, [id, date, '', '', me_(), now_(), 1, img]);
    writeChunks_(sh, row, DATA_COL_BOOK, JSON.stringify(obj));
    return id;
  } finally {
    lock.releaseLock();
  }
}

/* 교재 지우기: 시트에서 바로 없애지 않고 '지운 교재' 탭으로 옮겨 둡니다.
 * 되살리려면 '지운 교재' 탭의 그 줄에서 [번호]부터 끝까지 복사해 '교재' 탭 맨 아래에 붙여 넣으면 됩니다. */
function deleteBook(pw, id, name) {
  pass_(pw);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = books_(), row = findRow_(sh, id);
    if (!row) throw new Error('교재를 찾지 못했습니다. 이미 지워졌을 수 있습니다.');
    var vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    while (vals.length > 9 && vals[vals.length - 1] === '') vals.pop();
    var bin = sheet_('지운 교재', ['지운 사람', '지운 시각', '번호', '날짜', '한글 제목', '영어 제목', '저장한 사람', '저장 시각', '버전', '사진 번호', '내용(자동)']);
    var need = 2 + vals.length;
    if (bin.getMaxColumns() < need) bin.insertColumnsAfter(bin.getMaxColumns(), need - bin.getMaxColumns());
    setText_(bin, bin.getLastRow() + 1, 1, [name || me_(), now_()].concat(vals));
    sh.deleteRow(row);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/* 누가 같은 교재를 열어 두었는지 알려 주기 (시트에는 남기지 않고 잠깐만 기억)
 * 화면이 20초마다 알려 오고, 90초 넘게 소식이 없으면 나간 것으로 봅니다. */
function presence(pw, bookId, clientId, name, lang, editedAgo, leaving) {
  pass_(pw);
  if (!bookId || !clientId) return [];
  var cache = CacheService.getScriptCache(), key = 'pres_' + bookId, now = Date.now();
  var lock = LockService.getScriptLock();
  var locked = lock.tryLock(3000);
  try {
    var list = {};
    try { list = JSON.parse(cache.get(key) || '{}'); } catch (e) {}
    Object.keys(list).forEach(function (k) { if (now - list[k].at > 90000) delete list[k]; });
    if (leaving) delete list[clientId];
    else list[clientId] = { name: String(name || '').slice(0, 40), lang: lang === 'en' ? 'en' : 'ko', at: now,
                            edit: (editedAgo == null || editedAgo < 0) ? 0 : now - Number(editedAgo) };
    cache.put(key, JSON.stringify(list), 600);
    return Object.keys(list).filter(function (k) { return k !== clientId; }).map(function (k) {
      var x = list[k];
      return { name: x.name, lang: x.lang, editing: !!x.edit && now - x.edit < 120000 };
    });
  } finally {
    if (locked) lock.releaseLock();
  }
}

/* 글꼴과 원어 단어장은 모두가 함께 씀 */
function saveSettings(pw, json, name) {
  pass_(pw);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = settings_(), row = findRow_(sh, 'settings');
    if (!row) row = sh.getLastRow() + 1;
    setText_(sh, row, 1, ['settings', name || me_(), now_()]);
    writeChunks_(sh, row, DATA_COL_SET, json);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function verifyChurchCopy() {
 var ss=ss_(); if(ss.getId()!=='1xXoIcbgubCcDbeRtwD47mYD0I6VRVNEjyB5OKVwxeac')throw new Error('교회 사본 시트가 아닙니다.');
 console.log('교회 셀 교재 시트 연결 확인');console.log('교재 행 수: '+Math.max(0,books_().getLastRow()-1));console.log('사진 행 수: '+Math.max(0,images_().getLastRow()-1));console.log('편집 비밀번호 설정: '+!!PropertiesService.getScriptProperties().getProperty('PASSWORD'));
}

/* ==================== 완성본 PDF 드라이브 보관 ==================== */
/**
 * 셀 교재 완성본 PDF를 구글 드라이브에 보관합니다.
 *
 * 보관 위치:  [PDF_ROOT 폴더] › 2026 › 셀 교재(KR)        › 20261011_셀 교재 (KR).pdf
 *                                    › Cell Textbook(EN)  › 20261011_Cell Textbook (EN).pdf
 * - 연도 폴더와 언어 폴더가 없으면 자동으로 만듭니다. (내년에는 2027 폴더가 저절로 생깁니다)
 * - 같은 이름의 파일이 있으면 이전 파일은 휴지통으로 옮기고 새 파일로 바꿉니다.
 *
 * PDF_ROOT: '2026' 폴더가 들어 있는 "상위 폴더"의 주소(링크)나 ID를 넣어 주세요.
 * 이 Apps Script를 만든 계정이 그 폴더를 편집할 수 있어야 합니다.
 */
var PDF_ROOT = '';

var PDF_LANG = {
  ko: { folder: '셀 교재(KR)', name: '셀 교재 (KR)' },
  en: { folder: 'Cell Textbook(EN)', name: 'Cell Textbook (EN)' }
};

function pdfRoot_() {
  var v = String(PDF_ROOT || PropertiesService.getScriptProperties().getProperty('PDF_ROOT') || '').trim();
  if (!v) throw new Error('보관할 드라이브 폴더가 아직 정해지지 않았습니다. 담당 사역자에게 알려 주세요.');
  var m = /\/folders\/([\w-]+)/.exec(v) || /[?&]id=([\w-]+)/.exec(v);
  return DriveApp.getFolderById(m ? m[1] : v);
}

function subFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* 웹앱에서 부르는 기능: savePdf(비밀번호, 'ko'|'en', 'YYYY-MM-DD', PDF base64) */
function savePdf(pw, lang, date, base64) {
  pass_(pw);
  var L = PDF_LANG[lang];
  if (!L) throw new Error('교재 언어를 알 수 없습니다.');
  var d = String(date || '').replace(/\D/g, '');
  if (d.length !== 8) throw new Error('교재 날짜를 먼저 정해 주세요.');
  var bytes = Utilities.base64Decode(String(base64 || ''));
  if (bytes.length < 5 || String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) !== '%PDF') {
    throw new Error('PDF 파일이 아닙니다.');
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var year = d.slice(0, 4);
    var folder = subFolder_(subFolder_(pdfRoot_(), year), L.folder);
    var name = d + '_' + L.name + '.pdf';
    var old = folder.getFilesByName(name), replaced = false;
    while (old.hasNext()) { old.next().setTrashed(true); replaced = true; }
    var file = folder.createFile(Utilities.newBlob(bytes, 'application/pdf', name));
    return { name: name, url: file.getUrl(), folder: year + ' › ' + L.folder, replaced: replaced };
  } finally {
    lock.releaseLock();
  }
}

/* 설치할 때 한 번 실행: 드라이브 권한을 허용하고, 보관 폴더가 제대로 열리는지 확인합니다. */
function testPdfFolder() {
  var root = pdfRoot_();
  var y = String(new Date().getFullYear());
  Logger.log('상위 폴더: ' + root.getName());
  Logger.log('한글: ' + root.getName() + ' › ' + y + ' › ' + subFolder_(subFolder_(root, y), PDF_LANG.ko.folder).getName());
  Logger.log('영어: ' + root.getName() + ' › ' + y + ' › ' + subFolder_(subFolder_(root, y), PDF_LANG.en.folder).getName());
}
