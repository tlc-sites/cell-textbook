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
