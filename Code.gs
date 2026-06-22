/*** 이촌콘서바토리 음악학원 · 레슨 정산 저장 백엔드 (Google Apps Script) ***
 *
 * 사용법 (3단계)
 *  1) 구글 시트를 새로 하나 만든다.
 *  2) 상단 메뉴 [확장 프로그램] → [Apps Script] → 이 코드 전체를 붙여넣는다.
 *  3) [배포] → [새 배포] → 유형: '웹 앱'
 *        - 실행 계정: 나(본인)
 *        - 액세스 권한: 모든 사용자
 *     배포 후 나오는 URL(.../exec)을 index.html 의 API_URL 에 붙여넣는다.
 *
 *  ⚠ 아래 TOKEN 은 index.html 의 API_TOKEN 과 반드시 똑같이!
 *  데이터는 'data' 시트에 key / value / updatedAt 형식으로 저장됩니다.
 *  (이 코드는 SHEET_ID 로 지정한 파일에 직접 연동됩니다.)
 ************************************************************************/

var TOKEN = '이촌2026';   // index.html 의 API_TOKEN 과 동일하게
var SHEET = 'data';       // 앱 전용 저장 탭 (기존 정산 탭은 건드리지 않음)
var SHEET_ID = '16_noV7SPOBjYopp2gK2P9VHTyXmOaccpsYkixy3MTWs';  // 연동할 구글 시트 ID

function doPost(e){ return handle(e); }
function doGet(e){ return handle(e); }

function handle(e){
  var out = { ok:false };
  try{
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.payload) {
      body = JSON.parse(e.parameter.payload);
    }

    if (body.token !== TOKEN) { out.error = 'bad token'; return json(out); }

    var sh = getSheet();

    if (body.action === 'load') {
      out.data = readAll(sh);
      out.ok = true;
    } else if (body.action === 'save') {
      writeUpdates(sh, body.updates || {});
      out.ok = true;
    } else {
      out.error = 'unknown action';
    }
  } catch (err) {
    out.error = String(err);
  }
  return json(out);
}

function getSheet(){
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID)
                    : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['key', 'value', 'updatedAt']);
    sh.getRange('B:B').setNumberFormat('@'); // value 열은 텍스트로
  }
  return sh;
}

function readAll(sh){
  var data = {};
  var last = sh.getLastRow();
  if (last < 2) return data;
  var rows = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    var k = rows[i][0];
    if (k !== '' && k != null) data[k] = String(rows[i][1]);
  }
  return data;
}

function writeUpdates(sh, updates){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // 동시 저장 충돌 방지
  try{
    var last = sh.getLastRow();
    var keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
    var idx = {};
    for (var i = 0; i < keys.length; i++) idx[keys[i][0]] = i + 2;

    var now = new Date();
    Object.keys(updates).forEach(function(k){
      var v = updates[k];
      if (idx[k]) {
        sh.getRange(idx[k], 2, 1, 2).setValues([[v, now]]);
      } else {
        sh.appendRow([k, v, now]);
        idx[k] = sh.getLastRow();
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function json(o){
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
