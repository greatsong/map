/************ 오프닝 '전국 선생님 지도' 백엔드 (Google Apps Script) ************
 * 참가자 폰이 보낸 학교(POST)를 구글 시트에 쌓고,
 * 앞 화면 보드가 물어보면(GET) 접속한 학교 목록을 돌려줍니다.
 * 이 코드를 그대로 붙여넣고 '웹 앱'(액세스: 모든 사용자)으로 배포하세요.
 ***************************************************************************/

const SHEET_NAME = 'map';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['시각', '학교', '지역', '위도', '경도']);
  }
  return sh;
}

// 학교 1건 저장 (폰이 POST)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    getSheet_().appendRow([new Date(), d.school || '', d.region || '', d.lat, d.lon]);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 접속한 학교 목록 반환 (보드가 GET). 같은 학교는 1번만.
function doGet(e) {
  const sh = getSheet_();
  const schools = [];
  const seen = {};
  if (sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    rows.forEach(function (r) {
      const name = String(r[1]);
      if (!name || seen[name]) return;
      seen[name] = true;
      schools.push({ school: name, region: String(r[2]), lat: r[3], lon: r[4] });
    });
  }
  return out_({ ok: true, schools: schools });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
