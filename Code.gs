/************ 오프닝 '전국 선생님 지도' 백엔드 (Google Apps Script) ************
 * 참가자 폰이 보낸 학교·과목·교육경력(POST)을 구글 시트에 쌓고,
 * 보드가 물어보면(GET) 접속한 학교 목록을 돌려줍니다.
 * '웹 앱'(액세스: 모든 사용자)으로 배포하세요.
 ***************************************************************************/

const SHEET_NAME = 'map2';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['시각', '학교', '지역', '위도', '경도', '과목', '교육경력']);
  }
  return sh;
}

// 한 건 저장 (폰이 POST)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    getSheet_().appendRow([new Date(), d.school || '', d.region || '', d.lat, d.lon, d.subject || '', d.career || '']);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 접속한 학교 목록 반환 (보드가 GET). 같은 학교는 1건으로 묶고 제출 인원수(count)를 함께 돌려줍니다.
function doGet(e) {
  const order = [];
  const map = {};
  const sh = getSheet_();
  if (sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
    rows.forEach(function (r) {
      const name = String(r[1]);
      if (!name) return;
      if (!map[name]) {
        map[name] = { school: name, region: String(r[2]), lat: r[3], lon: r[4], subject: String(r[5]), career: String(r[6]), count: 0 };
        order.push(name);
      }
      map[name].count++;   // 같은 학교에서 여러 명 제출하면 인원수 누적
    });
  }
  const schools = order.map(function (n) { return map[n]; });
  return out_({ ok: true, schools: schools });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
