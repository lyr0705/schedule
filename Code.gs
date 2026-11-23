/**
 * 일정표 앱 ↔ 구글 스프레드시트 연동 스크립트
 *
 * 설치: 스프레드시트 메뉴 [확장 프로그램 › Apps Script]에 이 내용을 통째로 붙여넣고 저장
 *      → [배포 › 새 배포 › 유형 선택: 웹 앱]
 *        - 다음 사용자 인증정보로 실행: 나
 *        - 액세스 권한이 있는 사용자: 모든 사용자
 *      → 배포 후 나오는 "웹 앱 URL"(…/exec)을 앱의 [📊 구글 시트] 창에 붙여넣기
 *
 * 시트 구성
 *  - '일정' 시트: 한 줄에 일정 하나. 시트에서 직접 고치거나 줄을 추가해도 앱에 반영됩니다.
 *    (새 줄은 시작일·행사명만 있으면 되고, id는 비워 두면 자동으로 채워집니다)
 *  - '설정' 시트: 앱 설정(기관 이름, 요일 색 등). 앱이 관리하므로 직접 고치지 마세요.
 */
var SHEET_EVENTS = '일정';
var SHEET_CONFIG = '설정';
var HEAD = ['id', '시작일', '종료일', '요일', '구분', '행사명', '시간', '장소', '담당자', '비고', '반복ID', '예시'];
var DAYS = ['일', '월', '화', '수', '목', '금', '토'];

function doGet(e) {
  try {
    return json_(load_());
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    if (body.action !== 'sync') return json_({ ok: false, error: 'unknown action' });
    sync_(body);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function eventsSheet_() {
  var sh = ss_().getSheetByName(SHEET_EVENTS);
  if (!sh) {
    sh = ss_().insertSheet(SHEET_EVENTS, 0);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#eef2fb');
    sh.setFrozenRows(1);
    // 날짜·시간이 자동 변환되지 않도록 텍스트 서식
    sh.getRange('A:C').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('@');
    sh.setColumnWidths(1, 1, 90); sh.setColumnWidths(6, 1, 220); sh.setColumnWidths(10, 1, 220);
    sh.hideColumns(1); sh.hideColumns(11); sh.hideColumns(12);
  }
  return sh;
}

function configSheet_() {
  var sh = ss_().getSheetByName(SHEET_CONFIG);
  if (!sh) {
    sh = ss_().insertSheet(SHEET_CONFIG);
    sh.getRange('A1').setValue('앱 설정(JSON) — 앱이 관리합니다. 직접 고치지 마세요.');
  }
  return sh;
}

function fmt_(x) {
  if (x instanceof Date) return Utilities.formatDate(x, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(x == null ? '' : x).trim();
  var m = s.match(/^(\d{4})[.\-\/]\s*(\d{1,2})[.\-\/]\s*(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return s;
}
function str_(x) {
  if (x instanceof Date) return Utilities.formatDate(x, Session.getScriptTimeZone(), 'HH:mm');
  return String(x == null ? '' : x);
}

function load_() {
  var sh = eventsSheet_();
  var values = sh.getDataRange().getValues();
  var events = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var date = fmt_(r[1]), title = str_(r[5]).trim();
    if (!date || !title) continue;
    var id = String(r[0] || '').trim();
    if (!id) { // 시트에서 직접 추가한 줄
      id = 'g' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
      sh.getRange(i + 1, 1).setValue(id);
    }
    events.push({
      id: id, date: date, end: fmt_(r[2]), cat: str_(r[4]), title: title, time: str_(r[6]),
      place: str_(r[7]), owner: str_(r[8]), note: str_(r[9]), series: str_(r[10]),
      sample: r[11] === true || String(r[11]).toUpperCase() === 'Y'
    });
  }
  var raw = configSheet_().getRange('A2').getValue();
  var config = null;
  try { config = raw ? JSON.parse(raw) : null; } catch (e) {}
  return { ok: true, events: events, config: config, sheetUrl: ss_().getUrl() };
}

function rowOf_(e) {
  var d = new Date(e.date + 'T00:00:00');
  return [e.id, e.date, e.end || '', isNaN(d) ? '' : DAYS[d.getDay()], e.cat || '', e.title || '', e.time || '',
          e.place || '', e.owner || '', e.note || '', e.series || '', e.sample ? 'Y' : ''];
}

function sync_(b) {
  var sh = eventsSheet_();
  var values = sh.getDataRange().getValues();
  var index = {};
  for (var i = 1; i < values.length; i++) if (values[i][0]) index[String(values[i][0])] = i + 1;
  var appendRows = [];
  (b.put || []).forEach(function (e) {
    if (!e || !e.id) return;
    var row = rowOf_(e);
    if (index[e.id]) sh.getRange(index[e.id], 1, 1, row.length).setValues([row]);
    else appendRows.push(row);
  });
  if (appendRows.length) {
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, appendRows.length, HEAD.length).setNumberFormat('@').setValues(appendRows);
  }
  (b.del || []).map(function (id) { return index[id]; }).filter(Boolean)
    .sort(function (a, c) { return c - a; })
    .forEach(function (r) { sh.deleteRow(r); });
  if (b.config) configSheet_().getRange('A2').setValue(JSON.stringify(b.config));
  // 날짜순 정렬 (머리글 제외)
  var last = sh.getLastRow();
  if (last > 2) sh.getRange(2, 1, last - 1, HEAD.length).sort([{ column: 2, ascending: true }, { column: 7, ascending: true }]);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
