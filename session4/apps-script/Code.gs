/**
 * 4차시 완성 워크북 · 팀 간 메시지 서버
 *
 * 피드백 보드(구글 시트)에 붙여 웹 앱으로 배포합니다.
 *   확장 프로그램 → Apps Script → 이 코드 붙여넣기 → 배포 → 새 배포 → 웹 앱
 *   실행 사용자: 나 / 액세스 권한: 모든 사용자
 * 배포 후 나오는 /exec 주소를 워크북의 API_URL에 넣습니다.
 *
 * 메시지는 이 시트의 'messages' 탭에 쌓입니다 (없으면 자동으로 만듭니다).
 * 열: 시각(ms) · id · 보낸 팀 · 받는 팀 · 종류(prd/feedback/reply) · 내용(JSON)
 */

var SHEET_NAME = 'messages';
var KINDS = { prd: true, feedback: true, reply: true };
var MAX_BODY = 8000;

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['time', 'id', 'from', 'to', 'kind', 'body']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_() {
  var sh = sheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 6).getValues().map(function (r) {
    var body;
    try { body = JSON.parse(r[5]); } catch (e) { body = {}; }
    return { time: Number(r[0]), id: String(r[1]), from: String(r[2]), to: String(r[3]), kind: String(r[4]), body: body };
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(s, max) {
  // 시트가 수식으로 읽지 않도록 =, +, -, @로 시작하는 이름 앞에 ' 를 붙입니다
  return String(s == null ? '' : s).trim().slice(0, max || 60).replace(/^[=+\-@]/, "'$&");
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'teams') {
      // 팀마다 가장 최근에 보낸 PRD 하나
      var latest = {};
      rows_().forEach(function (m) {
        if (m.kind === 'prd' && (!latest[m.from] || latest[m.from].time < m.time)) latest[m.from] = m;
      });
      var teams = Object.keys(latest).map(function (k) {
        var m = latest[k];
        return { team: m.from, time: m.time, app: m.body.app || '', prd: m.body.prd || {} };
      }).sort(function (a, b) { return a.team.localeCompare(b.team, 'ko'); });
      return json_({ ok: true, teams: teams, now: Date.now() });
    }
    if (p.action === 'inbox') {
      var team = String(p.team || '').trim().slice(0, 60);
      var since = Number(p.since) || 0;
      if (!team) return json_({ ok: false, error: 'team_required' });
      var items = rows_().filter(function (m) {
        return m.to === team && m.time > since && (m.kind === 'feedback' || m.kind === 'reply');
      });
      return json_({ ok: true, items: items, now: Date.now() });
    }
    return json_({ ok: true, service: 'teic-s4-messages' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var from = clean_(d.from), to = clean_(d.to), kind = clean_(d.kind, 20);
    if (!from || !KINDS[kind]) return json_({ ok: false, error: 'bad_request' });
    if (kind !== 'prd' && !to) return json_({ ok: false, error: 'to_required' });
    var body = JSON.stringify(d.body || {});
    if (body.length > MAX_BODY) return json_({ ok: false, error: 'too_long' });
    lock.waitLock(10000);
    var id = Utilities.getUuid().slice(0, 8);
    var now = Date.now();
    sheet_().appendRow([now, id, from, kind === 'prd' ? '*' : to, kind, body]);
    return json_({ ok: true, id: id, time: now });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* not held */ }
  }
}
