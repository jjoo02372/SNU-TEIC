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
// 기록을 받을 스프레드시트 (교사 계정 소유). 비워 두면 처음 실행할 때 새로 만듭니다.
var SPREADSHEET_ID = '18fCgcJ1BtjySPSVrMyXcn2hGLpMNxhyaclVauyUtWUs';
// 시트에서 열지 않고 따로 만든 스크립트는 처음 실행할 때 메시지 시트를 직접 만들고 그 ID를 기억합니다.
function spreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.create('TEIC 4차시 팀 메시지');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function sheet_() {
  var ss = spreadsheet_();
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
    try { teacherRow_(now, from, to, kind, d.body || {}); } catch (e3) { /* 교사용 보기 실패는 메시지 전달을 막지 않습니다 */ }
    return json_({ ok: true, id: id, time: now });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* not held */ }
  }
}

// ---------- 교사용 탭: PRD 제출 / 피드백 ----------
var PRD_TAB = 'PRD 제출';
var FB_TAB = '피드백';
var PRD_FIELDS = [
  ['overview', '개요'], ['goal', '목표, 대상'], ['core', '핵심기능'], ['screens', '화면구성'], ['stack', '기술스택'],
  ['must', '꼭 있어야 할 기능'], ['wont', '만들지 않을 기능'], ['data', '입력, 저장하는 데이터'], ['link', '웹 앱 링크']
];
var PRD_HEAD = ['제출 시각', '팀', '앱 이름'].concat(PRD_FIELDS.map(function (f) { return f[1]; }));
var FB_HEAD = ['보낸 시각', '보낸 팀', '쓴 사람', '받은 팀', '이해됨', '막힘', '질문', '결정', '이유 / 메모', '답장 시각'];

function tab_(name, head) {
  var ss = spreadsheet_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name, 0);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#eaf0ec');
    sh.setColumnWidths(2, head.length - 1, 200);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setWrap(true).setVerticalAlignment('top');
  }
  return sh;
}

function cell_(v) {
  // 학생 입력이 수식으로 실행되지 않도록 막습니다
  return String(v == null ? '' : v).trim().replace(/^[=+\-@]/, "'$&");
}

function teacherRow_(time, from, to, kind, b) {
  if (kind === 'prd') {
    var prd = b.prd || {};
    tab_(PRD_TAB, PRD_HEAD).appendRow([new Date(time), from, cell_(b.app)].concat(PRD_FIELDS.map(function (f) { return cell_(prd[f[0]]); })));
  } else if (kind === 'feedback') {
    tab_(FB_TAB, FB_HEAD).appendRow([new Date(time), from, cell_(b.name), to, cell_(b.clear), cell_(b.stuck), cell_(b.q), '', '', '']);
  } else if (kind === 'reply') {
    // 답장: 원래 피드백 줄(보낸 팀 = 답장 받는 팀, 받은 팀 = 답장 보낸 팀, 같은 '막힘')에 결정을 채웁니다
    var sh = tab_(FB_TAB, FB_HEAD);
    var n = sh.getLastRow() - 1;
    var decision = b.status === 'done' ? '반영함' : b.status === 'keep' ? '반영 안 함' : '';
    var stuck = cell_(b.stuck);
    if (n > 0) {
      var vals = sh.getRange(2, 1, n, 6).getDisplayValues();
      for (var i = n - 1; i >= 0; i--) {
        if (vals[i][1] === to && vals[i][3] === from && vals[i][5] === stuck.replace(/^'/, '')) {
          sh.getRange(i + 2, 8, 1, 3).setValues([[decision, cell_(b.note), new Date(time)]]);
          sh.getRange(i + 2, 10).setNumberFormat('yyyy-mm-dd hh:mm');
          return;
        }
      }
    }
    // 원래 줄을 못 찾으면 새 줄로 남깁니다
    sh.appendRow([new Date(time), to, '', from, '', stuck, '', decision, cell_(b.note), new Date(time)]);
  }
}

// 이미 쌓인 메시지로 'PRD 제출'·'피드백' 탭을 처음부터 다시 만듭니다 (편집기에서 한 번 실행)
function rebuildTeacherView() {
  var ss = spreadsheet_();
  [PRD_TAB, FB_TAB, '교사용 보기'].forEach(function (name) {
    var old = ss.getSheetByName(name);
    if (old) ss.deleteSheet(old);
  });
  tab_(FB_TAB, FB_HEAD);
  tab_(PRD_TAB, PRD_HEAD);
  rows_().forEach(function (m) { teacherRow_(m.time, m.from, m.to, m.kind, m.body || {}); });
  Logger.log('교사용 시트: ' + ss.getUrl());
}

// 편집기에서 실행해 시트 접근을 확인하는 함수 (메시지 시트 주소도 로그에 남깁니다)
function testTeams() {
  var out = doGet({ parameter: { action: 'teams' } });
  Logger.log(out.getContent());
  Logger.log('메시지 시트: ' + spreadsheet_().getUrl());
}
