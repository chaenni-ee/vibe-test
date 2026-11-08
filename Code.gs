/**
 * 바이브 레벨 테스트 - 구글 앱스크립트 서버 코드 (예비용: 일반 호스팅 + 구글 시트 저장)
 *
 * 이 스크립트는 구글 시트에 붙여서(확장 프로그램 > Apps Script) 사용합니다.
 * 화면(index.html)은 Netlify 같은 일반 웹 호스팅에 따로 올리고,
 * 이 스크립트는 응답을 저장하고 돌려주는 일만 합니다. (앱스크립트에 HTML 파일은 넣지 않아요)
 * - doPost: 참여자 응답을 '응답' 시트에 한 줄씩 저장합니다.
 * - doGet(?action=list&r=회차): 교사 화면에 회차별 응답을 돌려줍니다.
 */

var SHEET_NAME = '응답';
var HEADER = ['제출 시각', '회차', '이름', '총점', '레벨',
  '개념 이해', '도구 경험', '프롬프트 감각', '수업 적용',
  'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12'];

/* 문항별 영역(index.html의 Q 배열 순서와 같아야 합니다) */
var Q_CATS = [0, 0, 0, 1, 1, 1, 2, 2, 2, 3, 3, 3];
/* 문항당 최고 점수(보기 3개 → 0~2점). 영역당 6점, 총 24점 */
var MAX_PT = 2;
var CAT_MAX = MAX_PT * 3;
var TOTAL_MAX = MAX_PT * Q_CATS.length;

/* 응답 목록 요청. 주소를 브라우저로 그냥 열면 연결 확인 메시지가 나옵니다. */
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  if (p.action === 'list') {
    try {
      return json_({ ok: true, data: listResponses(p.r) });
    } catch (err) {
      return json_({ ok: false, error: String(err && err.message || err) });
    }
  }
  return json_({ ok: true, message: '바이브 레벨 테스트 저장 서버가 동작 중이에요. 이 주소를 index.html의 API_URL에 넣어 주세요.' });
}

/* 응답 저장 요청 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action !== 'save') throw new Error('알 수 없는 요청이에요.');
    saveResponse(body.room, body.rec);
    return json_({ ok: true, data: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* 응답 저장 */
function saveResponse(room, rec) {
  room = cleanRoom_(room);
  if (!rec || !Array.isArray(rec.a) || rec.a.length !== Q_CATS.length) {
    throw new Error('응답 형식이 올바르지 않아요.');
  }
  // 점수는 서버에서 다시 계산합니다(문항별 점수만 믿고 나머지는 재계산).
  var a = rec.a.map(function (v) {
    v = Number(v);
    if (!(v >= 0 && v <= MAX_PT) || Math.floor(v) !== v) throw new Error('응답 값이 올바르지 않아요.');
    return v;
  });
  var c = [0, 0, 0, 0], s = 0;
  a.forEach(function (v, i) { c[Q_CATS[i]] += v; s += v; });
  var name = String(rec.n || '').trim().slice(0, 12);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet_();
    sh.appendRow([new Date(), safeText_(room), safeText_(name), s, levelOf_(s)]
      .concat(c).concat(a));
  } finally {
    lock.releaseLock();
  }
  return true;
}

/* 회차별 응답 목록 */
function listResponses(room) {
  room = cleanRoom_(room);
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, HEADER.length).getValues();
  var out = [];
  rows.forEach(function (r) {
    if (String(r[1]) !== room) return;
    var a = r.slice(9, 21).map(Number);
    var c = r.slice(5, 9).map(Number);
    var t = (r[0] instanceof Date) ? r[0].getTime() : 0; // Date는 그대로 넘길 수 없어서 숫자로 바꿉니다.
    out.push({ n: String(r[2] || ''), t: t, s: Number(r[3]), l: Number(r[4]), c: c, a: a });
  });
  return out;
}

/* ---------- 도우미 ---------- */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
  }
  return sh;
}

/* 24점 만점 기준 레벨 구간: 0–5 / 6–10 / 11–15 / 16–20 / 21–24 (index.html의 levelOf와 같아야 합니다) */
function levelOf_(s) { return s <= 5 ? 1 : s <= 10 ? 2 : s <= 15 ? 3 : s <= 20 ? 4 : 5; }

function cleanRoom_(r) {
  r = String(r || 'main').replace(/[^A-Za-z0-9가-힣_-]/g, '').slice(0, 20);
  return r || 'main';
}

/* 글자를 항상 '텍스트'로 저장합니다. 앞에 '를 붙이면 =수식, 5/3(날짜), 007(숫자) 같은 자동 변환을 막을 수 있어요.
   시트 화면과 읽어 온 값에는 '가 보이지 않습니다. */
function safeText_(v) {
  v = String(v);
  return v ? "'" + v : v;
}

/* 편집기에서 한 번 실행하면 시트 권한을 승인하고 '응답' 시트와 '요약 그래프' 시트를 만들어 둡니다. */
function setup() {
  getSheet_();
  buildDashboard();
}

/* 시트를 열면 위쪽 메뉴에 '바이브 테스트'가 생깁니다. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('바이브 테스트')
    .addItem('요약 그래프 만들기 / 다시 만들기', 'buildDashboard')
    .addToUi();
}

/* ============ 요약 그래프 시트 ============
   수식으로 집계하고 그 표를 차트로 그립니다. 응답이 쌓이면 표와 차트가 저절로 바뀝니다.
   B1 칸에서 회차(또는 '전체')를 고르면 그 회차만 집계합니다. */
var DASH_NAME = '요약 그래프';
var LEVEL_NAMES = ['말랑 새싹이', '두리번 탐험가', '척척 조립가', '뚝딱 메이커', '바이브 마법사'];
var CAT_NAMES = ['개념 이해', '도구 경험', '프롬프트 감각', '수업 적용'];
var Q_LABELS = ['바이브코딩 설명', '코드 살펴보기', 'AI 에이전트', 'AI 사용 경험', '링크로 공유·배포', '데이터 저장',
  '밋밋한 결과 대응', '기능이 많은 앱 요청', '오류 대응', '수업 적용 경험', '다른 선생님 웹앱 활용', '학생 사용 전 점검'];
var Q1_OPTS = ['처음 들어 보거나 이름만 들어 봤어요', '대략 설명할 수 있어요', '예를 들어 다른 선생님께 설명할 수 있어요'];

function buildDashboard() {
  getSheet_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DASH_NAME);
  if (sh) {
    sh.getCharts().forEach(function (c) { sh.removeChart(c); });
    sh.clear();
    sh.getDataRange().clearDataValidations();
  } else {
    sh = ss.insertSheet(DASH_NAME);
  }

  var R = "'" + SHEET_NAME + "'!";      // '응답'!
  var ROOM = R + '$B$2:$B';             // 회차 열
  var CRIT = '$I$1';                     // 집계 조건(숨긴 칸)
  var col = function (c) { return R + c + '$2:' + c; }; // 예: '응답'!E$2:E

  // 숨긴 도우미 칸: 회차 목록과 조건
  sh.getRange('H1').setValue('회차 목록');
  sh.getRange('H2').setValue('전체');
  sh.getRange('H3').setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + ROOM + ',' + ROOM + '<>""))),"")');
  sh.getRange('I1').setFormula('=IF(OR($B$1="",$B$1="전체"),"*",$B$1&"")');

  // 회차 선택과 요약 숫자
  sh.getRange('A1:B3').setValues([['회차 선택', '전체'], ['참여 인원', ''], ['평균 점수(' + TOTAL_MAX + '점)', '']]);
  sh.getRange('B1').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getRange('H2:H200'), true).setAllowInvalid(true).build())
    .setBackground('#fff7d6').setFontWeight('bold');
  sh.getRange('B2').setFormula('=COUNTIF(' + ROOM + ',' + CRIT + ')');
  sh.getRange('B3').setFormula('=IFERROR(ROUND(AVERAGEIF(' + ROOM + ',' + CRIT + ',' + col('D') + '),1),0)');

  // 표 1: 레벨 분포 (5~10행)
  var row = 5;
  sh.getRange(row, 1, 1, 2).setValues([['레벨', '인원']]);
  var lvTop = row + 1;
  LEVEL_NAMES.forEach(function (n, i) {
    sh.getRange(lvTop + i, 1).setValue('Lv.' + (i + 1) + ' ' + n);
    sh.getRange(lvTop + i, 2).setFormula('=COUNTIFS(' + ROOM + ',' + CRIT + ',' + col('E') + ',' + (i + 1) + ')');
  });

  // 표 2: 영역별 평균 (12~16행)
  row = 12;
  sh.getRange(row, 1, 1, 2).setValues([['영역', '평균(' + CAT_MAX + '점)']]);
  var catTop = row + 1;
  ['F', 'G', 'H', 'I'].forEach(function (c, i) {
    sh.getRange(catTop + i, 1).setValue(CAT_NAMES[i]);
    sh.getRange(catTop + i, 2).setFormula('=IFERROR(ROUND(AVERAGEIF(' + ROOM + ',' + CRIT + ',' + col(c) + '),1),0)');
  });

  // 표 3: 1번 문항 응답 (19~22행)
  row = 19;
  sh.getRange(row, 1, 1, 2).setValues([['1번: 바이브코딩을 설명할 수 있나요?', '인원']]);
  var q1Top = row + 1;
  Q1_OPTS.forEach(function (t, k) {
    sh.getRange(q1Top + k, 1).setValue(t);
    sh.getRange(q1Top + k, 2).setFormula('=COUNTIFS(' + ROOM + ',' + CRIT + ',' + col('J') + ',' + k + ')');
  });

  // 표 4: 문항별 평균 (25~37행)
  row = 25;
  sh.getRange(row, 1, 1, 2).setValues([['문항', '평균(' + MAX_PT + '점)']]);
  var qTop = row + 1;
  var qCols = ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U'];
  Q_LABELS.forEach(function (t, i) {
    sh.getRange(qTop + i, 1).setValue((i + 1) + '. ' + t);
    sh.getRange(qTop + i, 2).setFormula('=IFERROR(ROUND(AVERAGEIF(' + ROOM + ',' + CRIT + ',' + col(qCols[i]) + '),2),0)');
  });

  // 서식
  [5, 12, 19, 25].forEach(function (r) {
    sh.getRange(r, 1, 1, 2).setFontWeight('bold').setBackground('#1c2140').setFontColor('#ffffff');
  });
  sh.getRange('A1:A3').setFontWeight('bold');
  sh.setColumnWidth(1, 250);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 20);
  sh.hideColumns(8, 2); // H, I 도우미 칸 숨기기

  // 차트 (D열부터 세로로 배치)
  addChart_(sh, Charts.ChartType.COLUMN, sh.getRange(lvTop, 1, 5, 2), 1, '레벨 분포 (명)', null, '#ff5d73');
  addChart_(sh, Charts.ChartType.BAR, sh.getRange(catTop, 1, 4, 2), 20, '영역별 평균 (' + CAT_MAX + '점 만점)', CAT_MAX, '#1fb89a');
  addChart_(sh, Charts.ChartType.BAR, sh.getRange(q1Top, 1, Q1_OPTS.length, 2), 39, '바이브코딩을 설명할 수 있나요? (명)', null, '#0d7f6a');
  addChart_(sh, Charts.ChartType.BAR, sh.getRange(qTop, 1, 12, 2), 58, '문항별 평균 (' + MAX_PT + '점 만점)', MAX_PT, '#35b99c', 460);

  ss.setActiveSheet(sh);
}

function addChart_(sh, type, range, rowPos, title, max, color, height) {
  // 세로 막대는 세로축(vAxis), 가로 막대는 가로축(hAxis)이 값 축입니다.
  var axis = type === Charts.ChartType.COLUMN ? 'vAxis' : 'hAxis';
  var b = sh.newChart()
    .setChartType(type)
    .addRange(range)
    .setPosition(rowPos, 4, 0, 0)
    .setOption('title', title)
    .setOption('legend.position', 'none')
    .setOption('colors', [color])
    .setOption('width', 620)
    .setOption('height', height || 340)
    .setOption(axis + '.viewWindow.min', 0);
  if (max) b.setOption(axis + '.viewWindow.max', max);
  sh.insertChart(b.build());
}
