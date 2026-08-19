/**
 * ============================================================
 *  방문자 이탈 추적 + 리드 폼 수집 — Google Apps Script 백엔드
 * ============================================================
 *  사용 순서
 *   1) 새 구글 스프레드시트 생성 → 확장 프로그램 > Apps Script
 *   2) 이 파일 전체를 붙여넣고 저장
 *   3) 함수 목록에서 setupDashboard 선택 → 실행 (최초 1회, 권한 승인)
 *   4) 배포 > 새 배포 > 유형: 웹 앱
 *        - 실행 계정: 나
 *        - 액세스 권한: 모든 사용자   ← 반드시 이 옵션이어야 합니다
 *      배포 후 나오는 /exec URL 을 product.html 의
 *      TRACKER_ENDPOINT 상수에 붙여넣으세요.
 * ============================================================
 */

const SHEET_EXIT  = 'Raw_이탈';
const SHEET_LEAD  = 'Raw_폼제출';
const SHEET_DASH  = '대시보드';

const SECTIONS = [
  '1. 히어로', '2. 상품정보', '3. 예상 성과', '4. 액침냉각',
  '5. 상세 스펙', '6. 구매 안내/FAQ', '7. 계산기(게이트)', '8. 신청 폼'
];

const EXIT_HEADERS = [
  '수신시각', '세션ID', '방문자ID', '재방문', '기기', '유입경로',
  'utm_source', 'utm_medium', 'utm_campaign',
  '최대 스크롤(%)', '이탈 섹션', '체류(초)', '실제노출(초)',
  '25%(초)', '50%(초)', '75%(초)', '90%(초)', '100%(초)'
].concat(SECTIONS.map(function (s) { return s + ' 노출(초)'; }))
 .concat(['게이트 도달', '폼 노출', '폼 입력시작', '마지막 입력칸', '입력한 항목', '제출완료',
          '이탈유형', '화면크기', 'UA', '페이지']);

const LEAD_HEADERS = [
  '수신시각', '세션ID', '방문자ID', '성명', '전화번호', '도시', '연령대',
  '관심사', '관심사(기타)', '마케팅 수신동의',
  '기기', '유입경로', 'utm_source', 'utm_medium', 'utm_campaign',
  '최대 스크롤(%)', '체류(초)', '제출 시점 섹션', '페이지'
];

/* ============================================================
   수신 엔드포인트
   ============================================================ */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse(e.postData.contents);
    if (data.type === 'lead') writeLead_(data);
    else writeExit_(data);
    return json_({ ok: true });
  } catch (err) {
    logError_(err, e);
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** 브라우저에서 배포 URL을 직접 열었을 때의 헬스체크 */
function doGet() {
  var ss = SpreadsheetApp.getActive();
  var ex = ss.getSheetByName(SHEET_EXIT);
  var ld = ss.getSheetByName(SHEET_LEAD);
  return json_({
    ok: true,
    message: 'tracker endpoint alive',
    exitRows: ex ? Math.max(0, ex.getLastRow() - 1) : 0,
    leadRows: ld ? Math.max(0, ld.getLastRow() - 1) : 0
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   기록 — 세션ID 기준 upsert (한 방문 = 한 행)
   ============================================================ */
function writeExit_(d) {
  var sh = sheet_(SHEET_EXIT, EXIT_HEADERS);
  var sec = d.sectionSec || {};
  var row = [
    new Date(), d.sid || '', d.vid || '', isReturning_(d.vid, d.sid) ? '재방문' : '신규',
    d.device || '', d.referrer || '',
    d.utm_source || '', d.utm_medium || '', d.utm_campaign || '',
    num_(d.maxScroll), d.lastSection || '', num_(d.dwellSec), num_(d.activeSec),
    num_(d.d25), num_(d.d50), num_(d.d75), num_(d.d90), num_(d.d100)
  ];
  SECTIONS.forEach(function (s) { row.push(num_(sec[s])); });
  row = row.concat([
    num_(d.gateReached), num_(d.formViewed), num_(d.formStarted),
    d.lastField || '', d.fieldsFilled || '', num_(d.submitted),
    d.exitReason || '', d.viewport || '', d.ua || '', d.page || ''
  ]);

  var at = findRowBySid_(sh, d.sid);
  if (at > 0) sh.getRange(at, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

function writeLead_(d) {
  var sh = sheet_(SHEET_LEAD, LEAD_HEADERS);
  sh.appendRow([
    new Date(), d.sid || '', d.vid || '',
    d.name || '', "'" + (d.phone || ''), d.city || '', d.age || '',
    d.interest || '', d.interestEtc || '', d.agreeMarketing || '',
    d.device || '', d.referrer || '',
    d.utm_source || '', d.utm_medium || '', d.utm_campaign || '',
    num_(d.maxScroll), num_(d.dwellSec), d.lastSection || '', d.page || ''
  ]);
  // 이탈 시트에도 '제출완료' 반영
  var ex = SpreadsheetApp.getActive().getSheetByName(SHEET_EXIT);
  if (ex) {
    var at = findRowBySid_(ex, d.sid);
    var col = EXIT_HEADERS.indexOf('제출완료') + 1;
    if (at > 0 && col > 0) ex.getRange(at, col).setValue(1);
  }
}

function findRowBySid_(sh, sid) {
  if (!sid || sh.getLastRow() < 2) return 0;
  var vals = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) if (vals[i][0] === sid) return i + 2;
  return 0;
}

function isReturning_(vid, sid) {
  if (!vid) return false;
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_EXIT);
  if (!sh || sh.getLastRow() < 2) return false;
  var vals = sh.getRange(2, 2, sh.getLastRow() - 1, 2).getValues(); // sid, vid
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][1] === vid && vals[i][0] !== sid) return true;
  }
  return false;
}

function num_(v) { var n = Number(v); return isNaN(n) ? 0 : n; }

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#eef1f4').setWrap(true);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 145);
  }
  return sh;
}

function logError_(err, e) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('_errors') || ss.insertSheet('_errors');
  sh.appendRow([new Date(), String(err),
    (e && e.postData && e.postData.contents) ? e.postData.contents.slice(0, 4000) : '']);
}

/* ============================================================
   대시보드 자동 생성 (최초 1회 실행)
   ============================================================ */
function setupDashboard() {
  var ss = SpreadsheetApp.getActive();
  sheet_(SHEET_EXIT, EXIT_HEADERS);
  sheet_(SHEET_LEAD, LEAD_HEADERS);

  var old = ss.getSheetByName(SHEET_DASH);
  if (old) ss.deleteSheet(old);
  var d = ss.insertSheet(SHEET_DASH, 0);

  d.setHiddenGridlines(true);
  d.setColumnWidths(1, 1, 24);
  d.setColumnWidths(2, 6, 132);
  d.setColumnWidths(8, 1, 24);

  var E = "'" + SHEET_EXIT + "'!";
  var L = "'" + SHEET_LEAD + "'!";

  /* ---- 제목 ---- */
  d.getRange('B2').setValue('방문자 이탈 · 리드 대시보드')
    .setFontSize(20).setFontWeight('bold');
  d.getRange('B3').setFormula(
    '="마지막 갱신: "&TEXT(NOW(),"yyyy-mm-dd hh:mm")&"   ·   자동 갱신(1분)"')
    .setFontColor('#8b95a1').setFontSize(10);

  /* ---- KPI 타일 ---- */
  var kpis = [
    ['총 방문 세션',   '=COUNTA(' + E + 'B2:B)'],
    ['평균 스크롤(%)', '=IFERROR(ROUND(AVERAGE(' + E + 'J2:J),1),0)'],
    ['평균 체류(초)',  '=IFERROR(ROUND(AVERAGE(' + E + 'L2:L),1),0)'],
    ['게이트 도달률',  '=IFERROR(COUNTIF(' + E + 'AA2:AA,1)/COUNTA(' + E + 'B2:B),0)'],
    ['폼 입력 시작률', '=IFERROR(COUNTIF(' + E + 'AC2:AC,1)/COUNTA(' + E + 'B2:B),0)'],
    ['최종 전환율',    '=IFERROR(COUNTA(' + L + 'B2:B)/COUNTA(' + E + 'B2:B),0)']
  ];
  for (var i = 0; i < kpis.length; i++) {
    var c = 2 + i;
    d.getRange(5, c).setValue(kpis[i][0])
      .setFontSize(10).setFontColor('#4e5968').setHorizontalAlignment('center');
    var v = d.getRange(6, c).setFormula(kpis[i][1])
      .setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');
    if (i >= 3) v.setNumberFormat('0.0%').setFontColor('#0a8a5f');
  }
  d.getRange(5, 2, 2, 6)
    .setBackground('#f7f8fa').setBorder(true, true, true, true, true, false, '#e8ebee',
      SpreadsheetApp.BorderStyle.SOLID);
  d.setRowHeight(6, 38);

  /* ---- 섹션별 이탈 분포 ---- */
  d.getRange('B9').setValue('① 어느 섹션에서 이탈했나')
    .setFontSize(13).setFontWeight('bold');
  d.getRange('B10:D10').setValues([['섹션', '이탈 세션', '비중']])
    .setFontWeight('bold').setBackground('#eef1f4');
  for (var s = 0; s < SECTIONS.length; s++) {
    var r = 11 + s;
    d.getRange(r, 2).setValue(SECTIONS[s]);
    d.getRange(r, 3).setFormula('=COUNTIF(' + E + 'K:K,B' + r + ')');
    d.getRange(r, 4).setFormula('=IFERROR(C' + r + '/SUM($C$11:$C$18),0)')
      .setNumberFormat('0.0%');
  }
  d.getRange(11, 3, SECTIONS.length, 1).setHorizontalAlignment('center');

  /* ---- 스크롤 깊이 퍼널 ---- */
  d.getRange('F9').setValue('② 스크롤 깊이 도달률')
    .setFontSize(13).setFontWeight('bold');
  d.getRange('F10:G10').setValues([['구간', '도달률']])
    .setFontWeight('bold').setBackground('#eef1f4');
  var depths = [25, 50, 75, 90, 100];
  for (var k = 0; k < depths.length; k++) {
    var rr = 11 + k;
    d.getRange(rr, 6).setValue(depths[k] + '% 이상');
    d.getRange(rr, 7).setFormula(
      '=IFERROR(COUNTIFS(' + E + 'J:J,">=' + depths[k] + '")/COUNTA(' + E + 'B2:B),0)')
      .setNumberFormat('0.0%');
  }

  /* ---- 전환 퍼널 ---- */
  d.getRange('F17').setValue('③ 전환 퍼널').setFontSize(13).setFontWeight('bold');
  d.getRange('F18:G18').setValues([['단계', '세션']])
    .setFontWeight('bold').setBackground('#eef1f4');
  var funnel = [
    ['방문',        '=COUNTA(' + E + 'B2:B)'],
    ['게이트 도달', '=COUNTIF(' + E + 'AA:AA,1)'],
    ['폼 노출',     '=COUNTIF(' + E + 'AB:AB,1)'],
    ['폼 입력 시작','=COUNTIF(' + E + 'AC:AC,1)'],
    ['제출 완료',   '=COUNTA(' + L + 'B2:B)']
  ];
  for (var f = 0; f < funnel.length; f++) {
    d.getRange(19 + f, 6).setValue(funnel[f][0]);
    d.getRange(19 + f, 7).setFormula(funnel[f][1]).setHorizontalAlignment('center');
  }

  /* ---- 섹션별 평균 머문 시간 ---- */
  d.getRange('B21').setValue('④ 섹션별 평균 머문 시간(초)')
    .setFontSize(13).setFontWeight('bold');
  d.getRange('B22:C22').setValues([['섹션', '평균(초)']])
    .setFontWeight('bold').setBackground('#eef1f4');
  for (var t = 0; t < SECTIONS.length; t++) {
    var r2 = 23 + t;
    var col = colLetter_(19 + t); // S열부터 섹션 노출시간
    d.getRange(r2, 2).setValue(SECTIONS[t]);
    d.getRange(r2, 3).setFormula(
      '=IFERROR(ROUND(AVERAGEIF(' + E + col + '2:' + col + ',">0"),1),0)')
      .setHorizontalAlignment('center');
  }

  /* ---- 폼 이탈 지점 ---- */
  d.getRange('F26').setValue('⑤ 폼에서 멈춘 칸 (미제출)')
    .setFontSize(13).setFontWeight('bold');
  d.getRange('F27').setFormula(
    '=IFERROR(QUERY(' + E + 'AD2:AF, "select A, count(A) where A is not null ' +
    'and C = 0 group by A order by count(A) desc label A \'마지막 입력칸\', ' +
    'count(A) \'세션\'", 0), "아직 데이터가 없습니다")');

  /* ---- 최근 리드 ---- */
  d.getRange('B33').setValue('⑥ 최근 폼 제출 20건')
    .setFontSize(13).setFontWeight('bold');
  d.getRange('B34').setFormula(
    '=IFERROR(QUERY(' + L + 'A2:R, "select A, D, E, F, G, H, J, P, Q ' +
    'where A is not null order by A desc limit 20 ' +
    'label A \'시각\', D \'성명\', E \'전화번호\', F \'도시\', G \'연령대\', ' +
    'H \'관심사\', J \'마케팅동의\', P \'스크롤%\', Q \'체류(초)\'", 0), ' +
    '"아직 제출 데이터가 없습니다")');

  /* ---- 기기 / 유입 ---- */
  d.getRange('B57').setValue('⑦ 기기별').setFontSize(13).setFontWeight('bold');
  d.getRange('B58').setFormula(
    '=IFERROR(QUERY(' + E + 'E2:E, "select A, count(A) where A is not null ' +
    'group by A order by count(A) desc label A \'기기\', count(A) \'세션\'",0),"—")');
  d.getRange('E57').setValue('⑧ 유입 경로').setFontSize(13).setFontWeight('bold');
  d.getRange('E58').setFormula(
    '=IFERROR(QUERY(' + E + 'F2:F, "select A, count(A) where A is not null ' +
    'group by A order by count(A) desc limit 10 label A \'유입\', count(A) \'세션\'",0),"—")');

  buildCharts_(d);
  d.getRange('B2').activate();
  SpreadsheetApp.getActive().toast('대시보드를 생성했습니다.', '완료', 5);
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}

function buildCharts_(d) {
  d.getCharts().forEach(function (c) { d.removeChart(c); });

  d.insertChart(d.newChart().asColumnChart()
    .addRange(d.getRange('B10:C18'))
    .setPosition(9, 9, 0, 0)
    .setOption('title', '섹션별 이탈 세션 수')
    .setOption('colors', ['#4ade9f'])
    .setOption('legend', { position: 'none' })
    .setOption('width', 520).setOption('height', 260)
    .build());

  d.insertChart(d.newChart().asBarChart()
    .addRange(d.getRange('F10:G15'))
    .setPosition(24, 9, 0, 0)
    .setOption('title', '스크롤 깊이 도달률')
    .setOption('colors', ['#7dcbf5'])
    .setOption('legend', { position: 'none' })
    .setOption('width', 520).setOption('height', 240)
    .build());

  d.insertChart(d.newChart().asColumnChart()
    .addRange(d.getRange('F18:G23'))
    .setPosition(38, 9, 0, 0)
    .setOption('title', '전환 퍼널')
    .setOption('colors', ['#8e9ffb'])
    .setOption('legend', { position: 'none' })
    .setOption('width', 520).setOption('height', 240)
    .build());

  d.insertChart(d.newChart().asBarChart()
    .addRange(d.getRange('B22:C30'))
    .setPosition(52, 9, 0, 0)
    .setOption('title', '섹션별 평균 머문 시간(초)')
    .setOption('colors', ['#35e88e'])
    .setOption('legend', { position: 'none' })
    .setOption('width', 520).setOption('height', 260)
    .build());
}

/* ============================================================
   테스트용 더미 데이터 (선택) — 실행 후 대시보드 확인용
   지우려면 clearTestData 실행
   ============================================================ */
function insertTestData() {
  var devices = ['mobile', 'desktop', 'tablet'];
  var refs = ['직접유입', 'google.com', 'naver.com', 'instagram.com'];
  for (var i = 0; i < 40; i++) {
    var scroll = Math.floor(Math.random() * 101);
    var idx = Math.min(SECTIONS.length - 1, Math.floor(scroll / 100 * SECTIONS.length));
    var secSec = {};
    SECTIONS.forEach(function (s, j) { secSec[s] = j <= idx ? Math.round(Math.random() * 40) : 0; });
    var gate = scroll >= 70 ? 1 : 0;
    var started = gate && Math.random() < 0.5 ? 1 : 0;
    writeExit_({
      type: 'exit', sid: 'test-' + i, vid: 'testv-' + (i % 25),
      device: devices[i % 3], referrer: refs[i % 4],
      maxScroll: scroll, lastSection: SECTIONS[idx],
      dwellSec: 20 + Math.round(Math.random() * 300),
      activeSec: 15 + Math.round(Math.random() * 200),
      d25: scroll >= 25 ? 5 : 0, d50: scroll >= 50 ? 20 : 0, d75: scroll >= 75 ? 55 : 0,
      d90: scroll >= 90 ? 90 : 0, d100: scroll >= 100 ? 120 : 0,
      sectionSec: secSec,
      gateReached: gate, formViewed: gate, formStarted: started,
      lastField: started ? ['성명', '전화번호', '도시', '연령대', '관심사', '마케팅 수신'][i % 6] : '',
      fieldsFilled: started ? '성명, 전화번호' : '',
      submitted: 0, exitReason: 'pagehide', viewport: '390x844', ua: 'test', page: '/product.html'
    });
    if (started && Math.random() < 0.4) {
      writeLead_({
        type: 'lead', sid: 'test-' + i, vid: 'testv-' + (i % 25),
        name: '테스트' + i, phone: '010-0000-00' + ('0' + i).slice(-2),
        city: ['서울', '부산', '성남', '대구'][i % 4],
        age: ['40대', '50대', '60대', '30대 이하'][i % 4],
        interest: ['은퇴 후 노후준비', '직장인 투잡 부수익', '목돈 마련', '기타'][i % 4],
        interestEtc: '', agreeMarketing: i % 3 ? '동의합니다' : '동의하지 않습니다',
        device: devices[i % 3], referrer: refs[i % 4],
        maxScroll: scroll, dwellSec: 120, lastSection: SECTIONS[idx], page: '/product.html'
      });
    }
  }
  SpreadsheetApp.getActive().toast('테스트 데이터 40건을 넣었습니다.', '완료', 5);
}

function clearTestData() {
  [SHEET_EXIT, SHEET_LEAD].forEach(function (n) {
    var sh = SpreadsheetApp.getActive().getSheetByName(n);
    if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  });
  SpreadsheetApp.getActive().toast('원시 데이터를 비웠습니다.', '완료', 5);
}

/* 상단 메뉴 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 트래커')
    .addItem('대시보드 생성/재생성', 'setupDashboard')
    .addSeparator()
    .addItem('테스트 데이터 넣기', 'insertTestData')
    .addItem('원시 데이터 비우기', 'clearTestData')
    .addToUi();
}
