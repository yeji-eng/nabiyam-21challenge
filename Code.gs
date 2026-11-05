/**
 * 나비얌 × 현대해상 21일 챌린지 — 사진 인증 접수 백엔드
 *
 * 하는 일:
 *   1) 홈페이지에서 보낸 "미션 인증"을 받아서
 *   2) 사진을 내 구글드라이브 비공개 폴더에 저장하고 (외부에 공개되지 않음)
 *   3) 구글시트에 누가·며칠차·무슨 미션·사진링크를 한 줄씩 기록합니다.
 *
 * 설치 방법은 같은 폴더의 "설치안내.md" 파일을 참고하세요.
 */

// 사진이 저장될 드라이브 폴더 이름 (자동으로 만들어집니다. 운영팀만 보이는 비공개 폴더)
var FOLDER_NAME = '현대해상나비얌_챌린지_사진';

// 기록이 쌓일 시트 탭 이름
var SHEET_NAME = '제출현황';

// 시트 첫 줄(머리글)
var HEADERS = ['제출시각', '이름', '연락처뒷자리', '주차', '일차', '카테고리', '미션', '사진링크', '사진파일ID'];

/** 홈페이지에서 인증 사진을 보내면 실행됩니다. */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = JSON.parse(e.postData.contents);

    // 설문 접수: 사전(survey) / 사후(postsurvey)
    if (data.type === 'survey' || data.type === 'postsurvey') {
      var tabName = data.type === 'postsurvey' ? '사후설문' : '사전설문';
      var headColor = data.type === 'postsurvey' ? '#FFE8CC' : '#E3F5EE';
      var sv = getSurveySheet_(tabName, data.survey, headColor);
      var qKeys = Object.keys(data.survey || {});
      var answers = qKeys.map(function (k) { return data.survey[k]; });
      sv.appendRow([new Date(), data.name || '', data.contact || ''].concat(answers));
      return json_({ ok: true, type: data.type });
    }

    // 1) 사진을 비공개 폴더에 저장
    var fileUrl = '', fileId = '';
    if (data.imageBase64) {
      var folder = getFolder_();
      var ext = ((data.imageType || 'image/jpeg').split('/')[1] || 'jpg').replace('jpeg', 'jpg');
      var safeName = String(data.name || '익명').replace(/[^0-9A-Za-z가-힣]/g, '').slice(0, 20) || '익명';
      var fileName = 'D' + pad2_(data.day) + '_' + safeName + '_' + stamp_() + '.' + ext;
      var blob = Utilities.newBlob(Utilities.base64Decode(data.imageBase64), data.imageType || 'image/jpeg', fileName);
      var file = folder.createFile(blob);
      // 공유 설정을 건드리지 않음 = 소유자(운영팀)만 열람·다운로드 가능. 링크가 유출돼도 외부인은 못 봅니다.
      fileId = file.getId();
      fileUrl = file.getUrl();
    }

    // 2) 시트에 한 줄 기록
    var sheet = getSheet_();
    sheet.appendRow([
      new Date(), data.name || '', data.contact || '',
      data.week || '', data.day || '', data.category || '', data.mission || '',
      fileUrl, fileId
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 브라우저에서 주소를 열었을 때 (연결 확인용) */
function doGet() {
  return json_({ ok: true, message: '나비얌 21일 챌린지 접수 서버가 정상 동작 중입니다.' });
}

// ── 도우미 함수들 ──────────────────────────────────────────

function getFolder_() {
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#FFF2DC');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 설문 탭(사전/사후). 질문 개수에 맞춰 머리글(Q1, Q2...)을 자동 생성 */
function getSurveySheet_(tabName, survey, headColor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    var head = ['제출시각', '이름', '연락처뒷자리'].concat(Object.keys(survey || {}));
    sheet.appendRow(head);
    sheet.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground(headColor || '#E3F5EE');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function pad2_(n) { n = parseInt(n, 10) || 0; return (n < 10 ? '0' : '') + n; }

function stamp_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
}

/**
 * (선택) 지표 요약을 자동 계산해 '요약' 탭에 넣고 싶을 때 수동 실행하세요.
 * 참여율·주차별 유지율·미션별 완료율을 간단히 집계합니다.
 */
function buildSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = getSheet_().getDataRange().getValues();
  if (data.length < 2) return;

  var people = {};        // 이름+연락처 → {weeks:{1,2,3}, days:Set, lastDay}
  var missionCount = {};  // 일차 → 제출 수
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var key = (row[1] || '') + '|' + (row[2] || '');
    var week = parseInt(row[3], 10) || 0;
    var day = parseInt(row[4], 10) || 0;
    if (!people[key]) people[key] = { weeks: {}, days: {}, last: 0 };
    people[key].weeks[week] = true;
    people[key].days[day] = true;
    if (day > people[key].last) people[key].last = day;
    missionCount[day] = (missionCount[day] || 0) + 1;
  }

  var names = Object.keys(people);
  var total = names.length;
  var day1 = 0, w1 = 0, w2 = 0, w3 = 0;
  for (var k = 0; k < names.length; k++) {
    var p = people[names[k]];
    if (p.days[1]) day1++;
    if (p.weeks[1]) w1++;
    if (p.weeks[2]) w2++;
    if (p.weeks[3]) w3++;
  }

  var sum = ss.getSheetByName('요약') || ss.insertSheet('요약');
  sum.clear();
  sum.getRange(1, 1, 1, 2).setValues([['지표', '값']]).setFontWeight('bold').setBackground('#FFF2DC');
  var rows = [
    ['총 참여자 수(제출 기준)', total],
    ['최초 참여율(Day1)', pct_(day1, total)],
    ['1주차 유지율', pct_(w1, total)],
    ['2주차 유지율', pct_(w2, total)],
    ['3주차 유지율', pct_(w3, total)]
  ];
  for (var d = 1; d <= 21; d++) rows.push(['Day ' + d + ' 완료 수', missionCount[d] || 0]);
  sum.getRange(2, 1, rows.length, 2).setValues(rows);
  sum.autoResizeColumns(1, 2);
}

function pct_(a, b) { return b ? Math.round((a / b) * 1000) / 10 + '%' : '0%'; }
