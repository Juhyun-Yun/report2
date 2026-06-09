/**
 * 과학탐구 활동 보고서 - Google Apps Script Backend
 *
 * 이 앱은 "프론트엔드(웹 화면)"와 "백엔드(이 스크립트)"가 분리되어 있습니다.
 *   - 프론트엔드: GitHub Pages 등에 공개로 올라간 index.html (모두 같은 주소를 씁니다)
 *   - 백엔드: 선생님 각자의 "구글 시트 사본 + 이 Apps Script + 웹앱 배포 URL"
 *   → 그래서 웹앱 URL을 코드(공개 저장소)에 넣지 않고, 앱의 "선생님 설정" 화면에 붙여넣습니다.
 *
 * [선생님이 할 일 — 요약]
 *  1) 이 가이드 시트를 "사본 만들기"로 복사 (파일 → 사본 만들기)
 *  2) 시트를 열면 상단 메뉴에 [📘 탐구앱] 이 생깁니다 → [① 시트·가이드 만들기] 클릭
 *     (학생명단/활동기록/선생님 가이드 탭이 자동으로 만들어집니다)
 *  3) "학생명단" 탭에 우리 반 학생 이름 입력 (번호·팀은 선택)
 *  4) [확장 프로그램] → [Apps Script] → [배포] → [새 배포] → 유형 [웹 앱]
 *       · 실행: "나" / 액세스 권한: "모든 사용자"
 *  5) 배포 후 나온 "웹 앱 URL" 복사
 *  6) 공개된 앱 주소를 열고 → "선생님 설정" 화면에 그 URL 붙여넣기 → 저장
 *       · 학생에겐  (앱주소)?api=(웹앱URL)  형태 링크를 나눠주면 설정 없이 바로 사용
 *
 *  ※ 코드 수정 후 다시 배포: [배포] → [배포 관리] → 연필 → 버전 "새 버전"
 *  ※ 자세한 단계는 시트의 "📖 선생님 가이드" 탭에 적혀 있습니다.
 */

// (선택) 공개된 앱 주소를 적어두면 "선생님 가이드" 탭의 예시 링크에 자동으로 들어갑니다.
// 예: 'https://myname.github.io/science-report/'
const APP_URL = '';

const ROSTER_SHEET = '학생명단';
const RECORD_SHEET = '활동기록';
const GUIDE_SHEET = '📖 선생님 가이드';
const DRIVE_FOLDER_NAME = '탐구활동_보고서';

const RECORD_HEADERS = [
  '타임스탬프', '번호', '이름', '팀', '모드',
  '활동유형', '제목', '내용', '보고서 PDF',
  '그래프1', '그래프2', '그래프3'
];
const GRAPH_COLS = 3; // 그래프1~그래프3

// 시트를 열면 상단에 [📘 탐구앱] 메뉴를 추가한다.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📘 탐구앱')
    .addItem('① 시트·가이드 만들기', 'setupSheets')
    .addItem('📖 선생님 가이드 다시 만들기', 'createGuideSheet')
    .addToUi();
}

// 학생명단/활동기록/선생님 가이드 시트를 한 번에 준비한다.
function setupSheets() {
  ensureRosterSheet_();
  getOrCreateRecordSheet_();
  createGuideSheet_();
  try {
    SpreadsheetApp.getActive().setActiveSheet(
      SpreadsheetApp.getActive().getSheetByName(GUIDE_SHEET)
    );
    SpreadsheetApp.getUi().alert('준비 완료! "학생명단" 탭에 이름을 입력한 뒤, "📖 선생님 가이드"를 따라 배포하세요.');
  } catch (e) {}
}

// 메뉴에서 직접 호출 (가이드만 다시 생성)
function createGuideSheet() {
  createGuideSheet_();
  try {
    SpreadsheetApp.getActive().setActiveSheet(
      SpreadsheetApp.getActive().getSheetByName(GUIDE_SHEET)
    );
  } catch (e) {}
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getRoster';
    if (action === 'getRoster') return getRoster_();
    return jsonResponse_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    if (action === 'uploadImage') return uploadImage_(body);
    if (action === 'saveReport') return saveReport_(body);
    return jsonResponse_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function getRoster_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ROSTER_SHEET);
  if (!sheet) {
    return jsonResponse_({ ok: false, error: '"' + ROSTER_SHEET + '" 시트가 없습니다.' });
  }
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonResponse_({ ok: true, students: [] });

  const headers = values[0].map(function (h) { return String(h).trim(); });
  let idIdx = headers.indexOf('번호');
  if (idIdx < 0) idIdx = headers.indexOf('학번'); // 이전 명칭도 호환
  const nameIdx = headers.indexOf('이름');
  const teamIdx = headers.indexOf('팀');
  if (nameIdx < 0) {
    return jsonResponse_({ ok: false, error: '학생명단 1행에 "이름" 칼럼이 필요합니다.' });
  }

  const students = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const name = String(row[nameIdx] || '').trim();
    if (!name) continue;
    students.push({
      id: idIdx >= 0 ? String(row[idIdx] || '').trim() : '',
      name: name,
      team: teamIdx >= 0 ? String(row[teamIdx] || '').trim() : ''
    });
  }
  return jsonResponse_({ ok: true, students: students });
}

function uploadImage_(body) {
  if (!body.dataBase64) {
    return jsonResponse_({ ok: false, error: 'dataBase64 누락' });
  }
  const folder = getOrCreateFolder_(DRIVE_FOLDER_NAME);
  const decoded = Utilities.base64Decode(body.dataBase64);
  const mime = body.mimeType || 'image/png';
  const ext = mime === 'application/pdf' ? '.pdf'
            : mime === 'image/jpeg' ? '.jpg'
            : mime === 'image/png' ? '.png'
            : '';
  const filename = (body.filename || ('file_' + Date.now())) + ext;
  const blob = Utilities.newBlob(decoded, mime, filename);
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // 도메인 정책으로 외부 공유가 막혀 있으면 무시 — 선생님은 폴더 안에서 직접 보면 됨
  }
  return jsonResponse_({ ok: true, url: file.getUrl(), id: file.getId() });
}

function saveReport_(body) {
  const sheet = getOrCreateRecordSheet_();
  // 새 형식: reportUrl 단일 값 / 옛 형식: imageUrls 배열도 호환
  let reportLink = '';
  if (body.reportUrl) {
    reportLink = String(body.reportUrl);
  } else if (Array.isArray(body.imageUrls) && body.imageUrls.length) {
    reportLink = body.imageUrls.join('\n');
  }

  // 그래프 이미지 ID 배열 → IMAGE() 수식 셀에 들어감
  const graphIds = Array.isArray(body.graphImageIds) ? body.graphImageIds.slice(0, GRAPH_COLS) : [];
  const graphCells = [];
  for (var gi = 0; gi < GRAPH_COLS; gi++) {
    if (graphIds[gi]) {
      graphCells.push('=IMAGE("https://lh3.googleusercontent.com/d/' + graphIds[gi] + '=w400")');
    } else {
      graphCells.push('');
    }
  }

  sheet.appendRow([
    new Date(),
    body.studentId || '',
    body.studentName || '',
    body.team || '',
    body.mode || '',
    body.activityType || '',
    body.title || '',
    body.content || '',
    reportLink
  ].concat(graphCells));

  // 그래프 셀에 IMAGE 수식이 들어 있으면 행 높이 키워서 잘 보이도록
  if (graphIds.length) {
    var lastRow = sheet.getLastRow();
    sheet.setRowHeight(lastRow, 220);
  }

  return jsonResponse_({ ok: true });
}

function getOrCreateRecordSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(RECORD_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RECORD_SHEET);
    sheet.appendRow(RECORD_HEADERS);
    sheet.getRange(1, 1, 1, RECORD_HEADERS.length)
         .setFontWeight('bold')
         .setBackground('#e8eaf6');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(8, 400);
    sheet.setColumnWidth(9, 300);
    // 그래프1~3 열 폭
    sheet.setColumnWidth(10, 320);
    sheet.setColumnWidth(11, 320);
    sheet.setColumnWidth(12, 320);
  } else {
    // 옛 시트에 그래프 컬럼이 없으면 자동으로 추가
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('그래프1') === -1) {
      const startCol = sheet.getLastColumn() + 1;
      const newHeaders = [['그래프1', '그래프2', '그래프3']];
      sheet.getRange(1, startCol, 1, 3)
           .setValues(newHeaders)
           .setFontWeight('bold')
           .setBackground('#e8eaf6');
      sheet.setColumnWidth(startCol, 320);
      sheet.setColumnWidth(startCol + 1, 320);
      sheet.setColumnWidth(startCol + 2, 320);
    }
  }
  return sheet;
}

// "학생명단" 시트가 없으면 헤더(번호|이름|팀)와 예시 한 줄을 만든다.
function ensureRosterSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(ROSTER_SHEET);
  if (sheet) return sheet;
  sheet = ss.insertSheet(ROSTER_SHEET, 0);
  sheet.getRange(1, 1, 1, 3)
       .setValues([['번호', '이름', '팀']])
       .setFontWeight('bold')
       .setBackground('#e8f5e9');
  // 예시 행 (선생님이 지우고 실제 명단 입력)
  sheet.getRange(2, 1, 2, 3).setValues([
    ['1', '김탐구', '1모둠'],
    ['2', '이실험', '']
  ]);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  return sheet;
}

// 시트 안에 "📖 선생님 가이드" 탭을 만들어 설정 단계를 적어 둔다.
function createGuideSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(GUIDE_SHEET);
  if (sheet) ss.deleteSheet(sheet);          // 항상 최신 내용으로 새로 만든다
  sheet = ss.insertSheet(GUIDE_SHEET, 0);    // 맨 앞 탭으로

  const appUrl = APP_URL && APP_URL.indexOf('http') === 0 ? APP_URL : '(공개된 앱 주소)';
  const sampleLink = (APP_URL && APP_URL.indexOf('http') === 0 ? APP_URL : '(앱주소)')
                     + '?api=(여기에 웹앱 URL)';

  const rows = [
    ['📖 과학탐구활동 보고서 — 선생님 가이드', ''],
    ['', ''],
    ['이 앱은 화면(웹)은 모두가 같은 주소를 쓰고, 자료(시트)는 선생님마다 따로입니다.', ''],
    ['아래 순서대로 한 번만 설정하면 됩니다.', ''],
    ['', ''],
    ['STEP', '할 일'],
    ['0. 사본 준비', '이 스프레드시트를 [파일] → [사본 만들기]로 복사해 "내 시트"로 사용하세요.'],
    ['1. 시트 만들기', '상단 메뉴 [📘 탐구앱] → [① 시트·가이드 만들기] 클릭. (학생명단·활동기록 탭 자동 생성)'],
    ['2. 학생 입력', '"학생명단" 탭에 우리 반 이름 입력. "이름"만 필수, "번호"·"팀"은 선택. 팀이 비면 개인 모드.'],
    ['3. 스크립트 열기', '[확장 프로그램] → [Apps Script] 열기. (Code.gs 코드는 이미 들어 있음)'],
    ['4. 웹앱 배포', '[배포] → [새 배포] → 유형 [웹 앱] 선택.'],
    ['   - 실행 사용자', '"나(본인 계정)"'],
    ['   - 액세스 권한', '"모든 사용자"  ← 학생도 접속해야 하므로 꼭 이렇게!'],
    ['5. 권한 승인', '처음엔 권한 승인 창이 뜹니다. [고급] → [(프로젝트)로 이동] → [허용].'],
    ['6. URL 복사', '배포 완료 화면의 "웹 앱 URL"(.../exec 로 끝남)을 복사.'],
    ['7. 앱에 연결', appUrl + ' 접속 → "선생님 설정" 화면에 그 URL 붙여넣고 [저장].'],
    ['8. 학생 배포', '학생에겐 아래 형태의 링크를 나눠주면 설정 없이 바로 이름 선택으로 들어갑니다.'],
    ['   - 학생용 링크', sampleLink],
    ['', ''],
    ['📌 코드 수정 후 재배포', '[배포] → [배포 관리] → 연필(편집) → 버전 "새 버전" → [배포].'],
    ['📌 제출물 위치', '학생 보고서 PDF·사진은 내 구글 드라이브의 "' + DRIVE_FOLDER_NAME + '" 폴더에 저장됩니다.'],
    ['📌 기록 확인', '학생 제출 내역은 "활동기록" 탭에 한 줄씩 쌓입니다.'],
    ['📌 주의', '웹앱 URL은 내 시트에 쓰기 권한을 주는 주소입니다. 외부에 함부로 공개하지 마세요.'],
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);

  // 서식
  sheet.getRange('A1:B1').merge();
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#3f51b5').setFontColor('#ffffff');
  sheet.getRange('A6:B6').setFontWeight('bold').setBackground('#e8eaf6');
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 640);
  sheet.getRange(1, 1, rows.length, 2)
       .setVerticalAlignment('middle')
       .setWrap(true);
  sheet.setFrozenRows(1);
  return sheet;
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
