/**
 * ============================================================
 * © 2026 GEG 화성(깊이 e끌림). All rights reserved.
 *
 * 본 코드는 「저작권법」에 보호받는 저작물입니다.
 * - 복제권(제16조)·공중송신권(제18조)·배포권(제20조)이
 *   저작권자에게 있습니다.
 * - 정상 경로로 받은 이용자라도 코드의 무단 복제·재배포·
 *   재판매·리브랜딩은 허용되지 않습니다.
 * - 무단 이용 시 「저작권법」 제136조(5년 이하 징역 또는
 *   5천만 원 이하 벌금) 및 제125조(손해배상) 적용 대상이
 *   될 수 있습니다.
 * - 이용 문의: bacusiki777@gmail.com, for2102@jimj.kr
 * ============================================================
 */

// 빌드 서명
const _BUILD_SIG = 'GEGHS-DEEPE-2026';

// 출처 확인용 함수
function getBuildInfo() {
  return {
    sig: _BUILD_SIG,
    owner: 'GEG 화성(깊이 e끌림)',
    year: 2026
  };
}

/***********************************************************************
 * 과학탐구 활동 보고서 — 선생님 시트 연동 스크립트 (Code.gs)
 * ---------------------------------------------------------------------
 * 이 스크립트는 "선생님 본인의 스프레드시트"에 들어 있는 코드입니다.
 * 자기 시트만 읽고 쓰므로, 다른 사람의 데이터에는 접근하지 않습니다.
 *
 * [선생님이 한 번만 하는 일]
 *   1) 이 시트를 "사본 만들기" 해서 내 드라이브에 가져온다.
 *   2) 상단 메뉴 ▸ 탐구활동 보고서 ▸ 📋 탐구활동 보고서 ▸ 초기 설정 을 눌러
 *      "학생 명단"·"활동 기록" 탭을 만든다. (학생 명단은 학생1~학생30이 미리 들어 있음)
 *   3) "학생 명단" 탭에서 이름(B열)을 우리 반 이름으로 바꾼다. (필요하면 팀)
 *   4) 상단 메뉴 ▸ 확장 프로그램 ▸ Apps Script ▸ 배포 ▸ 새 배포
 *        ▸ 유형 "웹 앱" ▸ 실행 "나" ▸ 액세스 "모든 사용자" ▸ 배포
 *        ▸ (처음이면) 권한 검토 ▸ 고급 ▸ "(안전하지 않음)으로 이동"
 *   5) 나온 "웹 앱 URL"(끝이 /exec)을 복사해 앱의 [선생님 설정]에 붙여넣는다.
 *
 * ※ 코드를 수정한 경우에는 [배포 ▸ 배포 관리 ▸ 편집(연필) ▸ 버전 "새 버전" ▸ 배포]
 *   를 해야 변경이 반영됩니다. (수정하지 않으면 그대로 두면 됩니다.)
 ***********************************************************************/

// ── 탭 이름 (모두 한글) ───────────────────────────────────────────
var ROSTER_SHEET = '학생 명단';          // 기본 명단 탭 이름
// 아래 후보 중 "먼저 있는" 탭을 학생 명단으로 사용한다. (이미 만들어 둔 탭과 호환)
var ROSTER_ALIASES = ['학생 명단', '학생명단', '명단'];
var ACTIVITY_SHEET = '활동 기록';        // 제출물이 시간순으로 쌓이는 탭
var FOLDER_PROP_KEY = 'submitFolderId';

// 앱(index.html)의 [선생님 설정] 비밀번호와 똑같이 적어주세요.
var TEACHER_PASSCODE = '1234';

// ── 명단 탭 열 구성 ──────────────────────────────────────────────
// A=번호, B=이름, C=팀(선택), D=최근 활동(자동 기록)
var ROSTER_HEADERS = ['번호', '이름', '팀(선택)', '최근 활동'];
var ROSTER_DEFAULT_COUNT = 30; // 학생1~학생30 미리 채움

// ── 활동 기록 탭 열 순서 ─────────────────────────────────────────
var ACTIVITY_HEADERS = [
  '제출시각', '번호', '이름', '팀', '모드', '활동유형', '제목', '내용', '보고서(PDF)', '그래프'
];

// 명단을 읽을 때 자동으로 걸러낼 안내/예시 이름 (데이터가 아님)
var ROSTER_SKIP_NAMES = ['김과학', '이탐구', '박관찰', '홍길동'];

/*=====================================================================
 * 메뉴 & 초기 셋업
 *====================================================================*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var sub = ui.createMenu('📋 탐구활동 보고서')
    .addItem('초기 설정 (명단·활동기록 탭 만들기)', 'setupSheets');
  ui.createMenu('탐구활동 보고서')
    .addSubMenu(sub)
    .addToUi();
}

// "학생 명단", "활동 기록" 탭을 준비한다.
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var roster = ensureRosterSheet_(ss);
  ensureActivitySheet_(ss);
  cleanupSampleRoster_(ss);       // 예전에 자동 생성된 샘플 "명단" 탭 정리
  removeLegacySubmitSheet_(ss);   // 옛 "제출" 탭은 "활동 기록"으로 옮기고 삭제

  roster.activate();

  SpreadsheetApp.getUi().alert(
    '준비 완료',
    '“' + roster.getName() + '” 탭에서 학생 이름(B열)을 우리 반 이름으로 바꿔 주세요.\n' +
    '제출물은 “' + ACTIVITY_SHEET + '” 탭에 시간순으로 쌓이고, 각 학생의 최신 활동은\n' +
    '명단 탭의 “최근 활동” 칸에 자동으로 표시됩니다.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// 명단으로 쓸 탭을 후보 이름 순서대로 찾는다. (없으면 null)
function getRosterSheet_(ss) {
  for (var i = 0; i < ROSTER_ALIASES.length; i++) {
    var s = ss.getSheetByName(ROSTER_ALIASES[i]);
    if (s) return s;
  }
  return null;
}

// 학생 명단 탭을 준비한다. 이미 있으면 그대로 두고(데이터 보존), 없으면 학생1~학생30으로 채운다.
function ensureRosterSheet_(ss) {
  var sh = getRosterSheet_(ss);
  if (sh) return sh; // 이미 명단이 있으면 손대지 않는다.

  sh = ss.insertSheet(ROSTER_SHEET, 0);

  var rows = [ROSTER_HEADERS.slice()];
  for (var i = 1; i <= ROSTER_DEFAULT_COUNT; i++) {
    rows.push([i, '학생' + i, '', '']);
  }
  sh.getRange(1, 1, rows.length, ROSTER_HEADERS.length).setValues(rows);

  // 머리글 꾸미기
  sh.getRange(1, 1, 1, ROSTER_HEADERS.length)
    .setFontWeight('bold').setBackground('#E8EAF6').setFontColor('#1A237E');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 60);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 320);
  return sh;
}

// 활동 기록 탭을 준비한다.
function ensureActivitySheet_(ss) {
  var sh = ss.getSheetByName(ACTIVITY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ACTIVITY_SHEET);
    sh.getRange(1, 1, 1, ACTIVITY_HEADERS.length).setValues([ACTIVITY_HEADERS]);
    sh.getRange(1, 1, 1, ACTIVITY_HEADERS.length).setFontWeight('bold').setBackground('#E0F2F1');
    sh.setFrozenRows(1);
    sh.setColumnWidth(8, 320); // 내용 열 넓게
  }
  return sh;
}

// "학생 명단"이 있는데 옛 샘플 "명단" 탭(예시 이름만 들어있거나 비어 있음)이 남아 있으면 지운다.
// (실제 데이터가 들어 있으면 절대 지우지 않는다.)
function cleanupSampleRoster_(ss) {
  var real = ss.getSheetByName('학생 명단') || ss.getSheetByName('학생명단');
  var sample = ss.getSheetByName('명단');
  if (!real || !sample) return;
  if (real.getSheetId() === sample.getSheetId()) return;

  var vals = sample.getDataRange().getValues();
  var names = [];
  for (var r = 1; r < vals.length; r++) {
    var n = String(vals[r][1] || '').trim();
    if (n) names.push(n);
  }
  var onlySample = names.length > 0 && names.every(function (n) { return ROSTER_SKIP_NAMES.indexOf(n) >= 0; });
  if (names.length === 0 || onlySample) {
    ss.deleteSheet(sample);
  }
}

// 옛 버전이 만든 "제출" 탭을 정리한다.
//  - "활동 기록"과 내용이 겹치므로, 제출에 쌓인 데이터가 있으면 "활동 기록"으로 옮긴 뒤 "제출" 탭을 삭제한다.
function removeLegacySubmitSheet_(ss) {
  var legacy = ss.getSheetByName('제출');
  if (!legacy) return;

  var activity = ss.getSheetByName(ACTIVITY_SHEET);
  if (activity && legacy.getSheetId() === activity.getSheetId()) return;

  // 제출에 데이터 행(머리글 제외)이 있으면 활동 기록으로 옮긴다.
  var vals = legacy.getDataRange().getValues();
  if (activity && vals.length > 1) {
    var body = vals.slice(1).filter(function (row) {
      return row.some(function (c) { return String(c).trim() !== ''; });
    });
    if (body.length) {
      activity.getRange(activity.getLastRow() + 1, 1, body.length, body[0].length).setValues(body);
    }
  }

  ss.deleteSheet(legacy);
}

/*=====================================================================
 * 웹 요청 처리 (앱이 호출)
 *====================================================================*/
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || '';
  var result;
  try {
    if (action === 'getRoster') {
      result = { ok: true, students: getRoster_() };
    } else if (action === 'ping') {
      result = { ok: true, pong: true };
    } else {
      result = { ok: false, error: '알 수 없는 요청: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: String(err && err.message || err) };
  }
  return reply_(result, p.cb);
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action || '';

    if (action === 'uploadImage') {
      return reply_(uploadImage_(body));
    }
    if (action === 'saveReport') {
      return reply_(saveReport_(body));
    }
    return reply_({ ok: false, error: '알 수 없는 요청: ' + action });
  } catch (err) {
    return reply_({ ok: false, error: String(err && err.message || err) });
  }
}

// 응답을 돌려준다. cb(콜백 이름)가 있으면 스크립트 형태로, 없으면 JSON으로.
function reply_(obj, cb) {
  var s = JSON.stringify(obj);
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + s + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(s)
    .setMimeType(ContentService.MimeType.JSON);
}

/*=====================================================================
 * 기능 구현
 *====================================================================*/

// 명단 읽기 → [{id, name, team}]
function getRoster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getRosterSheet_(ss);
  if (!sh) throw new Error('“학생 명단” 탭이 없어요. 메뉴 ▸ 탐구활동 보고서 ▸ 📋 탐구활동 보고서 ▸ 초기 설정을 눌러 주세요.');

  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  var idx = rosterColumns_(values[0]);

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var name = String(values[r][idx.name] || '').trim();
    if (!name) continue;                              // 이름 없는 행은 건너뜀
    if (isGuideName_(name)) continue;                 // 안내/예시 이름은 데이터가 아님
    out.push({
      id: idx.id >= 0 ? String(values[r][idx.id] || '').trim() : '',
      name: name,
      team: idx.team >= 0 ? String(values[r][idx.team] || '').trim() : ''
    });
  }
  return out;
}

// 머리글에서 번호/이름/팀/최근 활동 열 위치를 찾는다. (열 순서가 달라도 됨)
function rosterColumns_(headerRow) {
  var header = headerRow.map(function (h) { return String(h).replace(/\s/g, ''); });
  var find = function (keys) {
    for (var i = 0; i < header.length; i++) {
      for (var k = 0; k < keys.length; k++) {
        if (header[i].indexOf(keys[k]) >= 0) return i;
      }
    }
    return -1;
  };
  var iName = find(['이름', '성명', 'name']);
  if (iName < 0) iName = 1; // 못 찾으면 두 번째 열을 이름으로 가정
  return {
    id: find(['번호', '학번']),
    name: iName,
    team: find(['팀', '모둠', '조']),
    latest: find(['최근활동', '최근', '활동'])
  };
}

// 안내/예시용으로 적어 둔 이름인지 판단 (데이터가 아니므로 명단에서 제외)
function isGuideName_(name) {
  var n = String(name).trim();
  if (!n) return true;
  if (ROSTER_SKIP_NAMES.indexOf(n) >= 0) return true;
  // ※, ▶, 안내, 예시 등으로 시작하는 줄은 설명용으로 본다.
  if (/^[※▶◀◆●·\-*]/.test(n)) return true;
  if (/^(안내|예시|설명|샘플)/.test(n)) return true;
  return false;
}

// base64 파일을 드라이브에 저장 → {ok, id, url}
function uploadImage_(p) {
  if (!p.dataBase64) throw new Error('파일 데이터가 없습니다.');
  var mime = p.mimeType || 'application/octet-stream';
  var name = sanitize_(p.filename || 'upload');
  if (mime === 'application/pdf' && !/\.pdf$/i.test(name)) name += '.pdf';
  if (mime === 'image/png' && !/\.png$/i.test(name)) name += '.png';

  var bytes = Utilities.base64Decode(p.dataBase64);
  var blob = Utilities.newBlob(bytes, mime, name);

  var folder = getSubmitFolder_();
  var file = folder.createFile(blob);
  // 링크가 있는 사람은 보기 가능 → 시트의 IMAGE()/링크가 정상 표시됨
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

  var id = file.getId();
  var url = (mime === 'application/pdf')
    ? file.getUrl()
    : ('https://drive.google.com/uc?export=view&id=' + id);
  return { ok: true, id: id, url: url };
}

// 제출 기록을 "활동 기록" 탭에 한 줄 추가하고, 학생 명단의 "최근 활동" 칸을 갱신 → {ok}
function saveReport_(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureActivitySheet_(ss);

  var when = new Date();
  var row = [
    when,
    p.studentId || '',
    p.studentName || '',
    p.team || '',
    p.mode || '',
    p.activityType || '',
    p.title || '',
    p.content || '',
    '', // 보고서(PDF) — 아래에서 하이퍼링크 수식으로 채움
    ''  // 그래프 — 아래에서 IMAGE 수식으로 채움
  ];
  sh.appendRow(row);
  var r = sh.getLastRow();

  // 보고서 PDF: 클릭하면 열리는 링크
  if (p.reportUrl) {
    sh.getRange(r, 9).setFormula('=HYPERLINK("' + p.reportUrl + '","PDF 열기")');
  }

  // 그래프 이미지: 여러 장이면 그래프 열 오른쪽으로 한 칸씩 펼쳐서 표시
  var ids = p.graphImageIds || [];
  for (var i = 0; i < ids.length; i++) {
    var imgUrl = 'https://drive.google.com/uc?export=view&id=' + ids[i];
    sh.getRange(r, 10 + i).setFormula('=IMAGE("' + imgUrl + '")');
  }

  // 학생 명단의 "최근 활동" 칸 갱신
  updateRosterLatest_(ss, p, when);

  return { ok: true };
}

// 학생 명단에서 해당 학생 줄을 찾아 "최근 활동" 칸을 채운다.
function updateRosterLatest_(ss, p, when) {
  try {
    var sh = getRosterSheet_(ss);
    if (!sh) return;
    var values = sh.getDataRange().getValues();
    if (values.length < 2) return;

    var idx = rosterColumns_(values[0]);
    var latestCol = idx.latest;
    if (latestCol < 0) {
      // "최근 활동" 열이 없으면 맨 오른쪽에 새로 만든다.
      latestCol = values[0].length;
      sh.getRange(1, latestCol + 1).setValue('최근 활동')
        .setFontWeight('bold').setBackground('#E8EAF6').setFontColor('#1A237E');
    }

    var wantId = String(p.studentId || '').trim();
    var wantName = String(p.studentName || '').trim();

    var found = -1;
    for (var r = 1; r < values.length; r++) {
      var name = String(values[r][idx.name] || '').trim();
      if (!name) continue;
      var id = idx.id >= 0 ? String(values[r][idx.id] || '').trim() : '';
      if (wantId && id && wantId === id && name === wantName) { found = r; break; }
      if (found < 0 && name === wantName) found = r; // 번호가 안 맞으면 이름으로
    }
    if (found < 0) return;

    var stamp = Utilities.formatDate(when, Session.getScriptTimeZone() || 'Asia/Seoul', 'M/d HH:mm');
    var parts = [];
    if (p.activityType) parts.push('[' + p.activityType + ']');
    if (p.title) parts.push(p.title);
    parts.push('(' + stamp + ')');
    sh.getRange(found + 1, latestCol + 1).setValue(parts.join(' '));
  } catch (e) {
    // 최근 활동 갱신 실패는 제출 자체를 막지 않는다.
  }
}

/*=====================================================================
 * 보조 함수
 *====================================================================*/
function getSubmitFolder_() {
  var props = PropertiesService.getDocumentProperties();
  var id = props.getProperty(FOLDER_PROP_KEY);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  var ssName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var folder = DriveApp.createFolder('📁 ' + ssName + ' - 제출물');
  props.setProperty(FOLDER_PROP_KEY, folder.getId());
  return folder;
}

function sanitize_(s) {
  return String(s).replace(/[\\/:*?"<>|]/g, '_').slice(0, 80) || 'file';
}
