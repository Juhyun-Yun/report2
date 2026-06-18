/**
 * ============================================================
 * © 2026 GEG 화성(깊이 e끌림). All rights reserved.
 *
 * 본 코드는 「저작권법」상 보호받는 저작물입니다.
 * - 복제권(제16조)·공중송신권(제18조)·배포권(제20조)은
 *   저작권자에게 있습니다.
 * - 정식 경로로 받은 이용자라도 코드의 무단 복제·재배포·
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
 *   2) "학생 명단" 탭에 번호(A열)·이름(B열)을 적는다. (필요하면 팀)
 *   3) 상단 메뉴 ▸ 확장 프로그램 ▸ Apps Script ▸ 배포 ▸ 새 배포
 *        ▸ 유형 "웹 앱" ▸ 실행 "나" ▸ 액세스 "모든 사용자" ▸ 배포
 *        ▸ (처음이면) 권한 검토 ▸ 고급 ▸ "(안전하지 않음)으로 이동"
 *   4) 나온 "웹 앱 URL"(끝이 /exec)을 복사해 앱의 [선생님 설정]에 붙여넣는다.
 *
 * ※ 코드를 수정한 경우에는 [배포 ▸ 배포 관리 ▸ 편집(연필) ▸ 버전 "새 버전" ▸ 배포]
 *   를 해야 변경이 반영됩니다. (수정하지 않으면 그대로 두면 됩니다.)
 ***********************************************************************/

// ── 시트/폴더 이름 (원하면 바꿔도 됩니다) ──────────────────────────
var ROSTER_SHEET = '학생 명단';        // 기본 명단 탭 이름
// 아래 후보 중 "먼저 있는" 탭을 학생 명단으로 사용한다. (이미 만들어 둔 탭과 호환)
var ROSTER_ALIASES = ['학생 명단', '학생명단', '명단'];
var SUBMIT_SHEET = '제출';   // 제출물이 기록되는 탭 이름 (없으면 자동 생성)
var FOLDER_PROP_KEY = 'submitFolderId';

// 앱(index.html)의 [선생님 설정] 비밀번호 (사용 설명 탭에 표시됩니다)
var TEACHER_PASSCODE = '1234';

// ── 제출 탭의 열 순서 ────────────────────────────────────────────
var SUBMIT_HEADERS = [
  '제출시각', '번호', '이름', '팀', '모드', '활동유형', '제목', '내용', '보고서(PDF)', '그래프'
];

/*=====================================================================
 * 메뉴 & 초기 셋업
 *====================================================================*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var sub = ui.createMenu('📋 탐구활동 보고서')
    .addItem('초기 설정 (사용 설명·명단·제출 탭 만들기)', 'setupSheets');
  ui.createMenu('탐구활동 보고서')
    .addSubMenu(sub)
    .addToUi();
}

// "사용 설명", "학생 명단", "제출" 탭을 준비한다.
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  setupGuideSheet(); // '사용 설명' 탭 생성 (맨 앞으로 이동)

  // 이미 명단 탭(학생 명단/명단 등)이 있으면 그대로 사용하고 건드리지 않는다.
  var roster = getRosterSheet_(ss);
  if (!roster) {
    roster = ss.insertSheet(ROSTER_SHEET, 1);
    roster.getRange(1, 1, 1, 3).setValues([['번호', '이름', '팀(선택)']]);
    roster.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E8EAF6').setFontColor('#1A237E');
    roster.setFrozenRows(1);
    roster.setColumnWidth(2, 140);
  }

  ensureSubmitSheet_(ss);
  cleanupSampleRoster_(ss); // 예전에 자동 생성된 샘플 "명단" 탭 정리

  var guideSheet = ss.getSheetByName('사용 설명');
  if (guideSheet) guideSheet.activate();

  SpreadsheetApp.getUi().alert(
    '준비 완료',
    '"사용 설명" 탭을 확인하여 설정을 완료해 주세요.\n' +
    '학생은 "' + roster.getName() + '" 탭의 명단으로 연동됩니다.',
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
  var SAMPLE = ['김과학', '이탐구', '박관찰'];
  var onlySample = names.length > 0 && names.every(function (n) { return SAMPLE.indexOf(n) >= 0; });
  if (names.length === 0 || onlySample) {
    ss.deleteSheet(sample);
  }
}

/*=====================================================================
 * 사용 설명 탭 생성
 *====================================================================*/
function setupGuideSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var OLD_NAMES = ['📋 사용법', '사용 설명', '📖 선생님 가이드'];
  var NEW_NAME  = '사용 설명';

  // 안전 삭제: 고유 임시 탭을 먼저 생성하여 유일 탭 삭제 오류 방지
  var tempName = '_guide_tmp_' + Date.now();
  var existingTemp = ss.getSheetByName(tempName);
  if (existingTemp) ss.deleteSheet(existingTemp);
  var tmpSheet = ss.insertSheet(tempName);

  for (var d = 0; d < OLD_NAMES.length; d++) {
    var old = ss.getSheetByName(OLD_NAMES[d]);
    if (old) ss.deleteSheet(old);
  }

  tmpSheet.setName(NEW_NAME);
  ss.setActiveSheet(tmpSheet);
  ss.moveActiveSheet(1);
  var sh = tmpSheet;

  sh.setHiddenGridlines(true);
  sh.setTabColor('#1A237E');

  // ── 내용 구성 (2D 배열, 열: A=항목/단계, B=내용/역할, C=주의사항) ──
  var rows = [];
  var pos  = {};

  // 제목
  pos.title = rows.length;
  rows.push(['과학탐구 활동 보고서 — 시트 사용 설명', '', '']);
  rows.push(['', '', '']);

  // 선생님 설정 비밀번호 안내
  pos.passcode = rows.length;
  rows.push(['앱 [선생님 설정] 비밀번호 :   ' + TEACHER_PASSCODE, '', '']);
  pos.passcodeNote = rows.length;
  rows.push(['학생이 선생님 설정에 들어가지 못하게 막는 비밀번호입니다. 학생에게는 알려주지 마세요.', '', '']);
  rows.push(['', '', '']);

  // 섹션 1: 사본을 만든 뒤 설정하기
  pos.sec1 = rows.length;
  rows.push(['1. 사본을 만든 뒤 설정하기', '', '']);
  pos.sec1th = rows.length;
  rows.push(['단계', '내용', '']);
  pos.sec1r0 = rows.length;
  rows.push(['①',
    '"학생 명단" 탭의 예시 내용을 지우고 우리 반 학생의 번호(A열)와 이름(B열)을 입력합니다. 팀·모둠으로 운영한다면 팀(C열)도 입력하세요.',
    '']);
  rows.push(['②',
    '확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹앱 → 액세스: 모든 사용자 → 배포',
    '']);
  rows.push(['③',
    '처음 1회만 진행합니다. 본인이 만든 사본이므로 안전합니다.\n' +
    '⑴ \'승인 필요\'에서 \'권한 검토\'를 선택합니다.\n' +
    '⑵ 본인 계정을 선택합니다.\n' +
    '⑶ \'확인되지 않은 앱\' 화면에서 왼쪽 아래 \'고급\'을 누른 뒤 \'(프로젝트 이름)(으)로 이동\'을 선택합니다.\n' +
    '⑷ 화면 맨 아래의 \'허용\'을 선택합니다.\n' +
    '⑸ 배포된 URL을 확인합니다.\n' +
    '※ \'고급\'이 보이지 않으면 창을 최대화합니다.\n' +
    '※ 반드시 해당 시트의 사본을 만든 계정으로 진행합니다.',
    '']);
  rows.push(['④',
    '학생들이 쓰는 앱 주소를 엽니다.\n' +
    '첫 화면의 [⚙ 선생님 설정] 버튼을 누릅니다. (위에 안내된 비밀번호 입력)\n' +
    '배포된 웹앱 URL을 붙여넣고 [저장]합니다.',
    '']);
  rows.push(['⑤',
    '[저장]하면 설정 화면에 "학생에게 공유할 링크"가 나타납니다. 옆의 [복사] 버튼을 누르세요.\n' +
    '복사된 링크를 학급 게시판·메신저 등으로 학생에게 보냅니다.\n' +
    '학생은 링크만 열면 설정 없이 바로 자기 이름을 고를 수 있습니다.\n' +
    '※ 반드시 "?api=" 가 포함된 이 링크를 공유하세요. 일반 앱 주소(?api= 없는 주소)만 주면 다른 기기에서는 명단이 보이지 않습니다.',
    '']);
  pos.sec1r1 = rows.length - 1;
  rows.push(['', '', '']);

  // 섹션 2: 탭 목록
  pos.sec2 = rows.length;
  rows.push(['2. 탭 목록', '', '']);
  pos.sec2th = rows.length;
  rows.push(['탭 이름', '역할', '주의사항']);
  pos.sec2r0 = rows.length;
  rows.push(['사용 설명',
    '앱 설정 방법과 각 탭의 사용 방법 안내',
    '탭 이름을 변경하거나 삭제하지 마세요.']);
  rows.push(['학생 명단',
    '학생 번호, 이름, 팀 정보를 저장하고 앱 화면의 학생 선택 목록에 연동합니다.',
    '헤더 행(번호·이름·팀)을 수정하지 마세요. 2행부터 학생 정보를 입력하세요. 탭 이름을 변경하면 안 됩니다.']);
  rows.push(['제출',
    '학생이 보고서를 제출하면 제출 시각, 이름, 활동 유형, 제목, 내용 등이 자동으로 기록됩니다.',
    '앱이 자동으로 기록하는 탭입니다. 데이터를 임의로 삭제하거나 열 순서를 변경하지 마세요. 탭 이름을 변경하면 안 됩니다.']);
  pos.sec2r1 = rows.length - 1;
  rows.push(['', '', '']);

  // 섹션 3: 시트의 내용 수정 안내
  pos.sec3 = rows.length;
  rows.push(['3. 시트의 내용 수정 안내', '', '']);
  pos.sec3r = rows.length;
  rows.push([
    '데이터나 설정을 변경할 때는 앱 화면이 아니라 해당 시트 탭에서 직접 수정하세요. ' +
    '탭 이름은 코드와 연결되어 있으므로 삭제하거나 변경하지 마세요.',
    '', '']);
  rows.push(['', '', '']);

  // 섹션 4: 메뉴·버튼 사용법
  pos.sec4 = rows.length;
  rows.push(['4. 메뉴·버튼 사용법', '', '']);
  pos.sec4th = rows.length;
  rows.push(['메뉴 항목', '하는 일', '실행 시점']);
  pos.sec4r0 = rows.length;
  rows.push([
    '탐구활동 보고서 → 📋 탐구활동 보고서 → 초기 설정 (사용 설명·명단·제출 탭 만들기)',
    '"사용 설명", "학생 명단", "제출" 탭을 생성합니다.',
    '처음 설정할 때 또는 안내 탭을 다시 만들 필요가 있을 때']);
  pos.sec4r1 = rows.length - 1;
  rows.push(['', '', '']);

  // 섹션 5: 저작권 안내
  pos.sec5 = rows.length;
  rows.push(['5. 저작권 안내', '', '']);
  pos.sec5r = rows.length;
  rows.push([
    '본 구글 시트 및 관련 자료(앱, 코드, 콘텐츠 포함)의 저작권은 GEG 화성(깊이 e끌림)에게 있습니다.\n\n' +
    '1. 본 자료는 책을 구입한 자에 한해 이용이 허락됩니다(교사일 경우는 해당 학급, 학부모일 경우 자녀). ' +
    '정상 경로로 구매하거나 배포받은 이용자라 하더라도 앱 코드의 무단 수정 및 2차 배포는 허용되지 않습니다.\n\n' +
    '2. 다음 행위를 금합니다.\n' +
    '· 무단 복제·전송·배포·공유(타인에게 시트 링크 또는 사본 전달 포함)\n' +
    '· 영리 목적의 사용 또는 배포(학원에서의 사용 포함)\n' +
    '· 영리 목적의 재판매 또는 재배포\n' +
    '· 무단 수정·편집을 통한 2차적 저작물 작성\n\n' +
    '3. 「저작권법」 제136조(벌칙) 제1항 제1호에 따라, 저작재산권을 복제·공연·공중송신·전시·배포·대여·2차적저작물 작성의 방법으로 침해한 자는 5년 이하의 징역 또는 5천만원 이하의 벌금에 처하거나 이를 병과할 수 있습니다.\n\n' +
    'ⓒ 2026 GEG 화성(깊이 e끌림)',
    '', '']);

  // ── setValues()로 한 번에 입력 ──
  sh.getRange(1, 1, rows.length, 3).setValues(rows);

  // ── 열 너비 ──
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 400);
  sh.setColumnWidth(3, 300);

  // ── 전체 줄바꿈 + 세로 정렬 위 ──
  sh.getRange(1, 1, rows.length, 3).setWrap(true).setVerticalAlignment('top');

  // ── 제목 서식 (14pt, 굵게, 강조색) ──
  sh.getRange(pos.title + 1, 1, 1, 3)
    .merge()
    .setFontSize(14).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#1A237E');

  // ── 비밀번호 강조 서식 ──
  sh.getRange(pos.passcode + 1, 1, 1, 3)
    .merge()
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#FFE082').setFontColor('#5D4037');
  sh.getRange(pos.passcodeNote + 1, 1, 1, 3)
    .merge()
    .setBackground('#FFF8E1').setFontColor('#8A6D3B').setFontSize(9);

  // ── 섹션 헤더 서식 (병합, 굵게, 배경색) ──
  var secList = [pos.sec1, pos.sec2, pos.sec3, pos.sec4, pos.sec5];
  for (var si = 0; si < secList.length; si++) {
    sh.getRange(secList[si] + 1, 1, 1, 3)
      .merge()
      .setFontWeight('bold')
      .setBackground('#C5CAE9').setFontColor('#1A237E');
  }

  // ── 표 헤더 서식 (옅은 배경, 굵게) ──
  sh.getRange(pos.sec1th + 1, 1, 1, 2).setFontWeight('bold').setBackground('#E8EAF6');
  sh.getRange(pos.sec2th + 1, 1, 1, 3).setFontWeight('bold').setBackground('#E8EAF6');
  sh.getRange(pos.sec4th + 1, 1, 1, 3).setFontWeight('bold').setBackground('#E8EAF6');

  // ── 표 전체 테두리 ──
  sh.getRange(pos.sec1th + 1, 1, pos.sec1r1 - pos.sec1th + 1, 2)
    .setBorder(true, true, true, true, true, true);
  sh.getRange(pos.sec2th + 1, 1, pos.sec2r1 - pos.sec2th + 1, 3)
    .setBorder(true, true, true, true, true, true);
  sh.getRange(pos.sec4th + 1, 1, pos.sec4r1 - pos.sec4th + 1, 3)
    .setBorder(true, true, true, true, true, true);

  // ── 섹션 3 내용 셀 3열 병합 ──
  sh.getRange(pos.sec3r + 1, 1, 1, 3).merge();

  // ── 저작권 셀 3열 병합 ──
  sh.getRange(pos.sec5r + 1, 1, 1, 3).merge();
}

/*=====================================================================
 * 웹 요청 처리 (앱이 호출)
 *====================================================================*/
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'getRoster') {
      return json_({ ok: true, students: getRoster_() });
    }
    if (action === 'ping') {
      return json_({ ok: true, pong: true });
    }
    return json_({ ok: false, error: '알 수 없는 요청: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action || '';

    if (action === 'uploadImage') {
      return json_(uploadImage_(body));
    }
    if (action === 'saveReport') {
      return json_(saveReport_(body));
    }
    return json_({ ok: false, error: '알 수 없는 요청: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/*=====================================================================
 * 기능 구현
 *====================================================================*/

// 명단 읽기 → [{id, name, team}]
function getRoster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getRosterSheet_(ss);
  if (!sh) throw new Error('"학생 명단" 탭이 없어요. 메뉴 ▸ 탐구활동 보고서 ▸ 📋 탐구활동 보고서 ▸ 초기 설정을 눌러 주세요.');

  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  // 1행(헤더)에서 번호/이름/팀 열 위치 찾기 (열 순서가 달라도 됨)
  var header = values[0].map(function (h) { return String(h).replace(/\s/g, ''); });
  var col = function (keys) {
    for (var i = 0; i < header.length; i++) {
      for (var k = 0; k < keys.length; k++) {
        if (header[i].indexOf(keys[k]) >= 0) return i;
      }
    }
    return -1;
  };
  var iId = col(['번호', '학번', 'id', 'ID']);
  var iName = col(['이름', '성명', 'name']);
  var iTeam = col(['팀', '모둠', '조', 'team']);
  if (iName < 0) iName = 1; // 못 찾으면 두 번째 열을 이름으로 가정

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var name = iName >= 0 ? String(values[r][iName]).trim() : '';
    if (!name) continue; // 이름 없는 행은 건너뜀
    out.push({
      id: iId >= 0 ? String(values[r][iId]).trim() : '',
      name: name,
      team: iTeam >= 0 ? String(values[r][iTeam]).trim() : ''
    });
  }
  return out;
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

// 제출 기록을 "제출" 탭에 한 줄 추가 → {ok}
function saveReport_(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSubmitSheet_(ss);

  var row = [
    new Date(),
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

  // 그래프 이미지: 여러 장이면 9열(그래프) 오른쪽으로 한 칸씩 펼쳐서 표시
  var ids = p.graphImageIds || [];
  for (var i = 0; i < ids.length; i++) {
    var imgUrl = 'https://drive.google.com/uc?export=view&id=' + ids[i];
    sh.getRange(r, 10 + i).setFormula('=IMAGE("' + imgUrl + '")');
  }
  if (ids.length > 0) sh.setRowHeight(r, 120);

  return { ok: true };
}

/*=====================================================================
 * 보조 함수
 *====================================================================*/
function ensureSubmitSheet_(ss) {
  var sh = ss.getSheetByName(SUBMIT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SUBMIT_SHEET);
    sh.getRange(1, 1, 1, SUBMIT_HEADERS.length).setValues([SUBMIT_HEADERS]);
    sh.getRange(1, 1, 1, SUBMIT_HEADERS.length).setFontWeight('bold').setBackground('#E0F2F1');
    sh.setFrozenRows(1);
    sh.setColumnWidth(8, 320); // 내용 열 넓게
  }
  return sh;
}

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

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
