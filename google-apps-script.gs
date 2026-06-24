/**
 * ESG EXPORT 수출역량진단 — Google Sheets 수집 백엔드
 * ─────────────────────────────────────────────────────────────
 * 이 코드는 진단 결과를 Google Sheet "고객DB" 시트에 한 행씩 누적 저장합니다.
 *
 * [설치 방법]
 *  1. 결과를 모을 Google Sheet 를 새로 만듭니다.
 *  2. 상단 메뉴 [확장 프로그램] → [Apps Script] 클릭.
 *  3. 기본 코드(Code.gs)를 모두 지우고 이 파일 내용을 그대로 붙여넣습니다.
 *  4. [배포] → [새 배포] → 유형 "웹 앱" 선택.
 *       - 실행 계정: 나(본인)
 *       - 액세스 권한: "모든 사용자" (Anyone)
 *     → [배포] 클릭 후 권한 승인.
 *  5. 생성된 "웹 앱 URL" 을 복사하여 index.html 의
 *       const GOOGLE_SCRIPT_URL = '...'  부분에 붙여넣습니다.
 *  6. 문항/항목을 수정한 뒤에는 index.html 에서 [진단 링크 생성]으로
 *     링크를 만들어 고객에게 전송하면, 응답이 이 시트에 자동 저장됩니다.
 *
 *  ※ 문항을 추가/수정해도 fullData(JSON) 컬럼에 전체 응답이 보존되므로
 *    데이터 유실 없이 누적됩니다.
 */

// 결과를 저장할 시트(탭) 이름
var SHEET_NAME = '고객DB';

// 컬럼 순서 정의 (헤더와 데이터가 동일한 키 순서로 매핑됩니다)
var COLUMNS = [
  { key: 'timestamp',        label: '접수시각' },
  { key: 'date',             label: '진단일자' },
  { key: 'source',           label: '구분' },
  { key: 'companyName',      label: '기업명' },
  { key: 'managerName',      label: '담당자' },
  { key: 'managerPos',       label: '직책' },
  { key: 'managerEmail',     label: '이메일' },
  { key: 'managerPhone',     label: '연락처' },
  { key: 'industry',         label: '업종' },
  { key: 'product',          label: '주요제품' },
  { key: 'revenue',          label: '매출규모' },
  { key: 'exportExp',        label: '수출경험' },
  { key: 'overseas',         label: '해외지사' },
  { key: 'hrTotal',          label: '전체임직원' },
  { key: 'hrExport',         label: '수출전담' },
  { key: 'hrRd',             label: 'R&D' },
  { key: 'hrSales',          label: '해외영업' },
  { key: 'hrMarketing',      label: '해외마케팅' },
  { key: 'hrProduction',     label: '생산품질' },
  { key: 'hrPurchase',       label: '구매SCM' },
  { key: 'totalScore',       label: '종합점수' },
  { key: 'grade',            label: '등급' },
  { key: 'gradeLabel',       label: '등급설명' },
  { key: 'cat0',             label: '경영진리더십' },
  { key: 'cat1',             label: '내부직원역량' },
  { key: 'cat2',             label: '제품경쟁력' },
  { key: 'cat3',             label: '기술경쟁력' },
  { key: 'cat4',             label: '수출인프라' },
  { key: 'cat5',             label: '마케팅역량' },
  { key: 'cat6',             label: '해외시장개척' },
  { key: 'catScores',        label: '영역별점수' },
  { key: 'infraSummary',     label: '인프라요약' },
  { key: 'consultantOpinion',label: '컨설턴트의견' },
  { key: 'fullData',         label: '전체데이터(JSON)' }
];

/**
 * 진단 페이지가 보내는 POST 요청을 처리하여 시트에 한 행 추가
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // 헤더가 없으면 최초 1회 생성
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUMNS.map(function (c) { return c.label; }));
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    }

    // 컬럼 순서에 맞춰 한 행 구성
    var row = COLUMNS.map(function (c) {
      var v = data[c.key];
      return (v === undefined || v === null) ? '' : v;
    });
    sheet.appendRow(row);

    return jsonOutput({ result: 'success', row: sheet.getLastRow() });
  } catch (err) {
    return jsonOutput({ result: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 브라우저로 웹 앱 URL을 직접 열었을 때(GET) 동작 확인용
 */
function doGet() {
  return jsonOutput({ result: 'ok', message: 'ESG EXPORT 진단 수집 서버가 정상 동작 중입니다.' });
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
