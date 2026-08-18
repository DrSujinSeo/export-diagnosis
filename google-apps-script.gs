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
 *
 * [KOTRA 수입규제 경보 연동 (선택)]
 *  - data.go.kr 에서 "수입규제품목(지역본부별) 정보" 활용신청 후 인증키를 발급받아
 *    아래 KOTRA_SERVICE_KEY 값에 붙여넣으세요. (진단앱 결과화면에 수입규제가 자동 표시됨)
 *  - 인증키를 넣지 않으면 수입규제 경보만 표시되지 않고, 나머지는 정상 동작합니다.
 */

// 결과를 저장할 시트(탭) 이름
var SHEET_NAME = '고객DB';

// ─── KOTRA 수입규제품목(지역본부별) Open API 연동 설정 ───
// data.go.kr 에서 발급받은 "인증키(Encoding)"를 아래에 붙여넣으세요. (URL 인코딩된 값 그대로)
var KOTRA_SERVICE_KEY = 'YOUR_DATA_GO_KR_SERVICE_KEY';
var IMPORT_REG_API = 'http://apis.data.go.kr/B410001/DS00000128/getDS00000128';
var REG_CACHE_SHEET = '_수입규제캐시';   // 전체 규제 데이터를 캐시하는 숨김 시트
var REG_CACHE_DAYS = 7;                  // 캐시 갱신 주기(일)

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
  { key: 'fullData',         label: '전체데이터(JSON)' },
  { key: 'govPrograms',      label: '추천정부지원사업' },
  { key: 'hsCode',           label: 'HS코드' },
  { key: 'productCategory',  label: '제품군' }
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
 * GET 요청 처리
 *  - ?hs=3304        → 해당 HS코드에 걸린 수입규제(반덤핑·세이프가드 등) 조회
 *  - ?action=refresh → 규제 데이터 캐시 강제 갱신
 *  - 그 외            → 상태 확인용 응답
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'refresh') {
      var n = refreshRegData();
      return jsonOutput({ result: 'success', refreshed: n });
    }
    if (p.hs) {
      return jsonOutput(queryImportReg(p.hs));
    }
  } catch (err) {
    return jsonOutput({ result: 'error', message: String(err) });
  }
  return jsonOutput({ result: 'ok', message: 'ESG EXPORT 진단 수집 서버가 정상 동작 중입니다.' });
}

/**
 * HS코드로 수입규제 조회 (캐시된 전체 데이터에서 필터)
 */
function queryImportReg(hs) {
  var digits = String(hs).replace(/[^0-9]/g, '');
  if (digits.length < 2) return { result: 'success', count: 0, records: [] };
  var q = digits.slice(0, 6);
  var data = getRegData();
  var out = [];
  for (var i = 0; i < data.length && out.length < 40; i++) {
    var r = data[i];
    if (codeMatches(r.hscd, r.hscdCn, q)) {
      var probe = r.probeTgt || '';
      out.push({
        item: r.cmdlt, hscd: r.hscd, type: r.reglCn,
        regNat: r.isoNat, probeTgt: probe,
        start: r.strDe, end: r.endDe, hq: r.hqurt,
        koreaAffected: (probe.indexOf('한국') >= 0 || probe.indexOf('전세계') >= 0)
      });
    }
  }
  // 한국 영향(조사대상=한국/전세계) 건을 위로 정렬
  out.sort(function (a, b) { return (b.koreaAffected ? 1 : 0) - (a.koreaAffected ? 1 : 0); });
  return { result: 'success', count: out.length, records: out };
}

// HS 코드 매칭: 대표 HSCD(6자리) 또는 HSCD_CN(10자리 목록) 중 하나라도 질의 접두사로 시작하면 true
function codeMatches(hscd, hscdCn, q) {
  var all = ((hscd || '') + ',' + (hscdCn || '')).split(',');
  for (var i = 0; i < all.length; i++) {
    var c = all[i].replace(/[^0-9]/g, '');
    if (c && c.indexOf(q) === 0) return true;
  }
  return false;
}

// 캐시 시트에서 규제 데이터 로드 (없거나 오래되면 API로 갱신)
function getRegData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(REG_CACHE_SHEET);
  var stale = true;
  if (sh && sh.getLastRow() > 1) {
    var ts = Number(PropertiesService.getScriptProperties().getProperty('reg_cache_ts') || 0);
    stale = (new Date().getTime() - ts) > REG_CACHE_DAYS * 86400000;
  }
  if (!sh || sh.getLastRow() <= 1 || stale) refreshRegData();
  sh = ss.getSheetByName(REG_CACHE_SHEET);
  var values = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var v = values[i];
    rows.push({ hqurt: v[0], cmdlt: v[1], hscd: String(v[2]), hscdCn: String(v[3]),
      reglCn: v[4], isoNat: v[5], probeTgt: v[6], strDe: String(v[7]), endDe: String(v[8]) });
  }
  return rows;
}

// KOTRA API에서 전체 규제 데이터를 받아 캐시 시트에 저장
function refreshRegData() {
  var all = [];
  var perPage = 1000, page = 1, total = Infinity;
  while ((page - 1) * perPage < total && page <= 30) {
    var url = IMPORT_REG_API + '?serviceKey=' + KOTRA_SERVICE_KEY +
      '&pageNo=' + page + '&numOfRows=' + perPage + '&type=json';
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(res.getContentText());
    total = Number(json.totalCount || 0);
    var recs = json.records || [];
    for (var i = 0; i < recs.length; i++) {
      var r = recs[i];
      all.push([r.HQURT_NAME || '', r.CMDLT_NAME || '', r.HSCD || '', r.HSCD_CN || '',
        r.REGL_CN || '', r.ISO_WD2_NAT_CD || '', r.PROBE_TGT_NAT_NAME || '',
        r.REGL_STR_DE || '', r.REGL_END_DE || '']);
    }
    if (recs.length < perPage) break;
    page++;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(REG_CACHE_SHEET);
  if (!sh) sh = ss.insertSheet(REG_CACHE_SHEET);
  sh.clear();
  sh.appendRow(['본부', '품목명', 'HSCD', 'HSCD_CN', '규제내용', '규제국', '조사대상국', '시작일', '종료일']);
  if (all.length) sh.getRange(2, 1, all.length, 9).setValues(all);
  try { sh.hideSheet(); } catch (e) {}
  PropertiesService.getScriptProperties().setProperty('reg_cache_ts', String(new Date().getTime()));
  return all.length;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
