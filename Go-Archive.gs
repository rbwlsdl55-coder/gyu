/**
 * Go Tennis — 컬렉션 사이트
 * match_history / format 시트와 실시간 연동되는 대회 목록 사이트
 *
 * [최초 1회 설정]
 * 1. 이 스프레드시트(Go_Tennis_site_collection)를 엽니다.
 * 2. 확장 프로그램 > Apps Script 로 들어갑니다.
 * 3. 기본 생성된 Code.gs 내용을 전부 지우고 이 파일 내용을 붙여넣습니다.
 * 4. 파일 추가(+) > HTML > 파일명 "index" 로 index.html 내용을 붙여넣습니다.
 * 5. 스프레드시트로 돌아와 새로고침하면 상단 메뉴에 "🎾 사이트 관리"가 생깁니다.
 * 6. "① 컬럼 정리 실행 (최초 1회만)" 을 클릭해 컬럼을 정리합니다. (아래 CONFIG 설명 참고)
 * 7. Apps Script 편집기에서 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행: 나(소유자)
 *    - 액세스 권한: 링크가 있는 모든 사용자 (또는 조직 내 모든 사용자)
 * 8. 배포하면 URL이 생성됩니다. 시트를 수정하면 새로고침 시 자동으로 반영됩니다 (실시간 연동).
 *
 * [컬럼 정리가 하는 일] (요청 3,4,7 반영)
 * - match_history 시트의 실제 링크가 들어있던 "URL" 컬럼 값을 "리다이렉트 사이트" 컬럼으로 이동
 * - 실제 링크가 들어있던 "sheet" 컬럼 값을 "리다이렉트 시트" 컬럼으로 이동
 * - 기존에 라벨(예: site_모비스)만 있던 값은 사라지기 전에 "대회명" 컬럼에 보존 (표시용, 사이트에서 이름으로 사용)
 * - 정리가 끝나면 기존 "URL", "sheet" 컬럼은 삭제
 * - 이미 정리된 시트에서 다시 실행해도 안전하도록 헤더를 확인 후 동작합니다.
 */

const CONFIG = {
  MATCH_SHEET: 'match_history',
  FORMAT_SHEET: 'format',
  TIMEZONE: 'Asia/Seoul',
  CACHE_SECONDS: 0 // 0 = 캐시 미사용(항상 실시간). 동시 접속 많으면 15~30으로.
};

/** 스프레드시트를 열 때 관리 메뉴 추가 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎾 사이트 관리')
    .addItem('① 컬럼 정리 실행 (최초 1회만)', 'restructureColumns')
    .addItem('데이터 미리보기 (로그 확인)', 'debugParse')
    .addToUi();
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setTitle('Go Tennis · 대회 컬렉션');
}

/* ───────────────────────── 컬럼 정리 (요청 3, 4, 7) ───────────────────────── */

function restructureColumns() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MATCH_SHEET);
  if (!sheet) { ui.alert('시트를 찾을 수 없습니다: ' + CONFIG.MATCH_SHEET); return; }

  const values = sheet.getDataRange().getValues();
  if (!values.length) { ui.alert('데이터가 없습니다.'); return; }

  const header = values[0].map(h => String(h).trim());
  const idx = name => header.indexOf(name);

  const colDate = idx('날짜');
  let colSite = idx('리다이렉트 사이트');
  let colSheet = idx('리다이렉트 시트');
  const colUrl = idx('URL');
  const colSheetLink = idx('sheet');

  if (colDate === -1 || colSite === -1 || colSheet === -1) {
    ui.alert('필수 컬럼(날짜/리다이렉트 사이트/리다이렉트 시트)을 찾을 수 없습니다. 헤더를 확인해 주세요.');
    return;
  }

  // 이미 정리된 상태(더 이상 URL/sheet 컬럼 없음)라면 중단
  if (colUrl === -1 && colSheetLink === -1) {
    ui.alert('이미 정리된 시트로 보입니다. (URL / sheet 컬럼이 없음)');
    return;
  }

  // "대회명" 컬럼이 없으면 새로 추가 (라벨 값 보존용)
  let colName = idx('대회명');
  if (colName === -1) {
    sheet.insertColumnAfter(colDate + 1);
    sheet.getRange(1, colDate + 2).setValue('대회명');
    colName = colDate + 1; // 0-indexed
    // 컬럼 삽입으로 뒤쪽 인덱스가 밀렸으므로 재계산
    return restructureColumns(); // 안전하게 재귀 재실행 (헤더 다시 읽음)
  }

  const lastRow = sheet.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    const label = sheet.getRange(r, colSite + 1).getValue();
    const url = colUrl !== -1 ? sheet.getRange(r, colUrl + 1).getValue() : '';
    const sheetLabel = sheet.getRange(r, colSheet + 1).getValue();
    const sheetLink = colSheetLink !== -1 ? sheet.getRange(r, colSheetLink + 1).getValue() : '';

    const cleanLabel = String(label || '').replace(/^site_/, '').trim();
    if (cleanLabel && !sheet.getRange(r, colName + 1).getValue()) {
      sheet.getRange(r, colName + 1).setValue(cleanLabel);
    }
    if (url && String(url).trim() !== '' && String(url).trim() !== '-') {
      sheet.getRange(r, colSite + 1).setValue(url);
    }
    if (sheetLink && String(sheetLink).trim() !== '' && String(sheetLink).trim() !== '-') {
      sheet.getRange(r, colSheet + 1).setValue(sheetLink);
    }
  }

  // 기존 URL / sheet 컬럼 삭제 (뒤에서부터 삭제해야 인덱스가 안 꼬임)
  const toDelete = [colUrl, colSheetLink].filter(c => c !== -1).sort((a, b) => b - a);
  toDelete.forEach(c => sheet.deleteColumn(c + 1));

  CacheService.getScriptCache().remove('collectionData');
  ui.alert('컬럼 정리가 완료되었습니다. 이제 웹 앱을 배포(또는 새로고침)하면 됩니다.');
}

/* ───────────────────────── 데이터 조회 (클라이언트에서 호출) ───────────────────────── */

function getCollectionData() {
  const cache = CacheService.getScriptCache();
  if (CONFIG.CACHE_SECONDS > 0) {
    const hit = cache.get('collectionData');
    if (hit) return JSON.parse(hit);
  }
  const data = parseMatchHistory_();
  if (CONFIG.CACHE_SECONDS > 0) {
    cache.put('collectionData', JSON.stringify(data), CONFIG.CACHE_SECONDS);
  }
  return data;
}

/** http(s)로 시작하는 실제 링크인지 확인 */
function isUrl_(v) {
  return !!v && /^https?:\/\//i.test(String(v).trim());
}

/** 'site_모비스' 같은 라벨에서 표시용 이름 추출 (URL이면 추출 불가하므로 빈 문자열) */
function deriveName_(label) {
  const s = String(label || '').trim();
  if (!s || isUrl_(s) || s === '-') return '';
  return s.replace(/^site_/i, '').trim();
}

function parseMatchHistory_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MATCH_SHEET);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + CONFIG.MATCH_SHEET);

  const values = sheet.getDataRange().getValues();
  const header = values[0].map(h => String(h).trim());
  const idx = name => header.indexOf(name);

  const colDate = idx('날짜');
  const colName = idx('대회명');
  const colSite = idx('리다이렉트 사이트');   // 정리 후: URL / 정리 전: 라벨(site_xxx)
  const colSheet = idx('리다이렉트 시트');    // 정리 후: URL / 정리 전: 라벨(sheet_xxx)
  const colUrlLegacy = idx('URL');            // 정리 전에만 존재
  const colSheetLegacy = idx('sheet');        // 정리 전에만 존재

  const norm = v => (v === null || v === undefined) ? '' : String(v).trim();

  // 오늘 00:00 기준 (Asia/Seoul)
  const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const today = new Date(todayStr + 'T00:00:00');

  const upcoming = [];
  const history = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rawDate = row[colDate];
    if (!rawDate) continue;

    const d = parseDateCell_(rawDate);

    const siteCell = colSite !== -1 ? norm(row[colSite]) : '';
    const legacyUrlCell = colUrlLegacy !== -1 ? norm(row[colUrlLegacy]) : '';
    // "컬럼 정리" 실행 여부와 무관하게 실제 URL을 찾아낸다: 리다이렉트 사이트가 이미 링크면 그것을,
    // 아니면(라벨만 있으면) 정리 전 컬럼인 URL을 대신 사용
    const siteUrl = isUrl_(siteCell) ? siteCell : (isUrl_(legacyUrlCell) ? legacyUrlCell : '');

    const sheetCell = colSheet !== -1 ? norm(row[colSheet]) : '';
    const legacySheetCell = colSheetLegacy !== -1 ? norm(row[colSheetLegacy]) : '';
    const sheetUrl = isUrl_(sheetCell) ? sheetCell : (isUrl_(legacySheetCell) ? legacySheetCell : '');

    // 이름: 대회명 컬럼 우선 → 없으면 리다이렉트 사이트(정리 전이면 라벨) → 없으면 sheet 라벨
    let name = colName !== -1 ? norm(row[colName]) : '';
    if (!name) name = deriveName_(siteCell);
    if (!name) name = deriveName_(sheetCell);
    if (!name) name = '(이름 미지정)';

    const item = {
      dateRaw: d ? Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy.MM.dd') : norm(rawDate),
      name: name,
      siteUrl: siteUrl,
      sheetUrl: sheetUrl,
      ready: !!siteUrl
    };

    if (d && d.getTime() < today.getTime()) {
      history.push(item);
    } else {
      upcoming.push(item);
    }
  }

  // upcoming: 날짜 오름차순(임박한 순), history: 날짜 내림차순(최근 순)
  upcoming.sort((a, b) => a.dateRaw.localeCompare(b.dateRaw));
  history.sort((a, b) => b.dateRaw.localeCompare(a.dateRaw));

  return {
    upcoming: upcoming,
    history: history,
    updatedAt: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm:ss')
  };
}

/** '26.02.14' 같은 문자열 또는 Date 객체를 Date로 변환 */
function parseDateCell_(v) {
  if (v instanceof Date) return v;
  const s = String(v).trim();
  const m = s.match(/^(\d{2,4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})$/);
  if (m) {
    let y = Number(m[1]);
    if (y < 100) y += 2000;
    return new Date(y, Number(m[2]) - 1, Number(m[3]));
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function getFormatData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.FORMAT_SHEET);
  if (!sheet) return { items: [] };

  const values = sheet.getDataRange().getValues();
  const items = [];
  values.forEach(row => {
    const label = String(row[0] || '').trim();
    const url = String(row[1] || '').trim();
    if (label && url) items.push({ label: label, url: url });
  });
  return { items: items };
}

/** 디버그용: Apps Script 편집기에서 직접 실행해 파싱 결과를 로그로 확인 */
function debugParse() {
  Logger.log(JSON.stringify(parseMatchHistory_(), null, 2));
  Logger.log(JSON.stringify(getFormatData(), null, 2));
}
