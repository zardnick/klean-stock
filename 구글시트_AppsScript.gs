/**
 * ============================================================
 * 클린 바코드 재고관리 - 구글시트 연동 Apps Script
 * ============================================================
 *
 * 사용 방법:
 *  1. 새 구글시트 생성
 *  2. 메뉴: 확장프로그램 > Apps Script
 *  3. 이 코드 전체 복사하여 붙여넣기 후 저장
 *  4. "배포 > 새 배포 > 유형: 웹 앱"
 *     - 다음 사용자로 실행: 본인
 *     - 액세스 권한: 모든 사용자
 *  5. 배포 후 나오는 "웹 앱 URL"을
 *     재고관리.html → 설정 탭 → 구글시트 URL에 입력
 *
 * 시트는 자동으로 다음 3개가 생성됩니다:
 *  - 상품목록 (products)
 *  - 재고현황 (inventory)
 *  - 거래내역 (history)
 * ============================================================
 */

const SHEET_PRODUCTS  = '상품목록';
const SHEET_HISTORY   = '거래내역';
const SHEET_INVENTORY = '재고현황';

/** 웹 앱 진입점 */
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const action = req.action;
    const data = req.data || {};
    let result;

    switch (action) {
      case 'ping':    result = { ok: true, message: '연결 정상 - ' + new Date().toLocaleString('ko-KR') }; break;
      case 'pushAll': result = pushAll(data); break;
      case 'pullAll': result = pullAll(); break;
      case 'addTx':   result = addTransaction(data); break;
      default: result = { ok: false, message: '알 수 없는 작업: ' + action };
    }
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      message: '클린 바코드 재고관리 시트 연동 서비스가 동작 중입니다.',
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 시트를 가져오거나 생성 */
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#2563eb').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 전체 업로드 (덮어쓰기) */
function pushAll(data) {
  const products = data.products || [];
  const history  = data.history  || [];

  // 상품목록
  const phHeaders = ['ID','상품명','바코드','검색어','기본수량','사무실재고','창고재고','합계','최종업데이트'];
  const ph = getOrCreateSheet(SHEET_PRODUCTS, phHeaders);
  // 옛 버전이면 누락 컬럼 자동 보강
  ensureSchemaColumns(ph);
  if (ph.getLastRow() > 1) ph.getRange(2, 1, ph.getLastRow()-1, ph.getLastColumn()).clearContent();
  if (products.length > 0) {
    const rows = products.map(p => [
      p.id, p.name, p.barcode, p.keywords || '',
      p.defaultQty || 0,
      p.office || 0, p.warehouse || 0,
      (p.office||0)+(p.warehouse||0),
      new Date(),
    ]);
    ph.getRange(2, 1, rows.length, phHeaders.length).setValues(rows);
  }

  // 거래내역
  const hhHeaders = ['일시','상품명','바코드','구분','위치','수량','메모','사무실잔여','창고잔여'];
  const hh = getOrCreateSheet(SHEET_HISTORY, hhHeaders);
  if (hh.getLastRow() > 1) hh.getRange(2, 1, hh.getLastRow()-1, hh.getLastColumn()).clearContent();
  if (history.length > 0) {
    const rows = history.map(h => [
      h.time, h.productName, h.barcode, h.type, h.location,
      h.qty, h.memo || '', h.officeAfter, h.warehouseAfter,
    ]);
    hh.getRange(2, 1, rows.length, hhHeaders.length).setValues(rows);
  }

  // 재고현황 요약
  buildInventorySheet(products);

  return { ok: true, message: `업로드 완료: 상품 ${products.length}개, 거래 ${history.length}건` };
}

/** 전체 다운로드 */
function pullAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ph = ss.getSheetByName(SHEET_PRODUCTS);
  const hh = ss.getSheetByName(SHEET_HISTORY);
  const products = [], history = [];

  if (ph && ph.getLastRow() > 1) {
    ensureSchemaColumns(ph);
    const vals = ph.getRange(2, 1, ph.getLastRow()-1, 9).getValues();
    vals.forEach(r => {
      if (!r[1]) return;
      products.push({
        id: r[0] || ('p' + Date.now()),
        name: r[1],
        barcode: String(r[2] || ''),
        keywords: String(r[3] || ''),
        defaultQty: Number(r[4]) || 0,
        office: Number(r[5]) || 0,
        warehouse: Number(r[6]) || 0,
      });
    });
  }
  if (hh && hh.getLastRow() > 1) {
    const vals = hh.getRange(2, 1, hh.getLastRow()-1, 9).getValues();
    vals.forEach(r => {
      if (!r[1]) return;
      history.push({
        id: 't' + Math.random().toString(36).slice(2),
        time: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(r[0]),
        productName: r[1],
        barcode: String(r[2] || ''),
        type: r[3],
        location: r[4],
        qty: Number(r[5]) || 0,
        memo: r[6] || '',
        officeAfter: Number(r[7]) || 0,
        warehouseAfter: Number(r[8]) || 0,
      });
    });
  }
  return { ok: true, data: { products, history } };
}

/** 단일 거래 추가 (자동 동기화용) */
function addTransaction(payload) {
  const tx = payload.tx;
  const product = payload.product;

  // 거래내역에 한 줄 추가
  const hhHeaders = ['일시','상품명','바코드','구분','위치','수량','메모','사무실잔여','창고잔여'];
  const hh = getOrCreateSheet(SHEET_HISTORY, hhHeaders);
  hh.insertRowAfter(1);
  hh.getRange(2, 1, 1, hhHeaders.length).setValues([[
    tx.time, tx.productName, tx.barcode, tx.type, tx.location,
    tx.qty, tx.memo || '', tx.officeAfter, tx.warehouseAfter,
  ]]);

  // 상품목록의 재고 업데이트 (해당 바코드 찾아 갱신)
  const phHeaders = ['ID','상품명','바코드','검색어','기본수량','사무실재고','창고재고','합계','최종업데이트'];
  const ph = getOrCreateSheet(SHEET_PRODUCTS, phHeaders);
  ensureSchemaColumns(ph);
  const last = ph.getLastRow();
  if (last >= 2) {
    const vals = ph.getRange(2, 1, last-1, phHeaders.length).getValues();
    let found = -1;
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][2]) === String(product.barcode)) { found = i; break; }
    }
    if (found >= 0) {
      // 검색어/기본수량은 시트에 비어있을 때만 새 값으로 채움 (사용자가 시트에서 설정한 값 보존)
      if (!vals[found][3] && product.keywords) ph.getRange(found+2, 4).setValue(product.keywords);
      if (!vals[found][4] && product.defaultQty) ph.getRange(found+2, 5).setValue(product.defaultQty);
      ph.getRange(found+2, 6).setValue(product.office);
      ph.getRange(found+2, 7).setValue(product.warehouse);
      ph.getRange(found+2, 8).setValue(product.office + product.warehouse);
      ph.getRange(found+2, 9).setValue(new Date());
    } else {
      ph.appendRow([product.id, product.name, product.barcode, product.keywords||'', product.defaultQty||0, product.office, product.warehouse, product.office+product.warehouse, new Date()]);
    }
  } else {
    ph.appendRow([product.id, product.name, product.barcode, product.keywords||'', product.defaultQty||0, product.office, product.warehouse, product.office+product.warehouse, new Date()]);
  }

  // 재고현황 시트 갱신
  buildInventorySheetFromProductSheet();

  return { ok: true, message: '거래 추가 완료' };
}

/** 상품 배열로부터 재고현황 시트 작성 */
function buildInventorySheet(products) {
  const headers = ['상품명','바코드','검색어','기본수량','사무실','창고','합계','상태'];
  const sh = getOrCreateSheet(SHEET_INVENTORY, headers);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).clearContent();
  if (!products.length) return;
  const rows = products.map(p => {
    const total = (p.office||0) + (p.warehouse||0);
    const status = total <= 0 ? '품절' : total <= 5 ? '부족' : '정상';
    return [p.name, p.barcode, p.keywords||'', p.defaultQty||0, p.office||0, p.warehouse||0, total, status];
  });
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // 조건부 서식 (상태 컬럼 - 8번)
  const range = sh.getRange(2, 8, rows.length, 1);
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('품절').setBackground('#fecaca').setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('부족').setBackground('#fde68a').setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('정상').setBackground('#bbf7d0').setRanges([range]).build());
  sh.setConditionalFormatRules(rules);
}

/** 상품목록 시트로부터 재고현황 시트 재작성 */
function buildInventorySheetFromProductSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ph = ss.getSheetByName(SHEET_PRODUCTS);
  if (!ph || ph.getLastRow() < 2) return;
  ensureSchemaColumns(ph);
  const vals = ph.getRange(2, 1, ph.getLastRow()-1, 9).getValues();
  const products = vals.map(r => ({
    name: r[1],
    barcode: r[2],
    keywords: r[3] || '',
    defaultQty: Number(r[4])||0,
    office: Number(r[5])||0,
    warehouse: Number(r[6])||0
  })).filter(p => p.name);
  buildInventorySheet(products);
}

/**
 * 상품목록 시트의 스키마를 최신 버전으로 자동 보강.
 * 컬럼 순서: ID, 상품명, 바코드, 검색어, 기본수량, 사무실재고, 창고재고, 합계, 최종업데이트
 * 구버전 시트도 손실 없이 업그레이드 가능.
 */
function ensureSchemaColumns(ph) {
  // 1) 검색어 컬럼이 4번 위치에 없으면 추가
  let headers = ph.getRange(1, 1, 1, Math.max(ph.getLastColumn(), 9)).getValues()[0];
  if (String(headers[3] || '').indexOf('검색어') < 0) {
    ph.insertColumnBefore(4);
    ph.getRange(1, 4).setValue('검색어')
      .setFontWeight('bold').setBackground('#2563eb').setFontColor('#fff');
  }
  // 2) 기본수량 컬럼이 5번 위치에 없으면 추가
  headers = ph.getRange(1, 1, 1, Math.max(ph.getLastColumn(), 9)).getValues()[0];
  if (String(headers[4] || '').indexOf('기본수량') < 0) {
    ph.insertColumnBefore(5);
    ph.getRange(1, 5).setValue('기본수량')
      .setFontWeight('bold').setBackground('#2563eb').setFontColor('#fff');
  }
}

/** 시트 초기화 - 메뉴에서 직접 실행 가능 */
function initializeSheets() {
  const ph = getOrCreateSheet(SHEET_PRODUCTS, ['ID','상품명','바코드','검색어','기본수량','사무실재고','창고재고','합계','최종업데이트']);
  ensureSchemaColumns(ph);
  getOrCreateSheet(SHEET_HISTORY,   ['일시','상품명','바코드','구분','위치','수량','메모','사무실잔여','창고잔여']);
  getOrCreateSheet(SHEET_INVENTORY, ['상품명','바코드','검색어','기본수량','사무실','창고','합계','상태']);
  SpreadsheetApp.getUi().alert('시트 초기화 완료!\n\n[상품목록], [거래내역], [재고현황] 시트가 준비되었습니다.\n검색어 + 기본수량 컬럼이 포함되어 있습니다.');
}

/** 메뉴 등록 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 재고관리')
    .addItem('시트 초기화', 'initializeSheets')
    .addItem('재고현황 다시 계산', 'buildInventorySheetFromProductSheet')
    .addToUi();
}
