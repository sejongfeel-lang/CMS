/**
 * MG DB 통합 관리 - 통합 시트 자동 생성
 *
 * 모든 시트(기존 + 앞으로 추가될 시트)에서 이름/학교/학년/연락처를 모아
 * 중복을 제거한 뒤 '통합' 시트를 만든다.
 * 기본 정보 4개 열 뒤에 각 시트명을 열로 두고, 해당 시트에 있으면 체크한다.
 *
 * 설치: 확장 프로그램 > Apps Script 에 이 파일을 붙여넣고
 *       setupTriggers 를 한 번 실행(권한 승인)하면 끝.
 */

const OUTPUT_SHEET = '통합';
// 통합 대상에서 뺄 시트명
const EXCLUDE_SHEETS = [OUTPUT_SHEET];
// 중복 판단 기준: 이름 + 연락처가 같으면 같은 사람.
// 학교·학년은 처음 나온 값을 쓰고, 비어 있으면 다른 시트의 값으로 채운다.
const KEY_FIELDS = ['이름', '연락처'];
// 헤더 행을 찾기 위해 위에서부터 살펴볼 행 수
const HEADER_SCAN_ROWS = 5;

const FIELDS = ['이름', '학교', '학년', '연락처'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('통합 시트')
    .addItem('지금 갱신', 'buildMergedSheet')
    .addItem('자동 갱신 설정', 'setupTriggers')
    .addToUi();
}

/** 시트 추가/변경, 설문 응답, 1시간마다 자동으로 통합 시트를 다시 만든다. */
function setupTriggers() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onChangeHandler' || t.getHandlerFunction() === 'buildMergedSheet')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onChangeHandler').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('buildMergedSheet').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('buildMergedSheet').timeBased().everyHours(1).create();
  buildMergedSheet();
}

function onChangeHandler(e) {
  // 시트 추가·삭제·이름변경(OTHER), 붙여넣기 등 모든 변경에 반응
  buildMergedSheet();
}

function buildMergedSheet() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30 * 1000)) return;
  try {
    const ss = SpreadsheetApp.getActive();
    const sources = ss.getSheets().filter(s => EXCLUDE_SHEETS.indexOf(s.getName()) === -1);
    const sourceNames = [];
    const people = new Map(); // key -> { info: [이름, 학교, 학년, 연락처], sheets: Set }

    sources.forEach(sheet => {
      const values = sheet.getDataRange().getDisplayValues();
      const header = findHeader(values);
      if (!header) {
        console.log('기본 정보 열을 찾지 못해 건너뜀: ' + sheet.getName());
        return;
      }
      const name = sheet.getName();
      sourceNames.push(name);
      for (let r = header.row + 1; r < values.length; r++) {
        const row = values[r];
        const info = {
          '이름': cleanText(row[header.cols['이름']]),
          '학교': cleanText(header.cols['학교'] == null ? '' : row[header.cols['학교']]),
          '학년': cleanText(header.cols['학년'] == null ? '' : row[header.cols['학년']]),
          '연락처': normalizePhone(row[header.cols['연락처']]),
        };
        if (!info['이름'] && !info['연락처']) continue;
        const key = KEY_FIELDS.map(f => f === '학교' ? normalizeSchool(info[f]) : info[f]).join('|');
        let person = people.get(key);
        if (!person) {
          person = { info: info, sheets: new Set() };
          people.set(key, person);
        } else {
          FIELDS.forEach(f => {
            if (!person.info[f] && info[f]) person.info[f] = info[f];
          });
        }
        person.sheets.add(name);
      }
    });

    const header = FIELDS.concat(sourceNames);
    const rows = [];
    people.forEach(p => {
      rows.push(FIELDS.map(f => p.info[f]).concat(sourceNames.map(n => p.sheets.has(n))));
    });

    let out = ss.getSheetByName(OUTPUT_SHEET);
    if (!out) out = ss.insertSheet(OUTPUT_SHEET, 0);
    out.clear();
    out.getRange(1, 1, out.getMaxRows(), out.getMaxColumns()).clearDataValidations();

    const numCols = header.length;
    if (out.getMaxColumns() < numCols) out.insertColumnsAfter(out.getMaxColumns(), numCols - out.getMaxColumns());
    if (out.getMaxRows() < rows.length + 1) out.insertRowsAfter(out.getMaxRows(), rows.length + 1 - out.getMaxRows());

    out.getRange(1, 1, 1, numCols).setValues([header]).setFontWeight('bold').setBackground('#e8eaed');
    out.setFrozenRows(1);
    out.setFrozenColumns(FIELDS.length);
    if (rows.length) {
      // 연락처의 앞자리 0이 사라지지 않도록 텍스트 서식
      out.getRange(2, 1, rows.length, FIELDS.length).setNumberFormat('@');
      out.getRange(2, 1, rows.length, numCols).setValues(rows);
      if (sourceNames.length) {
        out.getRange(2, FIELDS.length + 1, rows.length, sourceNames.length).insertCheckboxes();
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/** 위쪽 몇 행 중 이름·연락처 열이 모두 있는 행을 헤더로 본다. */
function findHeader(values) {
  const limit = Math.min(HEADER_SCAN_ROWS, values.length);
  for (let r = 0; r < limit; r++) {
    const cols = {};
    values[r].forEach((cell, c) => {
      const h = normalizeHeader(cell);
      if (!h || h.indexOf('개인정보') !== -1 || h.indexOf('동의') !== -1) return;
      FIELDS.forEach(f => {
        if (cols[f] == null && matchField(f, h)) cols[f] = c;
      });
    });
    if (cols['이름'] != null && cols['연락처'] != null) return { row: r, cols: cols };
  }
  return null;
}

function matchField(field, h) {
  switch (field) {
    case '이름': return h.indexOf('이름') !== -1 || h === '성명';
    case '학교': return /^학교(명)?/.test(h);
    case '학년': return h.indexOf('학년') !== -1;
    case '연락처': return /연락처|전화|휴대폰|핸드폰/.test(h);
  }
  return false;
}

/** '▣  학생 이름', 'Q. 학교명을 입력해 주세요.' 같은 헤더를 비교하기 쉽게 정리 */
function normalizeHeader(s) {
  return String(s).replace(/^[\s▣■□●◆\-]*(Q\.)?/, '').replace(/\[[^\]]*\]/g, '').replace(/\s+/g, '');
}

function cleanText(s) {
  return String(s == null ? '' : s).trim().replace(/\s+/g, ' ');
}

/** 숫자만 남기고 010-0000-0000 형태로. 앞 0이 빠진 1012345678 도 복원 */
function normalizePhone(s) {
  let d = String(s == null ? '' : s).replace(/\D/g, '');
  if (d.length === 10 && d.charAt(0) === '1') d = '0' + d;
  if (d.length === 11 && d.indexOf('01') === 0) return d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
  return d || cleanText(s);
}

/** 중복 판단용: '을숙도초등학교' 와 '을숙도초' 를 같게 본다 */
function normalizeSchool(s) {
  return s.replace(/\s+/g, '')
    .replace(/초등학교$/, '초')
    .replace(/중학교$/, '중')
    .replace(/고등학교$/, '고');
}
