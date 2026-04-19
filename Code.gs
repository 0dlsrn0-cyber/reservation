/**
 * 대광로제비앙 방문자 예약 프로그램
 * Google Apps Script 백엔드
 * - 현장별 예약 관리
 * - 시간 예약 (10시~16시)
 * - 자동 만료 처리
 */

// ============================================================
// 설정
// ============================================================
const CONFIG = {
  SPREADSHEET_ID: '',
  SHEET_NAME: '방문예약',
  SETTINGS_SHEET_NAME: '설정'
};

function getSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID !== '') {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ============================================================
// 웹앱 진입점
// ============================================================
function doGet(e) {
  // GitHub Pages 연동: action 파라미터가 있으면 JSON API로 응답
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e);
  }

  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';

  try { getSettingsSheet(); } catch(err) {}
  try { autoExpireReservations(); } catch(err) {}

  let template;
  switch(page) {
    case 'sites':      template = HtmlService.createTemplateFromFile('sites'); break;
    case 'reservation': template = HtmlService.createTemplateFromFile('reservation'); break;
    case 'inquiry':    template = HtmlService.createTemplateFromFile('inquiry'); break;
    case 'admin':      template = HtmlService.createTemplateFromFile('admin'); break;
    default:           template = HtmlService.createTemplateFromFile('index');
  }

  template.scriptUrl = ScriptApp.getService().getUrl();

  return template.evaluate()
    .setTitle('대광로제비앙 방문예약')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function handleApiRequest(e) {
  const action = e.parameter.action;
  let result;
  try {
    switch(action) {
      case 'getConfig':
        result = getConfig();
        break;
      case 'getConfigBySite':
        result = getConfigBySite(e.parameter.site);
        break;
      case 'submitReservation':
        result = submitReservation(JSON.parse(e.parameter.data));
        break;
      case 'getReservation':
        result = getReservation(e.parameter.name, e.parameter.phone);
        break;
      case 'cancelReservation':
        result = cancelReservation(e.parameter.id, e.parameter.name, e.parameter.phone);
        break;
      case 'verifyAdmin':
        result = { success: verifyAdmin(e.parameter.password) };
        break;
      case 'getAllReservations':
        result = getAllReservations(e.parameter.password);
        break;
      case 'deleteReservation':
        result = deleteReservation(e.parameter.password, e.parameter.id);
        break;
      case 'getAllTeams':
        result = { teams: getAllTeams() };
        break;
      case 'getTeamReservations':
        result = getTeamReservations(e.parameter.team, e.parameter.password);
        break;
      default:
        result = { success: false, message: '알 수 없는 action: ' + action };
    }
  } catch(err) {
    result = { success: false, message: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * ★ 초기 설정 함수 ★
 * Apps Script 에디터에서 initSetup 선택 후 ▶ 실행
 */
function initSetup() {
  const sheet = getSheet();
  const settingsSheet = getSettingsSheet();
  Logger.log('✅ 초기 설정 완료!');
  Logger.log('📋 방문예약 시트: ' + sheet.getName());
  Logger.log('⚙️ 설정 시트: ' + settingsSheet.getName());
}

function doPost(e) { return doGet(e); }

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// 스프레드시트 헬퍼
// ============================================================
// 시트 컬럼: 예약번호(0) | 접수일시(1) | 현장(2) | 담당팀(3) | 성명(4) | 휴대폰(5) | 주소(6) | 방문희망일(7) | 방문시간(8) | 상태(9)

function getSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(['예약번호', '접수일시', '현장', '담당팀', '성명', '휴대폰', '주소', '방문희망일', '방문시간', '상태']);
    
    const headerRange = sheet.getRange(1, 1, 1, 10);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1B2A4A');
    headerRange.setFontColor('#FFFFFF');
    
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 140);
    sheet.setColumnWidth(7, 250);
    sheet.setColumnWidth(8, 130);
    sheet.setColumnWidth(9, 100);
    sheet.setColumnWidth(10, 80);
  }
  
  return sheet;
}

// ============================================================
// 시간 값 변환 헬퍼 (스프레드시트가 시간을 Date로 읽을 때 처리)
// ============================================================
function formatVisitTime(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Seoul', 'HH:mm');
  }
  return String(val).trim();
}

// ============================================================
// 설정 시트
// ============================================================
function getSettingsSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SETTINGS_SHEET_NAME);
    
    sheet.getRange('A1').setValue('관리자비밀번호');
    sheet.getRange('B1').setValue('880831');
    
    sheet.getRange('A3').setValue('현장명');
    sheet.getRange('B3').setValue('담당팀1');
    sheet.getRange('C3').setValue('담당팀2');
    sheet.getRange('D3').setValue('담당팀3');
    
    sheet.getRange('A4').setValue('현장1');
    sheet.getRange('B4').setValue('총괄1팀');
    sheet.getRange('C4').setValue('총괄2팀');
    sheet.getRange('D4').setValue('총괄3팀');
    
    sheet.getRange('A5').setValue('현장2');
    sheet.getRange('B5').setValue('영업1팀');
    sheet.getRange('C5').setValue('영업2팀');
    
    sheet.getRange('A6').setValue('현장3');
    sheet.getRange('B6').setValue('관리1팀');
    
    const pw = sheet.getRange('A1');
    pw.setFontWeight('bold'); pw.setBackground('#1B2A4A'); pw.setFontColor('#FFFFFF');
    
    const siteHeader = sheet.getRange('A3:D3');
    siteHeader.setFontWeight('bold'); siteHeader.setBackground('#B8941F'); siteHeader.setFontColor('#FFFFFF');
    
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 120);
    
    sheet.getRange('C1').setValue('← 비밀번호 변경');
    sheet.getRange('C1').setFontColor('#999999');
    sheet.getRange('E3').setValue('← A열: 현장명, B~열: 해당 현장의 담당팀');
    sheet.getRange('E3').setFontColor('#999999');

    // 팀 비밀번호 섹션 (row 8~)
    sheet.getRange('A8').setValue('[팀비밀번호]');
    const teamPwHeader = sheet.getRange('A8');
    teamPwHeader.setFontWeight('bold'); teamPwHeader.setBackground('#1B2A4A'); teamPwHeader.setFontColor('#D4B44A');
    sheet.getRange('C8').setValue('← 각 팀의 조회 비밀번호. 팀명은 위 현장설정과 동일하게 입력');
    sheet.getRange('C8').setFontColor('#999999');

    sheet.getRange('A9').setValue('팀명');
    sheet.getRange('B9').setValue('비밀번호');
    const teamColHeader = sheet.getRange('A9:B9');
    teamColHeader.setFontWeight('bold'); teamColHeader.setBackground('#2C3E5E'); teamColHeader.setFontColor('#FFFFFF');

    sheet.getRange('A10').setValue('총괄1팀'); sheet.getRange('B10').setValue('1111');
    sheet.getRange('A11').setValue('총괄2팀'); sheet.getRange('B11').setValue('2222');
    sheet.getRange('A12').setValue('총괄3팀'); sheet.getRange('B12').setValue('3333');
    sheet.getRange('A13').setValue('영업1팀'); sheet.getRange('B13').setValue('4444');
    sheet.getRange('A14').setValue('영업2팀'); sheet.getRange('B14').setValue('5555');
    sheet.getRange('A15').setValue('관리1팀'); sheet.getRange('B15').setValue('6666');

    sheet.setColumnWidth(5, 350);
  }

  return sheet;
}

function getSites() {
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return ['현장1'];
  const values = sheet.getRange(4, 1, lastRow - 3, 1).getValues();
  const sites = [];
  for (let i = 0; i < values.length; i++) {
    const siteName = String(values[i][0]).trim();
    // 빈 셀, '비밀번호' 포함(마커), '팀명' 헤더 → 팀비밀번호 섹션 도달, 중단
    if (siteName === '' || siteName.includes('비밀번호') || siteName === '팀명') break;
    sites.push(siteName);
  }
  return sites.length > 0 ? sites : ['현장1'];
}

function getTeamsBySite(siteName) {
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return ['담당팀 없음'];
  const data = sheet.getRange(4, 1, lastRow - 3, sheet.getLastColumn()).getValues();
  const teams = [];
  for (let i = 0; i < data.length; i++) {
    const name = String(data[i][0]).trim();
    if (name === '' || name.startsWith('[')) break;
    if (name === String(siteName).trim()) {
      for (let j = 1; j < data[i].length; j++) {
        const teamName = String(data[i][j]).trim();
        if (teamName !== '') teams.push(teamName);
      }
      break;
    }
  }
  return teams.length > 0 ? teams : ['담당팀 없음'];
}

function getAdminPassword() {
  const sheet = getSettingsSheet();
  const password = sheet.getRange('B1').getValue();
  return password ? String(password).trim() : '880831';
}

// ============================================================
// 예약번호 생성
// ============================================================
function generateReservationId() {
  const sheet = getSheet();
  const today = new Date();
  const dateStr = Utilities.formatDate(today, 'Asia/Seoul', 'yyyyMMdd');
  const prefix = 'DK-' + dateStr + '-';
  
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]);
    if (id.startsWith(prefix)) {
      const num = parseInt(id.split('-')[2]);
      if (num > maxNum) maxNum = num;
    }
  }
  return prefix + String(maxNum + 1).padStart(3, '0');
}

// ============================================================
// 만료 자동 처리
// ============================================================
function autoExpireReservations() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][9]);
    if (status === '만료' || status === '취소') continue;
    
    let visitDate = data[i][7];
    if (visitDate instanceof Date) {
      visitDate.setHours(0, 0, 0, 0);
      if (visitDate < today) {
        sheet.getRange(i + 1, 10).setValue('만료');
      }
    } else if (typeof visitDate === 'string' && visitDate) {
      const parts = visitDate.split('-');
      if (parts.length === 3) {
        const vd = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        vd.setHours(0, 0, 0, 0);
        if (vd < today) {
          sheet.getRange(i + 1, 10).setValue('만료');
        }
      }
    }
  }
}

// ============================================================
// 예약 기능
// ============================================================
function submitReservation(formData) {
  try {
    const sheet = getSheet();
    const reservationId = generateReservationId();
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    
    sheet.appendRow([
      reservationId,
      now,
      formData.site,
      formData.team,
      formData.name,
      formData.phone,
      formData.address,
      formData.visitDate,
      formData.visitTime,
      '예약'
    ]);
    
    return { success: true, reservationId: reservationId, message: '예약이 완료되었습니다.' };
  } catch (error) {
    return { success: false, message: '예약 처리 중 오류: ' + error.message };
  }
}

// 예약 조회 (방문자용)
function getReservation(name, phone) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const results = [];
    
    const searchName = String(name).trim();
    const searchPhone = String(phone).trim().replace(/-/g, '');
    
    for (let i = 1; i < data.length; i++) {
      const rowName = String(data[i][4]).trim();
      const rowPhone = String(data[i][5]).trim().replace(/-/g, '');
      
      if (rowName === searchName && rowPhone === searchPhone) {
        // 날짜를 문자열로 변환
        let visitDate = data[i][7];
        if (visitDate instanceof Date) {
          visitDate = Utilities.formatDate(visitDate, 'Asia/Seoul', 'yyyy-MM-dd');
        }
        let datetime = data[i][1];
        if (datetime instanceof Date) {
          datetime = Utilities.formatDate(datetime, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        }
        
        results.push({
          reservationId: String(data[i][0]),
          datetime: String(datetime),
          site: String(data[i][2]),
          team: String(data[i][3]),
          name: String(data[i][4]),
          phone: String(data[i][5]),
          address: String(data[i][6]),
          visitDate: String(visitDate),
          visitTime: String(data[i][8] || ''),
          status: String(data[i][9])
        });
      }
    }
    
    return { success: true, data: results };
  } catch (error) {
    return { success: false, message: '조회 중 오류: ' + error.message };
  }
}

// 방문자 예약 취소
function cancelReservation(reservationId, name, phone) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    
    const searchName = String(name).trim();
    const searchPhone = String(phone).trim().replace(/-/g, '');
    
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]);
      const rowName = String(data[i][4]).trim();
      const rowPhone = String(data[i][5]).trim().replace(/-/g, '');
      
      if (rowId === reservationId && rowName === searchName && rowPhone === searchPhone) {
        sheet.getRange(i + 1, 10).setValue('취소');
        return { success: true, message: '예약이 취소되었습니다.' };
      }
    }
    
    return { success: false, message: '해당 예약을 찾을 수 없습니다.' };
  } catch (error) {
    return { success: false, message: '취소 중 오류: ' + error.message };
  }
}

// 관리자 비밀번호 확인
function verifyAdmin(password) {
  return password === getAdminPassword();
}

// 전체 예약 조회 (관리자용)
function getAllReservations(password) {
  if (!verifyAdmin(password)) {
    return { success: false, message: '비밀번호가 올바르지 않습니다.' };
  }
  
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const results = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === '') continue;
      
      let visitDate = data[i][7];
      if (visitDate instanceof Date) {
        visitDate = Utilities.formatDate(visitDate, 'Asia/Seoul', 'yyyy-MM-dd');
      }
      let datetime = data[i][1];
      if (datetime instanceof Date) {
        datetime = Utilities.formatDate(datetime, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      }
      
      results.push({
        reservationId: String(data[i][0]),
        datetime: String(datetime),
        site: String(data[i][2]),
        team: String(data[i][3]),
        name: String(data[i][4]),
        phone: String(data[i][5]),
        address: String(data[i][6]),
        visitDate: String(visitDate),
        visitTime: String(data[i][8] || ''),
        status: String(data[i][9])
      });
    }
    
    results.reverse();
    return { success: true, data: results };
  } catch (error) {
    return { success: false, message: '조회 중 오류: ' + error.message };
  }
}

// 관리자: 예약 삭제
function deleteReservation(password, reservationId) {
  if (!verifyAdmin(password)) {
    return { success: false, message: '비밀번호가 올바르지 않습니다.' };
  }
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === reservationId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: '삭제되었습니다.' };
      }
    }
    return { success: false, message: '해당 예약을 찾을 수 없습니다.' };
  } catch (error) {
    return { success: false, message: '삭제 중 오류: ' + error.message };
  }
}

// 설정값
function getConfig() {
  return { sites: getSites() };
}

function getConfigBySite(siteName) {
  return { site: siteName, teams: getTeamsBySite(siteName) };
}

// ============================================================
// 담당팀 비밀번호 관리
// ============================================================

/**
 * ★ 기존 시트에 팀비밀번호 섹션 추가 ★
 * 이미 설정 시트가 있는 경우 Apps Script 에디터에서 실행
 */
function setupTeamPasswords() {
  const sheet = getSettingsSheet();
  // 빈 행 찾기 (현장 데이터 아래)
  const lastRow = sheet.getLastRow();
  let insertRow = 8;
  for (let i = 4; i <= lastRow; i++) {
    const v = String(sheet.getRange(i, 1).getValue()).trim();
    if (v === '' || v === '[팀비밀번호]') { insertRow = i + 1; break; }
    insertRow = i + 2;
  }

  sheet.getRange(insertRow, 1).setValue('[팀비밀번호]');
  const h = sheet.getRange(insertRow, 1);
  h.setFontWeight('bold'); h.setBackground('#1B2A4A'); h.setFontColor('#D4B44A');
  sheet.getRange(insertRow, 3).setValue('← 팀명은 위 현장설정과 동일하게 입력');
  sheet.getRange(insertRow, 3).setFontColor('#999999');

  sheet.getRange(insertRow + 1, 1).setValue('팀명');
  sheet.getRange(insertRow + 1, 2).setValue('비밀번호');
  const ch = sheet.getRange(insertRow + 1, 1, 1, 2);
  ch.setFontWeight('bold'); ch.setBackground('#2C3E5E'); ch.setFontColor('#FFFFFF');

  sheet.getRange(insertRow + 2, 1).setValue('팀명예시'); sheet.getRange(insertRow + 2, 2).setValue('1234');
  sheet.setColumnWidth(5, 350);
  Logger.log('✅ 팀비밀번호 섹션이 추가되었습니다. 행 ' + insertRow + '부터 확인하세요.');
}

function getTeamPasswords() {
  const sheet = getSettingsSheet();
  const lastRow = sheet.getLastRow();
  const passwords = {};
  let inSection = false;
  for (let i = 1; i <= lastRow; i++) {
    const a = String(sheet.getRange(i, 1).getValue()).trim();
    if (a === '[팀비밀번호]') { inSection = true; continue; }
    if (inSection && a !== '' && a !== '팀명') {
      passwords[a] = String(sheet.getRange(i, 2).getValue()).trim();
    }
  }
  return passwords;
}

function getAllTeams() {
  return Object.keys(getTeamPasswords());
}

function verifyTeam(teamName, password) {
  const passwords = getTeamPasswords();
  const stored = passwords[String(teamName).trim()];
  return stored !== undefined && stored === String(password).trim();
}

function getTeamReservations(teamName, password) {
  if (!verifyTeam(teamName, password)) {
    return { success: false, message: '비밀번호가 올바르지 않습니다.' };
  }
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const results = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === '') continue;
      if (String(data[i][3]).trim() !== String(teamName).trim()) continue;

      let visitDate = data[i][7];
      if (visitDate instanceof Date) visitDate = Utilities.formatDate(visitDate, 'Asia/Seoul', 'yyyy-MM-dd');
      let datetime = data[i][1];
      if (datetime instanceof Date) datetime = Utilities.formatDate(datetime, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

      results.push({
        reservationId: String(data[i][0]),
        datetime: String(datetime),
        site: String(data[i][2]),
        team: String(data[i][3]),
        name: String(data[i][4]),
        phone: String(data[i][5]),
        address: String(data[i][6]),
        visitDate: String(visitDate),
        visitTime: formatVisitTime(data[i][8]),
        status: String(data[i][9])
      });
    }
    results.sort(function(a, b) { return a.visitDate < b.visitDate ? -1 : 1; });
    return { success: true, data: results, teamName: teamName };
  } catch (error) {
    return { success: false, message: '조회 중 오류: ' + error.message };
  }
}
