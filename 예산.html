// 시트 이름 (스프레드시트 탭 이름) - 본인 환경에 맞게 확인하세요 (보통 '시트1' 또는 'Sheet1')
const SHEET_NAME = '시트1';

// GET 요청 처리 (데이터 조회)
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'read') {
    return handleRead();
  }
  
  return createJsonResponse({ status: 'error', message: 'Invalid action for GET' });
}

// POST 요청 처리 (데이터 삽입, 삭제 등)
function doPost(e) {
  const action = e.parameter.action;
  
  try {
    if (action === 'insert') {
      return handleInsert(e.parameter);
    } else if (action === 'delete') {
      return handleDelete(e.parameter.id);
    } else if (action === 'clearAll') {
      return handleClearAll();
    }
  } catch(err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
  
  return createJsonResponse({ status: 'error', message: 'Invalid action for POST' });
}

// 데이터 읽기
function handleRead() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const result = [];
  
  // 첫 번째 행은 헤더이므로 두 번째 행부터 처리
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowObj = {};
    for (let j = 0; j < headers.length; j++) {
      rowObj[headers[j]] = row[j];
    }
    result.push(rowObj);
  }
  
  // 최신 데이터가 위로 오게 역순으로 정렬 (id가 보통 타임스탬프)
  result.reverse();
  
  return createJsonResponse({ status: 'success', data: result });
}

// 데이터 삽입
function handleInsert(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.appendRow([
    params.id,
    params.member,
    params.month,
    params.category,
    params.amount
  ]);
  
  return createJsonResponse({ status: 'success', message: 'Data inserted successfully' });
}

// 데이터 삭제
function handleDelete(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // 첫 행은 헤더
  for (let i = 1; i < data.length; i++) {
    // 스프레드시트의 값과 파라미터 값을 모두 문자열로 변환하여 비교
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1); // 배열 인덱스는 0부터, 시트 행 번호는 1부터 시작하므로 +1
      return createJsonResponse({ status: 'success', message: 'Row deleted' });
    }
  }
  
  return createJsonResponse({ status: 'error', message: 'ID not found' });
}

// 모든 데이터 삭제 (헤더 제외)
function handleClearAll() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    // 헤더(1행)를 제외하고 삭제
    sheet.deleteRows(2, lastRow - 1);
  }
  
  return createJsonResponse({ status: 'success', message: 'All data cleared' });
}

// JSON 응답 생성 유틸리티 
// 중요: CORS 해결을 위해 setMimeType(ContentService.MimeType.JSON) 사용 필수
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
