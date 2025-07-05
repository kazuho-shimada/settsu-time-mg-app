// 勤怠管理システム - メインコード
// スプレッドシートID（実際のIDに変更してください）
const SPREADSHEET_ID = '1Y2HJPknu0XhGRr2keVGei_b_X1Qn6Mmb3yj6VGBFNfw';

// シート名定数
const SHEETS = {
  EMPLOYEES: '社員マスタ',
  ATTENDANCE: '勤怠記録',
  WORK_RECORDS: '作業記録',
  WORK_TYPES: '作業項目マスタ'
};

// グローバル変数
let spreadsheet;

// 初期化処理
function initialize() {
  try {
    spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    setupSheets();
  } catch (e) {
    console.error('初期化エラー:', e);
  }
}

// シートの設定
function setupSheets() {
  // 社員マスタ
  const employeesSheet = getOrCreateSheet(SHEETS.EMPLOYEES);
  if (employeesSheet.getLastRow() === 0) {
    employeesSheet.appendRow(['社員ID', '社員名']);
    employeesSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f3f3f3');
    // サンプルデータ
    employeesSheet.appendRow(['E001', '山田太郎']);
    employeesSheet.appendRow(['E002', '鈴木花子']);
    employeesSheet.appendRow(['E003', '田中一郎']);
  }
  
  // 勤怠記録
  const attendanceSheet = getOrCreateSheet(SHEETS.ATTENDANCE);
  if (attendanceSheet.getLastRow() === 0) {
    attendanceSheet.appendRow(['日付', '社員ID', '社員名', '出勤時間', '退勤時間', '備考']);
    attendanceSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  // 作業記録
  const workRecordsSheet = getOrCreateSheet(SHEETS.WORK_RECORDS);
  if (workRecordsSheet.getLastRow() === 0) {
    workRecordsSheet.appendRow(['日付', '社員ID', '社員名', '作業項目', 'ロット番号', '数量']);
    workRecordsSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  // 作業項目マスタ
  const workTypesSheet = getOrCreateSheet(SHEETS.WORK_TYPES);
  if (workTypesSheet.getLastRow() === 0) {
    workTypesSheet.appendRow(['作業ID', '作業名']);
    workTypesSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f3f3f3');
    // サンプルデータ
    workTypesSheet.appendRow(['W001', '溶接']);
    workTypesSheet.appendRow(['W002', '塗装']);
    workTypesSheet.appendRow(['W003', '組立']);
    workTypesSheet.appendRow(['W004', '検査']);
  }
}

// シートの取得または作成
function getOrCreateSheet(sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

// Webアプリのエントリーポイント
function doGet(e) {
  initialize();
  
  const params = e.parameter || {};
  
  // 管理画面
  if (params.page === 'admin') {
    return HtmlService.createTemplateFromFile('admin')
      .evaluate()
      .setTitle('管理画面')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // メイン画面
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('勤怠管理システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// HTMLファイルのインクルード
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// === データ取得系関数 ===

// 社員一覧取得
function getEmployees() {
  try {
    initialize();
    const sheet = spreadsheet.getSheetByName(SHEETS.EMPLOYEES);
    const data = sheet.getDataRange().getValues();
    
    const employees = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        employees.push({
          id: String(data[i][0]),
          name: String(data[i][1])
        });
      }
    }
    
    return employees;
  } catch (e) {
    console.error('社員取得エラー:', e);
    return [];
  }
}

// 作業項目一覧取得
function getWorkTypes() {
  try {
    initialize();
    const sheet = spreadsheet.getSheetByName(SHEETS.WORK_TYPES);
    const data = sheet.getDataRange().getValues();
    
    const workTypes = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        workTypes.push({
          id: String(data[i][0]),
          name: String(data[i][1])
        });
      }
    }
    
    return workTypes;
  } catch (e) {
    console.error('作業項目取得エラー:', e);
    return [];
  }
}

// 当日の勤怠状況取得
function getTodayAttendance(employeeId) {
  try {
    initialize();
    const sheet = spreadsheet.getSheetByName(SHEETS.ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === today && data[i][1] === employeeId) {
        return {
          hasCheckedIn: !!data[i][3],
          hasCheckedOut: !!data[i][4],
          checkInTime: data[i][3] || '',
          checkOutTime: data[i][4] || ''
        };
      }
    }
    
    return {
      hasCheckedIn: false,
      hasCheckedOut: false,
      checkInTime: '',
      checkOutTime: ''
    };
  } catch (e) {
    console.error('勤怠状況取得エラー:', e);
    return { hasCheckedIn: false, hasCheckedOut: false };
  }
}

// === 勤怠記録系関数 ===

// 出勤記録
function recordCheckIn(employeeId) {
  try {
    initialize();
    
    // 既に出勤済みかチェック
    const attendance = getTodayAttendance(employeeId);
    if (attendance.hasCheckedIn) {
      return { success: false, message: '本日は既に出勤記録があります' };
    }
    
    const employee = getEmployees().find(emp => emp.id === employeeId);
    if (!employee) {
      return { success: false, message: '社員が見つかりません' };
    }
    
    const now = new Date();
    const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
    const time = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');
    
    const sheet = spreadsheet.getSheetByName(SHEETS.ATTENDANCE);
    sheet.appendRow([today, employeeId, employee.name, time, '', '']);
    
    return { 
      success: true, 
      message: `${employee.name}さんの出勤を記録しました（${time}）`
    };
  } catch (e) {
    console.error('出勤記録エラー:', e);
    return { success: false, message: 'エラーが発生しました' };
  }
}

// 退勤記録
function recordCheckOut(employeeId) {
  try {
    initialize();
    
    const employee = getEmployees().find(emp => emp.id === employeeId);
    if (!employee) {
      return { success: false, message: '社員が見つかりません' };
    }
    
    const now = new Date();
    const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
    const time = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');
    
    const sheet = spreadsheet.getSheetByName(SHEETS.ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    
    // 今日の出勤記録を探して退勤時間を更新
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === today && data[i][1] === employeeId) {
        sheet.getRange(i + 1, 5).setValue(time);
        return { 
          success: true, 
          message: `${employee.name}さんの退勤を記録しました（${time}）`
        };
      }
    }
    
    // 出勤記録がない場合は退勤のみ記録
    sheet.appendRow([today, employeeId, employee.name, '', time, '退勤のみ']);
    
    return { 
      success: true, 
      message: `${employee.name}さんの退勤を記録しました（${time}）`
    };
  } catch (e) {
    console.error('退勤記録エラー:', e);
    return { success: false, message: 'エラーが発生しました' };
  }
}

// === 作業記録系関数 ===

// 作業記録保存
function saveWorkRecord(employeeId, workTypeId, lotNumber, quantity) {
  try {
    initialize();
    
    const employee = getEmployees().find(emp => emp.id === employeeId);
    const workType = getWorkTypes().find(wt => wt.id === workTypeId);
    
    if (!employee || !workType) {
      return { success: false, message: 'データが見つかりません' };
    }
    
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    
    const sheet = spreadsheet.getSheetByName(SHEETS.WORK_RECORDS);
    sheet.appendRow([today, employeeId, employee.name, workType.name, lotNumber, quantity]);
    
    return { 
      success: true, 
      message: '作業記録を保存しました'
    };
  } catch (e) {
    console.error('作業記録保存エラー:', e);
    return { success: false, message: 'エラーが発生しました' };
  }
}

// === 管理機能 ===

// 社員追加
function addEmployee(id, name) {
  try {
    initialize();
    const sheet = spreadsheet.getSheetByName(SHEETS.EMPLOYEES);
    
    // 重複チェック
    const employees = getEmployees();
    if (employees.some(emp => emp.id === id)) {
      return { success: false, message: 'その社員IDは既に使用されています' };
    }
    
    sheet.appendRow([id, name]);
    return { success: true, message: '社員を追加しました' };
  } catch (e) {
    console.error('社員追加エラー:', e);
    return { success: false, message: 'エラーが発生しました' };
  }
}

// 社員削除
function deleteEmployee(id) {
  try {
    initialize();
    const sheet = spreadsheet.getSheetByName(SHEETS.EMPLOYEES);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: '社員を削除しました' };
      }
    }
    
    return { success: false, message: '社員が見つかりません' };
  } catch (e) {
    console.error('社員削除エラー:', e);
    return { success: false, message: 'エラーが発生しました' };
  }
}

// 作業項目追加
function addWorkType(id, name) {
  try {
    initialize();
    const sheet = spreadsheet.getSheetByName(SHEETS.WORK_TYPES);
    
    // 重複チェック
    const workTypes = getWorkTypes();
    if (workTypes.some(wt => wt.id === id)) {
      return { success: false, message: 'その作業IDは既に使用されています' };
    }
    
    sheet.appendRow([id, name]);
    return { success: true, message: '作業項目を追加しました' };
  } catch (e) {
    console.error('作業項目追加エラー:', e);
    return { success: false, message: 'エラーが発生しました' };
  }
}

// 作業項目削除
function deleteWorkType(id) {
  try {
    initialize();
    const sheet = spreadsheet.getSheetByName(SHEETS.WORK_TYPES);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: '作業項目を削除しました' };
      }
    }
    
    return { success: false, message: '作業項目が見つかりません' };
  } catch (e) {
    console.error('作業項目削除エラー:', e);
    return { success: false, message: 'エラーが発生しました' };
  }
}

// Excel出力用のデータ取得
function getAttendanceData(startDate, endDate) {
  try {
    initialize();
    const sheet = spreadsheet.getSheetByName(SHEETS.ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    
    const filteredData = [];
    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][0]);
      if (date >= new Date(startDate) && date <= new Date(endDate)) {
        filteredData.push(data[i]);
      }
    }
    
    return filteredData;
  } catch (e) {
    console.error('勤怠データ取得エラー:', e);
    return [];
  }
}

function getWorkRecordData(startDate, endDate) {
  try {
    initialize();
    const sheet = spreadsheet.getSheetByName(SHEETS.WORK_RECORDS);
    const data = sheet.getDataRange().getValues();
    
    const filteredData = [];
    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][0]);
      if (date >= new Date(startDate) && date <= new Date(endDate)) {
        filteredData.push(data[i]);
      }
    }
    
    return filteredData;
  } catch (e) {
    console.error('作業記録データ取得エラー:', e);
    return [];
  }
}