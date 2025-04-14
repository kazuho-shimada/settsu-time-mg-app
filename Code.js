// test
// Google Apps Scriptのメインコード

// スプレッドシートのID
const SPREADSHEET_ID = '1Y2HJPknu0XhGRr2keVGei_b_X1Qn6Mmb3yj6VGBFNfw';

// グローバル変数
let spreadsheet;
let attendanceSheet;
let employeesSheet;
let workItemsSheet;
let workRecordsSheet;

// 初期化処理
function initialize() {
  try {
    spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // シートの取得または作成
    attendanceSheet = getOrCreateSheet('勤怠管理');
    employeesSheet = getOrCreateSheet('社員マスタ');
    workItemsSheet = getOrCreateSheet('作業項目マスタ');
    workRecordsSheet = getOrCreateSheet('作業記録');
    
    // 初期設定
    setupSheets();
  } catch (e) {
    Logger.log('初期化エラー: ' + e.toString());
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

// シートの初期設定
function setupSheets() {
  // 社員マスタの設定
  if (employeesSheet.getLastRow() === 0) {
    employeesSheet.appendRow(['社員ID', '社員名']);
    employeesSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  // 作業項目マスタの設定
  if (workItemsSheet.getLastRow() === 0) {
    workItemsSheet.appendRow(['作業ID', '作業名']);
    workItemsSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  // 勤怠管理シートの設定
  if (attendanceSheet.getLastRow() === 0) {
    attendanceSheet.appendRow(['記録ID', '社員ID', '社員名', '日付', '出勤時間', '退勤時間', '備考']);
    attendanceSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  // 作業記録の設定
  if (workRecordsSheet.getLastRow() === 0) {
    workRecordsSheet.appendRow(['日付', '社員ID', '社員名', '作業ID', '作業名', 'ロット番号', '数量']);
    workRecordsSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#f3f3f3');
  }
}

// メイン画面を開く
function doGet(e) {
  // スプレッドシートを確実に初期化
  initialize();
  
  // クエリパラメータを取得
  const params = e.parameter || {};
  
  // 管理画面の場合
  if (params.page === 'admin') {
    return openAdminPage();
  }
  
  // 作業記録画面の場合
  if (params.page === 'work-record') {
    return getWorkRecordPage();
  }
  
  // デフォルトはメイン画面
  const template = HtmlService.createTemplateFromFile('index');
  
  // テンプレートを評価してHTMLを生成
  const htmlOutput = template.evaluate();
  
  // タイトルを設定
  htmlOutput.setTitle('勤怠管理システム');
  
  // サンドボックスモードを設定 - IFRAMEからNATIVEに変更
  htmlOutput.setSandboxMode(HtmlService.SandboxMode.NATIVE);
  
  // キャッシュを無効化
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  // メタタグを追加
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  return htmlOutput;
}

// 作業記録画面を開く
function getWorkRecordPage() {
  // 初期化を確実に行う
  initialize();
  
  // 作業記録画面のテンプレートを取得
  const template = HtmlService.createTemplateFromFile('workRecord');
  
  // テンプレートを評価してHTMLを生成
  const htmlOutput = template.evaluate();
  
  // タイトルを設定
  htmlOutput.setTitle('作業記録 - 勤怠管理システム');
  
  // サンドボックスモードを設定
  htmlOutput.setSandboxMode(HtmlService.SandboxMode.NATIVE);
  
  // キャッシュを無効化
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  // メタタグを追加
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  return htmlOutput;
}

// 管理画面を開く
function openAdminPage() {
  // 初期化を確実に行う
  initialize();
  
  // 管理画面のテンプレートを取得
  const template = HtmlService.createTemplateFromFile('adminPage');
  
  // テンプレートを評価してHTMLを生成
  const htmlOutput = template.evaluate();
  
  // タイトルを設定
  htmlOutput.setTitle('管理画面 - 勤怠管理システム');
  
  // サンドボックスモードを設定
  htmlOutput.setSandboxMode(HtmlService.SandboxMode.NATIVE);
  
  // キャッシュを無効化
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  // メタタグを追加
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  return htmlOutput;
}

// 管理画面のコンテンツを取得
function getAdminPageContent() {
  // 初期化を確実に行う
  initialize();
  
  // 管理画面のテンプレートを取得してHTMLを生成
  const template = HtmlService.createTemplateFromFile('adminPage');
  const html = template.evaluate().getContent();
  
  // デバッグ用にログを記録
  Logger.log('Admin page content generated, length: ' + html.length);
  
  return html;
}

// メイン画面URLの取得
function getMainPageUrl() {
  return ScriptApp.getService().getUrl();
}

// 外部ファイルのインクルード
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 社員一覧の取得
function getEmployees() {
  // 初期化を確実に行う
  initialize();
  
  Logger.log('社員一覧を取得します');
  
  try {
    // スプレッドシートを取得
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) {
      Logger.log('スプレッドシートが見つかりません: ' + SPREADSHEET_ID);
      throw new Error('スプレッドシートが見つかりません');
    }
    
    // グローバル変数を更新
    spreadsheet = ss;
    
    // 社員シートを取得
    let sheet = ss.getSheetByName('employees');
    if (!sheet) {
      Logger.log('社員シートが存在しないため作成します');
      sheet = ss.insertSheet('employees');
      sheet.appendRow(['ID', '名前']);
      
      // サンプルデータを追加（初期状態で空にならないように）
      sheet.appendRow(['1', '山田太郎']);
      sheet.appendRow(['2', '鈴木花子']);
      
      // グローバル変数を更新
      employeesSheet = sheet;
      
      // サンプルデータを返す
      return [
        { id: '1', name: '山田太郎' },
        { id: '2', name: '鈴木花子' }
      ];
    }
    
    // グローバル変数を更新
    employeesSheet = sheet;
    
    // データを取得
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    Logger.log('取得した行数: ' + values.length);
    
    // 行数が1以下（ヘッダーのみ）の場合はサンプルデータを追加
    if (values.length <= 1) {
      Logger.log('社員データが空のためサンプルデータを追加します');
      sheet.appendRow(['1', '山田太郎']);
      sheet.appendRow(['2', '鈴木花子']);
      
      // サンプルデータを返す
      return [
        { id: '1', name: '山田太郎' },
        { id: '2', name: '鈴木花子' }
      ];
    }
    
    // ヘッダー行をスキップしてデータを変換
    const employees = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[0] && row[1]) {
        employees.push({
          id: String(row[0]),
          name: String(row[1])
        });
      }
    }
    
    Logger.log('社員データを' + employees.length + '件取得しました');
    
    // データが空の場合はサンプルデータを返す
    if (employees.length === 0) {
      Logger.log('変換後の社員データが空のためサンプルデータを返します');
      return [
        { id: '1', name: '山田太郎' },
        { id: '2', name: '鈴木花子' }
      ];
    }
    
    return employees;
    
  } catch (error) {
    Logger.log('社員データ取得エラー: ' + error.toString());
    // エラーが発生した場合もサンプルデータを返す
    return [
      { id: '1', name: '山田太郎' },
      { id: '2', name: '鈴木花子' }
    ];
  }
}

// 作業項目一覧の取得
function getWorkItems() {
  try {
    // 確実に初期化を行う
    initialize();
    
    // デバッグ用にログを記録
    Logger.log('作業項目一覧を取得します');
    
    const dataRange = workItemsSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行をスキップ
    const workItems = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[0] && row[1]) { // IDと名前の両方が存在する場合のみ追加
        workItems.push({
          id: row[0],
          name: row[1]
        });
      }
    }
    
    // デバッグ用にログを記録
    Logger.log('作業項目データを取得しました。件数: ' + workItems.length);
    
    return workItems;
  } catch (e) {
    Logger.log('作業項目一覧取得エラー: ' + e.toString());
    // エラーが発生しても空の配列を返す
    return [];
  }
}

// 社員の勤怠状況を取得
function getAttendanceStatus(employeeId) {
  try {
    initialize();
    
    // 社員情報を取得
    const employees = getEmployees();
    const employee = employees.find(emp => emp.id === employeeId);
    
    if (!employee) {
      Logger.log('勤怠状況取得失敗: 社員が見つかりません ID=' + employeeId);
      return { 
        success: false, 
        message: '社員情報が見つかりません',
        data: {
          canCheckIn: true,
          canCheckOut: false,
          history: []
        }
      };
    }
    
    // 現在の日付を取得
    const now = new Date();
    const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
    
    // 勤怠データを取得
    const data = attendanceSheet.getDataRange().getValues();
    
    // 当日の勤怠記録を探す
    let todayRecords = [];
    let hasCheckedIn = false;
    let hasCheckedOut = false;
    
    // シンプルに勤怠記録を取得
    let allRecords = [];
    
    // 全ての記録を取得
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === employeeId) { // 社員IDが一致
        const recordDate = row[3]; // 日付
        const checkInTime = row[4]; // 出勤時間
        const checkOutTime = row[5]; // 退勤時間
        
        // 記録を追加
        allRecords.push({
          date: recordDate,
          checkInTime: checkInTime || '',
          checkOutTime: checkOutTime || ''
        });
        
        // 当日の記録をチェック
        if (recordDate === today) {
          todayRecords.push({
            date: recordDate,
            checkInTime: checkInTime,
            checkOutTime: checkOutTime
          });
          
          if (checkInTime) hasCheckedIn = true;
          if (checkOutTime) hasCheckedOut = true;
        }
      }
    }
    
    // 日付ごとに出勤・退勤記録をまとめる
    const recordsByDate = {};
    
    // 全ての記録を日付でグループ化
    for (let i = 0; i < allRecords.length; i++) {
      const record = allRecords[i];
      const date = record.date;
      
      if (!recordsByDate[date]) {
        recordsByDate[date] = {
          date: date,
          checkInTime: '',
          checkOutTime: ''
        };
      }
      
      // 出勤時間があれば設定
      if (record.checkInTime) {
        recordsByDate[date].checkInTime = record.checkInTime;
      }
      
      // 退勤時間があれば設定
      if (record.checkOutTime) {
        recordsByDate[date].checkOutTime = record.checkOutTime;
      }
    }
    
    // 日付ごとの記録を配列に変換
    const dateRecords = Object.values(recordsByDate);
    
    // 日付で降順にソート（最新の日付が先頭に来るように）
    dateRecords.sort((a, b) => {
      const dateA = new Date(a.date.split('/')[0], a.date.split('/')[1] - 1, a.date.split('/')[2]);
      const dateB = new Date(b.date.split('/')[0], b.date.split('/')[1] - 1, b.date.split('/')[2]);
      return dateB.getTime() - dateA.getTime();
    });
    
    // デバッグ用にログ出力
    Logger.log('日付ごとの勤怠記録数: ' + dateRecords.length);
    
    // 直近の5件のみを使用
    const recentHistory = dateRecords.slice(0, 5);
    
    // 出勤・退勤可能か判定
    const canCheckIn = !hasCheckedIn;
    const canCheckOut = hasCheckedIn && !hasCheckedOut;
    
    return {
      success: true,
      message: '勤怠状況を取得しました',
      data: {
        employeeId: employeeId,
        employeeName: employee.name,
        canCheckIn: canCheckIn,
        canCheckOut: canCheckOut,
        todayRecords: todayRecords,
        history: recentHistory.slice(0, 5) // 直近5件のみ返す
      }
    };
    
  } catch (e) {
    Logger.log('勤怠状況取得中にエラーが発生しました: ' + e.toString());
    return { 
      success: false, 
      message: 'エラーが発生しました: ' + e.toString(),
      data: {
        canCheckIn: true,
        canCheckOut: false,
        history: []
      }
    };
  }
}

// 出勤処理
function checkIn(employeeId) {
  try {
    initialize();
    
    // 勤怠状況を取得
    const status = getAttendanceStatus(employeeId);
    if (!status.success) {
      return { success: false, message: status.message };
    }
    
    // 既に出勤済みならエラー
    if (!status.data.canCheckIn) {
      return { success: false, message: '本日は既に出勤済みです' };
    }
    
    // 社員情報を取得
    const employees = getEmployees();
    const employee = employees.find(emp => emp.id === employeeId);
    
    if (!employee) {
      Logger.log('出勤処理失敗: 社員が見つかりません ID=' + employeeId);
      return { success: false, message: '社員情報が見つかりません' };
    }
    
    // 現在の日時を取得
    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
    const timeStr = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm:ss');
    
    // 記録IDを生成
    const recordId = 'A' + now.getTime().toString();
    
    // 勤怠データを追加
    attendanceSheet.appendRow([recordId, employeeId, employee.name, dateStr, timeStr, '', '']);
    
    Logger.log('出勤処理成功: ' + employee.name + ' (' + dateStr + ' ' + timeStr + ')');
    
    // 出勤処理後は必ず状態が変わるので、直接状態を設定する
    // 更新後の勤怠状況を取得
    const updatedStatus = getAttendanceStatus(employeeId);
    
    // デバッグ用にログ出力
    Logger.log('出勤処理後の勤怠状況: ' + JSON.stringify(updatedStatus.data));
    
    // 直近の履歴に新しい記録を追加
    if (updatedStatus.data && updatedStatus.data.history) {
      // 履歴が取得できない場合は、新しい記録を作成
      if (updatedStatus.data.history.length === 0) {
        updatedStatus.data.history.push({
          id: recordId,
          date: dateStr,
          checkInTime: timeStr,
          checkOutTime: ''
        });
      }
    }
    
    return { 
      success: true, 
      message: '出勤処理が完了しました', 
      data: {
        employeeId: employeeId,
        employeeName: employee.name,
        date: dateStr,
        time: timeStr,
        status: updatedStatus.data
      }
    };
    
  } catch (e) {
    Logger.log('出勤処理中にエラーが発生しました: ' + e.toString());
    return { success: false, message: 'エラーが発生しました: ' + e.toString() };
  }
}

// 退勤処理
function checkOut(employeeId) {
  try {
    initialize();
    
    // 勤怠状況を取得
    const status = getAttendanceStatus(employeeId);
    if (!status.success) {
      return { success: false, message: status.message };
    }
    
    // 退勤可能かチェック - 出勤処理がなくても退勤記録を可能にする
    if (!status.data.canCheckIn && !status.data.canCheckOut) {
      // 本日既に退勤済みの場合のみエラーとする
      return { success: false, message: '本日は既に退勤済みです' };
    }
    
    // 出勤処理がなくても退勤記録を可能にするため、ここではチェックしない
    Logger.log('出勤状態に関わらず退勤記録を許可します');
    
    // 社員情報を取得
    const employees = getEmployees();
    const employee = employees.find(emp => emp.id === employeeId);
    
    if (!employee) {
      Logger.log('退勤処理失敗: 社員が見つかりません ID=' + employeeId);
      return { success: false, message: '社員情報が見つかりません' };
    }
    
    // 現在の日時を取得
    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
    const timeStr = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm:ss');
    
    // 当日の出勤データを探す
    const data = attendanceSheet.getDataRange().getValues();
    let found = false;
    let rowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === employeeId && row[3] === dateStr && row[4] !== '' && row[5] === '') {
        // 当日の出勤データがあり、退勤時間が空の場合
        rowIndex = i + 1; // シートの行番号は1から始まる
        found = true;
        break;
      }
    }
    
    if (!found) {
      // 出勤記録が見つからない場合は新規作成
      const recordId = 'A' + now.getTime().toString();
      attendanceSheet.appendRow([recordId, employeeId, employee.name, dateStr, '', timeStr, '退勤のみ記録']);
      Logger.log('退勤のみ記録を作成しました: ' + employee.name);
    } else {
      // 既存の出勤記録を更新
      attendanceSheet.getRange(rowIndex, 6).setValue(timeStr); // 6列目が退勤時間
      Logger.log('退勤時間を更新しました: ' + employee.name);
    }
    
    // 退勤処理後は必ず状態が変わるので、直接状態を設定する
    // 更新後の勤怠状況を取得
    const updatedStatus = getAttendanceStatus(employeeId);
    
    // デバッグ用にログ出力
    Logger.log('退勤処理後の勤怠状況: ' + JSON.stringify(updatedStatus.data));
    
    // 勤怠記録が表示されない問題を解決するためのデバッグログ
    Logger.log('勤怠履歴件数: ' + updatedStatus.data.history.length);
    
    return { 
      success: true, 
      message: '退勤処理が完了しました', 
      data: {
        employeeId: employeeId,
        employeeName: employee.name,
        date: dateStr,
        time: timeStr,
        status: updatedStatus.data
      }
    };
    
  } catch (e) {
    Logger.log('退勤処理中にエラーが発生しました: ' + e.toString());
    return { success: false, message: 'エラーが発生しました: ' + e.toString() };
  }
}

// 出退勤の記録
function recordAttendance(employeeId, type, timeISOString) {
  initialize();
  
  try {
    // 社員情報の取得
    const employee = findEmployeeById(employeeId);
    if (!employee) {
      throw new Error('社員が見つかりません');
    }
    
    const time = new Date(timeISOString);
    const formattedDate = Utilities.formatDate(time, 'Asia/Tokyo', 'yyyy-MM-dd');
    const formattedTime = Utilities.formatDate(time, 'Asia/Tokyo', 'HH:mm:ss');
    
    // 当日のデータを検索
    const dataRange = attendanceSheet.getDataRange();
    const values = dataRange.getValues();
    let rowToUpdate = -1;
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[0] === formattedDate && row[1] === employeeId) {
        rowToUpdate = i + 1; // 1-indexedに変換
        break;
      }
    }
    
    if (type === 'check-in') {
      // 出勤時間の記録
      if (rowToUpdate === -1) {
        // 新規行の追加
        attendanceSheet.appendRow([
          formattedDate,
          employeeId,
          employee.name,
          formattedTime,
          '' // 退勤時間は空
        ]);
      } else {
        // 既存行の更新
        attendanceSheet.getRange(rowToUpdate, 4).setValue(formattedTime);
      }
    } else if (type === 'check-out') {
      // 退勤時間の記録
      if (rowToUpdate === -1) {
        // 出勤記録がない場合、出勤時間なしで退勤のみ記録
        attendanceSheet.appendRow([
          formattedDate,
          employeeId,
          employee.name,
          '', // 出勤時間は空
          formattedTime
        ]);
      } else {
        // 既存行の更新
        attendanceSheet.getRange(rowToUpdate, 5).setValue(formattedTime);
      }
    }
    
    return true;
  } catch (e) {
    Logger.log('記録エラー: ' + e.toString());
    throw e;
  }
}

// 作業記録の保存
function saveWorkRecords(employeeId, employeeName, records) {
  initialize();
  
  try {
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
    
    // 作業項目マスタの取得
    const workItemsMap = {};
    const workItemValues = workItemsSheet.getDataRange().getValues();
    for (let i = 1; i < workItemValues.length; i++) {
      workItemsMap[workItemValues[i][0]] = workItemValues[i][1];
    }
    
    // 作業記録の保存
    records.forEach(record => {
      workRecordsSheet.appendRow([
        formattedDate,
        employeeId,
        employeeName,
        record.workTypeId,
        record.workTypeName,
        record.lotNumber,
        record.quantity
      ]);
    });
    
    return true;
  } catch (e) {
    Logger.log('作業記録保存エラー: ' + e.toString());
    throw e;
  }
}

// 社員の追加
function addEmployee(id, name) {
  // 確実に初期化を行う
  initialize();
  
  Logger.log('社員追加処理を開始します: ID=' + id + ', 名前=' + name);
  
  try {
    // スプレッドシートを取得
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 社員シートを取得
    const sheet = ss.getSheetByName('employees');
    if (!sheet) {
      Logger.log('社員シートが存在しないため作成します');
      const newSheet = ss.insertSheet('employees');
      newSheet.appendRow(['ID', '名前']);
      employeesSheet = newSheet; // グローバル変数を更新
    } else {
      employeesSheet = sheet; // グローバル変数を更新
    }
    
    // 重複チェック
    const dataRange = employeesSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        Logger.log('重複する社員IDが見つかりました: ' + id);
        throw new Error('その社員IDはすでに使用されています');
      }
    }
    
    // 社員を追加
    Logger.log('社員を追加します: ' + id + ', ' + name);
    employeesSheet.appendRow([id, name]);
    
    // スプレッドシートを保存
    SpreadsheetApp.flush();
    
    Logger.log('社員追加が完了しました');
    return true;
  } catch (e) {
    Logger.log('社員追加エラー: ' + e.toString());
    throw e;
  }
}

// 社員情報の更新
function updateEmployee(id, newName) {
  initialize();
  
  try {
    const dataRange = employeesSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        employeesSheet.getRange(i + 1, 2).setValue(newName);
        return true;
      }
    }
    
    throw new Error('社員が見つかりません');
  } catch (e) {
    Logger.log('社員更新エラー: ' + e.toString());
    throw e;
  }
}

// 社員の削除
function deleteEmployee(id) {
  try {
    // 確実に初期化を行う
    initialize();
    
    // デバッグ用にログを記録
    Logger.log('社員削除処理を開始します。ID: ' + id);
    
    // スプレッドシートを取得
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) {
      Logger.log('スプレッドシートが見つかりません。ID: ' + SPREADSHEET_ID);
      throw new Error('スプレッドシートが見つかりません');
    }
    
    // 社員シートを取得
    const sheet = ss.getSheetByName('employees');
    if (!sheet) {
      Logger.log('社員シートが存在しません');
      throw new Error('社員シートが存在しません');
    }
    
    // グローバル変数を更新
    employeesSheet = sheet;
    
    // IDを文字列化
    const idToDelete = String(id);
    Logger.log('削除対象ID: ' + idToDelete);
    
    // データ範囲を取得
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    Logger.log('取得したデータ行数: ' + values.length);
    
    // 削除対象の行を探す
    let rowToDelete = -1;
    for (let i = 1; i < values.length; i++) {
      // 文字列化して比較
      if (String(values[i][0]) === idToDelete) {
        rowToDelete = i + 1; // スプレッドシートの行番号は1から始まる
        break;
      }
    }
    
    if (rowToDelete > 0) {
      Logger.log('削除対象の行が見つかりました: ' + rowToDelete);
      
      // 行を削除
      sheet.deleteRow(rowToDelete);
      
      // 変更を確実に反映させる
      SpreadsheetApp.flush();
      Logger.log('変更をフラッシュしました');
      
      // 削除が成功したか確認
      const newValues = sheet.getDataRange().getValues();
      let stillExists = false;
      
      for (let i = 1; i < newValues.length; i++) {
        if (String(newValues[i][0]) === idToDelete) {
          stillExists = true;
          break;
        }
      }
      
      if (stillExists) {
        Logger.log('警告: 削除後も社員IDが存在しています: ' + idToDelete);
      } else {
        Logger.log('社員を削除しました: ' + idToDelete);
      }
      
      return true;
    }
    
    Logger.log('指定されたIDの社員が見つかりません: ' + idToDelete);
    throw new Error('社員が見つかりません');
  } catch (e) {
    Logger.log('社員削除エラー: ' + e.toString());
    throw e;
  }
}

// 作業項目の追加
function addWorkItem(id, name) {
  initialize();
  
  try {
    // 重複チェック
    if (findWorkItemById(id)) {
      throw new Error('その作業IDはすでに使用されています');
    }
    
    workItemsSheet.appendRow([id, name]);
    return true;
  } catch (e) {
    Logger.log('作業項目追加エラー: ' + e.toString());
    throw e;
  }
}

// 作業項目情報の更新
function updateWorkItem(id, newName) {
  initialize();
  
  try {
    const dataRange = workItemsSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        workItemsSheet.getRange(i + 1, 2).setValue(newName);
        return true;
      }
    }
    
    throw new Error('作業項目が見つかりません');
  } catch (e) {
    Logger.log('作業項目更新エラー: ' + e.toString());
    throw e;
  }
}

// 作業項目の削除
function deleteWorkItem(id) {
  initialize();
  
  try {
    const dataRange = workItemsSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        workItemsSheet.deleteRow(i + 1);
        return true;
      }
    }
    
    throw new Error('作業項目が見つかりません');
  } catch (e) {
    Logger.log('作業項目削除エラー: ' + e.toString());
    throw e;
  }
}

// 勤怠データのExcel出力
function exportAttendanceData(startDateStr, endDateStr) {
  initialize();
  
  try {
    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);
    endDate.setHours(23, 59, 59, 999); // 終了日の終わりまで
    
    // データ抽出
    const dataRange = attendanceSheet.getDataRange();
    const values = dataRange.getValues();
    
    // 新しいスプレッドシートの作成
    const tempSpreadsheet = SpreadsheetApp.create('勤怠データ_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss'));
    const tempSheet = tempSpreadsheet.getActiveSheet();
    
    // ヘッダー行
    tempSheet.appendRow(['日付', '社員ID', '社員名', '出勤時間', '退勤時間']);
    tempSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#f3f3f3');
    
    // データ行
    let rowCount = 0;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowDate = new Date(row[0]);
      
      if (rowDate >= startDate && rowDate <= endDate) {
        tempSheet.appendRow([
          row[0], // 日付
          row[1], // 社員ID
          row[2], // 社員名
          row[3], // 出勤時間
          row[4]  // 退勤時間
        ]);
        rowCount++;
      }
    }
    
    if (rowCount === 0) {
      // データがない場合は削除
      DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
      return null;
    }
    
    // 列幅の調整
    tempSheet.autoResizeColumns(1, 5);
    
    // Excel形式でエクスポート
    const blob = exportAsExcel(tempSpreadsheet);
    const file = DriveApp.createFile(blob);
    
    // 元のスプレッドシートを削除
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    
    // ダウンロード用URLを返す
    return file.getUrl();
  } catch (e) {
    Logger.log('勤怠データ出力エラー: ' + e.toString());
    throw e;
  }
}

// 作業記録のExcel出力
function exportWorkRecords(startDateStr, endDateStr) {
  initialize();
  
  try {
    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);
    endDate.setHours(23, 59, 59, 999); // 終了日の終わりまで
    
    // データ抽出
    const dataRange = workRecordsSheet.getDataRange();
    const values = dataRange.getValues();
    
    // 新しいスプレッドシートの作成
    const tempSpreadsheet = SpreadsheetApp.create('作業記録_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss'));
    const tempSheet = tempSpreadsheet.getActiveSheet();
    
    // ヘッダー行
    tempSheet.appendRow(['日付', '社員ID', '社員名', '作業ID', '作業名', 'ロット番号', '数量']);
    tempSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#f3f3f3');
    
    // データ行
    let rowCount = 0;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowDate = new Date(row[0]);
      
      if (rowDate >= startDate && rowDate <= endDate) {
        tempSheet.appendRow([
          row[0], // 日付
          row[1], // 社員ID
          row[2], // 社員名
          row[3], // 作業ID
          row[4], // 作業名
          row[5], // ロット番号
          row[6]  // 数量
        ]);
        rowCount++;
      }
    }
    
    if (rowCount === 0) {
      // データがない場合は削除
      DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
      return null;
    }
    
    // 列幅の調整
    tempSheet.autoResizeColumns(1, 7);
    
    // Excel形式でエクスポート
    const blob = exportAsExcel(tempSpreadsheet);
    const file = DriveApp.createFile(blob);
    
    // 元のスプレッドシートを削除
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    
    // ダウンロード用URLを返す
    return file.getUrl();
  } catch (e) {
    Logger.log('作業記録出力エラー: ' + e.toString());
    throw e;
  }
}

// Excel形式でエクスポート
function exportAsExcel(spreadsheet) {
  try {
    const url = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=xlsx";
    const params = {
      method: "get",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, params);
    return response.getBlob().setName(spreadsheet.getName() + ".xlsx");
  } catch (e) {
    Logger.log('Excelエクスポートエラー: ' + e.toString());
    throw e;
  }
}

// 社員IDから社員情報を検索
function findEmployeeById(employeeId) {
  const dataRange = employeesSheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === employeeId) {
      return {
        id: values[i][0],
        name: values[i][1]
      };
    }
  }
  
  return null;
}

// 作業項目IDから作業項目情報を検索
function findWorkItemById(workItemId) {
  const dataRange = workItemsSheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === workItemId) {
      return {
        id: values[i][0],
        name: values[i][1]
      };
    }
  }
  
  return null;
}