/**
 * 生成AI部門事前課題進捗ダッシュボード - Google Apps Script
 * 
 * このスクリプトをGoogle Apps Scriptにデプロイして、
 * HTMLダッシュボードからスプレッドシートにデータを保存・読み込みできます。
 * 
 * 【設定手順】
 * 1. Google Driveで新しいスプレッドシートを作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードをコピー＆ペースト
 * 4. SPREADSHEET_IDを実際のスプレッドシートIDに置き換え
 * 5. デプロイ → 新しいデプロイ → ウェブアプリ
 * 6. アクセスできるユーザー: 「全員」を選択
 * 7. デプロイしてURLを取得
 * 8. HTMLダッシュボードのAPI設定にそのURLを入力
 */

// スプレッドシートのIDを設定してください
// URLの https://docs.google.com/spreadsheets/d/1Exqpklr9_1wN7pUsEJ9VvdbCk2YIjZEQSn7nEuY1rd8/edit
const SPREADSHEET_ID = '1Exqpklr9_1wN7pUsEJ9VvdbCk2YIjZEQSn7nEuY1rd8';

// シート名
const SHEET_NAME = '進捗データ';
const LOG_SHEET_NAME = '更新ログ';

/**
 * スプレッドシートを初期化（初回実行時）
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // メインデータシート
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    
    // ヘッダー行を設定
    const headers = [
      'ID',
      '最終更新日時',
      // ステップ1
      'Step1_Point1',
      'Step1_Criteria1', 'Step1_Criteria2', 'Step1_Criteria3', 'Step1_Criteria4', 'Step1_Criteria5',
      'Step1_Issue',
      // ステップ2
      'Step2_Point1',
      'Step2_Criteria1', 'Step2_Criteria2', 'Step2_Criteria3', 'Step2_Criteria4', 'Step2_Criteria5',
      'Step2_Issue',
      // ステップ3
      'Step3_Point1',
      'Step3_Criteria1', 'Step3_Criteria2', 'Step3_Criteria3', 'Step3_Criteria4', 'Step3_Criteria5', 'Step3_Criteria6',
      'Step3_Issue',
      // ステップ4
      'Step4_Point1',
      'Step4_Criteria1', 'Step4_Criteria2', 'Step4_Criteria3', 'Step4_Criteria4', 'Step4_Criteria5', 'Step4_Criteria6', 'Step4_Criteria7',
      'Step4_Issue',
      // ステップ5
      'Step5_Point1',
      'Step5_Criteria1', 'Step5_Criteria2', 'Step5_Criteria3', 'Step5_Criteria4', 'Step5_Criteria5', 'Step5_Criteria6', 'Step5_Criteria7',
      'Step5_Issue'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4a86e8');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('white');
    sheet.setFrozenRows(1);
  }
  
  // ログシート
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.getRange(1, 1, 1, 4).setValues([['タイムスタンプ', '操作', 'ユーザー', '詳細']]);
    logSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    logSheet.getRange(1, 1, 1, 4).setBackground('#6aa84f');
    logSheet.getRange(1, 1, 1, 4).setFontColor('white');
    logSheet.setFrozenRows(1);
  }
  
  return '初期化完了';
}

/**
 * POSTリクエストを処理（データ保存）
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    
    if (params.action === 'save') {
      return saveData(params.data);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: '不明なアクション' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GETリクエストを処理（データ読み込み）
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'load') {
      return loadData();
    }
    
    if (action === 'init') {
      initializeSpreadsheet();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, message: '初期化完了' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: '不明なアクション' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * データを保存
 */
function saveData(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    initializeSpreadsheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  
  const timestamp = new Date().toISOString();
  const id = 'main'; // 単一レコードとして管理
  
  // データ行を作成
  const rowData = [
    id,
    timestamp,
    // ステップ1
    data.steps['1'].points[0] || 'incomplete',
    data.steps['1'].criteria[0] || 'incomplete',
    data.steps['1'].criteria[1] || 'incomplete',
    data.steps['1'].criteria[2] || 'incomplete',
    data.steps['1'].criteria[3] || 'incomplete',
    data.steps['1'].criteria[4] || 'incomplete',
    data.steps['1'].issue || '',
    // ステップ2
    data.steps['2'].points[0] || 'incomplete',
    data.steps['2'].criteria[0] || 'incomplete',
    data.steps['2'].criteria[1] || 'incomplete',
    data.steps['2'].criteria[2] || 'incomplete',
    data.steps['2'].criteria[3] || 'incomplete',
    data.steps['2'].criteria[4] || 'incomplete',
    data.steps['2'].issue || '',
    // ステップ3
    data.steps['3'].points[0] || 'incomplete',
    data.steps['3'].criteria[0] || 'incomplete',
    data.steps['3'].criteria[1] || 'incomplete',
    data.steps['3'].criteria[2] || 'incomplete',
    data.steps['3'].criteria[3] || 'incomplete',
    data.steps['3'].criteria[4] || 'incomplete',
    data.steps['3'].criteria[5] || 'incomplete',
    data.steps['3'].issue || '',
    // ステップ4
    data.steps['4'].points[0] || 'incomplete',
    data.steps['4'].criteria[0] || 'incomplete',
    data.steps['4'].criteria[1] || 'incomplete',
    data.steps['4'].criteria[2] || 'incomplete',
    data.steps['4'].criteria[3] || 'incomplete',
    data.steps['4'].criteria[4] || 'incomplete',
    data.steps['4'].criteria[5] || 'incomplete',
    data.steps['4'].criteria[6] || 'incomplete',
    data.steps['4'].issue || '',
    // ステップ5
    data.steps['5'].points[0] || 'incomplete',
    data.steps['5'].criteria[0] || 'incomplete',
    data.steps['5'].criteria[1] || 'incomplete',
    data.steps['5'].criteria[2] || 'incomplete',
    data.steps['5'].criteria[3] || 'incomplete',
    data.steps['5'].criteria[4] || 'incomplete',
    data.steps['5'].criteria[5] || 'incomplete',
    data.steps['5'].criteria[6] || 'incomplete',
    data.steps['5'].issue || ''
  ];
  
  // 既存データを探す
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  let rowIndex = -1;
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex > 0) {
    // 既存行を更新
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // 新規行を追加
    sheet.appendRow(rowData);
  }
  
  // ログを記録
  addLog('保存', 'ダッシュボードデータを保存');
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: '保存完了' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * データを読み込み
 */
function loadData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'シートが見つかりません' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // mainレコードを探す
  let rowData = null;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === 'main') {
      rowData = values[i];
      break;
    }
  }
  
  if (!rowData) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'データが見つかりません' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // データを構造化
  const data = {
    lastUpdated: rowData[1],
    steps: {
      '1': {
        points: [rowData[2]],
        criteria: [rowData[3], rowData[4], rowData[5], rowData[6], rowData[7]],
        issue: rowData[8]
      },
      '2': {
        points: [rowData[9]],
        criteria: [rowData[10], rowData[11], rowData[12], rowData[13], rowData[14]],
        issue: rowData[15]
      },
      '3': {
        points: [rowData[16]],
        criteria: [rowData[17], rowData[18], rowData[19], rowData[20], rowData[21], rowData[22]],
        issue: rowData[23]
      },
      '4': {
        points: [rowData[24]],
        criteria: [rowData[25], rowData[26], rowData[27], rowData[28], rowData[29], rowData[30], rowData[31]],
        issue: rowData[32]
      },
      '5': {
        points: [rowData[33]],
        criteria: [rowData[34], rowData[35], rowData[36], rowData[37], rowData[38], rowData[39], rowData[40]],
        issue: rowData[41]
      }
    }
  };
  
  // ログを記録
  addLog('読込', 'ダッシュボードデータを読み込み');
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ログを追加
 */
function addLog(action, detail) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    
    if (!logSheet) {
      initializeSpreadsheet();
      logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    }
    
    const timestamp = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
    const user = Session.getActiveUser().getEmail() || 'anonymous';
    
    logSheet.appendRow([timestamp, action, user, detail]);
  } catch (e) {
    console.log('ログ記録エラー: ' + e.toString());
  }
}

/**
 * 進捗サマリーを取得（レポート用）
 */
function getProgressSummary() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return { error: 'シートが見つかりません' };
  }
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  let rowData = null;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === 'main') {
      rowData = values[i];
      break;
    }
  }
  
  if (!rowData) {
    return { error: 'データが見つかりません' };
  }
  
  // 進捗を計算
  const stepSizes = [6, 6, 7, 8, 8]; // 各ステップの項目数
  let totalComplete = 0;
  let totalItems = 0;
  const stepProgress = [];
  
  let colIndex = 2;
  for (let s = 0; s < 5; s++) {
    let stepComplete = 0;
    for (let i = 0; i < stepSizes[s]; i++) {
      if (rowData[colIndex] === 'complete') {
        stepComplete++;
        totalComplete++;
      }
      totalItems++;
      colIndex++;
    }
    stepProgress.push({
      step: s + 1,
      complete: stepComplete,
      total: stepSizes[s],
      percent: Math.round((stepComplete / stepSizes[s]) * 100)
    });
    colIndex++; // issueをスキップ
  }
  
  return {
    lastUpdated: rowData[1],
    totalComplete: totalComplete,
    totalItems: totalItems,
    totalPercent: Math.round((totalComplete / totalItems) * 100),
    stepProgress: stepProgress
  };
}

/**
 * テスト用関数
 */
function testInit() {
  console.log(initializeSpreadsheet());
}

function testSummary() {
  console.log(getProgressSummary());
}
