// スプレッドシート初期設定関数
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('医療費管理')
    .addItem('初期設定', 'initializeSheet')
    .addToUi();
}

// スプレッドシートの初期設定
function initializeSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // ヘッダー設定（日本語）
  const headers = [
    '記録日',
    '診療日',
    '病院名', 
    '患者名',
    '診療科',
    '初診料',
    '再診料',
    '医学管理料',
    '血液検査料',
    '尿検査料',
    'レントゲン料',
    '画像診断料',
    '処方箋料',
    '薬剤費',
    '注射料',
    '処置料',
    '手術料',
    '合計点数',
    '総医療費',
    '患者負担額'
  ];
  
  // 既存データをクリア
  sheet.clear();
  
  // ヘッダーを設定
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーの書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  // 列幅を自動調整
  sheet.autoResizeColumns(1, headers.length);
  
  SpreadsheetApp.getUi().alert('初期設定が完了しました！');
}

// Webアプリケーションエンドポイント
function doPost(e) {
  try {
    // POSTデータを解析
    const jsonData = JSON.parse(e.postData.contents);
    
    // スプレッドシートに追加
    addMedicalRecord(jsonData);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: '医療費明細が追加されました'
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error', 
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 医療費記録をスプレッドシートに追加
function addMedicalRecord(data) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // データを配列に変換（JSONのキー順序に対応）
  const rowData = [
    data.record_date,
    data.date,
    data.hospital_name,
    data.patient_name,
    data.department,
    data.first_visit_fee,
    data.re_visit_fee,
    data.medical_management_fee,
    data.blood_test_fee,
    data.urine_test_fee,
    data.xray_fee,
    data.image_diagnosis_fee,
    data.prescription_fee,
    data.medication_cost,
    data.injection_fee,
    data.treatment_fee,
    data.surgery_fee,
    data.total_points,
    data.total_medical_cost,
    data.patient_burden_amount
  ];
  
  // 最終行の次の行に追加
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
  
  // 数値データに書式設定（カンマ区切り）
  const numericColumns = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]; // 初診料以降
  numericColumns.forEach(col => {
    const cell = sheet.getRange(lastRow + 1, col);
    if (cell.getValue() !== '-') {
      cell.setNumberFormat('#,##0');
    }
  });
}

// テスト用関数（手動実行用）
function testAddRecord() {
  const testData = {
    "record_date": "2025-05-25",
    "date": "2025-05-15",
    "hospital_name": "さくら眼科クリニック", 
    "patient_name": "田中花子",
    "department": "眼科",
    "first_visit_fee": "-",
    "re_visit_fee": 73,
    "medical_management_fee": "-",
    "blood_test_fee": "-", 
    "urine_test_fee": "-",
    "xray_fee": "-",
    "image_diagnosis_fee": 200,
    "prescription_fee": 68,
    "medication_cost": 320,
    "injection_fee": 45,
    "treatment_fee": 180,
    "surgery_fee": "-",
    "total_points": 886,
    "total_medical_cost": 8860,
    "patient_burden_amount": 2658
  };
  
  addMedicalRecord(testData);
}