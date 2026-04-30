/**
 * 初期表示: WebアプリのHTMLを返す
 */
function doGet() {
  ensureSheets();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('初月引継太郎')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let casesSheet = ss.getSheetByName('Cases');
  const caseHeaders = ["ID", "作成日時", "企業名", "設立", "従業員数", "売上", "利益", "事業内容", "本社住所", "本社都道府県", "報告資料URL", "実施月", "単価", "訪問 or ZOOM", "プロジェクトの目的", "主な課題・解決すべきテーマ", "初動で着手するテーマ", "ツール指定", "内製化の温度感", "進めるうえでの留意点", "その他引継ぎ事項", "備考", "初月担当", "メイン担当", "サブ担当", "ステータス"];
  
  if (!casesSheet) {
    casesSheet = ss.insertSheet('Cases');
    casesSheet.appendRow(caseHeaders);
  } else if (casesSheet.getLastRow() === 0) {
    casesSheet.appendRow(caseHeaders);
  }

  let cautionSheet = ss.getSheetByName('CautionMaster');
  if (!cautionSheet) {
    cautionSheet = ss.insertSheet('CautionMaster');
    cautionSheet.appendRow(["留意点ID", "留意点項目"]);
  }
  let staffSheet = ss.getSheetByName('StaffMaster');
  if (!staffSheet) {
    staffSheet = ss.insertSheet('StaffMaster');
    // ふりがな列を追加
    staffSheet.appendRow(["メンバーID", "担当者名", "メールアドレス", "ふりがな"]);
  }
  let priceSheet = ss.getSheetByName('PriceMaster');
  if (!priceSheet) {
    priceSheet = ss.insertSheet('PriceMaster');
    priceSheet.appendRow(["単価ID", "単価"]);
  }
}

function getMasters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cautionSheet = ss.getSheetByName('CautionMaster');
  const staffSheet = ss.getSheetByName('StaffMaster');
  const priceSheet = ss.getSheetByName('PriceMaster');

  let cautions = [];
  if (cautionSheet && cautionSheet.getLastRow() > 1) {
    const vals = cautionSheet.getDataRange().getDisplayValues();
    vals.shift();
    // 最初の正常な状態に戻す
    cautions = vals.map(r => ({ id: r, name: r }));
  }

  let staffs = [];
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const vals = staffSheet.getDataRange().getDisplayValues();
    vals.shift();
    // ここだけ、ふりがなを取得するように追加
    staffs = vals.map(r => ({ id: r, name: r, email: r || "", furigana: r || "" }));
  }

  let prices = [];
  if (priceSheet && priceSheet.getLastRow() > 1) {
    const vals = priceSheet.getDataRange().getDisplayValues();
    vals.shift();
    // 最初の正常な状態に戻す
    prices = vals.map(r => ({ id: r, name: r }));
  }

  const currentUserEmail = Session.getActiveUser().getEmail();
  return { cautions, staffs, prices, currentUserEmail };
}

function getCases() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cases');
    if(!sheet || sheet.getLastRow() <= 1) return [];

    const values = sheet.getDataRange().getDisplayValues();
    const headers = values.shift();
    const idIdx = headers.indexOf('ID');

    const result = [];
    for (let i = 0; i < values.length; i++) {
      let row = values[i];
      if (idIdx === -1 || !row[idIdx] || row[idIdx].trim() === "") continue;
      let obj = {};
      headers.forEach((h, j) => { obj[h] = row[j] || ""; });
      result.push(obj);
    }
    return result;
  } catch (e) {
    throw new Error("データ取得エラー: " + e.message);
  }
}

function analyzeFilesWithGemini(fileDataArray, slideUrl) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('APIキーが設定されていません。');

  let promptText = typeof getAiPrompt === 'function' ? getAiPrompt(slideUrl) : "";
  
  let parts = [{ text: promptText }];
  if (fileDataArray && fileDataArray.length > 0) {
    fileDataArray.forEach(file => { parts.push({ inline_data: { mime_type: file.mimeType, data: file.data } }); });
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const response = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: parts }], generationConfig: { responseMimeType: "application/json" } }),
    muteHttpExceptions: true
  });
  const resJson = JSON.parse(response.getContentText());
  if (resJson.error) throw new Error(resJson.error.message);
  return JSON.parse(resJson.candidates.content.parts.text);
}

function saveCase(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cases');
  const id = "ID-" + new Date().getTime();
  const d = new Date();
  const timestamp = `${d.getFullYear()}/${(d.getMonth()+1).toString().padStart(2,'0')}/${d.getDate().toString().padStart(2,'0')} ${d.getHours().toString().padStart(2,'0')}:${d.getMinutes().toString().padStart(2,'0')}:${d.getSeconds().toString().padStart(2,'0')}`;

  const row = [
    id, timestamp, formData['企業名']||"", formData['設立']||"", formData['従業員数']||"", formData['売上']||"", formData['利益']||"", formData['事業内容']||"", formData['本社住所']||"", formData['本社都道府県']||"", formData['報告資料URL']||"", formData['実施月']||"", formData['単価']||"", formData['訪問 or ZOOM']||"", formData['プロジェクトの目的']||"", formData['主な課題・解決すべきテーマ']||"", formData['初動で着手するテーマ']||"", formData['ツール指定']||"", formData['内製化の温度感']||"", formData['進めるうえでの留意点']||"", formData['その他引継ぎ事項']||"", formData['備考']||"", formData['初月担当']||"", "", "", "未差配"
  ];
  sheet.appendRow(row);
  return id;
}

function updateAssignment(id, mainStaff, subStaff, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cases');
  const data = sheet.getDataRange().getValues();
  const headers = data;
  const idIdx = headers.indexOf('ID');
  const mainIdx = headers.indexOf('メイン担当');
  const subIdx = headers.indexOf('サブ担当');
  const statusIdx = headers.indexOf('ステータス');
  const autoStatus = (mainStaff && mainStaff.trim() !== "") ? "差配済" : "未差配";

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === id) {
      sheet.getRange(i + 1, mainIdx + 1).setValue(mainStaff);
      sheet.getRange(i + 1, subIdx + 1).setValue(subStaff);
      sheet.getRange(i + 1, statusIdx + 1).setValue(autoStatus);
      return true;
    }
  }
  return false;
}

function updateCaseDetail(id, formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cases');
  const data = sheet.getDataRange().getValues();
  const headers = data;
  const idIdx = headers.indexOf('ID');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === id) {
      const rowNum = i + 1;
      const range = sheet.getRange(rowNum, 1, 1, headers.length);
      const rowValues = range.getValues();
      headers.forEach((header, colIdx) => {
        if (header !== "ID" && header !== "作成日時" && header !== "メイン担当" && header !== "サブ担当" && header !== "ステータス") {
          rowValues[colIdx] = formData[header] || ""; 
        }
      });
      range.setValues([rowValues]);
      return true;
    }
  }
  return false;
}
