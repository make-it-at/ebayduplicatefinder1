/**
 * スプレッドシートツール
 * 
 * メインエントリーポイント
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('スプレッドシートツール')
    .addItem('ツール実行', 'runTool')
    .addSeparator()
    .addItem('設定', 'showSettings')
    .addToUi();
}

/**
 * ツールのメイン実行関数
 */
function runTool() {
  var sheet = SpreadsheetApp.getActiveSheet();
  Logger.log('ツールを実行しています...');
  
  // ここにツールのロジックを実装
  Browser.msgBox('ツールが実行されました！');
}

/**
 * 設定画面を表示
 */
function showSettings() {
  var html = HtmlService.createHtmlOutputFromFile('settings')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '設定');
}

/**
 * 現在のスプレッドシートの情報を取得
 */
function getSpreadsheetInfo() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return {
    id: spreadsheet.getId(),
    name: spreadsheet.getName(),
    sheets: spreadsheet.getSheets().map(function(sheet) {
      return {
        name: sheet.getName(),
        rows: sheet.getLastRow(),
        cols: sheet.getLastColumn()
      };
    })
  };
}

/**
 * ユーザー設定を保存
 */
function saveUserSettings(settings) {
  try {
    var properties = PropertiesService.getUserProperties();
    properties.setProperties(settings);
    Logger.log('設定を保存しました: ' + JSON.stringify(settings));
    return { success: true };
  } catch (error) {
    Logger.log('設定保存エラー: ' + error.toString());
    throw error;
  }
}

/**
 * ユーザー設定を取得
 */
function getUserSettings() {
  try {
    var properties = PropertiesService.getUserProperties();
    var settings = properties.getProperties();
    Logger.log('設定を読み込みました: ' + JSON.stringify(settings));
    return settings;
  } catch (error) {
    Logger.log('設定読み込みエラー: ' + error.toString());
    return {};
  }
} 