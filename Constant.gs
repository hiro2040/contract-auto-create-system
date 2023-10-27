// 埋め込み文字列と契約書の列数
let BIND_STRING = [];

// 読込リストスプレッドURL
let INPUT_SPREAD_URL = "";

// 読込リストスプレッド入力列数
let INPUT_SPREAD_COLUMN = 0;

// 出力フォルダ名列数
let OUTPUT_FOLDER_NAME_COLUMN = 0;

// サンプルファイルID
let OUTPUT_DOCUMENT_ID = [];

// ルートフォルダのID
let ROOT_FOLDER_ID = "";

function set_constant() {
  // 埋め込み文字列と契約書の列数の設定
  BIND_STRING = input_spread_def("埋込設定", 1, 2);

  // 読込設定情報を取得
  let input_setting_info = input_spread_def("入力設定", 1, 3);
  // 読込リストスプレッドURLの設定
  INPUT_SPREAD_URL = input_setting_info[0][0];
  // 読込リストスプレッド入力列数の設定
  INPUT_SPREAD_COLUMN = input_setting_info[0][1];
  // 出力フォルダ名列数の設定
  OUTPUT_FOLDER_NAME_COLUMN = input_setting_info[0][2];

  // サンプルファイルURLの設定
  let spread_url_org = input_spread_def("出力設定", 1, 1);
  for(let i in spread_url_org) {
    let doc = DocumentApp.openByUrl(spread_url_org[i][0]);
    OUTPUT_DOCUMENT_ID.push(doc.getId())
  }

  // スプレッドシートのIDを取得
  let id = SpreadsheetApp.getActiveSpreadsheet().getId();
 
  // 格納されているフォルダIDを取得
  let file = DriveApp.getFileById(id);
  let folder = file.getParents().next();

  // ルートフォルダのIDの設定
  ROOT_FOLDER_ID = folder.getId();
}