/**
 * システム変数
 */

/** 契約先リスト情報 */ 
let inputList = {
  name: '契約先リスト',
  /** URL */
  url: '',
  /** シート名 */
  sheetName: 'シート1',
  /** 読み込み開始行 */
  startRow: 2,
  /** 読み込み最終列 */
  endCol: 0
}

/** 出力情報 */
let outputInfo = {
  /** 出力フォルダの名前として出力する契約先リストの列 */
  folderNameCol: 0,
  /** ルートフォルダのID */
  rootFolderId:'',
  /** サンプルファイルのドキュメントID */
  sampleFileIdList:[],
  /** 埋め込み文字列と契約先リストの列数のマッピング情報 */
  bindList:[]
}

/**
 * システム変数を設定
 */
function setSysVariable() {
  // 埋め込み文字列と契約先リストの列数のマッピング情報を取得
  inputSpread(SETTING_BIND.name, SETTING_BIND.sheetName, SETTING_BIND.startRow, SETTING_BIND.endCol).forEach(list => {
    const obj = new Bind();
    obj.str = list[0]
    obj.mapCol = list[1]
    outputInfo.bindList.push(obj)
  })

  // 読込リストファイルの読込設定情報を取得する。
  const inputListInfo = inputSpread(SETTING_INPUT.name, SETTING_INPUT.sheetName, SETTING_INPUT.startRow, SETTING_INPUT.endCol).shift()
  inputList.url = inputListInfo[0]
  inputList.endCol = inputListInfo[1]
  outputInfo.folderNameCol = inputListInfo[2]

  // サンプルファイルURLの設定
  inputSpread(SETTING_OUTPUT.name, SETTING_OUTPUT.sheetName, SETTING_OUTPUT.startRow, SETTING_OUTPUT.endCol).forEach(url =>{
    // サンプルファイルのドキュメントIDを取得
    outputInfo.sampleFileIdList.push(DocumentApp.openByUrl(url[0]).getId());
  })

  // スプレッドシートのIDを取得
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
 
  // 格納されているフォルダIDを取得
  const file = DriveApp.getFileById(id);
  const folder = file.getParents().next();

  // ルートフォルダのIDの設定
  outputInfo.rootFolderId = folder.getId();
}