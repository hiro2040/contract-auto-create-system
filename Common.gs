/**
 * 共通関数
 */

/**
 * スプレッドシート情報の取得
 *
 * @param alertName エラーの際に出力する名前
 * @param sheetName 取得対象のシート名
 * @param startRow 取得開始行
 * @param endCol 取得終了列
 * @param url 取得するスプレッドシートのurl
 * @return 取得したデータ
 */
function inputSpread(alertName, sheetName, startRow, endCol, url='') {
  // urlが空の場合は本スプレッドシート、空でない場合は指定されたurlのスプレッドシート情報を取得
  const spread = (url) ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.getActiveSpreadsheet();

  // シート情報を取得
  const sheet = spread.getSheetByName(sheetName);

  // 取得開始行からの取得データの行数を設定
  const rowCount = (startRow == 1) ? sheet.getLastRow() : sheet.getLastRow()-startRow+1;
  if(!rowCount) {
    throw new Error(`${alertName}を記入してください。`)
  } 

  return sheet.getRange(startRow, 1, rowCount, endCol).getValues();
}

/**
 * 新規フォルダの作成
 *
 * @param name フォルダ名
 * @return フォルダID
 */
function createFolder(folderName) {
  // 新規作成したフォルダの情報を取得
  const folder = DriveApp.getFolderById(outputInfo.rootFolderId).createFolder(folderName);

  return folder.getId()
}

/**
 * ドキュメントのコピー
 *
 * @param sampleFileId サンプルファイルのID
 * @param folderId 出力先フォルダID
 * @return コピーしたファイルのurl
 */
function docCopy(sampleFileId, folderId) {
  // サンプルファイルのドキュメント情報を取得
  const doc = DriveApp.getFileById(sampleFileId);
  // 出力先フォルダ情報を取得
  const folder = DriveApp.getFolderById(folderId);
  // 出力先フォルダにサンプルファイルをコピー
  const newfile = doc.makeCopy(doc.getName().replace("ひな形",""), folder);

  return newfile.getUrl()
}

/**
 * ドキュメントのデータを取得
 *
 * @param url ドキュメントのurl
 * @return ドキュメントのデータ
 */
function openDoc(url){
  const basedoc = DocumentApp.openByUrl(url);
  return basedoc.getBody()
}

/**
 * リスト内のデータをすべてString型に変換する。
 *
 * @param list String型に変換するリスト
 * @return content String型に変換後のリスト
 */
function toString(list) {
  return list.map(rec => {
    return rec.map(data => {
      // データの型を取得
      const type = Object.prototype.toString.call(data)
      // 数値、日付をString型に変換
      if(type == '[object Date]') {
        return Utilities.formatDate(data, 'JST', 'yyyy/MM/dd')
      } else if(type == '[object Number]') {
        return data.toLocaleString()
      } else {
        return data
      }
    })
  })
}