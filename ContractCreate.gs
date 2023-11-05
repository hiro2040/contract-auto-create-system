/**
 * メイン処理
 */

function main() {
  // システム変数を設定
  setSysVariable()

  // 契約書リストを取得
  const contractList = toString(inputSpread(inputList.name, inputList.sheetName, inputList.startRow, inputList.endCol, inputList.url));

  // 各フォルダおよび契約書を作成
  contractList.forEach(contract => {
    // 会社毎のフォルダを作成
    const folderId = createFolder(contract[outputInfo.folderNameCol-1]);

    // 契約書の作成
    outputInfo.sampleFileIdList.forEach(sampleFileId => {
      // サンプルファイルをコピーする
      const newFileUrl = docCopy(sampleFileId, folderId);

      // コピーしたドキュメントを開く
      let basebody = openDoc(newFileUrl);

      // 埋め込み文字をバインド
      outputInfo.bindList.forEach(bind => {
        basebody.replaceText(bind.str, contract[bind.mapCol-1])
      });
    });
  });
}
