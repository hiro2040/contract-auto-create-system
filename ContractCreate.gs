function main() {
  // 定数を設定
  set_constant();

  // 契約先リストを取得
  let contractor_setting_list = input_spread_external(INPUT_SPREAD_URL, "シート1", 1, INPUT_SPREAD_COLUMN);
  for(let i in contractor_setting_list) {
    for(let s in contractor_setting_list[i]) {
      contractor_setting_list[i][s] = string_format(contractor_setting_list[i][s])
    }
  }

  for(let i in contractor_setting_list) {
    // 会社毎のフォルダを作成
    let folder_id = create_folder(contractor_setting_list[i][OUTPUT_FOLDER_NAME_COLUMN-1])

    for(let s in OUTPUT_DOCUMENT_ID) {
      // サンプルファイルをコピーする
      let new_sample_file_url = doc_copy(OUTPUT_DOCUMENT_ID[s], folder_id)

      // コピーしたファイル情報を取得
      let basebody = openDoc(new_sample_file_url);

      // 埋め込み文字を埋め込む
      for(let s in BIND_STRING) {
        let output_deta = contractor_setting_list[i][BIND_STRING[s][1]-1]
        basebody.replaceText(BIND_STRING[s][0],output_deta)
      }
    }
  }
}
