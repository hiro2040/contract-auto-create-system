# contract_auto_create_system  
契約書自動作成システムの使用方法  
1.準備  
①以下のエクセルをダウンロード  
　・設定ファイル.xlsx  
　・契約先リスト.xlsx  
②上記2ファイル用のスプレッドシートを自作し、上記のすべてのシートを転記する。  
③「設定ファイル」のスプレッドシートのスクリプトで以下のファイルを作成し、GitHubのコードを転記する。  
　・ContractCreate.gs  
　・Common.gs  
　・Constant.gs  
　※appsscript.jsonはデフォルトのままで良い。  
④「設定ファイル」の入力設定sheetのリストファイル欄に先ほど作成した「契約先リスト」のURLを記入  
　※一旦、入力列数、出力フォルダ名列数欄は無視で良いです。  
⑤「設定ファイル」の実行sheetの契約書作成実行ボタンを右クリックし、3点リーダー→スクリプトを割り当てをクリックする。  
⑥mainと記入し、確定ボタンを押下。  
  
2.契約書自動作成のための設定  
①作成したい契約書のひな型をGoogleドキュメントで作成する。  
　※いくつ作成しても大丈夫です。  
　情報を埋め込みたい位置に[変数名]という形式で記載する。  
　例えば会社名を埋め込みたければ、埋め込みたい位置に［会社名］と記載する。  
②「設定ファイル」の埋込設定sheetを開き、①にて設定した埋込文字を埋込文字列欄に記載する。  
　［会社名］という埋め込み文字を記載している場合は［会社名］と記載する。  
 　※一旦、ここでは対象列欄は無視で良いです。  
③「設定ファイル」の出力設定sheetのひな型ファイル欄に作成したひな型ファイルのURLを縦向きに記載していく。  
④「契約先リスト」のシート1sheetに契約書に埋め込みたい情報がまとめられている表を作成する。  
　※埋め込みたい情報があれば問題ないので関係のない情報の列が存在していても問題ありません。  
　※「契約先リスト」の1行毎にフォルダを作成してその中に契約書を生成するため、そのフォルダ名を記載する列も用意してください。  
　　すでに用意している情報の中からフォルダ名を決定する場合は別で作成する必要はありません。  
⑤「設定ファイル」の埋込設定sheetを開き、対象列欄に埋め込み文字列に対して④にて作成した「契約先リスト」のどの列の情報を埋め込みたいか列番号で記載する。  
　※Aから1スタートで数えた時の列数を記載する。  
⑥「設定ファイル」の入力設定sheetの入力列数欄に「契約先リスト」の列数を記載。  
　関係のない情報の列も含めた総列数を記載してください。  
⑦「設定ファイル」の入力設定sheetの出力フォルダ名列数欄に④にて作成したフォルダ名に使用したい列数を記載。  
⑧上記すべて完了したら「設定ファイル」の実行sheetの契約書作成実行ボタンを押下すると自動作成されます。  
  
「Sample.zip」にサンプルファイルが格納されております。  
「Sample.zip」をダウンロードして1.の手順、2.③の手順のみ行えば自動生成の挙動確認ができます。  
以下のファイルがサンプルとして用意しているひな型ファイルです。  
・覚書ひな形（サンプル）.docx  
・契約書ひな形（サンプル）.docx  
