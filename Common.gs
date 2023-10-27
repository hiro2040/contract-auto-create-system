/**
 * webページのHTMLの取得(cookie有,js有)
 *
 * @param url webページのURL
 * @param cookie_domain cookieのドメイン
 * @param cookie_name cookieの名前
 * @param cookie_value cookieの値
 * @return content webページのHTML
 */
function input_spread_def(sheet_name, start_col, end_col) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheet = ss.getSheetByName(sheet_name);

    let lastRow = sheet.getLastRow()-1;

    return sheet.getRange(2, start_col, lastRow, end_col).getValues();
  } catch(e) {
    if(e.message === PropertiesService.getScriptProperties().getProperty('sheet_emp_msg')) {
      return []
    }
  }
}

function input_spread_external(url, sheet_name, start_col, end_col) {
  try {
    const ss = SpreadsheetApp.openByUrl(url);

    const sheet = ss.getSheetByName(sheet_name);

    let lastRow = sheet.getLastRow()-1;

    return sheet.getRange(2, start_col, lastRow, end_col).getValues();
  } catch(e) {
    if(e.message === PropertiesService.getScriptProperties().getProperty('sheet_emp_msg')) {
      return []
    }
  }
}

// フォルダの作成
function create_folder(name) {
  var folder = DriveApp.getFolderById(ROOT_FOLDER_ID);
  var house_folder = folder.createFolder(name);

  return house_folder.getId()
}

function doc_read(sample_url) {
  let doc = DocumentApp.openByUrl(sample_url);
  return [doc, doc.getName()]
}

function doc_copy(sample_id, folder_id) {
  let doc = DriveApp.getFileById(sample_id);
  let folder = DriveApp.getFolderById(folder_id);
  let newfile = doc.makeCopy(doc.getName().replace("ひな形",""), folder);

  return newfile.getUrl()
}

function openDoc(url){
  const basedoc = DocumentApp.openByUrl(url);

  return basedoc.getBody()
}

function replace_string(basebody) {
  for(let i in BIND_STRING) {
    basebody.replaceText(BIND_STRING[i][0],BIND_STRING[i][1])
  }
}

function string_format(data) {
  let type = Object.prototype.toString.call(data)
  let format_data = data
  if(type == '[object Date]') {
    format_data = Utilities.formatDate(format_data, 'JST', 'yyyy/MM/dd')
  } else if(type == '[object Number]') {
    format_data = format_data.toLocaleString()
  }

  return format_data;
}