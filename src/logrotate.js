//バックアップ元ファイルID
const FILE_ID = '<ファイルID>';
//バックアップ先フォルダー名
const DIR_NAME = 'org'
//スプレッドシート名
const SHEET_NAME = 'DATA'
//バックアップ・ローテーションを実行する行数
const LIMIT_ROWS = 10
//削除するレコードの日にち（指定した日数以前のログが削除）
const DEL_DAY = 10


function renameFile(filename) {
  var date = new Date();
  var t = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMddhhmmss');
  var newName = filename + '_' + String(t);
  return newName.trim()
}

function bkCellCopy(targetFileId, sheetName, targetDirName) {
  //バックアップ元ファイルは ID 指定
  var targetFile = DriveApp.getFileById(targetFileId);
  //ファイルルートからバックアップ先のフォルダを確認
  var root = DriveApp.getRootFolder(); 
  var targetDirs = root.getFoldersByName(targetDirName);
  
  if (targetDirs.hasNext()) {
    //フォルダが存在する場合は最初のフォルダをバックアップ先とする
    var bkDir = targetDirs.next();
  } else {
    //フォルダーが存在しない場合は作成
    var bkDir = DriveApp.createFolder(targetDirName);
  }
  var bkFolderId = bkDir.getId();
  //バックアップファイル名の作成
  var bkFileName = renameFile(targetFile);
  //makeCopy だとGoogle App のプロジェクトも複製されるため、シートの中身をコピー
  // targetFile.makeCopy(bkFileName, targetDir);
  
  //空のバックアップファイルを作成しIDを取得
  var bkFileId = SpreadsheetApp.create(bkFileName).getId();
  //ファイルを移動させるために File オブジェクトを取得
  var bkFile = DriveApp.getFileById(bkFileId);

  var copyToSheet = SpreadsheetApp.openById(bkFileId).getActiveSheet();
  var copyFromSheet = SpreadsheetApp.openById(targetFileId).getSheetByName(sheetName);
  //コピー先のシートをクリア  
  copyToSheet.clear();
  var lastRow = copyFromSheet.getLastRow();
  var lastColumn = copyFromSheet.getLastColumn();
  var copyValues = copyFromSheet.getRange(1,1,lastRow,lastColumn).getValues();
  copyToSheet.getRange(1,1,lastRow,lastColumn).setValues(copyValues);
  
  //コピー先ファイルをバックアップフォルダーに移動
  DriveApp.getFolderById(bkFolderId).addFile(bkFile);
  //rootディレクトリーからコピー先ファイルを削除
  root.removeFile(bkFile);
  return 1;
}

function bkFileCopy(targetFileId, sheetName, targetDirName) {
  //バックアップ元ファイルは ID 指定
  var targetFile = DriveApp.getFileById(targetFileId);
  //ファイルルートからバックアップ先のフォルダを確認
  var root = DriveApp.getRootFolder(); 
  var targetDirs = root.getFoldersByName(targetDirName);
  
  if (targetDirs.hasNext()) {
    //フォルダが存在する場合は最初のフォルダをバックアップ先とする
    var bkDir = targetDirs.next();
  } else {
    //フォルダーが存在しない場合は作成
    var bkDir = DriveApp.createFolder(targetDirName);
  }
  //バックアップファイル名の作成
  var bkFileName = renameFile(targetFile);
  //makeCopy だとGoogle App のプロジェクトも複製される
  targetFile.makeCopy(bkFileName, bkDir);
  return 1;
}

function deleteRecords(targetFileId, sheetName, delDay) {
  var sheet = SpreadsheetApp.openById(targetFileId).getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  //日付の列を取得
  //時刻列を2行目から読みとる必要はないので、最新データから2割を getValues する
  var rRow = Math.floor((lastRow) * 0.8);
  var dt = sheet.getRange(rRow, 1, lastRow - rRow + 1).getValues();
  //Logger.log("rRow: " + rRow + " lastRow: " + lastRow + " dt.length: " + dt.length);
  
  //バックアップ対象の日付を取得
  var bkd = new Date();
  bkd.setDate(bkd.getDate() - delDay);
  //Logger.log(Utilities.formatDate(bkd, 'Asia/Tokyo', 'yyyyMMdd'));
  
  var bkdFormat_yyyyMMdd = String(Utilities.formatDate(bkd, 'Asia/Tokyo', 'yyyyMMdd')).trim();
  //バックアップ対象の行数を格納する変数
  var delRow = null;

  //シートの最終行と配列dtの後ろからチェック
  for (var i = lastRow, j = dt.length - 1; i != -1 && j != -1; i--, j--) {
    var tempFormat_yyyyMMdd = String(Utilities.formatDate(dt[j][0], 'Asia/Tokyo', 'yyyyMMdd')).trim();
    //バックアップ対象と同じ日付
    if (tempFormat_yyyyMMdd === bkdFormat_yyyyMMdd) {
      Logger.log(bkdFormat_yyyyMMdd + " が一致しました");
      //最終行からカウントした値
      delRow = i;
      //Logger.log(delRow);
      break;
    }
  }
  //deleteRows で古いデータを削除
  if (delRow != null) {
    //2列目から delRow - 2 列を削除する
    sheet.deleteRows(2, delRow - 2);
    Logger.log("bkdFormat_yyyyMMdd: " + bkdFormat_yyyyMMdd + " の行より以前を削除しました");
    return 1;
  } else {
    //バックアップ対象がない場合は終了
    Logger.log("bkdFormat_yyyyMMdd: " + bkdFormat_yyyyMMdd + " の行が存在しません");
    return 0;
  }
}

function isBackupTarget(targetFileId, sheetName, limitRows) {
  var sheet = SpreadsheetApp.openById(targetFileId).getSheetByName(sheetName);
  if (sheet === null) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow <= limitRows) {
    Logger.log('バックアップ対象なし');
    return 0
  } 
  return 1;
}

function main() {
  Logger.log("START");
  if(isBackupTarget(FILE_ID, SHEET_NAME, LIMIT_ROWS)) {
    //bkFileCopy(FILE_ID, SHEET_NAME, DIR_NAME);
    //元ファイルのデータを削除
    deleteRecords(FILE_ID, SHEET_NAME, DEL_DAY);
  }
  //UnitTest
  //deleteRecords(FILE_ID, SHEET_NAME, DEL_DAY);
  Logger.log("END");
}
