//バックアップ元ファイルID
const FILE_ID = '<ファイルID>';
//バックアップ先フォルダー名
const DIR_NAME = '<フォルダー名>'
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
  var targetFile = DriveApp.getFileById(targetFileId);
  var root = DriveApp.getRootFolder(); 
  var targetDirs = root.getFoldersByName(targetDirName);
  
  if (targetDirs.hasNext()) {
    var bkDir = targetDirs.next();
  } else {
    var bkDir = DriveApp.createFolder(targetDirName);
  }
  var bkFolderId = bkDir.getId();
  var bkFileName = renameFile(targetFile);
  //makeCopy だとGoogle App のプロジェクトも複製されるため、シートの中身をコピー
  // targetFile.makeCopy(bkFileName, targetDir);
  
  var bkFileId = SpreadsheetApp.create(bkFileName).getId();
  var bkFile = DriveApp.getFileById(bkFileId);

  var copyToSheet = SpreadsheetApp.openById(bkFileId).getActiveSheet();
  var copyFromSheet = SpreadsheetApp.openById(targetFileId).getSheetByName(sheetName);

  copyToSheet.clear();
  var lastRow = copyFromSheet.getLastRow();
  var lastColumn = copyFromSheet.getLastColumn();
  var copyValues = copyFromSheet.getRange(1,1,lastRow,lastColumn).getValues();
  copyToSheet.getRange(1,1,lastRow,lastColumn).setValues(copyValues);
  
  DriveApp.getFolderById(bkFolderId).addFile(bkFile);
  root.removeFile(bkFile);
  return 1;
}

function bkFileCopy(targetFileId, sheetName, targetDirName) {
  var targetFile = DriveApp.getFileById(targetFileId);
  var root = DriveApp.getRootFolder(); 
  var targetDirs = root.getFoldersByName(targetDirName);
  
  if (targetDirs.hasNext()) {
    var bkDir = targetDirs.next();
  } else {
    var bkDir = DriveApp.createFolder(targetDirName);
  }
  var bkFileName = renameFile(targetFile);
  targetFile.makeCopy(bkFileName, bkDir);
  return 1;
}

function deleteRecords(targetFileId, sheetName, delDay) {
  var sheet = SpreadsheetApp.openById(targetFileId).getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  var rRow = Math.floor((lastRow) * 0.8);
  var dt = sheet.getRange(rRow, 1, lastRow - rRow + 1).getValues();
  //Logger.log("rRow: " + rRow + " lastRow: " + lastRow + " dt.length: " + dt.length);
  
  var bkd = new Date();
  bkd.setDate(bkd.getDate() - delDay);
  //Logger.log(Utilities.formatDate(bkd, 'Asia/Tokyo', 'yyyyMMdd'));
  
  var bkdFormat_yyyyMMdd = String(Utilities.formatDate(bkd, 'Asia/Tokyo', 'yyyyMMdd')).trim();
  var delRow = null;

  for (var i = lastRow, j = dt.length - 1; i != -1 && j != -1; i--, j--) {
    var tempFormat_yyyyMMdd = String(Utilities.formatDate(dt[j][0], 'Asia/Tokyo', 'yyyyMMdd')).trim();

    if (tempFormat_yyyyMMdd === bkdFormat_yyyyMMdd) {
      Logger.log(bkdFormat_yyyyMMdd + " が一致しました");
      delRow = i;
      //Logger.log(delRow);
      break;
    }
  }

  if (delRow != null) {
    sheet.deleteRows(2, delRow - 2);
    Logger.log("bkdFormat_yyyyMMdd: " + bkdFormat_yyyyMMdd + " の行より以前を削除しました");
    return 1;
  } else {
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
    bkFileCopy(FILE_ID, SHEET_NAME, DIR_NAME);
    deleteRecords(FILE_ID, SHEET_NAME, DEL_DAY);
  }
  Logger.log("END");
}
