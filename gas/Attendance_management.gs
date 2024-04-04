
//スタンドアロンスクリプト

//各種フォルダのIDを定義
const adminFolder = DriveApp.getFolderById('フォルダIDを設定');
const timeSheetFolder = DriveApp.getFolderById('フォルダIDを設定');
const pdfFolder = DriveApp.getFolderById('フォルダIDを設定');

function main() {
  //管理者用シートを作成
  Logger.log("管理者用シート作成：開始");
  const adminFileId = createAdminFile();
  Logger.log("管理者用シート作成：完了");

  //勤務データの収集
  Logger.log("勤務データ収集：開始");
  const data = getDataFromEmployeeFiles();
  Logger.log("勤務データ収集：完了");

  //勤務データを管理者用シートへ転記
  Logger.log("勤務データの転記：開始");
  writeDataToAdminSheet(adminFileId, data);
  Logger.log("勤務データの転記：完了");

  //勤務表をPDFで出力
  Logger.log("PDF出力：開始");
  exportPdf();
  Logger.log("PDF出力：完了");

  //今月の勤務表をoldフォルダへ
  Logger.log("oldフォルダへの移動：開始");
  movePastFiles();
  Logger.log("oldフォルダへの移動：完了");

  //来月の勤務表を生成
  Logger.log("来月の勤務表生成：開始");
  makeEmployeeFiles();
  Logger.log("来月の勤務表生成：完了");
}

function createAdminFile(){
  //今月の管理者用シートの名前を定義
  const adminSheetName = "管理者用シート_" + dayjs.dayjs().format('YYYYMM');

  //ひな型をコピーして今月の管理者用シートを生成
  const fieldId = adminFolder.getFilesByName("管理者用シート_ひな型").next().makeCopy(adminSheetName).getId();
  return fieldId;
}

function getDataFromEmployeeFiles(){
  //勤務表フォルダから全ファイルを取得
  const files = timeSheetFolder.getFiles();
  const data = [];

  //読み込めるファイルがなくなるまで
  while(files.hasNext()){
    const file = files.next();
    const sheet = SpreadsheetApp.open(file).getActiveSheet();
    //セル範囲からデータを取得し、1次元配列にして「data」に格納
    data.push(Array.prototype.concat.apply([],sheet.getRange("A40:C40").getValues()));
  }
  return data;
}

function writeDataToAdminSheet(adminFileId, data){
  const adminSheet = SpreadsheetApp.open(DriveApp.getFileById(adminFileId)).getActiveSheet();

  //管理者用シートの書き込みエリアを取得
  const range = adminSheet.getRange(2,1,data.length,data[0].length);
  //勤務データを管理者用シートに転記
  range.setValues(data);
}

function exportPdf(){
  const files = timeSheetFolder.getFiles();
  while(files.hasNext()){
    const file = files.next();

    //Blobオブジェクトとしてファイル取得（この時点でコンテンツタイプ=PDFに）
    const blob = file.getBlob();
    //PDFフォルダにPDFファイルを生成
    pdfFolder.createFile(blob);
  }
}

function makeEmployeeFiles(){
  //来月の年、月、年月を取得（dayjs関数）
  const year =  dayjs.dayjs().add(1,'M').format("YYYY");
  const nextMonthStr =  dayjs.dayjs().add(1,'M').format("MM");
  const yyyymm = dayjs.dayjs().add(1,'M').format("YYYYMM");

  const empMasterFile = SpreadsheetApp.open(adminFolder.getFilesByName("社員マスタ").next());
  const empMasterSheet = empMasterFile.getActiveSheet();

  //社員マスタから社員データを取得
  const empData = empMasterSheet.getRange(2,1,empMasterSheet.getLastRow()-1,empMasterSheet.getLastColumn()).getValues();
  const nextMonth = new Date().getMonth()+2;
  const monthDays = new Date(nextMonthStr,nextMonth,0).getDate();

  //社員数分、来月の勤務表を生成（中身も設定）
  for(let i = 0; i < empData.length; i++){
    const empId = empData[i][0];
    const empName = empData[i][1];
    const newFileName = "勤務表_" + yyyymm + "_" + empName;
    const templateFolder = DriveApp.getFolderById('フォルダのIDを設定');
    const newFileId = templateFolder.getFilesByName("勤務表_ひな型").next().makeCopy(newFileName, timeSheetFolder).getId();
    const newFile = SpreadsheetApp.openById(newFileId).getActiveSheet();
    newFile.getRange("A1").setValue(year);
    newFile.getRange("C1").setValue(nextMonth);
    newFile.getRange("A2").setValue(empId);
    newFile.getRange("B2").setValue(empName);

    //A5セル（「上のセルの日付の1日後」の関数あり）を日数分コピー）
    const a5Cell = newFile.getRange("A5");
    const row = 4;
    for(let i = 1; i < monthDays; i++){
      const range = newFile.getRange(row+i,1,1,1);
      a5Cell.copyTo(range);
    }
  }
}

function movePastFiles(){
  const oldFolder = DriveApp.getFolderById('フォルダのIDを設定');

  //oldフォルダ内に今月の年月でフォルダ生成
  const yyyymmFolder = oldFolder.createFolder(dayjs.dayjs().format("YYYYMM"));
  const files = timeSheetFolder.getFiles();
  while(files.hasNext()){
    const file = files.next();
    //ファイルを移動
    file.moveTo(yyyymmFolder);
  }
}



