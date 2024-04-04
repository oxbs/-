
//コンテナバインドスクリプト（スプレッドシートと連携）

//イベントハンドラ：イベント発生時に処理する（onOpenはスプレッドシートを開くと実行）
function onOpen() { 

  let ui = SpreadsheetApp.getUi()

  //「追加メニュー」というメニューがスプレッドシートに追加される
  let menu = ui.createMenu("追加メニュー"); 

  //「カレンダー登録」というアイテム名を設定、calenderAddの関数を実行
  menu.addItem("カレンダー登録", "calenderAdd"); 

  //「クリア」というアイテム名を設定、clearCellの関数を実行
  menu.addItem("クリア", "clearCell");

  //画面上のメニューとして追加するために必要な処理
  menu.addToUi(); 
}

function calenderAdd() {

  var data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();

  //1行目は項目名、2行目はサンプル入力行なので削除
  data.shift();
  data.shift();

  for (var row of data){
    var title = row[0];
    var date = row[1];
    var startTime = row[2];
    var endTime = row[3];

    //new Dateで新しい日付オブジェクトとして定義しないと、他の日付変数と競合するっぽい
    var startDate = new Date(date); 

    //日付に時間情報を付与
    startDate.setHours(startTime.getHours()); 

    //日付に分情報を付与　⇒カレンダー登録できる日時形式になる
    startDate.setMinutes(startTime.getMinutes()); 

    var endDate = new Date(date);
    endDate.setHours(endTime.getHours());
    endDate.setMinutes(endTime.getMinutes());

    //オプションとして「説明」「場所」項目を取得
    var option = {
      description: row[4],
      location: row[5]
    }

    //Googleカレンダーを呼び出す（引数には自身のID（Googleメールアドレス）を設定）
    let calender = CalendarApp.getCalendarById("oxbs2005@gmail.com"); 

    //カレンダーにイベント登録
    calender.createEvent( 
      title,
      startDate,
      endDate,
      option
    );

  }

  //ブラウザ上にポップアップメッセージを出力
  Browser.msgBox("カレンダーに登録しました");
}

function clearCell(){

  let rSheet = SpreadsheetApp.getActiveSheet();
  let lastRow = rSheet.getLastRow(); 

  //削除範囲の設定
  var clearData = rSheet.getRange(3,1,lastRow-2,6);

  //セルの入力内容だけを削除（書式などは変えない）
  clearData.clearContent();
}

