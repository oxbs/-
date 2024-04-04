
//コンテナバインドスクリプト（スプレッドシートと連携）

// LINE Developerのアクセストークン
var access_token = "WxE11DiRHD5ONOOKUe6UP3MRqkFIsQzik2sx+fu9mtiFGTWaWYqepaWFlSJ8IdEYfWiGkRTQJ+3C3uNcbpX/K2r58MJUvCTfsRudRyHlfwd2gRcbp2IMZJowONX/QWw9mFqXCkClHdcPk3dyHWjObwdB04t89/1O/w1cDnyilFU="

/**
 LINEからのPOST受け取り
 */
//クライアント（LINE）からのPOSTメソッドに応答（シンプルトリガー：ある操作をきっかけに関数が実行される仕組み）
//引数eは「イベントオブジェクト」
function doPost(e) {
  
  var json = JSON.parse(e.postData.contents);
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log').getRange(1, 1).setValue(json.events);

  var event = JSON.parse(e.postData.contents).events[0];
  var user_id = event.source.userId;
  var eventType = event.type;
  var nickname = getUserProfile(user_id);
   
  // botが友達追加された場合に起きる処理（LINEIDとニックネームの取得）
  if(eventType == "follow") {
    var data2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LINEID');
    var last_row = data2.getLastRow();
    for(var i = last_row; i >= 1; i--) {
      if(data2.getRange(i,1).getValue() != '') {
        var j = i + 1;
        data2.getRange(j,1).setValue(nickname);
        data2.getRange(j,2).setValue(user_id);
        data2.getDataRange().removeDuplicates([2]);
        break;
      }
    }
  }
  //reply関数呼び出し
  reply(json);
  
}

// profileを取得してくる関数
function getUserProfile(user_id){ 
  var url = 'https://api.line.me/v2/bot/profile/' + user_id;
  var userProfile = UrlFetchApp.fetch(url,{
    'headers': {
      'Authorization' :  'Bearer ' + access_token,
    },
  })
  return JSON.parse(userProfile).displayName;
}

/**
 reply(LINEに返す内容の定義)
 */
function reply(data) {
  var url = "https://api.line.me/v2/bot/message/reply";
  var headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    'Authorization': 'Bearer ' + access_token,
  };

  // 取得したメッセージを改行コードで分割
  var message = data.events[0].message.text.split("\n");
  
  // 分割後のメッセージを取得(1行目：inpmsg1、2行目：inpmsg2、3行目：inpmsg3に格納)
  var inpmsg1 = message[0];
  var inpmsg2 = message[1];
  var inpmsg3 = message[2];
  var text = "";

  switch(inpmsg1){
  
    // 1行目に「名言」を入力した場合
    case "名言":

      // アクティブスプレッドシートの名言シートを取得
      var meigen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("名言");
      // A1セルから入力されている最終行まで一気に取得
      var meigenData = meigen.getRange(1, 1, meigen.getLastRow());
      // ランダムで候補を選ぶ
      var intRandomNum = Math.floor(Math.random()*meigen.getLastRow());
  
      text = meigenData.getValues()[intRandomNum][0];
    
      break;  

    // 1行目に「追加」を入力した場合
    // ※2行目には「追加したい名言」を引数で渡す
    case "追加":
      text = meigenPlus(inpmsg2);
      break;  

    // 1行目に「削除」を入力した場合
    // ※2行目には「削除したい名言」を引数で渡す
    case "削除":
      text = meigenDelete(inpmsg2);
      break;  
      
    // 1行目に「誕生日」を入力した場合
    // ※2行目には「一覧」を引数で渡すか（登録している誕生日一覧が返ってくる）
    // ※2行目に「名前」、3行目に「誕生日（MMDD形式）」を引数で渡せば登録できる
    case "誕生日":
      text = birthdayLog(inpmsg2,inpmsg3);
      break;  

    // 1行目に「予定」を入力した場合
    case "予定":
      //Googleカレンダーから今日の予定取得
      let TodayEvents = GetEventsFunction();
      let TodayEventsList = "";
      text = "＜今日の予定＞" + "\n";

      //予定がないときは「特になし」を返す
      if (TodayEvents == "特になし"){
        text += "■" + "特になし";
      }else{

        //予定の数だけループ（
        for(let i = 0;i < TodayEvents.length;i++){

          //最終行の時だけ改行を入れない
          if(i == TodayEvents.length - 1){
            text += "■" + TodayEvents[i][0] + " " + TodayEvents[i][1];
          //最終行以外は改行を入れる
          }else{
            text += "■" + TodayEvents[i][0] + " " + TodayEvents[i][1] + "\n";
          }
        }
      }
      break;  
    
    // 1行目に「天気」を入力した場合
    case "天気":
      let TodayWeather = GetWeatherFunction();
      text = "＜今日の天気(" + TodayWeather[6] +")＞\n●天気：" + TodayWeather[0] + "\n●最高気温：" + TodayWeather[1] + "\n●最低気温：" + TodayWeather[2] + "\n●6-12時の降水確率：" + TodayWeather[3] + "\n●12-18時の降水確率：" + TodayWeather[4] + "\n●18-24時の降水確率：" + TodayWeather[5];
      break;  

    // 1行目に「電車遅延」を入力した場合　※機能見直し中！
    case "電車遅延":
      text = trainDelayInfo2();
      break;

    // 別途LINEAPI側の機能で返答するのでGAS側では何もしない
    case "機能":
      break;  

    default:
      
      //1行目に「YYYY年MM月DD日」と入力すると年齢と生まれてからの日数を返す
      if(inpmsg1.substring(4,5) === "年" && inpmsg1.substring(7,8) === "月" && inpmsg1.substring(10,11) === "日"){
        var birthday = inpmsg1.substring(0,4) + inpmsg1.substring(5,7) + inpmsg1.substring(8,10);
        
        if(!isNaN(birthday)){
          text = birthdayCul(birthday);   
        }
        else{
          text = "生年月日は「YYYY年MM月DD日」の形式で、YYYYMMDD部分は数字で入力ください！"
        }
      }      
  }  

  //APIリクエスト時にセットするデータを設定
  var postData = {
    "replyToken" : data.events[0].replyToken,
    "messages" : [
      {
        'type':'text',
        // textの内容をLINEの画面上に出力
        'text':text,
      }
    ]
  };

  //HTTPSのPOST時のオプションパラメータを設定
  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };

  // LINE Messaging APIにHTMLを投げる（ユーザーからの投稿に返答）
  return UrlFetchApp.fetch(url, options);
}

/**
 名言リスト追加
 */
function meigenPlus(add) {

  var meigen = ""
  // 名言リストスプレッドシートを取得
  meigen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("名言");

  if(!add){
    return "追加したい言葉を入力ください！";  
  }else{

    // 最終行の次行に名言追加
    var lastRowNext = meigen.getLastRow(); 
    lastRowNext = lastRowNext + 1;
    meigen.getRange(lastRowNext,1).setValue(add);
    return "名言リストに「" +add+ "」を追加しました！";
  }
}

/**
 名言リスト削除
 */
function meigenDelete(del) {

  var meigen = ""

  // 名言リストスプレッドシートを取得
  meigen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("名言");

  if(!del){
    return "削除したい言葉を入力ください！";  
  }else{
    var lastRow = meigen.getLastRow(); 

    // 削除したい行をリストから索引
    for(let i = 1; i<lastRow+1 ;i++ ){
      var meigenData = meigen.getRange(i,1).getValue();
      if(meigenData === del){
        meigen.deleteRows(i);
        return "名言リストから「" +del+ "」を削除しました！";
        break;
      }
    }
    return "名言リストに「" +del+ "」はありません！";
  }
}

/**
 誕生日登録
 */
function birthdayLog(inpmsg2,inpmsg3) {

  var data = SpreadsheetApp.openById("1_sJUfmcBvOBmMH3MDT3VSSRfXv-1UANVLWBKXz6Sc_w").getSheetByName("誕生日");
  var last_row = data.getLastRow();
  var birthdays = data.getRange(2, 1, last_row,2).getValues();
  // birthday2には名前が入る
  var birthday2 = "";
  var birthday3 = "＜誕生日一覧＞";
  
  if(!inpmsg2){
      return "改行して情報入力してください！";  
  }
  else if(inpmsg2==="一覧"){
    
    for(let i = 0; i<last_row-1 ;i++ ){
      for(let j = 0; j<2 ;j++ ){
        birthday2 = String(birthdays[i][j]);
        
        if(j===0){
          // birthday3に改行と「名前：」を追加
          birthday3 = birthday3+"\n★"+birthday2+ "：";
        }
        else{
          if(birthday2.length===4){
            var month = birthday2.substring(0,2);

            //日の頭に0があれば表示されないようにする
            if(birthday2.substring(2,3) == 0){
              var day = birthday2.substring(3,4);
            }else{
              var day = birthday2.substring(2,4);
            }
          }
          else{
            // 3桁で誕生日が登録されている場合を考慮（MDD）
            var month = birthday2.substring(0,1);

            //日の頭に0があれば表示されないようにする
            if(birthday2.substring(2,3) == 0){
              var day = birthday2.substring(3,4);
            }else{
              var day = birthday2.substring(2,4);
            }
          }
          // birthday3に「X月X日」を追加
          birthday3 = birthday3+month+"月"+day+"日";
        }
      } 
    }
    return birthday3;
    /**
    こんな感じで出力
    ＜誕生日一覧＞
    A：1月1日
    B：3月3日
    C：5月5日
    */
    
  }  
  else{  
    if(!inpmsg3){
      return inpmsg2+ "の誕生日を入力してください！";  
    }
    else{

      if(!isNaN(inpmsg3) && inpmsg3.length===4){

        var month = inpmsg3.substring(0,2);
        var day = inpmsg3.substring(2,4);

        for(var i = last_row; i >= 1; i--) {
          if(data.getRange(i,1).getValue() != '') {
            var j = i + 1;
            data.getRange(j,1).setValue(inpmsg2);
            data.getRange(j,2).setValue(inpmsg3);
            data.getDataRange().removeDuplicates([1]);
            return inpmsg2+"の誕生日（"+month+"月"+day+"日）を登録しました！";
            break;
          }
        }        
      }
      else{
        return "誕生日は数字4桁で入力ください！";
      }
    }
  }
}


/**
 誕生日オプション（年齢と生まれたからの日数計算）
 */
function birthdayCul(birthday) {
  
  var today = new Date();
  var strToday = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  today2 = strToday.substring(0,4) + strToday.substring(5,7) + strToday.substring(8,10);
  
  if(birthday < today2){

    var nenrei;
    
    var birthdayY = birthday.substring(0,4);
    birthdayY = Number(birthdayY);

    var birthdayY2 = birthdayY;

    var birthdayM = birthday.substring(4,6);
    if(birthdayM.substring(0,1) === "0"){
      birthdayM = birthdayM.substring(1,2);
    }
    birthdayM = Number(birthdayM);

      var birthdayD = birthday.substring(6,8);
    if(birthdayD.substring(0,1) === "0"){
      birthdayD = birthdayD.substring(1,2);
    }
    birthdayD = Number(birthdayD);

    var today2Y = today2.substring(0,4);
    today2Y = Number(today2Y);

    var today2M = today2.substring(4,6);
    if(today2M.substring(0,1) === "0"){
      today2M = today2M.substring(1,2);
    }
    today2M = Number(today2M);
    
    var today2D = today2.substring(6,8);
    if(today2D.substring(0,1) === "0"){
      today2D = today2D.substring(1,2);
    }
    today2D = Number(today2D);

    var totaldays = 0
    var tsuki = new Array(1,2,3,4,5,6,7,8,9,10,11,12);
    var nissu = new Array(31,28,31,30,31,30,31,31,30,31,30,31);

    if(birthdayY === today2Y){

      if((birthdayY % 4 === 0 && birthdayY % 100 != 0) || birthdayY % 400 === 0){
        nissu[1] = 29;
      }
      else{
        nissu[1] = 28;
      }
      for(let i = 0; i<today2M-1 ;i++ ){
        totaldays = totaldays + nissu[i];          
      }
      totaldays = totaldays + today2D;

      for(let i = 0; i<birthdayM-1 ;i++ ){
        totaldays = totaldays - nissu[i];          
      }
      totaldays = totaldays - birthdayD;
      
    }
    else{
      
      while(birthdayY <= today2Y){
        
        //生まれた年の日数をカウント
        if(birthdayY === birthdayY2){

          if((birthdayY % 4 === 0 && birthdayY % 100 != 0) || birthdayY % 400 === 0){
            nissu[1] = 29;
          }
        
          for(let i = birthdayM + 1; i<13 ;i++ ){
            totaldays = totaldays + nissu[i-1];
          }       
          totaldays = totaldays + nissu[birthdayM-1]-birthdayD;
        }

        //生まれた翌年から今年までの日数をカウント
        else if(birthdayY < today2Y){
        
          if((birthdayY % 4 === 0 && birthdayY % 100 != 0) || birthdayY % 400 === 0){
            totaldays = totaldays + 366;           
          }
          else{
            totaldays = totaldays + 365; 
          }
        }

        //今年の日数をカウント
        else{
          if((birthdayY % 4 === 0 && birthdayY % 100 != 0) || birthdayY % 400 === 0){
            nissu[1] = 29;
          }
          else{
            nissu[1] = 28;
          }
          for(let i = 0; i<today2M-1 ;i++ ){
            totaldays = totaldays + nissu[i];          
          }
          totaldays = totaldays + today2D;
        }
        birthdayY = birthdayY + 1;      
      }
    }
    nenrei = Math.floor(totaldays / 365);
    return "あなたは今、" +nenrei+ "歳で\n今日はあなたが生まれてから\n" +totaldays+ "日目の記念日です！"
  }

  else if(birthday == today2){
    return "今日生まれたんですね！おめでとうございます！"    
  }  
  else{
    return "あなたは未来から来たのですか？"    
  }
  console.log(totaldays);
}


//天気予報の取得
function GetWeatherFunction(){
  //指定したURLから天気情報をJSON形式で取得。それを「JSON.parse」で配列化。
  //URLの「130010」は東京を表す。数字を変えると別の地域の天気予報も取得できる。（数字は右記URLを参照⇒https://weather.tsukumijima.net/primary_area.xml）
  //名古屋は「230010」、大阪は「270000」
  //取得できる情報の詳細は「https://weather.tsukumijima.net/」を参照。
  const API_Data = JSON.parse(UrlFetchApp.fetch("https://weather.tsukumijima.net/api/forecast/city/130010").getContentText());

  let WeatherData = [];
  {
    //配列情報の中から必要な要素を取得
    WeatherData[0] = API_Data["forecasts"][0]["telop"] //天気
    WeatherData[1] = API_Data["forecasts"][0]["temperature"]["max"]["celsius"] //最高気温
    if(WeatherData[1] == null){
      WeatherData[1] = "不明";
    }else{
      WeatherData[1] += "℃";
    }
    WeatherData[2] = API_Data["forecasts"][0]["temperature"]["min"]["celsius"] //最低気温
    if(WeatherData[2] == null){
      WeatherData[2] = "不明";
    }else{
      WeatherData[2] += "℃";
    }
    WeatherData[3] = API_Data["forecasts"][0]["chanceOfRain"]["T06_12"] //降水確率（6-12時）
    WeatherData[4] = API_Data["forecasts"][0]["chanceOfRain"]["T12_18"] //降水確率（6-12時）
    WeatherData[5] = API_Data["forecasts"][0]["chanceOfRain"]["T18_24"] //降水確率（6-12時）
    WeatherData[6] = API_Data["location"]["city"] //場所
    //location={area=近畿, city=大阪, district=大阪府, prefecture=大阪府}

  }
  return WeatherData;
}

//天気の自動通知機能
function push_weather() {

  let TodayWeather = GetWeatherFunction();
  text = "＜今日の天気(" + TodayWeather[6] +")＞\n●天気：" + TodayWeather[0] + "\n●最高気温：" + TodayWeather[1] + "\n●最低気温：" + TodayWeather[2] + "\n●6-12時の降水確率：" + TodayWeather[3] + "\n●12-18時の降水確率：" + TodayWeather[4] + "\n●18-24時の降水確率：" + TodayWeather[5];

  var postData = {
    //自分のLINEID（他の人のIDに変えれば、他の人に届く。IDをリスト化してループすれば複数人にまとめて送れる）
    "to": "Ub34e68348214e6644a53738f6d0d1c3e",
    "messages": [
      {
        "type": "text",
        "text": text,
      }
    ]
  }

  var headers = {
    "Content-Type": "application/json",
    'Authorization': 'Bearer ' + access_token,
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(postData)
  };
  var response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
}

//Googleカレンダーから今日の予定を取得して返す
function GetEventsFunction() {

  //Googleカレンダーの情報取得（引数を他の人のGoogleカレンダーIDにすれば、そのカレンダー情報を取得できる）
  const myCalendar = CalendarApp.getCalendarById("oxbs2005@gmail.com")
  const today = new Date;

  //今日の予定を取得
  const myEvents = myCalendar.getEventsForDay(today);

  //予定が無ければ「特になし」を返す
  if(myEvents[0] === undefined){
    return ["特になし"];
  } else {

    let ReturnData = [];

    //予定の数だけループ（タイトルと開始時間をリスト格納）
    for (let i = 0;i < myEvents.length;i++){
      startTime = dayjs.dayjs(myEvents[i].getStartTime()).format("HH:mm");
      title = myEvents[i].getTitle();
      ReturnData.push([startTime,title]);
    }

    return ReturnData;
  }
}

//今日の予定の自動通知機能
function push_schedule() {

  let TodayEvents = GetEventsFunction();
  let TodayEventsList = "";
  text = "＜今日の予定＞" + "\n";

  //予定がないときは「特になし」を返す
  if (TodayEvents == "特になし"){
    text += "■" + "特になし";
  }else{

    //予定の数だけループ（
    for(let i = 0;i < TodayEvents.length;i++){

      //最終行の時だけ改行を入れない
      if(i == TodayEvents.length - 1){
        text += "■" + TodayEvents[i][0] + " " + TodayEvents[i][1];
      //最終行以外は改行を入れる
      }else{
        text += "■" + TodayEvents[i][0] + " " + TodayEvents[i][1] + "\n";
      }
    }
  }

  var postData = {
    //自分のLINEID（他の人のIDに変えれば、他の人に届く。IDをリスト化してループすれば複数人にまとめて送れる）
    "to": "Ub34e68348214e6644a53738f6d0d1c3e",
    "messages": [
      {
        "type": "text",
        "text": text,
      }
    ]
  }

  var headers = {
    "Content-Type": "application/json",
    'Authorization': 'Bearer ' + access_token,
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(postData)
  };
  var response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
}


//　曜日の日本語変換
function toWD(date){
  var myTbl = new Array("日","月","火","水","木","金","土","日"); 
  var myDay = Utilities.formatDate(date, "JST", "u");
  return "(" + myTbl[myDay] + ")";
}




