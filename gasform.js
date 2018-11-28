/* カレンダーへイベントを登録する */
function getsheet() {

/*-前準備-*/

　　　//シートの項目を以下変数定義
　　　 　var sht, i, eventname, eventcontent, eventday, start, end, profile, added;

　　　//shtを定義
　　　　sht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("form");

　　　//シートの2行目〜最終行まで処理を繰り返す
　　　　for(i = 2; i <= sht.getLastRow(); i++) {

/*-前準備ここまで-*/


/*--スプレッドシートの値を取得して変数へ格納--*/ 

//i行2列目の値(イベントの名前)をeventnameへ格納
　eventname = sht.getRange(i,2).getValue();

//i行3列目の値（イベントの内容）をplaceへ格納
　eventcontent = sht.getRange(i,3).getValue();
//i行4列目の値(イベント日)をeventdayへ格納
　eventday = sht.getRange(i, 4).getValue();

//開始日をUtilities.formatDateでフォーマットしてbへ格納
　var b = Utilities.formatDate(eventday,"JST","yyyy/MM/dd");

//i行5列目の値(開始時刻)をstartへ格納
　var starttime = sht.getRange(i,5).getValue();

　var H = starttime.getHours();//starttimeの時間を取得してHへ格納
　var M = starttime.getMinutes();//starttimeの時間を取得してMへ格納
　var S = starttime.getSeconds();//starttimeの時間を取得してSへ格納

//new Dateメソッドで開始日時「yyMMdd hh:mm」をstartへ格納
　var start = new Date(b+" "+H+":"+M+":"+S);　

//i行6列目の値(終了時刻)をendへ格納
　var endtime = sht.getRange(i,6).getValue();

　var H1 = endtime.getHours();//endtimeの時間を取得してH1へ格納
　var M1 = endtime.getMinutes();//endtimeの分を取得してM1へ格納
　var S1 = endtime.getSeconds();//endtimeの秒を取得してS1へ格納

//new Dateメソッドで終了日時「yyMMdd hh:mm」をendへ格納
　var end = new Date(b+" "+H1+":"+M1+":"+S1);

//i行7列目の値(代表者の名前)をprofileへ格納
　profile = sht.getRange(i,7).getValue();


/*--カレンダーへ登録--*/

//i行6列目の値(イベント登録有無)をaddedへ格納
　added = sht.getRange(i,8).getValue();

//addedの値が空白だったらカレンダー登録を実行
　if(added == "") {
　　Cal = CalendarApp.getCalendarById("********************");//<---ここをカレンダーIDへ変更する

//指定のカレンダーIDへインベント登録
   Cal.createEvent(eventname, start, end, {description: eventcontent,location: profile});//createEvent(タイトル、開始日時、終了日時、オプション）

//カレンダー登録が終わったイベントのaddedへ「登録完了」を記入
　sht.getRange(i,8).setValue("登録完了");

　　}　　　//ifを閉じる
　}　　　　　　//forを閉じる
}　　　　　　　　//functionを閉じる
