var manager = '';
var apps = {};

// メイン処理
function main() {

  prepare();
  control();

}


// 初期化
function prepare() {

  // 各種設定
  var subdomain = PropertiesService.getScriptProperties().getProperty("subdomain");
  var user      = PropertiesService.getScriptProperties().getProperty("user");
  var pass      = PropertiesService.getScriptProperties().getProperty("pass");

  // アプリ情報の設定
  var sheet     = SpreadsheetApp.getActive().getSheetByName('マスタ');
  var data      = sheet.getDataRange().getValues();
  var appName   = '';

  // 2行目から検索
  for (var i = 1; i < data.length; i++) {

    // 1列目が空ではないとき
    appName = data[i][0];
    if('' !== appName){
      apps[appName] = {};
      apps[appName]['appid'] = data[i][1];
      apps[appName]['name']  = data[i][2];
    }
  }
  Logger.log(apps);

  // パスワード認証
  manager = new KintoneManager.KintoneManager(subdomain, apps, user, pass);
}

// 検索
function control() {

  // 操作シートの情報を取得
  var sheet      = SpreadsheetApp.getActive().getSheetByName('操作');
  var data       = sheet.getDataRange().getValues();
  var utillList  = '';
  var query      = '';
  var records    = '';
  var sheetWrite = '';
  var appName    = '';

  // 2行目から検索
  for (var i = 1; i < data.length; i++) {

    // 1列目がtrueの場合
    utillList = data[i];
    if (utillList[0]) {

      // kintoneの対象アプリ名を取得
      appName = searchApp(utillList[1]);

      // クエリを取得
      query = utillList[2];

      // 検索処理
      records = search(appName, query);

      // 書き込みシートを取得
      sheetWrite = SpreadsheetApp.getActive().getSheetByName(utillList[3]);
      writeSheet(sheetWrite, records);
    }
  }
}



// 検索処理
function search(appName, query){

  // 検索を実行
  var response = manager.search(appName, query);

  // 結果コード
  Logger.log('ステータスコード：' + response.getResponseCode());

  // レコードの配列を返却
  return JSON.parse(response.getContentText()).records;
}

// シートに書き込み
function writeSheet(sheet, records){


  // 列見出しを取得
  var array_kintone_fields = sheet.getRange("1:1").getValues()[0];
  array_kintone_fields = array_kintone_fields.filter(Boolean);

  // 書き込み行
  var row = 2;

  // 100行の内容を削除
  sheet.getRange(row, 1, 100, array_kintone_fields.length).clearContent();

  // レコードが取得された場合
  if(typeof records !== 'undefined'){

    // 値設定
    records.forEach(function(record){
      array_kintone_fields.forEach(function(kintone_field,index){
        sheet.getRange(row,index+1).setValue(record[kintone_field].value);
      })
       row++;
    })

  // レコードが取得されなかった場合
  }else{
    sheet.getRange(row,1).setValue('あれ？取得できませんでした…');
  }



}

// アプリケーション名を検索
function searchApp(appId){

  // 対象アプリを検索
  for (var key in apps) {

    // アプリIDが一致した場合
    if(apps[key].appid === parseInt(appId, 10)){

      // アプリ名を返却
      Logger.log('アプリ→' + key);
      return key;
    }
  }

  // 見つからなかった場合
  Logger.log('アプリが見つかりませんでした');
  return "";
}
