// スプレッドシート上のデータを元にリスト（プルダウンリスト）更新
// （事前にformを作成しておき、タイトル「担当者」でプルダウンリスト（中身はなんでも良い）を作成しておくこと
function overwriteFormList() {

  // 取得元のスプレッドシートを開く
  const SS_ID   = PropertiesService.getScriptProperties().getProperty('SS_ID');
  console.log(SpreadsheetApp.openById(SS_ID));
  const sheet   = SpreadsheetApp.openById(SS_ID).getSheetByName('質問');

  // 更新対象フォームを取得
  const FORM_ID =  PropertiesService.getScriptProperties().getProperty('FORM_ID');
  const form = FormApp.openById(FORM_ID);

  // 対象者(スプレッドシート上の担当者列タイトル行の位置を指定して取得)
  if("担当者" == sheet.getRange("A1").getValue()){
    var staffList = sheet.getRange(1, 1, sheet.getLastRow() - 1).getValues();
  }

  // 質問項目がプルダウンのもののみ取得
  var items = form.getItems(FormApp.ItemType.LIST);
  console.log(items);
  items.forEach(function(item){

    if(item.getTitle().match(/担当者.*$/)){
      var listItemQuestion = item.asListItem();
      var choices = [];

      staffList.forEach(function(name){
        if(name != ""){
          choices.push(listItemQuestion.createChoice(name));
        }
      });
      // プルダウンの選択肢を上書きする
      listItemQuestion.setChoices(choices);
    }
  });

}
// プロパティファイルへの設定メソッド
// 新エディタ（デフォルト）ではGUI操作できないため、コードから実行（以前のエディタを使用→GUI操作も可能）
function setScriptProperty() {
  PropertiesService.getScriptProperties().setProperty('SS_ID', '**************'); // スプレッドシートID
  PropertiesService.getScriptProperties().setProperty('FORM_ID', '**************'); // フォームID
}

