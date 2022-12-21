// Yahoo商品検索APIを利用するためのClient ID
const appid= 'YahooAPIのClient ID';

function fetchInventoryIndex(janCode) {
  const url = 'https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch?appid=' + appid + '&jan_code=' + janCode;
  const res = UrlFetchApp.fetch(url,{muteHttpExceptions: true});

  const json = JSON.parse(res.getContentText());
  if (!json.totalResultsReturned === 0) return {"name":"", "image":{"medium":""}, "brand": {"name":""}};
  return json.hits[0];
}

// Created atに日本時間で「2022/12/24 09:15:30」の形式で登録日時を返す
function getNowDate(){
  let d = new Date();
  return Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
}

// シートが更新された時に呼び出される関数として定義
function onChangeSheet(e) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Inventory');
  
  sheet.getDataRange().getValues().forEach((row, i) => {
    const janCode = row[1], name = row[2];
    if (!janCode || name) return;
    const s = fetchInventoryIndex(janCode);
    sheet.getRange(i + 1, 1, 1, row.length).setValues([[i, janCode, s.name, s.image.medium, s.brand.name, getNowDate()]]); 
  }); 
}
