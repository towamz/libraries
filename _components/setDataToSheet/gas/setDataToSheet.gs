// sh:書き込み対象シート
// data:書き込みデータ(2次元配列)
// dataType: 0=上書き, 1=追記
function setDataToSheet(sh, data, dataType){
    // データチェック(sh)
    if (typeof sh.getName !== 'function' || typeof sh.getRange !== 'function') {
      console.log('データが不正です');
      return
    }
  
    // データチェック(data)
    if (!Array.isArray(data)) {
      console.log('データが不正です');
      return
    }
    if (data.length === 0) {
      console.log('データが不正です');
      return
    }
    const aryLength = data[0].length;
  
    for(i=1; i<data.length; i++){
      if(data[i].length != aryLength){
      console.log('データが不正です');
      return
      }
    }
  
    switch (dataType) {
      case 0:
        // 現在の登録データは削除
        sh.clear();
        sh.getRange(1, 1, data.length, data[0].length).setValues(data);
        break;
      case 1:
        // データがタイトル行しかないときは処理しない
        if(data.length==1){
          return
        }
        // データからタイトル行を削除
        data.shift();
        // 現在あるデータの末尾の+1行目からデータを貼り付ける
        sh.getRange(sh.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
        break;
      default:
        console.log('データが不正です');
        return
    }
  }
