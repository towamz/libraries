class CopyColumns{
    // コンストラクタ
    constructor(shIdOrig, shIdDest,ssIdOrig = -1,ssIdDest = -1) {
      let ssOrig;
      let ssDest;
  
      if(ssIdOrig == -1){
        ssOrig = SpreadsheetApp.getActiveSpreadsheet();
      }else{
        ssOrig = SpreadsheetApp.openById(ssIdOrig);
      }
  
      if(ssIdDest == -1){
        ssDest = SpreadsheetApp.getActiveSpreadsheet();
      }else{
        ssDest = SpreadsheetApp.openById(ssIdDest);   
      }
  
      this._shOrig = ssOrig.getSheetById(shIdOrig);
      this._shDest = ssDest.getSheetById(shIdDest);
      this._targetColumns = []; //要素なしの配列を宣言
    }
    
  
  
    // setter/getter
    set targetColumns(columnStr) {
      // 列名の有効性確認(エラーハンドルなし、プログラムを止める)
      console.log(columnStr);
      let testRange = this._shOrig.getRange(columnStr + "1");
      this._targetColumns.push(columnStr);
    }
  
    get targetColumns() {
      return this._targetColumns;
    }
  
  
    // メソッド
    copyColumns(){
      for (var i = 0; i < this._targetColumns.length; i++) {
        let colNumOrig = this._shOrig.getRange(this._targetColumns[i] + '1').getColumn();
        let colData = this._shOrig.getRange(1,colNumOrig,this._shOrig.getMaxRows(),1).getValues();
        this._shDest.getRange(1,i+1,colData.length,1).setValues(colData);
      }
    }
  }