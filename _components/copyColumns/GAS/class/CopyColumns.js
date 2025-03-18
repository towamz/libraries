class CopyColumns{
  /**
   * 指定された列をデータ用シートから結果用シートへコピーします
   * 
   * 
   * @param {string|number} shIdNameOrig - データ用シート名またはシートID
   * @param {string} shNameDest - 結果用シート名(シート新規追加)(既定値:'結果')
   * @param {number} ssIdOrig - データ用スプレッドシートID(既定値:このスプレッドシート)
   * @param {number} ssIdDest - 結果用スプレッドシートID(既定値:このスプレッドシート)
   */
  constructor(shIdNameOrig, shNameDest = '結果',ssIdOrig = -1,ssIdDest = -1) {
    this._shOrig = this._getSheet(ssIdOrig, shIdNameOrig);
    this._shDest = this._createSheet(ssIdDest, shNameDest);

    this._targetColumns = []; //要素なしの配列を宣言
  }


  // setter/getter
  set targetColumns(columnStr) {
    // 列名の有効性確認
    // 列記号(アルファベット)から列番号に変換して格納する
    try {
      let colNum = this._shOrig.getRange(columnStr + "1").getColumn();
      this._targetColumns.push(colNum);
    } catch (error) {
      console.log('無効な列名です: ' + columnStr);
    }
  }

  get targetColumns() {
    // 列番号を列記号(アルファベット)に変換して返す
    let targetColumnsStr = this._targetColumns.map((columnNum) => {
      return this._shOrig.getRange(1, columnNum).getA1Notation().replace(/[0-9]/g, '');
    });
    return targetColumnsStr;
  }


  // メソッド
  _getSheet(ssId, shIdName) {
    let ss = ssId == -1 ? SpreadsheetApp.getActiveSpreadsheet() : SpreadsheetApp.openById(ssId);
    let sh = typeof shIdName  === 'number' ? ss.getSheetById(shIdName) : ss.getSheetByName(shIdName);

    if (!sh) throw new Error('指定されたデータシート「' + shIdName + '」が存在しません。');
    
    return sh;
  }

  _createSheet(ssId, shIdName) {
    let ss = ssId == -1 ? SpreadsheetApp.getActiveSpreadsheet() : SpreadsheetApp.openById(ssId);
    let sh = typeof shIdName === 'number' ? ss.getSheetById(shIdName) : ss.getSheetByName(shIdName);

    if (sh) {
      let dateTimeString = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyMMdd-HHmmss");
      return ss.insertSheet(shIdName + "-" + dateTimeString);  // 名前にタイムスタンプを追加
    } else {
      return ss.insertSheet(shIdName);  // 新しいシートをそのまま作成
    }
  }

  copyColumns(){
    let dataOrig = this._shOrig.getDataRange().getValues();
    // 配列に格納されるとA列(列番号1)が要素0になるので、(各列番号-1)する
    let dataResult = dataOrig.map(row => this._targetColumns.map(columnNum => row[columnNum - 1]));

    this._shDest.getRange(1,1,dataResult.length,dataResult[0].length).setValues(dataResult);
  }
}
