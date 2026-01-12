class CreateEvents {
  constructor(calendarId = null, allDayEvent = false, sheetName = null) {
    // 格納[0][0]がA1(1,1)なのでデータをシートに転記するときは行列が各+1
    this._offsetArySh = 1 
    this._sleepMs = 1000;
    this._calendar = calendarId
      ? CalendarApp.getCalendarById(calendarId)
      : CalendarApp.getDefaultCalendar();
    // シートはsetterで設定する
    this.sheetName = sheetName;
    this._allDayEvent = allDayEvent;
    this._tags = [];

    //各列の既定値を設定する
    this._rowIndexDataFrom = 1 //データ開始行(シート2行目,配列1)
    this._columnIndexEventId = 0 //イベントID列
    this._columnIndexTitle = 1 //タイトル列
    this._columnIndexStartDate = 2
    this._columnIndexStartTime = 2
    this._columnIndexEndDate = 3
    this._columnIndexEndTime = 3
    //任意の引数は-1(設定しない)
    this._color = -1
    this._columnIndexDescription = -1 
    this._columnIndexLocation = -1
  }

  set sheetName(sheetStr) {
    this._sh1 = sheetStr === null
      ? SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
      : SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetStr);

    if (!this._sh1) {
      throw new Error(`指定されたシート "${sheetStr}" が見つかりません。`);
    }
  }
  get sheetName() {
    return this._sh1.getName();
  }

  // イベントのカラー
  set color(colorIndex) {

    if (Object.values(CalendarApp.EventColor).includes(colorIndex)) {
      this._color = colorIndex;
    } else if(colorIndex == -1){
      this._color = -1;
    } else {
      throw new Error("無効な色が指定されました \n An invalid color is specified");
    }
  }
  get color() {
    switch (this._color) {
      case CalendarApp.EventColor.BLUE:
        return 'BLUE';
      case CalendarApp.EventColor.CYAN:
        return 'CYAN';
      case CalendarApp.EventColor.GRAY:
        return 'GRAY';
      case CalendarApp.EventColor.GREEN:
        return 'GREEN';
      case CalendarApp.EventColor.MAUVE:
        return 'MAUVE';
      case CalendarApp.EventColor.ORANGE:
        return 'ORANGE';
      case CalendarApp.EventColor.PALE_BLUE:
        return 'PALE_BLUE';
      case CalendarApp.EventColor.PALE_GREEN:
        return 'PALE_GREEN';
      case CalendarApp.EventColor.PALE_RED:
        return 'PALE_RED';
      case CalendarApp.EventColor.RED:
        return 'RED';
      case CalendarApp.EventColor.YELLOW:
        return 'YELLOW';
      case -1:
        return 'DEFAULT';
      default:
        throw new Error('色の設定が間違っています \ncolor setting failure');
    }
  }

  // データ開始行
  set rowIndexDataFrom(rowNum) {
    if (rowNum<1) {
      throw new Error('行番号は1以上です。\n A row number is 1 or higher');
    }
    // シートの1行目は配列の0になるので-1する
    this._rowIndexDataFrom = rowNum - 1;
  }
  get rowIndexDataFrom() {
    // 配列の0はシートの1行目になるので+1する
    return this._rowIndexDataFrom + 1;
  }

  // イベントID列
  set columnStrEventId(columnStr) {
    this._columnIndexEventId = this.columnToIndex_(columnStr);
  }
  get columnStrEventId() {
    return this.indexToColumn_(this._columnIndexEventId);
  }

  // タイトル列
  set columnStrTitle(columnStr) {
    this._columnIndexTitle = this.columnToIndex_(columnStr);
  }
  get columnStrTitle() {
    return this.indexToColumn_(this._columnIndexTitle);
  }

  // 開始日列
  set columnStrStartDate(columnStr) {
    this._columnIndexStartDate = this.columnToIndex_(columnStr);
  }
  get columnStrStartDate() {
    return this.indexToColumn_(this._columnIndexStartDate);
  }

  // 開始時刻列
  set columnStrStartTime(columnStr) {
    if(this._allDayEvent){
      throw new Error('終日イベントでは時間の指定はできません')
    }
    this._columnIndexStartTime = this.columnToIndex_(columnStr);
  }
  get columnStrStartTime() {
    if(this._allDayEvent){
      throw new Error('終日イベントでは時間の指定はできません')
    }
    return this.indexToColumn_(this._columnIndexStartTime);
  }

  // 終了日列
  set columnStrEndDate(columnStr) {
    if(this._allDayEvent){
      throw new Error('終日イベントでは時間の指定はできません')
    }
    this._columnIndexEndDate = this.columnToIndex_(columnStr);
  }
  get columnStrEndDate() {
    if(this._allDayEvent){
      throw new Error('終日イベントでは時間の指定はできません')
    }
    return this.indexToColumn_(this._columnIndexEndDate);
  }

  // 終了時刻列
  set columnStrEndTime(columnStr) {
    this._columnIndexEndTime = this.columnToIndex_(columnStr);
  }
  get columnStrEndTime() {
    return this.indexToColumn_(this._columnIndexEndTime);
  }

  //　説明列 
  set columnStrDescription(columnStr) {
    this._columnIndexDescription = this.columnToIndex_(columnStr);
  }
  get columnStrDescription() {
    return this.indexToColumn_(this._columnIndexDescription);
  }

  // 場所列
  set columnStrLocation(columnStr) {
    this._columnIndexLocation = this.columnToIndex_(columnStr);
  }
  get columnStrLocation() {
    return this.indexToColumn_(this._columnIndexLocation);
  }

  // 列名をインデックス番号に変換するプライベートメソッド 
  columnToIndex_(column) {
    // 列名が英文字（A〜Zまたはa〜z）のみで構成されている1文字のみをチェック
    if (!/^[a-zA-Z]$/.test(column)) {
      throw new Error('列名はA〜Zの範囲で指定してください。\nThe column letters are in the range from A to Z.');
    }

    // 大文字に変換（例：'a' → 'A'）
    const columnName = column.toUpperCase();

    // A列->0, B列->1
    return columnName.charCodeAt(0) - 'A'.charCodeAt(0); 
  }

  // インデックス番号を列名に変換するプライベートメソッド
  indexToColumn_(index) {
    if (typeof index !== 'number' || !Number.isInteger(index)) {
      throw new Error('インデックスは整数で指定してください。\n The index is specified by an integer.');
    }

    if (index < 0 || index > 25) {
      throw new Error('インデックスは0〜25の範囲で指定してください（A〜Zに対応）。');
    }
    return String.fromCharCode('A'.charCodeAt(0) + index);
  }

  set　tags(tag){
    // tag オブジェクトが name と value の2つのキーのみを持っているかチェック
    const keys = Object.keys(tag); // オブジェクトのキーを配列で取得
    
    // キーが 2 つであることを確認
    if (keys.length !== 2) {
      throw new Error('タグのキーはname,valueの２つである必要があります');
    }

    // tag オブジェクトが "name" と "value" を持っているかを確認
    if (!tag.hasOwnProperty('name') || !tag.hasOwnProperty('value')) {
      throw new Error('タグのキーはname,valueの２つである必要があります');
    }
    
    // 正しい形式であれば、タグを配列に追加
    this._tags.push(tag);
  }
  // 設定されたタグを展開して文字列として返す
  get tags() {
    return this._tags.map(tag => `${tag.name}: ${tag.value}`).join("\n");
  }

  getDateAllDay(date){
    let dateObj = new Date(date);

    if (isNaN(dateObj)) {
      throw new Error('無効な日付です: ');
    }
    dateObj.setHours(0,0,0,0);

    return new Date(dateObj);    
  }

  getDateTime(date, time){
    let dateObj = new Date(date);
    let timeObj = new Date(time);

    if (isNaN(dateObj) || isNaN(timeObj)) {
      throw new Error('無効な日付です: ');
    }

    // 同じ日時の場合は結合しないでそのまま返す
    if(dateObj.getTime()==timeObj.getTime()){
      return dateObj
    }

    let dateTime = new Date()

    dateTime.setFullYear(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
    dateTime.setHours(timeObj.getHours(),timeObj.getMinutes(),0,0);

    return new Date(dateTime);
  }

  getStartEndDateTime(startDate, endDate, startTime, endTime){
    // 終日イベントの時は時分秒を0にして返す
    if(this._allDayEvent){
      let startDateObj = this.getDateAllDay(startDate);

      let endDateObj;
      if (endDate === undefined) {
        endDateObj = new Date(startDateObj);      
      }else{
        endDateObj = this.getDateAllDay(endDate);
      }

      endDateObj.setDate(endDateObj.getDate()+1);

      return [new Date(startDateObj), new Date(endDateObj)]

    // 通常イベントの時は、日と時間を結合して返す
    }else{
      let startDateObj = new Date(this.getDateTime(startDate, startTime));
      let endDateObj = new Date(this.getDateTime(endDate, endTime));

      return [new Date(startDateObj), new Date(endDateObj)]
    }
  }


  // 1行データを取得してデータ編集後のデータを返す
  getEventInfo_(row, index){
    let dataEditedRow = [];

    // eventIdに入力がある場合は登録データなしで返す
    if(row[this._columnIndexEventId]!=''){
      dataEditedRow.push(
                    row[this._columnIndexEventId],
                    '',
                    '',
                    '',
                    '',
                    '');
      return dataEditedRow;
    }

    // eventIdが空白の時はevent生成用データを作成する
    try{
      let title = row[this._columnIndexTitle];

      let startEndDateTime = this.getStartEndDateTime(
                                      row[this._columnIndexStartDate],
                                      row[this._columnIndexEndDate],
                                      row[this._columnIndexStartTime],              
                                      row[this._columnIndexEndTime]);
      let startDateTime = startEndDateTime[0];
      let endDateTime = startEndDateTime[1];

      let description = ""
      if(this._columnIndexDescription != -1){
        description = row[this._columnIndexDescription]
      }

      let location = ""
      if(this._columnIndexLocation != -1){
        location = row[this._columnIndexLocation]
      }          

      dataEditedRow.push(
                        '', 
                        title, 
                        startDateTime, 
                        endDateTime,
                        description,
                        location);

    }catch(e){
        throw new Error(`${e.message},${index}`);
    }

    return dataEditedRow;
  }

  // シートから取得したデータをもとにカレンダー登録用の配列を生成する
  // 0列=evnetID, 空白の場合は登録処理実行する, 登録成功した場合はそのID・失敗した場合は空白のまま 
  getEventsInfo(){
    let dataOrig = this._sh1.getDataRange().getValues();
    let dataReg = [];

    // データ開始行から最終行まで１つずつ処理する
    dataOrig.slice(this._rowIndexDataFrom).forEach((row, rowIndex) => {
      dataReg.push(this.getEventInfo_(row, rowIndex));
    });

    return dataReg
  }

  createEvents(){
    let data = this.getEventsInfo();

    for(let i=0; i<data.length; i++){
      console.log(i,data[i][0]);
      if(data[i][0]!=''){
        console.log("skip->", data[i]);
        continue;
      }else{
        // GASは大量に実行するとエラーとなる場合があるので指定ミリ秒止める
        Utilities.sleep(this._sleepMs);
        console.log("start->",i, data[i]);
        try{
          let event;
          if(this._allDayEvent){  
            event = this._calendar.createAllDayEvent(data[i][1], data[i][2], data[i][3]);
          }else{
            event = this._calendar.createEvent(data[i][1], data[i][2], data[i][3]);
          }

          event.setDescription(data[i][4])
          event.setLocation(data[i][5])
          if(this._color != -1){
            event.setColor(this._color);
          }

          event.removeAllReminders();

          this._tags.forEach(tag => {
            console.log(tag.name, tag.value);
            event.setTag(tag.name, tag.value);
          });

          //今生成したイベントのeventIdを取得する
          data[i][0] = event.getId();

          console.log("->create");
        }catch(e){
          // イベント生成に失敗したらeventIdを空白にする
          data[i][0] = '';
          console.log("->error", e.message);
        }
      }
    }

    // 実行結果を反映する
    // 配列からeventid(0列目)を切り取る
    let dataEventIds = data.map(function(row) {
      return [row[0]]; // 各行の0列目を取り出す
    });

    // eventIdの最初のセル取得(配列用のindexなのでそれぞれ+1する)
    let rangeFirstEventId = this._sh1.getRange(this._rowIndexDataFrom + this._offsetArySh,
                                               this._columnIndexEventId + this._offsetArySh);

    rangeFirstEventId.offset(0, 0, dataEventIds.length, dataEventIds[0].length).
      setValues(dataEventIds);
  }
}
