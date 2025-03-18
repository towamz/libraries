function main(){
  makeCardsMany(2);
}


function makeCards(pageNum = 1) {
  const sheetsId = '';
  const folderId = '';
  var sheets = SpreadsheetApp.openById(sheetsId);

  getPicturesIdList(sheets, folderId);

  const docID = "";
  var doc = DocumentApp.openById(docID);
  var body = doc.getBody();
  
  var rowNum = 6;
  var colNum = 4;

  setPageMargin(body);

  for(var i=0; i<pageNum; i++){
    table = addTable(body, rowNum, colNum);
    setTableLength(table, rowNum, colNum);

    insertContents(table, sheets, folderId, rowNum, colNum);

    body.appendPageBreak();
    body.appendPageBreak();

  }
}
