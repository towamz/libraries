function setPageMargin(body) {
  // console.log(body.getMarginTop());
  // console.log(body.getMarginBottom());
  // console.log(body.getMarginRight());
  // console.log(body.getMarginLeft());

  body.setMarginTop(0)
  body.setMarginBottom(0)
  body.setMarginRight(0)
  body.setMarginLeft(0)

  // console.log(body.getMarginTop());
  // console.log(body.getMarginBottom());
  // console.log(body.getMarginRight());
  // console.log(body.getMarginLeft());

}

function addTable(body, rowNum, colNum){
  var rowsData = make2DArray(rowNum, colNum);

  var table = body.appendTable(rowsData);

  table.setBorderColor('#CCCCCC');
  table.setBorderWidth(0.5); 

  return table;
}

function setTableLength(table, rowNum, colNum){
  //  91mm× 52mm (91mm× 55mm)
  // 257(207+50)pt x 147pt

  var cellHeight = [207, 50];

  for(var i=0; i<rowNum; i++){
    var row = table.getRow(i);
    row.setMinimumHeight(cellHeight[i%2]);

    for(var j=0; j<colNum; j++){
      if(i>0) break;

      var cell = row.getCell(j);
      cell.setWidth(147);
    }
  }
}
