
function setDataValid_(range, sourceRange) {
  //需加過濾null
  if(range != null && sourceRange != null){
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange, true).build();
    range.setDataValidation(rule);
  }
}
function onEdit(){
  var aSheet = SpreadsheetApp.getActiveSheet(); //使用表格
  var aSheetName = aSheet.getName();  //表格名稱
  var aCell = aSheet.getActiveCell(); //選取中欄位
  var aColumn = aCell.getColumn();    //選取中欄位為第幾欄
  
  //過濾可套用表格
  if(aSheetName == "開銷表" || aSheetName == "基金" || aSheetName == "花旗"){
    var range = aSheet.getRange(aCell.getRow(), aColumn + 1);
    var sourceRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(aCell.getValue());
    setDataValid_(range, sourceRange);
  }
}