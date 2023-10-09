function count_difference_time() {
  
  //let pasteSheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Result"); // active sheet can not find when the sheet is closed
  let pasteSheet1 = SpreadsheetApp.openById("1Jusj7GoGdqmSdenPagZXAdxmhqsCfz_QdPzOekJN_rc").getSheetByName("Result"); // must use the ID of the sheet
  // 
  //var date = new Date();
  //var currentD = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM");
  //console.log(date);

  if (pasteSheet1.getRange("J1").getValue()!="Assigned date"){
    
    pasteSheet1.insertColumns(10,3);
    pasteSheet1.getRange(1,10).setValue("Assigned date");
    pasteSheet1.getRange(1,11).setValue("Completed date");
    pasteSheet1.getRange(1,12).setValue("Difference date");
    pasteSheet1.getRange('J:L').setBackground("#ffe599");

    pasteSheet1.getRange(2,10).setFormula('=IF(len(H2)>0,H2,DATE(2022,10,30))');
    pasteSheet1.getRange(2,11).setFormula('=IF(len(I2)>0,I2,NOW())');
    pasteSheet1.getRange(2,12).setFormula('=IF(K2-J2>0,K2-J2,0)');
    var destination = pasteSheet1.getRange(2,10,pasteSheet1.getLastRow()-1,3);
    pasteSheet1.getRange(2,10,1,3).autoFill(destination,SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    pasteSheet1.getRange('L:L').setNumberFormat('#,##0.00');

  }
  
  
  // min = pasteSheet1.getRange(8,11).getValue() - pasteSheet1.getRange(8,10).getValue();
  // pasteSheet1.getRange(2,11).setValue(min);
}

