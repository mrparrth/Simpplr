function UntitledMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D6:J11').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('D7:J11').activate();
  spreadsheet.getActiveRangeList().setBorder(null, null, null, null, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  // .setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID)
  // .setBorder(null, null, null, null, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange('D13:J17').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('D14:J17').activate();
  spreadsheet.getActiveRangeList().setBorder(null, null, null, null, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('L14').activate();
};