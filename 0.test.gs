const testGoals = () => new App().showGoalsPrompt('Effective Communication')
const testSolutions = () => showSolutionsPrompt('Effective Communication')
const testSaveSolutions = () => {
  let data = JSON.parse(DriveApp.getFileById('1QVb9YikeiAVCIMAapXvP28ay4F6O43uo').getBlob().getDataAsString())
  let exCorePillar = "2-exceptional-employee-experience"

  new App().saveSolutions(data, exCorePillar)
}

function unprotectSpecificRanges() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Success Plan"); // Change to your sheet name

  // Protect the entire sheet
  let sheetprotections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  let rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  ([...sheetprotections, ...rangeProtections]).forEach(protection => protection.remove())
  var protection = sheet.protect().setDescription('Sheet Protection').setWarningOnly(true);

  // Define ranges to unprotect
  var unprotectedRanges = [
    sheet.getRange('A1:A10'),
    sheet.getRange('B1:B10') // Add more ranges as needed
  ];

  // Remove protection from specific ranges
  protection.setUnprotectedRanges(unprotectedRanges);
}

function downloadXLS_GUI() {
  var ss = SpreadsheetApp.getActive();
  // var nSheet = SpreadsheetApp.create(sh.getName() + ": copy");

  // var d = sh.getDataRange();
  // nSheet.getSheets()[0].getRange(1, 1, d.getLastRow(), d.getLastColumn()).setValues(d.getValues());

  var URL = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="' + URL + '">Click to download</a>')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(80)
    .setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download XLS');
}

function testPDF() {
  _getAsBlob(SpreadsheetApp.getActive().getSheetByName('Success Plan').getRange('A:L'))
}

function _getAsBlob(range, isProtrait = true) {
  sheet = range.getSheet()
  let url = sheet.getParent().getUrl()

  var rangeParam = ''
  var sheetParam = ''
  if (range) {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  if (sheet) {
    sheetParam = '&gid=' + sheet.getSheetId()
  }
  // A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  // these parameters are reverse-engineered (not officially documented by Google)
  // they may break overtime.
  var exportUrl = url.replace(/\/edit.*$/, '')
    + '/export?exportFormat=pdf&format=pdf'
    + '&size=LETTER'
    + `&portrait=${isProtrait}`
    // + '&fitw=true'
    + '&top_margin=0.75'
    + '&bottom_margin=0.75'
    + '&left_margin=0.7'
    + '&right_margin=0.7'
    + '&sheetnames=false&printtitle=false'
    + '&pagenum=CENTER' // change it to CENTER to print page numbers
    + '&gridlines=false'
    // + '&fzr=true'
    // + '&scale=4'
    // + '&fith=true'
    // + '&fitw=true'
    + sheetParam
  // + rangeParam

  console.log(exportUrl)
  // Logger.log('exportUrl=' + exportUrl)
  var response
  var i = 0
  for (; i < 5; i += 1) {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      },
    })
    if (response.getResponseCode() === 429) {
      // printing too fast, retrying
      Utilities.sleep(3000)
    } else {
      break
    }
  }

  if (i === 5) {
    throw new Error('Printing failed. Too many sheets to print.')
  }

  return response.getBlob()
}
