
function automateDaily(e) {
  var monitoredSpreadsheetName = 'automate-report-test';
  var monitoredRange = 'E9:E17';
  
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  if (e.source.getName() === monitoredSpreadsheetName) {
    var monitoredCells = sheet.getRange(monitoredRange);
    
    if (range.getRow() >= monitoredCells.getRow() && range.getRow() <= monitoredCells.getLastRow() &&
        range.getColumn() >= monitoredCells.getColumn() && range.getColumn() <= monitoredCells.getLastColumn()) {
      
      var dropdownValue = sheet.getRange('D5').getValue();
      
      if (dropdownValue === 'Daily Report ONLY' || dropdownValue === 'Both') {
        
        var newValue = range.getValue();
        var dailyReportId = '1hmWEh-H1wraA_cLSNqvPr1a85kz2ZAmwSZou0a1LMZs'; //change this
        var dailyReport = SpreadsheetApp.openById(dailyReportId);
        var dailySheet = dailyReport.getSheetByName('Sheet1'); //change this
        
        if (dailySheet) {
          var startCellReference = sheet.getRange('D6').getValue();
          
          var targetRange = getMappedRange(range, monitoredCells, dailySheet, startCellReference);
          targetRange.setValue(newValue);

          targetRange.setWrap(true);
        }
      }
    }
  }
}

function getMappedRange(sourceRange, monitoredCells, targetSheet, startCellReference) {
  var offsetRow = sourceRange.getRow() - monitoredCells.getRow(); // Calculate row offset
  
  var columnLetter = startCellReference.match(/[A-Z]+/)[0];
  var startRow = parseInt(startCellReference.match(/\d+/)[0], 10); // Extract row number (e.g., 12)

  var startColumn = columnLetter.charCodeAt(0) - 'A'.charCodeAt(0) + 1;

  return targetSheet.getRange(startRow + offsetRow, startColumn);
}
