function main(e) {
  const monitoredSpreadsheetName = "automate-report-test";
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  const dropdownValue = sheet.getRange("D5").getValue();

  if (e.source.getName() === monitoredSpreadsheetName) {
    if (["Daily Report ONLY"].includes(dropdownValue)) {
      automateDaily(
        "E9:E17",
        "1hmWEh-H1wraA_cLSNqvPr1a85kz2ZAmwSZou0a1LMZs", // change this
        "Sheet1", // change this
        range,
        sheet
      );
    }
  }
}

function automateDaily(
  monitoredDescriptions,
  dailyReportId,
  dailySheetName,
  range,
  sheet
) {
  const monitoredCells = sheet.getRange(monitoredDescriptions);
  if (
    range.getRow() >= monitoredCells.getRow() &&
    range.getRow() <= monitoredCells.getLastRow() &&
    range.getColumn() >= monitoredCells.getColumn() &&
    range.getColumn() <= monitoredCells.getLastColumn()
  ) {
    const newValue = range.getValue();
    const dailyReport = SpreadsheetApp.openById(dailyReportId);
    const dailySheet = dailyReport.getSheetByName(dailySheetName);

    if (dailySheet) {
      const startCellReference = sheet.getRange("D6").getValue();
      const targetRange = getMappedDaily(
        range,
        monitoredCells,
        dailySheet,
        startCellReference
      );
      targetRange.setValue(newValue).setWrap(true);
    }
  }
}

function getMappedDaily(
  sourceRange,
  monitoredCells,
  targetSheet,
  startCellReference
) {
  const offsetRow = sourceRange.getRow() - monitoredCells.getRow(); // Calculate row offset
  const columnLetter = startCellReference.match(/[A-Z]+/)[0];
  const startRow = parseInt(startCellReference.match(/\d+/)[0], 10); // Extract row number
  const startColumn = columnLetter.charCodeAt(0) - "A".charCodeAt(0) + 1;

  return targetSheet.getRange(startRow + offsetRow, startColumn);
}
