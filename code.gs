function matchColumnDWithDataHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const reportSheet = ss.getSheetByName('Credit Card Report');
  const dataSheet = ss.getSheetByName('Data');

  const col = 4;       // Column D
  const targetCol = 5; // Column E (where dropdown goes)
  const startRow = 13;
  const lastRow = reportSheet.getLastRow();
  const numRows = lastRow - startRow + 1;

  const columnDValues = reportSheet
    .getRange(startRow, col, numRows)
    .getValues()
    .flat();

  const dataHeaders = dataSheet
    .getRange(1, 1, 1, dataSheet.getLastColumn())
    .getValues()[0];

  const dataLastRow = dataSheet.getLastRow();

  columnDValues.forEach((cellValue, index) => {
    const actualRow = startRow + index;
    if (cellValue === "" || cellValue === null) return;

    dataHeaders.forEach((header, colIndex) => {
      if (cellValue === header) {

        // pull values from header
        const dropdownValues = dataSheet
          .getRange(2, colIndex + 1, dataLastRow - 1)
          .getValues()
          .flat()
          .filter(v => v !== "");

        // build dropdown rule
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(dropdownValues, true)
          .setAllowInvalid(false)
          .build();

        // apply dropdown to Column E on the same row
        reportSheet
          .getRange(actualRow, targetCol)
          .setDataValidation(rule);

        Logger.log(
          `Dropdown added to E${actualRow} from Data column ${colIndex + 1}`
        );
      }
    });
  });
}
