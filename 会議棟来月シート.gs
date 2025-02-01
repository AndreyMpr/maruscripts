function duplicateActiveSheet() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCopy = activeSpreadsheet.duplicateActiveSheet();

  // Get the current sheet's name to determine the current month.
  const currentSheetName = activeSpreadsheet.getActiveSheet().getName();
  const currentSheetDateMatch = currentSheetName.match(/(\d{4})年(\d{2})月/);

  let currentYear, currentMonth;
  if (currentSheetDateMatch) {
    currentYear = parseInt(currentSheetDateMatch[1]);
    currentMonth = parseInt(currentSheetDateMatch[2]);
  } else {
    // If the sheet name doesn't match the pattern, use the current date as a fallback.
    const today = new Date();
    currentYear = today.getFullYear();
    currentMonth = today.getMonth() + 1; // getMonth() is 0-indexed
  }

  // Calculate the next month and year
  let nextMonth = currentMonth + 1;
  let nextYear = currentYear;
  if (nextMonth > 12) {
    nextMonth = 1;
    nextYear++;
  }

  // Format the next month and year into the desired "yyyy年mm月" format.
  const formattedNextMonth = (nextMonth < 10 ? "0" : "") + nextMonth;
  const nextMonthName = `${nextYear}年${formattedNextMonth}月`;

  // Rename the duplicated sheet
  sheetCopy.setName(nextMonthName);

  // Set the date in cell A1 of the new sheet to the 1st of the new month
  const newDate = new Date(nextYear, nextMonth - 1, 2); // Month is 0-indexed
  sheetCopy.getRange("A1").setValue(newDate).setNumberFormat("yyyy\"年\"mm\"月\"");

  // Return the newly created sheet
  return sheetCopy;
}

function clearNewSheet(sheetToClear) {
  const rangeToClear = sheetToClear.getRange("D2:Z99");
  rangeToClear.clearContent();
  rangeToClear.setBackground(null);
  Logger.log(`Range D2:Z99 cleared in sheet "${sheetToClear.getName()}".`);
}

// Add menu element for the script.

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('まるやま')
    .addItem('来月の表作成', 'main')
    .addToUi();
}

function main() {
  const newSheet = duplicateActiveSheet(); // Call duplicateActiveSheet and store the returned sheet
  clearNewSheet(newSheet); // Pass the new sheet to clearNewSheet

}
