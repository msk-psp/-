/** @OnlyCurrentDoc */
function extractMonth() {
  // Get the active spreadsheet and the first sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Get the value from R1C1 (Row 1, Column 1)
  var cellValue = sheet.getRange(1, 1).getValue();

  Logger.log("Cell value is : " + cellValue);

  // Check if the cell contains a valid date
  if (cellValue instanceof Date) {
    // Extract the month (0-based, so add 1)

    var currentDate = cellValue

    currentDate.setDate(currentDate.getDate() + 1);

    var month = currentDate.getMonth() + 1;

    // Log the month
    Logger.log("The month is: " + month);
    
    // Optionally write it back to the sheet (e.g., B1)
    // sheet.getRange(1, 2).setValue(month); // Writes the month to cell B1
    return month
  } else {
    Logger.log("The value in R1C1 is not a valid date.");
    return null
  }
}

function copyValuesToSheet(sourceRange) {
  // Get the active spreadsheet and the active sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculateSheet = ss.getActiveSheet();

  if (calculateSheet.getName() != "식수계산기") {
    SpreadsheetApp.getUi().alert("식수 계산기에서 실행시켜주세요.");
    return;
  }

  var month = extractMonth()

  if (month == null) {
    SpreadsheetApp.getUi().alert("날짜가 올바르지 않습니다.");
    return;
  }

  Logger.log("1")
  var siksooAndCustomerRange = calculateSheet.getRange(1, 2, 35, calculateSheet.getLastColumn());
  var siksooAndCustomerValues = siksooAndCustomerRange.getValues();

  // for(var row = 0; row<siksooValues.length; row++) {
    
  // }

  for(var col = 0; col<siksooAndCustomerValues[0].length; col++) {
    var customerName = siksooAndCustomerValues[0][col] 

    Logger.log("customerName" + String(customerName))

    Logger.log("column " + String(col))

    if (!customerName || typeof customerName !== "string") { continue } // Check if it's non-blank

    var sheetName = customerName + " " + month + "월"

    Logger.log("sheet name" + sheetName)
    // Check if the sheet already exists
    var existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }

    var baseSheet = ss.getSheetByName("간이영수증");

    var copiedSheet = baseSheet.copyTo(ss);

    copiedSheet.setName(sheetName);

    var siksooValues = siksooAndCustomerValues.slice(1, 32).map((row) => row[col]) 

    var dhangha = siksooAndCustomerValues[33][col]
    var isTaxIncluded = String(siksooAndCustomerValues[34][col]).toLocaleLowerCase() == "y"

    var startRow = 14
    var siksooColum = 8
    var numberOfRows = 31


    var targetCustomerNameRange = copiedSheet.getRange(4,5)
    var targetSiksooRange = copiedSheet.getRange(startRow, siksooColum, numberOfRows, 1)
    var targetDhanghaRange = copiedSheet.getRange(startRow, siksooColum + 1, numberOfRows, 1)
    var targetSumFormulaRange = copiedSheet.getRange(45,3)
    
    Logger.log("lastrow", targetSiksooRange.getValues().length)
    Logger.log("lastcolumn", targetSiksooRange.getValues()[0].length)

    targetCustomerNameRange.setValue(customerName)
    targetSiksooRange.setValues(siksooValues.map((value)=>[value]))
    targetDhanghaRange.setValues(new Array(31).fill([dhangha]))
    
    if (isTaxIncluded) {
      targetSumFormulaRange.setValue("=SUM(J14:L44) * 1.1")
    } else {
      targetSumFormulaRange.setValue("=SUM(J14:L44)")
    }

    var sicksooValuesString = siksooValues.map((value) => String(value[0])).reduce((pv, cv) => pv + cv)
  
    Logger.log(sheetName+ " 식수 값", sicksooValuesString)
  }

  SpreadsheetApp.getUi().alert("완료했습니다");
}

