/** @OnlyCurrentDoc */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Vidyagrama Inventory Manager")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. SEARCH: Find item by ID (Col 1) or SKU (Col 13)
function searchItem(searchText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    // Check Column A (ID) or Column M (SKU)
    if (data[i][0].toString().trim() == searchText.toString().trim() ||
      data[i][12].toString().trim() == searchText.toString().trim()) {

      // CLEAN THE DATA: Convert Dates to Strings so they don't crash the return
      var cleanData = data[i].map(function (cell) {
        if (cell instanceof Date) {
          // Converts date to YYYY-MM-DD format for the HTML input
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        return cell;
      });

      return {
        row: i + 1,
        data: cleanData
      };
    }
  }
  return null;
}

// 2. CREATE or UPDATE: Decision logic
function processForm(formObject) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
    var rowNumber = formObject.rowNumber;

    // We need to know the target row for the formula logic
    var targetRow = rowNumber ? Number(rowNumber) : sheet.getLastRow() + 1;

    // Automated Formulas (ensure these match your Column letters H and I)
    var salePriceFormula = "=F" + targetRow + "*(1 + (G" + targetRow + "/100))";
    var stockValueFormula = "=E" + targetRow + "*F" + targetRow;

    var formData = [
      formObject.itemId || "",
      formObject.category,
      formObject.itemName,
      formObject.uom,
      formObject.stock,
      formObject.purchasePrice,
      formObject.priceMarkupPercentage,
      salePriceFormula,   // Column H
      stockValueFormula,  // Column I
      formObject.reorderPoint,
      formObject.vendorID,
      formObject.status,
      formObject.mfgDate,
      formObject.expiryDate,
      formObject.sku
    ];

    if (rowNumber) {
      // UPDATE: Changed 13 to 15 here to match your new column count
      sheet.getRange(rowNumber, 1, 1, 15).setValues([formData]);
      return "Item " + formObject.itemId + " updated successfully!";
    } else {
      // CREATE NEW
      var lastRow = sheet.getLastRow();
      var nextId = 1001;
      if (lastRow > 1) {
        var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        var maxId = Math.max(...idValues.map(r => isNaN(r[0]) ? 0 : Number(r[0])));
        if (maxId >= 1001) nextId = maxId + 1;
      }
      formData[0] = nextId;
      sheet.appendRow(formData);
      return "New Item " + nextId + " added successfully!";
    }
  } finally {
    lock.releaseLock();
  }
}

// 3. Auto-ID for Manual Spreadsheet Entries
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() !== "main") return;
  var row = range.getRow();
  var col = range.getColumn();
  if (row > 1 && col > 1) {
    var idCell = sheet.getRange(row, 1);
    if (idCell.getValue() === "") {
      var lastRow = sheet.getLastRow();
      var idValues = sheet.getRange(2, 1, lastRow, 1).getValues();
      var maxId = 1000;
      for (var i = 0; i < idValues.length; i++) {
        var val = Number(idValues[i][0]);
        if (!isNaN(val) && val > maxId) maxId = val;
      }
      idCell.setValue(maxId + 1);
    }
  }
}