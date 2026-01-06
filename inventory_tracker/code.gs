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
      var cleanData = data[i].map(function(cell) {
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
  var rowNumber = formObject.rowNumber; // Hidden field from HTML
  
  var formData = [
    formObject.itemId || "", // This will be filled by logic below if empty
    formObject.itemName,
    formObject.category,
    formObject.uom,
    formObject.salePrice,
    formObject.purchasePrice,
    formObject.stock,
    formObject.reorderPoint,
    formObject.stockValue,
    formObject.vendorID,
    formObject.status,
    formObject.expiryDate,
    formObject.sku
  ];

  if (rowNumber) {
    // UPDATE EXISTING
    sheet.getRange(rowNumber, 1, 1, 13).setValues([formData]);
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
    formData[0] = nextId; // Assign the new ID
    sheet.appendRow(formData);
    return "New Item " + nextId + " added successfully!";
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