/** @OnlyCurrentDoc */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle("Vidyagrama Inventory Manager")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Logic for Inventory Form Submissions
function processForm(formObject) {
  // Ensure the sheet name matches exactly
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
  if (!sheet) return "Error: Sheet 'main' not found!";
  
  // Logic to generate the next Item ID
  var lastRow = sheet.getLastRow();
  var nextId = 1001; // Starting ID for inventory items
  
  if (lastRow > 1) {
    var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var maxId = Math.max(...idValues.map(r => isNaN(r[0]) ? 0 : Number(r[0])));
    if (maxId >= 1001) nextId = maxId + 1;
  }

  // Append data mapping to your 13 columns
  sheet.appendRow([
    nextId,                 // itemId
    formObject.itemName,    // itemName
    formObject.category,    // category
    formObject.uom,         // uom
    formObject.salePrice,   // salePrice
    formObject.purchasePrice,// purchasePrice
    formObject.stock,       // stock
    formObject.reorderPoint,// reorderPoint
    formObject.stockValue,  // stockValue (Recommended: calculation or form value)
    formObject.vendorID,    // vendorID
    formObject.status,      // status
    formObject.expiryDate,  // expiryDate
    formObject.sku          // sku
  ]);
  
  return "Item " + nextId + " added successfully!"; 
}

// 2. Auto-ID for Manual Entries
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  
  if (sheet.getName() !== "main") return;

  var row = range.getRow();
  var col = range.getColumn();

  // If user edits any column except ID (1) and it's not the header
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