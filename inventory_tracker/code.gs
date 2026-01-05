function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle("Vidyagrama Inventory Manager")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processForm(formObject) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
  var data = sheet.getDataRange().getValues();
  
  // 1. Search for existing item by SKU (Column M / Index 12)
  var rowIndex = -1;
  var searchSku = formObject.sku;
  
  if (searchSku) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][12] == searchSku) { // Check Column M
        rowIndex = i + 1;
        break;
      }
    }
  }

  // 2. Logic for Item ID (Column A)
  var itemId = formObject.itemId;
  if (!itemId && rowIndex === -1) {
    // Generate next ID if it's a brand new entry
    var ids = data.slice(1).map(r => isNaN(r[0]) ? 0 : Number(r[0]));
    itemId = ids.length > 0 ? Math.max(...ids) + 1 : 1;
  } else if (rowIndex !== -1) {
    // Keep existing ID if updating
    itemId = data[rowIndex - 1][0];
  }

  // 3. Prepare the row data based on your specific columns
  // Order: Item ID, Item name, Category, UOM, Sale Price, Purchase Price, Stock, 
  // Reorder Point, Stock value (calc), VendorID, Status (calc), Expiry Date, SKU
  var stockVal = Number(formObject.salePrice) * Number(formObject.stock);
  var status = Number(formObject.stock) <= Number(formObject.reorderPoint) ? "Low Stock" : "In Stock";

  // Prepare the row data based on your specific columns
  // Note: We leave Column I (index 8) blank or as null so the ArrayFormula handles it
  var rowData = [
    itemId,                 // Col A: Item ID
    formObject.itemName,    // Col B: Item name
    formObject.category,    // Col C: Category
    formObject.uom,         // Col D: UOM
    formObject.salePrice,   // Col E: Sale Price
    formObject.purchasePrice, // Col F: Purchase Price
    formObject.stock,       // Col G: Stock
    formObject.reorderPoint, // Col H: Reorder Point
    "",                     // Col I: Stock Value (Leave empty for ArrayFormula)
    formObject.vendorID,    // Col J: VendorID
    "",                     // Col K: Status (You can also use an ArrayFormula for this!)
    formObject.expiryDate,  // Col L: Expiry Date
    formObject.sku          // Col M: SKU/Barcode
  ];

  if (rowIndex > -1) {
    // UPDATE existing row
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    return "Inventory Updated!";
  } else {
    // ADD new row
    sheet.appendRow(rowData);
    return "New Item Added!";
  }
}

// Keep your manual entry ID logic - it's great for direct sheet edits!
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "main") return;
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();

  if (row > 1 && col > 1) {
    var idCell = sheet.getRange(row, 1);
    if (idCell.getValue() === "") {
      var lastRow = sheet.getLastRow();
      var idValues = sheet.getRange(2, 1, lastRow, 1).getValues();
      var maxId = 0;
      for (var i = 0; i < idValues.length; i++) {
        var val = Number(idValues[i][0]);
        if (!isNaN(val) && val > maxId) maxId = val;
      }
      idCell.setValue(maxId + 1);
    }
  }
}