/** @OnlyCurrentDoc */

function doGet() {
  // .addMetaTag is essential for mobile responsiveness
  // .setFaviconUrl adds a professional touch when saved as a mobile bookmark
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Vidyagrama Inventory Manager")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}

/** * NEW FUNCTION: Fetches the last 10 items for the sidebar list
 */
function getRecentItems() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return []; // Return empty if only header exists

  // Get the last 10 rows (or fewer if the sheet is small)
  var numItems = Math.min(10, lastRow - 1);
  var startRow = lastRow - numItems + 1;
  var data = sheet.getRange(startRow, 1, numItems, 15).getValues();

  // Map to a clean object for the HTML list, reversed so newest is on top
  return data.map(function(row) {
    return {
      id: row[0],   // Column A
      name: row[2], // Column C
      sku: row[14]  // Column O
    };
  }); 
}

// 1. SEARCH: Find item by ID (Col 1) or SKU (Col 15/Index 14)
function searchItem(searchText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
  var data = sheet.getDataRange().getValues();
  
  // Clean the incoming search text for mobile keyboard compatibility
  var cleanSearch = searchText.toString().trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var idInSheet = data[i][0].toString().trim().toLowerCase();
    
    // Column O is index 14 (SKU)
    var skuValue = data[i][14] || "";
    var skuInSheet = skuValue.toString().trim().toLowerCase();

    if (idInSheet === cleanSearch || skuInSheet === cleanSearch) {
      // Convert Dates so mobile HTML5 date inputs can read them (yyyy-MM-dd)
      var cleanData = data[i].map(function (cellValue) {
        if (cellValue instanceof Date) {
          return Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        return cellValue;
      });

      return {
        row: i + 1,
        data: cleanData
      };
    }
  }
  return null;
}

// 2. CREATE or UPDATE: Decision logic with LockService for multi-user mobile safety
function processForm(formObject) {
  var lock = LockService.getScriptLock();
  try {
    // Mobile connections can be flaky; 15 seconds wait is safer than 10
    lock.waitLock(15000); 
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
      // UPDATE: Matches the 15 columns in your HTML form
      sheet.getRange(rowNumber, 1, 1, 15).setValues([formData]);
      return "Item " + formObject.itemName + " updated successfully!";
    } else {
      // CREATE NEW: Auto-ID logic
      var lastRow = sheet.getLastRow();
      var nextId = 1001;
      if (lastRow > 1) {
        var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        var maxId = Math.max(...idValues.map(function(r) { 
          return isNaN(r[0]) || r[0] === "" ? 0 : Number(r[0]); 
        }));
        if (maxId >= 1001) nextId = maxId + 1;
      }
      formData[0] = nextId;
      sheet.appendRow(formData);
      return "New Item " + formObject.itemName + " added successfully!";
    }
  } catch (e) {
    return "Error: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

// 3. Auto-ID for Manual Spreadsheet Entries (Optional for Mobile App, but good for Sheet)
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
      if (lastRow <= 1) { idCell.setValue(1001); return; }
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