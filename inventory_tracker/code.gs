/** @OnlyCurrentDoc */

// --- CONFIGURATION ---
var ID_VENDORS = "188U_8Catanggeycs_VY2kisIaZl1uUi4KYpOC2qyh8g";
var VALID_SHEETS = ["dhanyam", "varnam", "vastram", "gavya", "soaps","snacks"];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Vidyagrama Inventory Manager")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}

/**
 * TRIGGER: Auto-Serial No for Manual Spreadsheet Entries
 */
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();

  if (VALID_SHEETS.indexOf(sheetName) === -1) return;

  var row = range.getRow();
  var col = range.getColumn();

  if (row > 1 && col > 1) {
    var slNoCell = sheet.getRange(row, 1);
    if (slNoCell.getValue() === "") {
      var lastRow = sheet.getLastRow();
      var slNoValues = sheet.getRange(2, 1, lastRow, 1).getValues();
      var maxNo = 0;
      for (var i = 0; i < slNoValues.length; i++) {
        var val = Number(slNoValues[i][0]);
        if (!isNaN(val) && val > maxNo) maxNo = val;
      }
      slNoCell.setValue(maxNo + 1);
    }
  }
}

function getRecentItems(sheetName) {
  var targetSheet = sheetName || "dhanyam";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheet);
  if (!sheet) return [];

  SpreadsheetApp.flush();

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // SAFE CHECK: Get the actual number of columns available
  var actualCols = sheet.getLastColumn();
  // Ensure we don't try to pull 19 columns if only 17 exist
  var columnsToPull = Math.min(actualCols, 19);

  var data = sheet.getRange(2, 1, lastRow - 1, columnsToPull).getValues();

  return data.map(function (row) {
    return {
      slNo: row[0] ? row[0].toString() : "",
      name: row[2] || "Unnamed Item",
      sku: row[15] || "",
      // Capture Column Q (Index 16) for the image URL
      image_url: row[16] || "",
      updated: (row.length >= 19 && row[18]) ? row[18].toString() : "No Date",
      stock: Number(row[4]) || 0,
      reorder: Number(row[9]) || 0,
      sheetOrigin: targetSheet
    };
  }).filter(item => item.slNo !== "");
}

// Search across ALL defined sheets to find the item
function searchItem(searchText) {
  var cleanSearch = searchText.toString().trim().toLowerCase();

  for (var s = 0; s < VALID_SHEETS.length; s++) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VALID_SHEETS[s]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var skuInSheet = (data[i][15] || "").toString().trim().toLowerCase();

      if (skuInSheet === cleanSearch) {
        var cleanData = data[i].map(function (cellValue) {
          if (cellValue instanceof Date) {
            return Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          return cellValue;
        });
        return { row: i + 1, data: cleanData, sheetName: VALID_SHEETS[s] };
      }
    }
  }
  return null;
}

function checkSkuExists(sku, currentSlNo) {
  if (!sku) return null;
  var cleanSku = sku.toString().trim().toLowerCase();

  for (var s = 0; s < VALID_SHEETS.length; s++) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VALID_SHEETS[s]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var skuInSheet = (data[i][15] || "").toString().trim().toLowerCase();
      var slNoInSheet = data[i][0].toString();

      if (skuInSheet === cleanSku && slNoInSheet !== currentSlNo) {
        return { name: data[i][2], sheet: VALID_SHEETS[s] };
      }
    }
  }
  return null;
}

function processForm(formObject) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var duplicate = checkSkuExists(formObject.sku, formObject.slNo);
    if (duplicate) {
      throw new Error("Duplicate SKU! This SKU is already assigned to '" + duplicate.name + "' in " + duplicate.sheet);
    }

    var sheetName = formObject.mainCategory;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet not found.");

    var rowNumber = formObject.rowNumber;
    var targetRow = rowNumber ? Number(rowNumber) : sheet.getLastRow() + 1;

    var salePriceFormula = "=F" + targetRow + "*(1 + (G" + targetRow + "/100))";
    var stockValueFormula = "=E" + targetRow + "*F" + targetRow;

    // Updated formData structure: 
    // image_url in Column Q (16), delete_url in Column R (17), Timestamp in Column S (18)
    var formData = [
      formObject.slNo || "",
      formObject.category,
      formObject.itemName,
      formObject.uom,
      formObject.stock,
      formObject.purchasePrice,
      formObject.priceMarkupPercentage,
      salePriceFormula,
      stockValueFormula,
      formObject.reorderPoint,
      formObject.moq,
      formObject.vendorID,
      formObject.status,
      formObject.mfgDate,
      formObject.expiryDate,
      formObject.sku,
      formObject.image_url, // Column Q
      formObject.delete_url, // Column R (New)
      new Date()           // Column S
    ];

    if (rowNumber) {
      sheet.getRange(rowNumber, 1, 1, 19).setValues([formData]);
      return "Updated successfully!";
    } else {
      var lastRow = sheet.getLastRow();
      var nextSlNo = 1;
      if (lastRow > 1) {
        var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        var maxNo = Math.max(...values.map(r => isNaN(r[0]) || r[0] === "" ? 0 : Number(r[0])));
        nextSlNo = maxNo + 1;
      }
      formData[0] = nextSlNo;
      sheet.appendRow(formData);
      return "Added successfully!";
    }
  } catch (e) {
    return "Error: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

function getVendorList() {
  try {
    const ss = SpreadsheetApp.openById(ID_VENDORS);
    const sheet = ss.getSheetByName('main');
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    return data.map(row => ({ id: row[0].toString(), name: row[1].toString() })).filter(v => v.id !== "");
  } catch (e) { return []; }
}

function getSheetSummary(sheetName) {
  var targetSheet = sheetName || "dhanyam";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheet);
  if (!sheet) return { totalValue: 0, lowStockCount: 0 };

  var data = sheet.getDataRange().getValues();
  var totalValue = 0;
  var lowStockCount = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === "") continue;

    var stock = Number(data[i][4]) || 0;
    var purchasePrice = Number(data[i][5]) || 0;
    var reorderPoint = Number(data[i][9]) || 0;

    totalValue += (stock * purchasePrice);
    if (stock <= reorderPoint) {
      lowStockCount++;
    }
  }

  return {
    totalValue: totalValue.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
    lowStockCount: lowStockCount
  };
}

function saveBarcodeToDrive(sku, itemName) {
  const FOLDER_ID = '1xRpSS39qScUQp-0U4yPGRktxKyTSJzlW';
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const barcodeUrl = "https://bwipjs-api.metafloor.com/?bcid=code128" +
      "&text=" + encodeURIComponent(sku) +
      "&scale=3&rotate=N&includetext&textsize=10" +
      "&textxalign=center" +
      "&alttext=" + encodeURIComponent("vidyagrama | " + itemName + "\n*" + sku + "*");
    const response = UrlFetchApp.fetch(barcodeUrl);
    const blob = response.getBlob().setName(sku + "_" + itemName.replace(/\s+/g, '_') + ".png");
    const file = folder.createFile(blob);
    return file.getUrl();
  } catch (e) {
    throw new Error("Label Generation Failed: " + e.message);
  }
}

/**
 * GENERATE BULK PDF: Prints all barcodes in the current sheet
 */
function generateBulkBarcodePDF(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet '" + sheetName + "' not found.");

    const data = sheet.getDataRange().getValues();

    // 1. Process items and filter out anything without a SKU
    const items = data.slice(1).map(row => ({
      name: row[2] ? row[2].toString().trim() : "Unknown",
      sku: row[15] ? row[15].toString().trim() : ""
    })).filter(item => item.sku !== "");

    if (items.length === 0) return null;

    const FOLDER_ID = '1xRpSS39qScUQp-0U4yPGRktxKyTSJzlW';
    const folder = DriveApp.getFolderById(FOLDER_ID);

    const tempDoc = DocumentApp.create('Print_Sheet_' + sheetName);
    const body = tempDoc.getBody();
    body.setMarginTop(30).setMarginBottom(30).setMarginLeft(30).setMarginRight(30);

    const table = body.appendTable();
    const columns = 3;
    let currentRow;
    let addedCount = 0; // Track how many we actually add

    items.forEach((item) => {
      // Reconstruct the exact filename format you use: SKU_ItemName.png
      // Replacing spaces with underscores as per your saveBarcodeToDrive function
      const expectedFileName = item.sku + "_" + item.name.replace(/\s+/g, '_') + ".png";
      const files = folder.getFilesByName(expectedFileName);

      // SKIP LOGIC: Only proceed if the file actually exists
      if (files.hasNext()) {
        // Create a new row every 3 items that we ACTUALLY add
        if (addedCount % columns === 0) {
          currentRow = table.appendTableRow();
        }

        const cell = currentRow.appendTableCell();
        const blob = files.next().getBlob();

        // Add Name
        cell.appendParagraph(item.name.substring(0, 25))
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
          .setFontSize(9).setBold(true);

        // Add Barcode Image
        const img = cell.appendImage(blob);
        img.setWidth(140).setHeight(55);

        // Add SKU text
        cell.appendParagraph(item.sku)
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
          .setFontSize(8);

        cell.setPaddingBottom(10).setPaddingTop(10);

        addedCount++;
      }
      // If file doesn't exist, we do nothing (item is skipped)
    });

    // If no barcodes were found at all, don't generate an empty PDF
    if (addedCount === 0) {
      DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
      return null;
    }

    tempDoc.saveAndClose();

    const pdfBlob = tempDoc.getAs('application/pdf');
    const pdfFile = DriveApp.createFile(pdfBlob).setName("Print_" + sheetName + ".pdf");

    // Cleanup
    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return pdfFile.getUrl();

  } catch (e) {
    console.error("PDF Error: " + e.toString());
    throw new Error("Could not generate PDF. Please check folder permissions.");
  }
}

function deleteItemRecord(sheetName, rowNumber) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) throw new Error("Category sheet not found.");
    
    var rowIdx = Number(rowNumber);
    if (rowIdx <= 1) throw new Error("Cannot delete header row.");

    // --- NEW: BARCODE DELETION LOGIC ---
    // 1. Get SKU and Name from the row before deleting it
    // Based on your structure: Name is Col 3 (Index 2), SKU is Col 16 (Index 15)
    var rowData = sheet.getRange(rowIdx, 1, 1, 16).getValues()[0];
    var itemName = rowData[2];
    var sku = rowData[15];

    if (sku) {
      const FOLDER_ID = '1xRpSS39qScUQp-0U4yPGRktxKyTSJzlW';
      const folder = DriveApp.getFolderById(FOLDER_ID);
      
      // Reconstruct the filename to match your saving convention
      const fileNameToDelete = sku + "_" + itemName.toString().replace(/\s+/g, '_') + ".png";
      const files = folder.getFilesByName(fileNameToDelete);
      
      while (files.hasNext()) {
        var file = files.next();
        file.setTrashed(true); // Moves the barcode to Google Drive trash
      }
    }
    // -----------------------------------

    // 2. Delete the actual row from the sheet
    sheet.deleteRow(rowIdx);

    // 3. Re-index the Serial Numbers (Column A)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var range = sheet.getRange(2, 1, lastRow - 1, 1);
      var newSlNos = [];
      for (var i = 1; i <= (lastRow - 1); i++) {
        newSlNos.push([i]);
      }
      range.setValues(newSlNos);
    }

    return "Item and associated barcode deleted successfully.";
  } catch (e) {
    console.error("Delete Error: " + e.toString());
    return "Error: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * TEST FUNCTION: debugGetRecentItems
 * Run this to check if your data is being pulled correctly from the sheet.
 */
function debugGetRecentItems() {
  // 1. Set the sheet you want to test
  var testSheet = "dhanyam";

  try {
    console.log("--- Starting Test for: " + testSheet + " ---");

    var items = getRecentItems(testSheet);

    if (items.length === 0) {
      console.warn("No items found. Check if the sheet exists or if it's empty.");
      return;
    }

    // 2. Log the first item found to check mapping
    var firstItem = items[0];
    console.log("Total Items Found: " + items.length);
    console.log("First Item Details:");
    console.log("- SL No: " + firstItem.slNo);
    console.log("- Name: " + firstItem.name);
    console.log("- SKU: " + firstItem.sku);
    console.log("- Stock: " + firstItem.stock);
    console.log("- Last Updated (Col S): " + firstItem.updated);

    // 3. Range Verification
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(testSheet);
    console.log("Actual Sheet Column Count: " + sheet.getLastColumn());

    if (sheet.getLastColumn() < 19) {
      console.error("CRITICAL: Your sheet only has " + sheet.getLastColumn() +
        " columns. getRecentItems needs 19 (up to Column S) to work!");
    } else {
      console.log("✅ Column count is correct (19 or more).");
    }

    console.log("--- Test Complete ---");

  } catch (e) {
    console.error("Test Failed with Error: " + e.toString());
  }
}