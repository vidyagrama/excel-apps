/** @OnlyCurrentDoc */

// --- CONFIGURATION ---
var ID_VENDORS = "188U_8Catanggeycs_VY2kisIaZl1uUi4KYpOC2qyh8g";
var ID_ADMINS = "1iiZtZclKgr7G7ISZFlM1We4LTmMLNkZLp_x4gP2DoOM";
var ID_INVENTORY = "1YDiJsrkNEj4HxDaNlirGIczAX4h7FExpb3XNs9Xu5co";
var ID_BARCODES = "1xRpSS39qScUQp-0U4yPGRktxKyTSJzlW";
var ID_BARCODES_PDF = "1DMNF_rgQNLUPTc1P2_kb8Dy4bWUIdLsT";

var TAB_ENABLE_CATEGORY = "enable_maincategory";

var VALID_SHEETS = ["Shridhanya", "Varnam", "Vastram", "GauAmruth", "Tejas", "Madhuram"];
var Default_Sheet = "Shridhanya";

function doGet() {

  var template = HtmlService.createTemplateFromFile('Index');

  // 2. Evaluate the template to execute <?!= include('Styles'); ?>
  return template.evaluate()
    .setTitle("Vidyagrama  Inventory Manager")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}

/*This requires if you like seperate Styles,Scripts to sepearte html as template loading */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * TRIGGER: Auto-Serial No for Manual Spreadsheet Entries
 */
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();

  // 1. Safety Check: Only run on valid inventory sheets
  if (VALID_SHEETS.indexOf(sheetName) === -1) return;

  var row = range.getRow();
  var col = range.getColumn();
  if (row <= 1) return; // Skip header row

  // --- PART A: AUTO-INCREMENT SL NO ---
  if (col > 1) {
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

  // --- PART B: DYNAMIC DROPDOWN (Column B) ---
  const subCatCol = 2; // Column B

  // If you edit any column OTHER than the subcategory itself,
  // we ensure the dropdown is present in Column B for that row.
  if (col !== subCatCol) {
    const subCatCell = sheet.getRange(row, subCatCol);

    // Check if validation already exists to prevent redundant slow calls
    if (!subCatCell.getDataValidation()) {
      // Pass the sheetName (e.g., "Vastram") to fetch the right list
      updateSubCategoryDropdown(sheetName, subCatCell,false);
    }
  }

}

/**
 * Helper to fetch mapping and apply validation
 */
/**
 * Helper to fetch mapping and apply validation
 * Added forceRefresh parameter to bypass cache when updates occur
 */
function updateSubCategoryDropdown(mainCat, cell, forceRefresh = false) {
  if (!mainCat) {
    cell.clearDataValidations();
    return;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = "subcats_" + mainCat.toLowerCase().replace(/\s+/g, '_');
  
  // 1. If forceRefresh is true, we ignore the cache and set subCatString to null
  let subCatString = forceRefresh ? null : cache.get(cacheKey);

  // 2. Fetch from Admin if not cached OR if we are forcing a refresh
  if (subCatString === null) {
    const adminSS = SpreadsheetApp.openById(ID_ADMINS);
    const adminSheet = adminSS.getSheetByName(TAB_ENABLE_CATEGORY);
    const adminData = adminSheet.getDataRange().getValues();

    // Loop through Admin data and update ALL category caches at once
    for (let i = 1; i < adminData.length; i++) {
      let catName = String(adminData[i][0]);
      let catSubs = String(adminData[i][4] || ""); // Column E

      // Update the cache for every category found
      cache.put("subcats_" + catName.toLowerCase().replace(/\s+/g, '_'), catSubs, 1500);
      
      if (catName.toLowerCase() === mainCat.toLowerCase()) {
        subCatString = catSubs;
      }
    }
    console.log("Subcategory cache refreshed from Admin Sheet for: " + mainCat);
  }

  // 3. Apply validation
  if (subCatString) {
    const options = subCatString.split(',').map(item => item.trim()).filter(String);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .setAllowInvalid(false)
      .build();

    cell.setDataValidation(rule);
  } else {
    cell.clearDataValidations();
  }
}

/**
 * Fetches the full subcategory map with an optional cache bypass
 * @param {boolean} forceRefresh - If true, ignores cache and fetches from Admin Sheet
 */
function getSubCategoryMap(forceRefresh = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "full_subcategory_map";
  
  // 1. Check cache only if we aren't forcing a refresh
  let cachedMap = forceRefresh ? null : cache.get(cacheKey);

  if (cachedMap) {
    console.log("Sidebar: Loading map from cache");
    return JSON.parse(cachedMap);
  }

  // 2. Fetch from Admin Sheet if cache is empty or forced
  console.log("Sidebar: Fetching fresh map from Admin Sheet");
  const adminSS = SpreadsheetApp.openById(ID_ADMINS);
  const adminSheet = adminSS.getSheetByName(TAB_ENABLE_CATEGORY);
  const adminData = adminSheet.getDataRange().getValues();

  let map = {};

  for (let i = 1; i < adminData.length; i++) {
    let catName = String(adminData[i][0]).trim();
    let catSubs = String(adminData[i][4] || ""); // Column E

    if (catName) {
      map[catName] = catSubs.split(',')
        .map(s => s.trim())
        .filter(String);
    }
  }

  // 3. Update the cache with the fresh map
  cache.put(cacheKey, JSON.stringify(map), 1500);

  return map;
}

/*Fetch Recent item list */
function getRecentItems(sheetName) {
  var targetSheet = sheetName || Default_Sheet;
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

/* Check SKU exists in sheets */
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

/*Main Function to Add Data to Google sheets */
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
      // APPLY DROPDOWN VALIDATION TO COLUMN B (Index 2)
      updateSubCategoryDropdown(sheetName, sheet.getRange(rowNumber, 2));

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

      // APPLY DROPDOWN VALIDATION TO COLUMN B OF THE NEWLY APPENDED ROW
      var newRowIndex = sheet.getLastRow();
      updateSubCategoryDropdown(sheetName, sheet.getRange(newRowIndex, 2),false);

      return "Added successfully!";
    }
  } catch (e) {
    return "Error: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/* Get Vendors list from Vendors google sheet */
function getVendorList() {
  try {
    const ss = SpreadsheetApp.openById(ID_VENDORS);
    const sheet = ss.getSheetByName('main');
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    return data.map(row => ({ id: row[0].toString(), name: row[1].toString() })).filter(v => v.id !== "");
  } catch (e) { return []; }
}

/*We are not using this function, its used by Admin portal */
function getSheetSummary(sheetName) {
  var targetSheet = sheetName || Default_Sheet;
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

/*This helps to generate SKU for the next new item based on main category */
function getNextSku(category) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(category);
  if (!sheet) return "VG??-007";

  // 1. Define Short Codes
  const shortCodes = {
    'Shridhanya': 'SD',
    'Varnam': 'VN',
    'Vastram': 'VS',
    'GauAmruth': 'GA',
    'Tejas': 'TJ',
    'Madhuram': 'MD'
  };

  const prefix = "VG" + (shortCodes[category] || category.substring(0, 2).toUpperCase()) + "-";
  const lastRow = sheet.getLastRow();

  // 2. Default if sheet is empty (starting at 007)
  if (lastRow < 2) return prefix + "007";

  // 3. Get all SKUs from Column P (Index 15)
  const skuValues = sheet.getRange(2, 16, lastRow - 1, 1).getValues().flat();

  let maxNum = 6; // Start below 007 so the first increment hits 007

  skuValues.forEach(sku => {
    if (sku && typeof sku === 'string' && sku.includes('-')) {
      const parts = sku.split('-');
      const num = parseInt(parts[parts.length - 1], 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  });

  // 4. Increment and Pad with zeros (e.g., 008, 012, 105)
  const nextNum = (maxNum + 1).toString().padStart(3, '0');
  return prefix + nextNum;
}

/* Save Barcode to backend google drive */
/**
 * Generates a clean, rectangular barcode without text and with extra white padding.
 * @param {string} sku The product SKU.
 * @param {string} itemName The name of the item.
 * @return {string} The URL of the saved file.
 */
function saveBarcodeToDrive(sku, itemName) {
  try {
    const folder = DriveApp.getFolderById(ID_BARCODES);
    
    // Updated BWIP-JS Parameters:
    // bcid=code128      : Standard industrial barcode
    // scale=4           : High resolution
    // height=12         : Reduced height relative to width for a sleek rectangular look
    // paddingwidth=20   : Adds significant white space on the left and right
    // paddingheight=10  : Adds white space on top and bottom
    // backgroundcolor=ffffff : Solid white background
    // (Notice: 'includetext' is REMOVED to keep it pure barcode)
    
    const barcodeUrl = "https://bwipjs-api.metafloor.com/?bcid=code128" +
                       "&text=" + encodeURIComponent(sku) +
                       "&scale=4" +
                       "&height=12" + 
                       "&paddingwidth=2" +
                       "&paddingheight=3" +
                       "&backgroundcolor=ffffff";

    const response = UrlFetchApp.fetch(barcodeUrl);
    
    // Clean file naming
   // const safeItemName = itemName.replace(/[^a-z0-9]/gi, '_');
    const fileName = `${sku}.png`;
    
    const blob = response.getBlob().setName(fileName);
    const file = folder.createFile(blob);
    
    return file.getUrl();
    
  } catch (e) {
    throw new Error("Barcode Generation Failed: " + e.message);
  }
}

/**
 * GENERATE BULK PDF: Prints all barcodes in the current sheet
 */
function generateBulkBarcodePDF(ids, sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet '" + sheetName + "' not found.");

    const data = sheet.getDataRange().getValues();
    const idSet = new Set(ids.map(id => id.toString()));

    // 1. Filter items based on the SELECTED IDs from the UI
    const items = data.slice(1).map(row => ({
      slNo: row[0] ? row[0].toString() : "",
      name: row[2] ? row[2].toString().trim() : "Unknown",
      sku: row[15] ? row[15].toString().trim() : ""
    })).filter(item => idSet.has(item.slNo) && item.sku !== "");

    if (items.length === 0) return "Error: No barcodes found for selected items.";
    
    const barCodefolder = DriveApp.getFolderById(ID_BARCODES);

    const tempDoc = DocumentApp.create('Print_Sheet_' + sheetName);
    const body = tempDoc.getBody();
    body.setMarginTop(30).setMarginBottom(30).setMarginLeft(30).setMarginRight(30);

    const table = body.appendTable();
    const columns = 2; // CHANGED: 2 columns for better mobile scanning size
    let currentRow;
    let addedCount = 0;

    items.forEach((item) => {
      const expectedFileName = item.sku + ".png";
      const files = barCodefolder.getFilesByName(expectedFileName);

      if (files.hasNext()) {
        if (addedCount % columns === 0) {
          currentRow = table.appendTableRow();
        }

        const cell = currentRow.appendTableCell();
        const blob = files.next().getBlob();

        // 1. TOP: Item Name (Centered, Bold)
        // Allowing for longer names with a slightly smaller font if needed
        cell.appendParagraph(item.name)
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
          .setFontSize(10)
          .setBold(true)
          .setSpacingAfter(2); // Tight spacing to the barcode

        // 2. MIDDLE: Barcode Image (Centered)
        const imgPara = cell.appendParagraph(""); 
        const img = imgPara.appendInlineImage(blob);
        imgPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        
        // Sized for scannability
        img.setWidth(220).setHeight(80); 

        // 3. BOTTOM: Vidyagrama - SKU (Centered)
        cell.appendParagraph("Vidyagrama - " + item.sku)
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
          .setFontSize(9)
          .setBold(false)
          .setSpacingBefore(2); // Tight spacing from the barcode

        // Cell Styling for a clean box look
        cell.setPaddingBottom(10).setPaddingTop(10);
        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        
        addedCount++;
      }
    });

    if (addedCount === 0) {
      DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
      return "Error: No matching barcode images found in Drive.";
    }

    // Fill last row if it has only one item to keep borders clean
    if (addedCount % columns !== 0) {
      currentRow.appendTableCell().setText("");
    }

    // 1. Get the PDF blob from the temp document
    tempDoc.saveAndClose();
    const pdfBlob = tempDoc.getAs('application/pdf');
    const pdfName = "Print_Barcodes_" + sheetName + "_" + new Date().toLocaleDateString() + ".pdf";

    // 2. Access the specific PDF folder
    const pdfFolder = DriveApp.getFolderById(ID_BARCODES_PDF);

    // 3. Create the file DIRECTLY in that folder
    const pdfFile = pdfFolder.createFile(pdfBlob).setName(pdfName);

    // 4. Cleanup: Move the temporary Google Doc to trash immediately
    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);

    // 5. Set permissions for the user to view/print
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Return the URL so the UI can open it in a new tab
    return pdfFile.getUrl();

  } catch (e) {
    console.error("PDF Generation Error: " + e.toString());
    return "PDF Error: " + e.message;
  }
}

/*Delete selected item with Bardcode */
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
      const folder = DriveApp.getFolderById(ID_BARCODES);

      // Reconstruct the filename to match your saving convention
      const fileNameToDelete = sku + ".png";
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
/********************************************* Bulk Operation functions ********************************************************** */
/**
 * Deletes multiple items based on their slNo (Column A).
 * @param {Array} ids Array of serial numbers to delete.
 * @param {string} sheetName The main category (e.g., "Vastram").
 */
function processBulkDelete(ids, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Sheet not found.";

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const values = sheet.getDataRange().getValues();
    const idSet = new Set(ids.map(id => id.toString().trim()));
    const folder = DriveApp.getFolderById(ID_BARCODES);
    
    let deletedCount = 0;

    // Start from the bottom, stop at the first row after header (index 1)
    for (let i = values.length - 1; i >= 1; i--) {
      let currentId = values[i][0].toString().trim();
      
      if (idSet.has(currentId)) {
        // Barcode Deletion
        var itemName = values[i][2];
        var sku = values[i][15];

        if (sku && itemName) {
          var fileNameToDelete = sku + ".png";
          var files = folder.getFilesByName(fileNameToDelete);
          while (files.hasNext()) {
            files.next().setTrashed(true);
          }
        }

        // Inside your processBulkDelete loop for the last row:
        if (sheet.getLastRow() === 2) {
          // Clear the entire row so the filter ignores it on next refresh
          sheet.getRange(2, 1, 1, sheet.getLastColumn()).clearContent();
        } else {
           // Delete Row
           sheet.deleteRow(i + 1);
        }
        deletedCount++;
      }
    }
    // Safety: Only re-index if there is at least one data row remaining
    var lastRowAfterDelete = sheet.getLastRow();
    if (deletedCount > 0 && lastRowAfterDelete > 1) {
      var range = sheet.getRange(2, 1, lastRowAfterDelete - 1, 1);
      var newSlNos = [];
      for (var j = 1; j <= (lastRowAfterDelete - 1); j++) {
        newSlNos.push([j]);
      }
      range.setValues(newSlNos);
    }

    return "Successfully deleted " + deletedCount + " items.";
  } catch (e) {
    // If the error persists, it may be due to row protections
    return "Error: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fetches current inventory for the bulk table based on the processForm structure.
 * Mapping:
 * Index 0 (Col A): slNo (ID)
 * Index 1 (Col B): category (Sub-Category)
 * Index 2 (Col C): itemName
 * Index 3 (Col D): uom
 * Index 4 (Col E): stock
 * Index 15 (Col P): sku
 */
function getBulkInventoryData(sheetName) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    console.error("Sheet not found: " + sheetName);
    return [];
  }

  SpreadsheetApp.flush();

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // SAFE CHECK: Get the actual number of columns available
  var actualCols = sheet.getLastColumn();
  // Ensure we don't try to pull 19 columns if only 17 exist
  var columnsToPull = Math.min(actualCols, 19);

  var data = sheet.getRange(2, 1, lastRow - 1, columnsToPull).getValues();

  return data.map(function (r) {
    return {
      id: r[0],           // slNo (Internal ID for deletion/updates)
    subCategory: r[1],  // Column B: Sub-Category
    name: r[2],         // Column C: Item Name
    uom: r[3],          // Column D: Unit of Measure
    stock: r[4],        // Column E: Stock quantity
    sku: r[15]          // Column P: SKU
    };
  }).filter(item => item.slNo !== "");
}

/********************************************* Test/Debug functions ********************************************************** */
/**
 * TEST FUNCTION: debugGetRecentItems
 * Run this to check if your data is being pulled correctly from the sheet.
 */
function debugGetRecentItems() {
  // 1. Set the sheet you want to test
  var testSheet = Default_Sheet;

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

/**
 * RUN THIS TO TEST: This will generate a barcode image and save it to your folder.
 * You can then check the folder to see if the layout is correct.
 */
function debugBarcodeDesign() {
  var testSku = "VGDH-002";
  var testItemName = "Kodo Millet Idli Rava"; // Will be used for bottom-left

  console.log("Starting barcode test for: " + testSku);

  try {
    var resultUrl = saveBarcodeToDrive(testSku, testItemName);
    console.log("✅ Success! Barcode generated.");
    console.log("View it here: " + resultUrl);
  } catch (e) {
    console.error("❌ Test Failed: " + e.message);
  }
}

// test subcategory dropdown
function testSubCategoryUpdate() {

  clearAllSubCategoryCaches();

  const testCategory = "Vastram"; // <--- Change to one of your real categories

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(testCategory);
  const testCell = sheet.getRange("B9"); // <--- Change to your subcategory column

  console.log("Starting test for category: " + testCategory);

  try {
    updateSubCategoryDropdown(testCategory, testCell,true);

    // Verification
    const validation = testCell.getDataValidation();
    if (validation) {
      console.log("TEST PASSED: Data validation is now present in " + testCell.getA1Notation());
    } else {
      console.warn("TEST FAILED: No validation found. Check if category exists in Admin Sheet.");
    }
  } catch (e) {
    console.error("TEST CRASHED: " + e.message);
  }
}