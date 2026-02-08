/** @OnlyCurrentDoc */

// --- CONFIGURATION ---
var ID_VENDORS = "188U_8Catanggeycs_VY2kisIaZl1uUi4KYpOC2qyh8g";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Vidyagrama Inventory Manager")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}

function getRecentItems() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return []; 
  var data = sheet.getRange(2, 1, lastRow - 1, 18).getValues(); 

  return data.map(function(row) {
    return {
      id: row[0] ? row[0].toString() : "",      
      name: row[2] || "Unnamed Item",    
      sku: row[15] || "",    
      updated: row[17] ? row[17].toString() : "" 
    };
  }).filter(item => item.id !== "").reverse(); 
}

// 1. SEARCH: Find item by ID (Col 1) or SKU (Col 16/Index 15)
function searchItem(searchText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
  var data = sheet.getDataRange().getValues();
  
  // Clean the incoming search text for mobile keyboard compatibility
  var cleanSearch = searchText.toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var idInSheet = data[i][0].toString().trim().toLowerCase();
    
    // Column P is index 15 (SKU)
    var skuValue = data[i][15] || "";
    var skuInSheet = skuValue.toString().trim().toLowerCase();
    if (idInSheet === cleanSearch || skuInSheet === cleanSearch) {
      var cleanData = data[i].map(function (cellValue) {
        if (cellValue instanceof Date) {
          return Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        return cellValue;
      });
      return { row: i + 1, data: cleanData };
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
    var timestamp = new Date(); 

    var formData = [
      formObject.itemID || "",
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
      formObject.imageUrl,
      timestamp            
    ];

    if (rowNumber) {
      sheet.getRange(rowNumber, 1, 1, 18).setValues([formData]);
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

/** @OnlyCurrentDoc */


// --- EXISTING FUNCTIONS (doGet, getRecentItems, searchItem, processForm) REMAIN SAME ---

/**
 * Fetches vendor data from the external 'vendors_list' spreadsheet.
 * Mapping: Column A (0) = vendorID, Column B (1) = buissnessName
 */
function getVendorList() {
  try {

    const ss = SpreadsheetApp.openById(ID_VENDORS);
    const sheet = ss.getSheetByName('main');

    if (!sheet) {
      console.error("Sheet 'main' not found.");
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    // Fetch columns A and B
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    console.log("Raw Data from Sheet:", data); // Check this in Apps Script 'Executions'

    const vendors = data.map(function(row) {
      return {
        id: row[0] ? row[0].toString().trim() : "",
        name: row[1] ? row[1].toString().trim() : ""
      };
    }).filter(v => v.id !== ""); // Only require an ID to count as a vendor

    console.log("Filtered Vendors:", vendors);
    return vendors;
    
  } catch (e) {
    console.error("Critical Vendor Load Error: " + e.toString());
    return []; 
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