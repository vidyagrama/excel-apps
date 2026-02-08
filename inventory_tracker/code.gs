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

// Dynamically fetch items based on the selected sheet
function getRecentItems(sheetName) {
  var targetSheet = sheetName || "dhanyam";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheet);
  if (!sheet) return [];
  
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

// Search across ALL defined sheets to find the item
function searchItem(searchText) {
  var sheets = ["dhanyam", "varnam", "vastram", "gavya", "soaps"];
  var cleanSearch = searchText.toString().trim().toLowerCase();
  
  for (var s = 0; s < sheets.length; s++) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheets[s]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var idInSheet = data[i][0].toString().trim().toLowerCase();
      var skuInSheet = (data[i][15] || "").toString().trim().toLowerCase();
      
      if (idInSheet === cleanSearch || skuInSheet === cleanSearch) {
        var cleanData = data[i].map(function (cellValue) {
          if (cellValue instanceof Date) {
            return Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          return cellValue;
        });
        return { row: i + 1, data: cleanData, sheetName: sheets[s] };
      }
    }
  }
  return null;
}

function processForm(formObject) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    
    // USES SELECTED CATEGORY AS SHEET NAME
    var sheetName = formObject.mainCategory; 
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet '" + sheetName + "' not found.");

    var rowNumber = formObject.rowNumber;
    var targetRow = rowNumber ? Number(rowNumber) : sheet.getLastRow() + 1;

    var salePriceFormula = "=F" + targetRow + "*(1 + (G" + targetRow + "/100))";
    var stockValueFormula = "=E" + targetRow + "*F" + targetRow;
    var timestamp = new Date(); 

    var formData = [
      formObject.itemID || "",
      formObject.category, // Sub-category
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
      return "Updated in " + sheetName + " successfully!";
    } else {
      var lastRow = sheet.getLastRow();
      var nextId = 1001;
      if (lastRow > 1) {
        var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        var maxId = Math.max(...idValues.map(r => isNaN(r[0]) || r[0] === "" ? 0 : Number(r[0])));
        if (maxId >= 1001) nextId = maxId + 1;
      }
      formData[0] = nextId;
      sheet.appendRow(formData);
      return "Added to " + sheetName + " successfully!";
    }
  } catch (e) {
    return "Error: " + e.toString();
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

  / 3. Auto-ID for Manual Spreadsheet Entries (Optional for Mobile App, but good for Sheet)
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
}