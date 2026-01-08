/** @OnlyCurrentDoc */
const SHEET_NAME = 'main'; 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle("Vidyagrama Vendor Manager")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. SEARCH: Find Vendor by manual String ID (Case-Insensitive)
function searchVendor(searchText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  
  // Normalize search text
  var searchClean = searchText.toString().toLowerCase().trim();
  
  for (var i = 1; i < data.length; i++) {
    // Exact match comparison using lowercase
    if (data[i][0].toString().toLowerCase().trim() === searchClean) {
      
      var cleanData = data[i].map(function(cell) {
        if (cell instanceof Date) {
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

// 2. CREATE or UPDATE: Case-Insensitive logic
function processForm(formObject) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    
    // Normalize user input
    var manualId = formObject.vendorId.toString().trim();
    var manualIdLower = manualId.toLowerCase();
    
    if (!manualId) return "Error: Vendor ID is required.";

    var rowIndex = -1;
    // Check if this String ID already exists (Case-Insensitive check)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase().trim() === manualIdLower) {
        rowIndex = i + 1;
        break;
      }
    }

    var formData = [
      manualId,                   // Col 1: Manual String ID (Keeps user's casing)
      formObject.businessName,    // Col 2
      formObject.identity,        // Col 3
      formObject.prodCategory,    // Col 4
      formObject.status,          // Col 5
      formObject.contactPerson,   // Col 6
      "'" + formObject.phone,     // Col 7
      formObject.email,           // Col 8
      formObject.address,         // Col 9
      formObject.taxId,           // Col 10
      formObject.moq,             // Col 11
      formObject.leadTime,        // Col 12
      formObject.bankDetails      // Col 13
    ];

    if (rowIndex > -1) {
      // UPDATE EXISTING
      sheet.getRange(rowIndex, 1, 1, 13).setValues([formData]);
      return "Vendor " + manualId + " updated successfully!";
    } else {
      // CREATE NEW
      sheet.appendRow(formData);
      return "New Vendor " + manualId + " onboarded successfully!";
    }
  } catch (e) {
    return "Error: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

function deleteVendorByRow(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sheet.deleteRow(row);
  return "Vendor deleted successfully.";
}