var ID_ADMINS = "1iiZtZclKgr7G7ISZFlM1We4LTmMLNkZLp_x4gP2DoOM";

var TAB_VENDORS_MAIN = "main";
var TAB_ADMINS_VARGA = "varga"
var TAB_ADMINS_ENABLE_CATEGORY = "enable_maincategory";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Vidyagrama Vendor Manager")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}

// 1. SEARCH: Find Vendor by manual String ID (Case-Insensitive)
function searchVendor(searchText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_VENDORS_MAIN);
  var data = sheet.getDataRange().getValues();
  var searchClean = searchText.toString().toLowerCase().trim();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase().trim() === searchClean) {
      
      var cleanData = data[i].map(function (cell) {
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        return cell;
      });

      return {
        row: i + 1,
        data: cleanData // This now contains all 14 columns
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_VENDORS_MAIN);
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

    // UPDATED MAPPING: Now 14 Columns total
    var formData = [
      manualId,                   // Col 1: Vendor ID
      formObject.businessName,    // Col 2: Business Name
      formObject.identity,        // Col 3: Identity
      formObject.mainCategory,    // Col 4: Main Category
      formObject.subCategory,     // Col 5: Sub Category
      formObject.status,          // Col 6: Vendor Status
      formObject.contactPerson,   // Col 7: Contact Person
      "'" + formObject.phone,     // Col 8: Mobile
      formObject.email,           // Col 9: Email
      formObject.address,         // Col 10: Address
      formObject.taxId,           // Col 11: Tax_No
      formObject.moq,             // Col 12: MOQ
      formObject.leadTime,        // Col 13: LeadTime
      formObject.bankDetails      // Col 14: Bank Details
    ];

    if (rowIndex > -1) {
      // UPDATE EXISTING - Note: getRange width updated to 14
      sheet.getRange(rowIndex, 1, 1, 14).setValues([formData]);
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
  try {
    // Ensure you use TAB_VENDORS_MAIN, not SHEET_NAME
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_VENDORS_MAIN);
    sheet.deleteRow(parseInt(row)); 
    return "Vendor record has been permanently removed.";
  } catch (e) {
    throw new Error("Could not delete row: " + e.toString());
  }
}

function getVendorIdentities() {
  try {
    // Replace with your actual Admin Spreadsheet ID
    var adminSs = SpreadsheetApp.openById(ID_ADMINS);
    var vargaSheet = adminSs.getSheetByName(TAB_ADMINS_VARGA);

    // Get data from Column 3 (starting from Row 2 to skip header)
    var lastRow = vargaSheet.getLastRow();
    if (lastRow < 2) return ["No Categories Found"];

    var data = vargaSheet.getRange(2, 3, lastRow - 1, 1).getValues();

    // Flatten the 2D array and remove empty values
    var identities = data.map(r => r[0]).filter(item => item !== "");

    // Optional: Sort alphabetically
    return identities.sort();
  } catch (e) {
    console.log("Error loading identities: " + e.toString());
    return ["Error loading list"];
  }
}

function getCategoryMap() {
  try {
    var adminSs = SpreadsheetApp.openById(ID_ADMINS);
    var sheet = adminSs.getSheetByName(TAB_ADMINS_ENABLE_CATEGORY);
    var data = sheet.getDataRange().getValues();

    var categoryMap = {};

    // Start from Row 2 to skip headers
    for (var i = 1; i < data.length; i++) {
      var mainCat = data[i][0]; // Column 1
      var subCatsRaw = data[i][4]; // Column 5 (index 4)

      if (mainCat) {
        // Split comma-separated string into an array and clean whitespace
        var subCatArray = subCatsRaw ? subCatsRaw.split(',').map(s => s.trim()) : [];
        categoryMap[mainCat] = subCatArray;
      }
    }
    return categoryMap;
  } catch (e) {
    return { "Error": [e.toString()] };
  }
}

/** Get Onboarded Vendors **/
function getOnboardedVendors() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_VENDORS_MAIN);
  var data = sheet.getDataRange().getValues();
  var list = [];
  
  for (var i = 1; i < data.length; i++) {
    list.push({
      id: data[i][0],      // Col 1
      name: data[i][1],    // Col 2
      category: data[i][3] // Col 4
    });
  }
  return list;
}
