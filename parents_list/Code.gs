var ID_PARENTS = "1xgcQfWYczXmkwpQsbonkRUraAMvlWExNRtm7D_iSJbk";
var ID_ADMINS = "1iiZtZclKgr7G7ISZFlM1We4LTmMLNkZLp_x4gP2DoOM";

var TAB_PARENTS_MAIN = "main";
var TAB_ADMINS_ACTIVITY_LOGS = "activitiy_logs";
var TAB_ADMINS_VARGA = "varga";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Vidyagrama Registration")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}

function processForm(formObject) {
  var lock = LockService.getScriptLock();
  try {
    // Wait for up to 30 seconds for other processes to finish
    lock.waitLock(30000);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_PARENTS_MAIN);

    // 1. Find the last row that actually has data in the ID column (Column A)
    var idColumnValues = sheet.getRange("A:A").getValues();
    var maxId = 0;

    // Loop backwards to find the highest ID number currently in use
    for (var i = idColumnValues.length - 1; i >= 0; i--) {
      var val = idColumnValues[i][0];
      if (val !== "" && !isNaN(val)) {
        maxId = Number(val);
        break;
      }
    }

    var nextId = maxId + 1;

    // CLEAN MOBILE NUMBER: Remove +91 and any whitespace
    var cleanMobile = formObject.mobile ? formObject.mobile.replace('+91', '').trim() : "";

    // TYPE CONVERSIONS: Enforce matching types cleanly
    var numericDiscount = formObject.discount ? Number(formObject.discount) : 0;
    var numericCredit = formObject.credit ? Number(formObject.credit) : 0;
    var numericBalance = formObject.balance ? Number(formObject.balance) : 0;

    // Must match exact column ordering: id, varga, name, father, mother, mobile, email, discount, credit, balance, Notes
    var rowData = [
      nextId,
      formObject.varga || "",
      formObject.name || "",
      formObject.father || "",
      formObject.mother || "",
      cleanMobile,
      formObject.email || "",
      numericDiscount, 
      numericCredit,   
      numericBalance,  
      formObject.notes || ""
    ];
    
    // Append row cleanly into your Data Table
    sheet.appendRow(rowData);

    // REMOVED: The updateVargaDropdown logic here was causing the crash.
    // The Data Table will automatically inherit the "Varga" dropdown configuration setup in Column B.

    return "Success!"; // This will now successfully trigger the web app success handler!
  } catch (e) {
    return "Error: " + e.toString();
  } finally {
    lock.releaseLock(); // Always release the lock
  }
}


// 2. Logic for Manual Entries (onEdit)
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "main") return;

  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();

  var vargaCol = 2; // Column B 

  if (row > 1) {
    if (col > 1) {
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

    if (col !== vargaCol) {
      const vargaColCell = sheet.getRange(row, vargaCol);
      if (!vargaColCell.getDataValidation()) {
        updateVargaDropdown(vargaColCell);
      }
    }
  }
}

/**Update Varga from admins sheet, so it avoids manually adding varga in future*/

/**
 * Updates the Varga dropdown in the Parents_List spreadsheet.
 * Fetches data from Admin SS > 'varga' sheet.
 */
function updateVargaDropdown(cell, forceRefresh = false) {
  cell.clearDataValidations();

  const cache = CacheService.getScriptCache();
  const cacheKey = "all_vargas_list";

  let vargaString = forceRefresh ? null : cache.get(cacheKey);

  if (vargaString === null) {
    const adminSS = SpreadsheetApp.openById(ID_ADMINS);
    const vargaSheet = adminSS.getSheetByName(TAB_ADMINS_VARGA);
    const data = vargaSheet.getDataRange().getValues();

    // Extract, filter out empty rows, and sort alphabetically
    const vargas = data.slice(1)
      .map(row => row[0])
      .filter(String)
      .sort(function (a, b) {
        return a.localeCompare(b, undefined, { sensitivity: 'base' });
      });

    vargaString = vargas.join(',');
    cache.put(cacheKey, vargaString, 1500);
    console.log("Varga list refreshed from Admin and sorted alphabetically.");
  }

  if (vargaString) {
    const options = vargaString.split(',');
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .setAllowInvalid(true) 
      .build();

    cell.setDataValidation(rule);
  } else {
    cell.clearDataValidations();
  }
}

// Add this function to return the Varga list as an array and force to load fresh data based on forcerefresh flag status
function getVargaList(forceRefresh = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "all_vargas_list";

  // If forceRefresh is true, we ignore the cache and fetch fresh
  let vargaString = forceRefresh ? null : cache.get(cacheKey);

  if (!vargaString) {
    try {
      const adminSS = SpreadsheetApp.openById(ID_ADMINS);
      const vargaSheet = adminSS.getSheetByName(TAB_ADMINS_VARGA);
      const data = vargaSheet.getDataRange().getValues();
      
      // Extract, filter out empty rows, and sort alphabetically
      const vargas = data.slice(1)
        .map(r => r[0])
        .filter(String)
        .sort(function (a, b) {
          return a.localeCompare(b, undefined, { sensitivity: 'base' });
        });

      vargaString = vargas.join(',');
      cache.put(cacheKey, vargaString, 1500);
    } catch (err) {
      return ["Error loading list"];
    }
  }
  return vargaString.split(',');
}
