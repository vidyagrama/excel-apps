var ID_PARENTS = "1xgcQfWYczXmkwpQsbonkRUraAMvlWExNRtm7D_iSJbk";
var ID_ADMINS = "1iiZtZclKgr7G7ISZFlM1We4LTmMLNkZLp_x4gP2DoOM";


var TAB_PARENTS_MAIN = "main";
var TAB_ADMINS_ACTIVITY_LOGS = "activitiy_logs";
var TAB_ADMINS_VARGA = "varga"

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Vidyagrama Registration")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}

// 1. Logic for Form Submissions
function processForm(formObject) {
  var lock = LockService.getScriptLock();
  try {
    // Wait for up to 30 seconds for other processes to finish
    lock.waitLock(30000);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_PARENTS_MAIN);

    // 1. Find the last row that actually has data in the ID column (Column A)
    var idColumnValues = sheet.getRange("A:A").getValues();
    var lastDataRow = 0;
    var maxId = 0;

    // Loop backwards to find the last row with a number
    for (var i = idColumnValues.length - 1; i >= 0; i--) {
      var val = idColumnValues[i][0];
      if (val !== "" && !isNaN(val)) {
        lastDataRow = i + 1; // Index is 0-based, rows are 1-based
        maxId = Number(val);
        break;
      }
    }

    var nextId = maxId + 1;
    var targetRow = lastDataRow + 1;

    // CLEAN MOBILE NUMBER: Remove +91 and any whitespace
    var cleanMobile = formObject.mobile.replace('+91', '').trim();

    var rowData = [
      nextId,
      formObject.varga,
      formObject.name,
      formObject.father,
      formObject.mother,
      cleanMobile,
      formObject.email,
      formObject.discount,
      formObject.credit,
      formObject.balance,
      formObject.notes
    ];
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

    var vargaCol = 2; // Column B 
    const vargaColCell = sheet.getRange(targetRow, vargaCol); 
    if (!vargaColCell.getDataValidation()) {
        updateVargaDropdown(vargaColCell);
     }

    sheet.appendRow();

    return "Success!";
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

  // ADJUST THIS: Set to the actual column index of 'varga' in your Parents_List
  var vargaCol = 2; // Column B 

  if (row > 1) {
    // 1. YOUR EXISTING AUTO-ID LOGIC
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

    // we ensure the dropdown is present in Column B for that row.
    // Varga dropdown on edit on spreadsheet manually
    if (col !== vargaCol) {
      const vargaColCell = sheet.getRange(row, vargaCol);

      // Check if validation already exists to prevent redundant slow calls
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
    // ID_ADMINS should be defined globally in your code.gs
    const adminSS = SpreadsheetApp.openById(ID_ADMINS);
    const vargaSheet = adminSS.getSheetByName(TAB_ADMINS_VARGA);
    const data = vargaSheet.getDataRange().getValues();

    // Extract first column (Varga), skip header, filter out empty rows
    const vargas = data.slice(1)
      .map(row => row[0])
      .filter(String);

    vargaString = vargas.join(',');

    // Cache for 25 minutes (1500 seconds)
    cache.put(cacheKey, vargaString, 1500);
    console.log("Varga list refreshed from Admin.");
  }

  if (vargaString) {
    const options = vargaString.split(',');
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .setAllowInvalid(false)
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
      // Skip header, take first column, filter empty rows
      const vargas = data.slice(1).map(r => r[0]).filter(String);
      vargaString = vargas.join(',');
      cache.put(cacheKey, vargaString, 1500);
    } catch (err) {
      return ["Error loading list"];
    }
  }
  return vargaString.split(',');
}