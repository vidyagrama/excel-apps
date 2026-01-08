function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle("Vidyagrama Registration")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Logic for Form Submissions
function processForm(formObject) {
  var lock = LockService.getScriptLock();
  try {
    // Wait for up to 30 seconds for other processes to finish
    lock.waitLock(30000); 
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
  
  // Get all IDs from Column A to find the highest number
    var lastRow = sheet.getLastRow();
  var nextId = 1; // Default for first entry

    if (lastRow > 1) {
      var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var maxId = Math.max(...idValues.map(r => isNaN(r[0]) ? 0 : Number(r[0])));
      nextId = maxId + 1;
    }

    sheet.appendRow([
      nextId,
      formObject.varga,
      formObject.name,
      formObject.father,
      formObject.mother,
    "'" + formObject.mobile, // Added ' to keep +91 formatting
      formObject.email,
      formObject.discount,
      formObject.notes
    ]);

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

  // If we edited any column other than ID (Col 1) and Row 1 (Header)
  // AND the ID cell is currently empty
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