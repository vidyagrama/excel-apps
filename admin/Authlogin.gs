/**
 * Admin Login Logic
 * Checks against the "main" sheet of the current Admin spreadsheet
 */
function checkLogin(userId, mobile) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("main"); 
  const data = sheet.getDataRange().getValues();
  
  const cleanId = String(userId).trim();
  const cleanMob = String(mobile).trim();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == cleanId && data[i][2] == cleanMob) {
      // Record login time
      const timestamp = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy HH:mm:ss");
      sheet.getRange(i + 1, 4).setValue(timestamp);
      
      return { success: true, userName: data[i][1] };
    }
  }
  return { success: false, message: "Invalid Credentials. Please try again." };
}
