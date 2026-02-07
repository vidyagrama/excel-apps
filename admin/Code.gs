function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Vidyagrama Admin Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkLogin(userId, mobile) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("main"); 
  const data = sheet.getDataRange().getValues();
  
  // Clean inputs to avoid whitespace issues
  const cleanId = String(userId).trim();
  const cleanMob = String(mobile).trim();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == cleanId && data[i][2] == cleanMob) {
      // Record login time: Format: DD/MM/YYYY HH:MM:SS
      const timestamp = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy HH:mm:ss");
      sheet.getRange(i + 1, 4).setValue(timestamp);
      
      return { success: true, userName: data[i][1] };
    }
  }
  return { success: false, message: "Invalid Credentials. Please try again." };
}