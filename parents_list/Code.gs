function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function processForm(formObject) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main');
  
  // Appends a new row with the Name and Email from the form
  sheet.appendRow([
    formObject.id,
    formObject.varga,
    formObject.name,
    formObject.father,
    formObject.mother,
    formObject.mobile,
    formObject.email,
    formObject.discount,
    formObject.notes
  ]);
  
  return "Success!"; 
}