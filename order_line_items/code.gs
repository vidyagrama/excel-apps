// --- CONFIGURATION ---
var ID_PARENTS = "1xgcQfWYczXmkwpQsbonkRUraAMvlWExNRtm7D_iSJbk";
var ID_INVENTORY = "1YDiJsrkNEj4HxDaNlirGIczAX4h7FExpb3XNs9Xu5co";
var ID_ORDERS_LINE_ITEMS = "1j5ma5hH1vKaoNW0O3JrYL19FZvPLBXMOyN5_0efP0e8";
var ID_ORDERS = "1i3XQ7tfoKKb6RH8CjyP0fryMnbuOthbXnb26-FCa0MU";

var TAB_PARENTS = "main";
var TAB_INVENTORY = "main";
var TAB_LINE_ITEMS = "main";
var TAB_ORDERS = "main";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Grocery Shop")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}

function getVargas() {
  const ss = SpreadsheetApp.openById(ID_PARENTS);
  const data = ss.getSheetByName(TAB_PARENTS).getDataRange().getValues();
  return [...new Set(data.slice(1).map(row => row[1]))].filter(v => v).sort();
}

function getNamesByVarga(varga) {
  const ss = SpreadsheetApp.openById(ID_PARENTS);
  const data = ss.getSheetByName(TAB_PARENTS).getDataRange().getValues();
  return data.filter(row => row[1] === varga).map(row => row[2]);
}

function validateLogin(varga, name, mobile) {
  const ss = SpreadsheetApp.openById(ID_PARENTS);
  const data = ss.getSheetByName(TAB_PARENTS).getDataRange().getValues();
  const user = data.find(row => row[1] === varga && row[2] === name && String(row[5]).trim() === String(mobile).trim());
  return user ? { success: true, email: user[6], discount: user[7] || 0, name: user[2], id: user[0] } : { success: false };
}

function getInventoryData() {
  const ss = SpreadsheetApp.openById(ID_INVENTORY);
  const data = ss.getSheetByName(TAB_INVENTORY).getDataRange().getValues().slice(1);
  return data.map(row => ({
    itemId: row[0], category: row[1], itemName: row[2], uom: row[3], salePrice: row[7],
    imageUrl: row[15] || "https://via.placeholder.com/150" 
  }));
}

function finalizeOrderBulk(summary, fullCart) {
  try {
    const liSheet = SpreadsheetApp.openById(ID_ORDERS_LINE_ITEMS).getSheetByName(TAB_LINE_ITEMS);
    const ordSheet = SpreadsheetApp.openById(ID_ORDERS).getSheetByName(TAB_ORDERS);

    // Ensure all 10 columns are mapped correctly for Line Items
    const lineRows = fullCart.map((item, index) => [
      index + 1,            // Column A: Sr No
      summary.orderId,      // Column B: Order ID
      item.category,        // Column C: Category
      item.itemId,          // Column D: Item ID
      item.itemName,        // Column E: Name
      item.quantity,        // Column F: Qty
      item.uom,             // Column G: UOM (FIXED)
      item.salePrice,       // Column H: Unit Price
      item.fullSubtotal,    // Column I: Subtotal
      ""                    // Column J: Notes
    ]);

    liSheet.getRange(getFirstEmptyRowInColumn(liSheet, 2), 1, lineRows.length, 10).setValues(lineRows);

    // Ensure 9 columns are mapped for Orders Summary
    const ordRow = [[
      "P0",                 // Column A: Priority
      summary.orderId,      // Column B: Order ID
      summary.customerId,   // Column C: Customer ID
      summary.customerName, // Column D: Name
      new Date(),           // Column E: Date
      "Received",           // Column F: Status
      summary.finalTotal,   // Column G: Total
      "Not Recieved",       // Column H: Payment
      summary.notes         // Column I: Notes
    ]];
    
    ordSheet.getRange(getFirstEmptyRowInColumn(ordSheet, 2), 1, 1, 9).setValues(ordRow);
    SpreadsheetApp.flush(); 
    return true;
  } catch (e) { return e.toString(); }
}

function getFirstEmptyRowInColumn(sheet, col) {
  const range = sheet.getRange(1, col, sheet.getMaxRows()).getValues();
  for (let i = 0; i < range.length; i++) { 
    if (range[i][0] === "" || range[i][0] === null) return i + 1; 
  }
  return sheet.getLastRow() + 1;
}