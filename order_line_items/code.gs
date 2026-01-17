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
  const vargas = data.slice(1).map(row => row[1]); 
  return [...new Set(vargas)].filter(v => v && v.toString().trim() !== "").sort();
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

function finalizeOrderBulk(orderSummary, fullCart) {
  try {
    const ssLineItems = SpreadsheetApp.openById(ID_ORDERS_LINE_ITEMS);
    const ssOrders = SpreadsheetApp.openById(ID_ORDERS);
    const lineItemSheet = ssLineItems.getSheetByName(TAB_LINE_ITEMS);
    const summarySheet = ssOrders.getSheetByName(TAB_ORDERS);

    // MAPPING UNIT PRICE: Column H from fullCart (index 7) is passed here
    const lineItemRows = fullCart.map((item, index) => [
      index + 1,            // Column A: Sr No
      orderSummary.orderId, // Column B: Order ID
      item.category,        // Column C: Category
      item.itemId,          // Column D: ID
      item.itemName,        // Column E: Name
      item.quantity,        // Column F: Qty
      item.uom,             // Column G: UOM
      item.salePrice,       // Column H: UNIT PRICE (Added as requested)
      item.subtotal,        // Column I: Subtotal
      ""                    // Column J: Notes
    ]);

    const liStartRow = getFirstEmptyRowInColumn(lineItemSheet, 2);
    const summaryStartRow = getFirstEmptyRowInColumn(summarySheet, 2);

    lineItemSheet.getRange(liStartRow, 1, lineItemRows.length, 10).setValues(lineItemRows);

    const summaryData = [[
      "P0", orderSummary.orderId, orderSummary.customerId, orderSummary.customerName,
      new Date(), "Received", orderSummary.total, "Not Recieved", ""
    ]];
    
    summarySheet.getRange(summaryStartRow, 1, 1, 9).setValues(summaryData);
    SpreadsheetApp.flush(); 
    return true;
  } catch (e) {
    return e.toString();
  }
}

function getFirstEmptyRowInColumn(sheet, col) {
  const range = sheet.getRange(1, col, sheet.getMaxRows()).getValues();
  for (let i = 0; i < range.length; i++) {
    if (range[i][0] === "" || range[i][0] === null) return i + 1;
  }
  return sheet.getLastRow() + 1;
}