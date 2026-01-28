/** @OnlyCurrentDoc */

// --- CONFIGURATION ---
var ID_INVENTORY = "1YDiJsrkNEj4HxDaNlirGIczAX4h7FExpb3XNs9Xu5co";
var ID_PO_ORDERS_LINE_ITEMS = "1xB4hkA3W8AScB7MIU0eBVCCi8cuRfQVcE8aIhwq8J6Y";
var ID_PO_ORDERS = "1ryofXlAQ0REZ65sWTIwQ_V-nj7RnfCWBfsQtjoO_bvI";
var ID_VENDORS = "188U_8Catanggeycs_VY2kisIaZl1uUi4KYpOC2qyh8g";

var TAB_VENDORS = "main"; 
var TAB_INVENTORY = "main";
var TAB_LINE_ITEMS = "main"; 
var TAB_ORDERS = "main";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Vidyagrama Procurement")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// FETCH VENDORS
function getVendors() {
  const ss = SpreadsheetApp.openById(ID_VENDORS);
  const data = ss.getSheetByName(TAB_VENDORS).getDataRange().getValues().slice(1);
  return data.map(row => ({ vendorID: row[0], businessName: row[1], poc: row[7] }));
}

// FETCH STOCK ALERTS
function getStockAlerts() {
  const ss = SpreadsheetApp.openById(ID_INVENTORY);
  const data = ss.getSheetByName(TAB_INVENTORY).getDataRange().getValues().slice(1);
  return data.map(row => ({
    itemId: row[0], itemName: row[2], uom: row[3], currentStock: parseFloat(row[4]) || 0,
    unitPrice: parseFloat(row[5]) || 0, reorderPoint: parseFloat(row[9]) || 0,
    moq: parseFloat(row[10]) || 1, vendorID: String(row[11]).trim(), status: row[12], sku: row[15]
  })).filter(item => item.itemName && (item.status === "Sold out" || item.currentStock <= item.reorderPoint));
}

// FETCH HISTORY (Adaptive Range)
function getPOHistory() {
  const ss = SpreadsheetApp.openById(ID_PO_ORDERS);
  const sheet = ss.getSheetByName(TAB_ORDERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  return data.map(row => ({
    priority: row[0], poNumber: row[1], vendorID: row[2], status: row[3],
    date: row[4] instanceof Date ? row[4].toLocaleDateString() : row[4],
    cost: row[6]
  })).filter(po => po.poNumber).reverse();
}

// UPDATE STATUS
function updatePOStatus(poNumber, newStatus) {
  const ss = SpreadsheetApp.openById(ID_PO_ORDERS);
  const sheet = ss.getSheetByName(TAB_ORDERS);
  const data = sheet.getRange("B:B").getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === poNumber) {
      sheet.getRange(i + 1, 4).setValue(newStatus);
      return true;
    }
  }
  return false;
}

function finalizePurchaseOrder(poHeader, lineItems) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
    const ssOrders = SpreadsheetApp.openById(ID_PO_ORDERS);
    const ssLines = SpreadsheetApp.openById(ID_PO_ORDERS_LINE_ITEMS);
    const sheetHeader = ssOrders.getSheetByName(TAB_ORDERS);
    const sheetLines = ssLines.getSheetByName(TAB_LINE_ITEMS);

    // 1. DYNAMIC ROW FINDER: Find the first empty slot in Column B
    const colB = sheetHeader.getRange("B:B").getValues();
    let targetRow = 2; // Default starting point after header
    for (let i = 1; i < colB.length; i++) {
      if (colB[i][0] === "" || colB[i][0] === null) {
        targetRow = i + 1;
        break;
      }
      if (i === colB.length - 1) targetRow = colB.length + 1;
    }

    // 2. AUTO-INSERT ROW: This pushes the rest of the template down
    // This maintains the "design" of your excel template
    sheetHeader.insertRowBefore(targetRow);

    // 3. PREPARE DATA
    const headerData = [
      poHeader.priority, 
      poHeader.poNumber, 
      poHeader.vendorID, 
      "Ordered",
      new Date(), 
      poHeader.arriveBy, 
      poHeader.cost, 
      poHeader.poc, 
      poHeader.notes
    ];

    // 4. WRITE DATA to the newly inserted row
    sheetHeader.getRange(targetRow, 1, 1, headerData.length).setValues([headerData]);
    
    // 5. SAVE LINE ITEMS (Standard append)
    lineItems.forEach((item, idx) => {
      sheetLines.appendRow([
        (idx + 1), 
        poHeader.poNumber, 
        item.itemName, 
        item.quantity, 
        item.uom, 
        item.unitPrice, 
        item.subtotal, 
        item.sku, 
        ""
      ]);
    });
    
    SpreadsheetApp.flush(); 
    return { success: true };
  } catch (e) { 
    return { success: false, error: e.toString() }; 
  } finally { 
    lock.releaseLock(); 
  }
}

function generatePOPreview(poHeader, lineItems) {
  let html = `<html><head><style>body{font-family:sans-serif;padding:30px}table{width:100%;border-collapse:collapse;margin:20px 0}th,td{border:1px solid #ddd;padding:8px;text-align:left}th{background:#f4f4f4}.total{text-align:right;font-size:1.5em;color:#673ab7}</style></head><body>
  <div style="display:flex;justify-content:space-between;border-bottom:2px solid #673ab7"><div><h1>PURCHASE ORDER</h1><p>Vendor: ${poHeader.vendorID}</p></div><div style="text-align:right"><h2>${poHeader.poNumber}</h2><p>Date: ${poHeader.date || new Date().toLocaleDateString()}</p></div></div>
  <table><thead><tr><th>#</th><th>Item</th><th>SKU</th><th>Qty</th><th>Price</th><th>Total</th></tr></thead><tbody>
  ${lineItems.map((item, i) => `<tr><td>${i+1}</td><td>${item.itemName}</td><td>${item.sku}</td><td>${item.quantity}</td><td>₹${item.unitPrice}</td><td>₹${item.subtotal}</td></tr>`).join('')}
  </tbody></table><div class="total">Grand Total: ₹${poHeader.cost}</div></body></html>`;
  return html;
}