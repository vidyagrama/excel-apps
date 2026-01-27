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

// 1. FETCH VENDORS
function getVendors() {
  const ss = SpreadsheetApp.openById(ID_VENDORS);
  const sheet = ss.getSheetByName(TAB_VENDORS);
  const data = sheet.getDataRange().getValues().slice(1);
  return data.map(row => ({
    vendorID: row[0],
    businessName: row[1],
    poc: row[7],
    phone: row[8],
    leadTime: row[13] || 0
  }));
}

// 2. FETCH STOCK ALERTS
function getStockAlerts() {
  const ss = SpreadsheetApp.openById(ID_INVENTORY);
  const data = ss.getSheetByName(TAB_INVENTORY).getDataRange().getValues().slice(1);
  
  return data.map(row => ({
    itemId: row[0],
    itemName: row[2],
    uom: row[3],
    currentStock: parseFloat(row[4]) || 0,
    unitPrice: parseFloat(row[5]) || 0,
    reorderPoint: parseFloat(row[9]) || 0,
    moq: parseFloat(row[10]) || 1,
    vendorID: String(row[11]).trim(),
    status: row[12],
    sku: row[15]
  })).filter(item => item.itemName && (item.status === "Sold out" || item.currentStock <= item.reorderPoint));
}

// 3. FINAL SAVE (With Row 12 Fix & Sequential IDs)
function finalizePurchaseOrder(poHeader, lineItems) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
    
    const ssOrders = SpreadsheetApp.openById(ID_PO_ORDERS);
    const ssLines = SpreadsheetApp.openById(ID_PO_ORDERS_LINE_ITEMS);
    const sheetHeader = ssOrders.getSheetByName(TAB_ORDERS);
    const sheetLines = ssLines.getSheetByName(TAB_LINE_ITEMS);

    // FIX: Check if Row 12 is the placeholder "po_number"
    let targetRow = sheetHeader.getLastRow() + 1;
    const checkValue = sheetHeader.getRange("B12").getValue();
    if (checkValue === "po_number" || checkValue === "") {
      targetRow = 12;
    }

    // Save Header
    const headerData = [
      poHeader.priority, poHeader.poNumber, poHeader.vendorID, "Pending",
      new Date(), poHeader.arriveBy, poHeader.cost, poHeader.poc, poHeader.notes
    ];
    sheetHeader.getRange(targetRow, 1, 1, headerData.length).setValues([headerData]);
    
    // Save Lines with Sequential ID (1, 2, 3...)
    lineItems.forEach((item, idx) => {
      sheetLines.appendRow([
        (idx + 1), poHeader.poNumber, item.itemName,
        item.quantity, item.uom, item.unitPrice, item.subtotal, item.sku, ""
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

// 4. PRINT PREVIEW GENERATOR
function generatePOPreview(poHeader, lineItems) {
  let html = `
    <html>
    <head>
      <style>
        body { font-family: 'Segoe UI', sans-serif; padding: 40px; color: #333; }
        .header-table { width: 100%; border-bottom: 2px solid #673ab7; margin-bottom: 20px; }
        table.items { width: 100%; border-collapse: collapse; margin-top: 20px; }
        table.items th, table.items td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        table.items th { background-color: #f8f9fa; }
        .total-box { text-align: right; margin-top: 20px; font-size: 1.4em; font-weight: bold; color: #673ab7; }
        .footer { margin-top: 50px; font-size: 0.8em; color: #777; border-top: 1px solid #eee; padding-top: 10px; }
      </style>
    </head>
    <body>
      <table class="header-table">
        <tr>
          <td><h1>PURCHASE ORDER</h1><p><strong>Vendor:</strong> ${poHeader.vendorID}</p></td>
          <td style="text-align:right"><h3>${poHeader.poNumber}</h3><p>Date: ${new Date().toLocaleDateString()}<br>Expected: ${poHeader.arriveBy}</p></td>
        </tr>
      </table>
      <table class="items">
        <thead><tr><th>#</th><th>Item Name</th><th>SKU</th><th>Qty</th><th>Price</th><th>Total</th></tr></thead>
        <tbody>
          ${lineItems.map((item, idx) => `<tr><td>${idx+1}</td><td>${item.itemName}</td><td>${item.sku}</td><td>${item.quantity}</td><td>₹${item.unitPrice}</td><td>₹${item.subtotal}</td></tr>`).join('')}
        </tbody>
      </table>
      <div class="total-box">Grand Total: ₹${poHeader.cost}</div>
      <div class="footer">Note: ${poHeader.notes} | POC: ${poHeader.poc}</div>
    </body>
    </html>`;
  return html;
}