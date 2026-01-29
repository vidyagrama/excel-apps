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

function getVendors() {
  const ss = SpreadsheetApp.openById(ID_VENDORS);
  const data = ss.getSheetByName(TAB_VENDORS).getDataRange().getValues().slice(1);
  return data.map(row => ({ vendorID: row[0], businessName: row[1], poc: row[7] }));
}

function getStockAlerts() {
  const ss = SpreadsheetApp.openById(ID_INVENTORY);
  const data = ss.getSheetByName(TAB_INVENTORY).getDataRange().getValues().slice(1);
  return data.map(row => ({
    itemId: row[0], itemName: row[2], uom: row[3], currentStock: parseFloat(row[4]) || 0,
    unitPrice: parseFloat(row[5]) || 0, reorderPoint: parseFloat(row[9]) || 0,
    moq: parseFloat(row[10]) || 1, vendorID: String(row[11]).trim(), status: row[12], sku: row[15]
  })).filter(item => item.itemName && (item.status === "Sold out" || item.currentStock <= item.reorderPoint));
}

function getPOHistory(offset) {
  offset = offset || 0;
  const limit = 50; 
  const ss = SpreadsheetApp.openById(ID_PO_ORDERS);
  const sheet = ss.getSheetByName(TAB_ORDERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { data: [], hasMore: false };

  const endRow = Math.max(2, lastRow - offset);
  const startRow = Math.max(2, endRow - limit + 1);
  const numRows = endRow - startRow + 1;
  if (numRows <= 0) return { data: [], hasMore: false };
  
  const data = sheet.getRange(startRow, 1, numRows, 10).getValues();
  const formattedData = data.map(row => {
    if (!row[1]) return null;
    let leadTime = "N/A";
    if (row[4] instanceof Date && row[9] instanceof Date) {
      const diff = Math.abs(row[9] - row[4]);
      leadTime = Math.ceil(diff / (1000 * 60 * 60 * 24)) + " days";
    }
    return {
      priority: String(row[0] || ""), poNumber: String(row[1] || ""), vendorID: String(row[2] || ""),
      status: String(row[3] || "Ordered"), date: row[4] instanceof Date ? row[4].toLocaleDateString() : String(row[4]),
      cost: row[6] || 0, leadTime: leadTime
    };
  }).filter(x => x !== null).reverse();
  return { data: formattedData, hasMore: startRow > 2 };
}

function updatePOStatus(poNumber, newStatus) {
  const ss = SpreadsheetApp.openById(ID_PO_ORDERS);
  const sheet = ss.getSheetByName(TAB_ORDERS);
  const data = sheet.getRange("B:B").getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === poNumber) {
      sheet.getRange(i + 1, 4).setValue(newStatus);
      if (newStatus === "Delivered") sheet.getRange(i + 1, 10).setValue(new Date());
      return true;
    }
  }
  return false;
}

/**
 * UPDATED: Physically Deletes PO from Header and Line Item sheets
 */
function deletePurchaseOrder(poNumber) {
  const ssOrders = SpreadsheetApp.openById(ID_PO_ORDERS);
  const ssLines = SpreadsheetApp.openById(ID_PO_ORDERS_LINE_ITEMS);
  const sheetHeader = ssOrders.getSheetByName(TAB_ORDERS);
  const sheetLines = ssLines.getSheetByName(TAB_LINE_ITEMS);

  // Delete from Orders Header
  const headerData = sheetHeader.getRange("B:B").getValues();
  for (let i = headerData.length - 1; i >= 0; i--) {
    if (headerData[i][0] === poNumber) { sheetHeader.deleteRow(i + 1); break; }
  }

  // Delete all matching Line Items
  const lineData = sheetLines.getRange("B:B").getValues();
  for (let j = lineData.length - 1; j >= 0; j--) {
    if (lineData[j][0] === poNumber) { sheetLines.deleteRow(j + 1); }
  }
  return { success: true };
}

function finalizePurchaseOrder(poHeader, lineItems) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ssOrders = SpreadsheetApp.openById(ID_PO_ORDERS);
    const ssLines = SpreadsheetApp.openById(ID_PO_ORDERS_LINE_ITEMS);
    const sheetHeader = ssOrders.getSheetByName(TAB_ORDERS);
    const sheetLines = ssLines.getSheetByName(TAB_LINE_ITEMS);

    const colB = sheetHeader.getRange("B:B").getValues();
    let targetRow = 2;
    for (let i = 1; i < colB.length; i++) {
      if (colB[i][0] === "" || colB[i][0] === null) { targetRow = i + 1; break; }
      if (i === colB.length - 1) targetRow = colB.length + 1;
    }

    sheetHeader.insertRowBefore(targetRow);
    const headerData = [poHeader.priority, poHeader.poNumber, poHeader.vendorID, "Ordered", new Date(), poHeader.arriveBy, poHeader.cost, poHeader.poc, ""];
    sheetHeader.getRange(targetRow, 1, 1, headerData.length).setValues([headerData]);
    
    lineItems.forEach((item, idx) => {
      sheetLines.appendRow([(idx + 1), poHeader.poNumber, item.itemName, item.quantity, item.uom, item.unitPrice, item.subtotal, item.sku, ""]);
    });
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

/**
 * UPDATED: Includes UOM in the print preview
 */
function getSpecificPODetails(poNumber) {
  const ssOrders = SpreadsheetApp.openById(ID_PO_ORDERS);
  const ssLines = SpreadsheetApp.openById(ID_PO_ORDERS_LINE_ITEMS);
  const ssVendors = SpreadsheetApp.openById(ID_VENDORS);
  
  const header = ssOrders.getSheetByName(TAB_ORDERS).getDataRange().getValues().find(r => r[1] === poNumber);
  const lineRows = ssLines.getSheetByName(TAB_LINE_ITEMS).getDataRange().getValues();
  const vendorRows = ssVendors.getSheetByName(TAB_VENDORS).getDataRange().getValues();
  
  // Pulling UOM (Index 4) from Line Items sheet
  const items = lineRows.filter(r => r[1] === poNumber).map(r => ({ 
    itemName: r[2], 
    quantity: r[3],
    uom: r[4] 
  }));
  
  const vendorID = header[2];
  const vendorMatch = vendorRows.find(v => v[0] === vendorID);
  const businessName = vendorMatch ? vendorMatch[1] : "N/A";

  const h = { 
    poNumber: header[1], 
    vendorID: vendorID, 
    businessName: businessName,
    date: header[4] instanceof Date ? header[4].toLocaleDateString() : header[4]
  };
  
  return `<html><body style="font-family:sans-serif;padding:30px">
    <div style="border-bottom:2px solid #673ab7;padding-bottom:10px;margin-bottom:20px">
      <h2 style="color:#673ab7;margin:0">PURCHASE ORDER</h2>
      <p style="margin:5px 0"><b>PO#:</b> ${h.poNumber} | <b>Date:</b> ${h.date}</p>
    </div>
    <div style="margin-bottom:20px; line-height: 1.5;">
      <b>Vendor ID:</b> ${h.vendorID}<br>
      <b>Business Name:</b> ${h.businessName}
    </div>
    <table border="1" style="width:100%;border-collapse:collapse">
      <tr style="background:#f2f2f2">
        <th width="50">S.No</th>
        <th style="text-align:left;padding:8px">Item Description</th>
        <th width="100" style="padding:8px">Quantity</th>
        <th width="80" style="padding:8px">UOM</th>
      </tr>
      ${items.map((i, idx) => `<tr>
        <td style="text-align:center;padding:8px">${idx + 1}</td>
        <td style="padding:8px">${i.itemName}</td>
        <td style="text-align:center;padding:8px">${i.quantity}</td>
        <td style="text-align:center;padding:8px">${i.uom || ''}</td>
      </tr>`).join('')}
    </table>
    <div style="margin-top:50px;font-size:11px;color:#666">Generated via Vidyagrama Procurement System</div>
  </body></html>`;
}