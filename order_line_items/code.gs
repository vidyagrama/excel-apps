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
    .setTitle("Vidyagrama Online Order")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
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
    itemId: row[0], 
    category: row[1], 
    itemName: row[2], 
    uom: row[3], 
    salePrice: row[7],
    imageUrl: row[15] || "https://via.placeholder.com/150" 
  }));
}

function finalizeOrderBulk(summary, fullCart) {
  try {
    const liSheet = SpreadsheetApp.openById(ID_ORDERS_LINE_ITEMS).getSheetByName(TAB_LINE_ITEMS);
    const ordSheet = SpreadsheetApp.openById(ID_ORDERS).getSheetByName(TAB_ORDERS);
    const invSheet = SpreadsheetApp.openById(ID_INVENTORY).getSheetByName(TAB_INVENTORY);

    // 1. Save Line Items (Mapping preserved from your working code)
    const lineRows = fullCart.map((item, index) => [
      index + 1,            // Column A: Sr No
      summary.orderId,      // Column B: Order ID
      item.category,        // Column C: Category
      item.itemId,          // Column D: Item ID
      item.itemName,        // Column E: Name
      item.quantity,        // Column F: Qty
      item.uom,             // Column G: UOM
      item.salePrice,       // Column H: Unit Price
      item.fullSubtotal,    // Column I: Subtotal
      ""                    // Column J: Notes
    ]);
    liSheet.getRange(getFirstEmptyRowInColumn(liSheet, 2), 1, lineRows.length, 10).setValues(lineRows);

    // 2. Save Order Summary (Mapping preserved from your working code)
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

    // 3. INVENTORY SYNC (Fixed Column Mappings)
    const invData = invSheet.getDataRange().getValues();
    
    fullCart.forEach(cartItem => {
      for (let i = 1; i < invData.length; i++) {
        if (invData[i][0] == cartItem.itemId) {
          let currentStock = parseFloat(invData[i][4]) || 0; 
          let reorderPoint = parseFloat(invData[i][9]) || 0; 
          let newStock = currentStock - cartItem.quantity;
          invSheet.getRange(i + 1, 5).setValue(newStock);
          
          let status = "In stock";
          if (newStock <= 0) status = "Sold out";
          else if (newStock <= reorderPoint) status = "Repurchase needed";
          invSheet.getRange(i + 1, 12).setValue(status);
          break;
        }
      }
    });

    // 4. EMAIL RECEIPT
    sendReceiptEmail(summary, fullCart);

    SpreadsheetApp.flush(); 
    return true;
  } catch (e) { 
    return e.toString(); 
  }
}

function sendReceiptEmail(summary, cart) {
  try {
    const ss = SpreadsheetApp.openById(ID_PARENTS);
    const data = ss.getSheetByName(TAB_PARENTS).getDataRange().getValues();
    
    // Improved matching logic: trims whitespace to prevent lookup failure
    const user = data.find(r => String(r[0]).trim() === String(summary.customerId).trim());
    const userEmail = user ? user[6] : null;

    if (!userEmail) {
      console.log("No email found for ID: " + summary.customerId);
      return;
    }

    let itemTable = cart.map(i => `
      <tr>
        <td style="padding: 8px; border-bottom: 1px solid #eee;">${i.itemName}</td>
        <td style="padding: 8px; border-bottom: 1px solid #eee;">${i.quantity} ${i.uom}</td>
        <td style="padding: 8px; border-bottom: 1px solid #eee;">₹${i.fullSubtotal}</td>
      </tr>`).join("");
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #ddd; padding: 20px;">
        <h2 style="color: #2e7d32;">Order Confirmation</h2>
        <p>Namaste <b>${summary.customerName}</b>,</p>
        <p>Your Order <b>${summary.orderId}</b> has been placed successfully.</p>
        <table style="width: 100%; border-collapse: collapse;">
          <tr style="background: #f4f4f4;">
            <th style="text-align: left; padding: 8px;">Item</th>
            <th style="text-align: left; padding: 8px;">Qty</th>
            <th style="text-align: left; padding: 8px;">Amount</th>
          </tr>
          ${itemTable}
        </table>
        <p style="font-size: 18px; margin-top: 20px;"><b>Final Total: ₹${summary.finalTotal}</b></p>
        <p><small>Note: ${summary.notes || "None"}</small></p>
      </div>
    `;

    MailApp.sendEmail({
      to: userEmail,
      bcc: "writetovidyagrama@gmail.com", // Keeping a copy for your records
      subject: "New Grocery Order - " + summary.orderId,
      htmlBody: htmlBody
    });
  } catch (e) {
    console.log("Email Error: " + e.toString());
  }
}

function testEmail() {
  const summary = {
    orderId: "TEST-123",
    customerName: "Admin Test",
    customerId: "1", // REPLACE THIS with a real ID from your spreadsheet Column A
    finalTotal: "100",
    notes: "Testing"
  };
  const cart = [{itemName: "Test Item", quantity: 1, uom: "kg", fullSubtotal: 100}];
  
  sendReceiptEmail(summary, cart);
}

function getFirstEmptyRowInColumn(sheet, col) {
  const range = sheet.getRange(1, col, sheet.getMaxRows()).getValues();
  for (let i = 0; i < range.length; i++) { 
    if (range[i][0] === "" || range[i][0] === null) return i + 1; 
  }
  return sheet.getLastRow() + 1;
}