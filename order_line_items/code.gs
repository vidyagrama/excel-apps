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
    moq: parseFloat(row[10]) || 0.5,
    imageUrl: row[16] || "https://via.placeholder.com/150" 
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

    // 4. EMAIL INVOICE (Newly Integrated Format)
    sendReceiptEmail(summary, fullCart);

    SpreadsheetApp.flush(); 
    return true;
  } catch (e) { 
    return e.toString(); 
  }
}

/**
 * Integrated Tax Invoice Email Logic
 */
function sendReceiptEmail(summary, cart) {
  try {
    const parentSS = SpreadsheetApp.openById(ID_PARENTS);
    const parentData = parentSS.getSheetByName(TAB_PARENTS).getDataRange().getValues();
    const user = parentData.find(r => String(r[0]).trim() === String(summary.customerId).trim());
    const userEmail = user ? user[6] : null;

    if (!userEmail) return;

    // --- Configuration for Invoice ---
    const logoUrl = "https://i.ibb.co/3mk7ddzj/vidyagrama-logo.png";
    const upiId = "9035734752@icici";
    
    let tableRows = "";
    let overallTotal = 0;

    cart.forEach(item => {
      let qty = parseFloat(item.quantity);
      let price = parseFloat(item.salePrice);
      let unit = item.uom;
      
      // Unit conversion
      if (unit.toLowerCase() === 'gms') {
        qty = qty / 1000;
        unit = 'kg';
      }

      let lineTotal = qty * price;
      overallTotal += lineTotal;

      tableRows += `
        <tr>
          <td style="border: 1px solid #cccccc; padding: 10px;">${item.itemName}</td>
          <td align="right" style="border: 1px solid #cccccc; padding: 10px;">${qty} ${unit}</td>
          <td align="right" style="border: 1px solid #cccccc; padding: 10px;">₹ ${price.toFixed(2)}</td>
          <td align="right" style="border: 1px solid #cccccc; padding: 10px;">₹ ${lineTotal.toFixed(2)}</td>
        </tr>`;
    });

    const discountRate = parseFloat(user[7] || 0);
    const discountAmount = overallTotal * (discountRate / 100);
    const finalAmount = overallTotal - discountAmount;
    
    const upiLink = `upi://pay?pa=${upiId}&pn=Vidyakshetra&am=${finalAmount.toFixed(2)}&cu=INR`;
    const qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(upiLink)}`;

    const htmlInvoice = `
      <!DOCTYPE html>
      <html>
      <body style="font-family: sans-serif; padding: 20px; color: #333;">
        <table width="100%" style="margin-bottom: 20px; border-bottom: 2px solid #444; padding-bottom: 10px;">
          <tr>
            <td><img src="${logoUrl}" height="70" alt="Logo"></td>
            <td align="right">
              <h1 style="margin:0; font-size: 24px;">TAX INVOICE</h1>
              <p style="margin:5px 0;">No: <strong>${summary.orderId}</strong></p>
              <p style="margin:5px 0;">Date: ${new Date().toLocaleDateString('en-IN')}</p>
            </td>
          </tr>
        </table>
        <p><strong>Billed To:</strong> ${summary.customerName}</p>
        <table width="100%" style="border-collapse: collapse;">
          <thead>
            <tr style="background: #f4f4f4;">
              <th align="left" style="padding: 10px; border: 1px solid #ccc;">Description</th>
              <th align="right" style="padding: 10px; border: 1px solid #ccc;">Qty</th>
              <th align="right" style="padding: 10px; border: 1px solid #ccc;">Price</th>
              <th align="right" style="padding: 10px; border: 1px solid #ccc;">Total</th>
            </tr>
          </thead>
          <tbody>${tableRows}</tbody>
          <tfoot>
            <tr><td colspan="3" align="right" style="padding: 10px;">Subtotal</td><td align="right" style="padding: 10px;">₹ ${overallTotal.toFixed(2)}</td></tr>
            ${discountRate > 0 ? `<tr><td colspan="3" align="right" style="padding: 10px;">Discount (${discountRate}%)</td><td align="right" style="padding: 10px; color: #1e88e5;">- ₹ ${discountAmount.toFixed(2)}</td></tr>` : ''}
            <tr style="font-size: 18px;">
              <td colspan="3" align="right" style="padding: 10px; font-weight: bold;">Final Amount Due</td>
              <td align="right" style="padding: 10px; font-weight: bold; color: #d32f2f;">₹ ${finalAmount.toFixed(2)}</td>
            </tr>
          </tfoot>
        </table>
        <div style="margin-top: 40px; border-top: 1px solid #eee; padding-top: 20px;">
          <table width="100%">
            <tr>
              <td width="70%" style="vertical-align: top;">
                 <p style="font-size: 13px; font-weight: bold; margin-bottom: 5px;">A COMMUNITY ENTERPRISE INSPIRED BY THE VISION OF VIDYAKSHETRA</p>
                 <p style="font-size: 11px; color: #666;">Thank you for your support!</p>
              </td>
              <td width="30%" align="right">
                <p style="font-size: 11px; margin-bottom: 5px; font-weight: bold;">Scan to Pay via UPI</p>
                <img src="${qrCodeUrl}" width="130" height="130" style="border: 1px solid #ccc; padding: 5px;">
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>`;

    MailApp.sendEmail({
      to: userEmail,
      bcc: "writetovidyagrama@gmail.com",
      subject: "Tax Invoice - " + summary.orderId,
      htmlBody: htmlInvoice
    });

  } catch (e) {
    console.log("Email Error: " + e.toString());
  }
}

function getFirstEmptyRowInColumn(sheet, col) {
  const range = sheet.getRange(1, col, sheet.getMaxRows()).getValues();
  for (let i = 0; i < range.length; i++) { 
    if (range[i][0] === "" || range[i][0] === null) return i + 1; 
  }
  return sheet.getLastRow() + 1;
}