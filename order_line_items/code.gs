// --- CONFIGURATION ---
var ID_PARENTS = "1xgcQfWYczXmkwpQsbonkRUraAMvlWExNRtm7D_iSJbk";
var ID_INVENTORY = "1YDiJsrkNEj4HxDaNlirGIczAX4h7FExpb3XNs9Xu5co";
var ID_ORDERS_LINE_ITEMS = "1j5ma5hH1vKaoNW0O3JrYL19FZvPLBXMOyN5_0efP0e8";
var ID_ORDERS = "1i3XQ7tfoKKb6RH8CjyP0fryMnbuOthbXnb26-FCa0MU";

var TAB_PARENTS = "main";
var TAB_LINE_ITEMS = "main";
var TAB_ORDERS = "main";

var VALID_SHEETS = ["dhanyam", "varnam", "vastram", "gavya", "soaps", "snacks"];

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
  let allItems = [];

  VALID_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const items = data.slice(1).map(row => ({
      sku: String(row[15]),   // Column P
      mainCategory: sheetName,
      subCategory: row[1],
      itemName: row[2],
      uom: row[3],
      stock: parseFloat(row[4]) || 0,
      salePrice: parseFloat(row[7]) || 0, // Ensure this is a number
      moq: parseFloat(row[10]) || 0.5,
      imageUrl: row[16] || "https://via.placeholder.com/150"
    })).filter(item => item.sku && item.sku !== "undefined" && item.stock > 0);

    allItems = allItems.concat(items);
  });
  return allItems;
}

function finalizeOrderBulk(summary, fullCart) {
  try {
    const liSheet = SpreadsheetApp.openById(ID_ORDERS_LINE_ITEMS).getSheetByName(TAB_LINE_ITEMS);
    const ordSheet = SpreadsheetApp.openById(ID_ORDERS).getSheetByName(TAB_ORDERS);
    const invSS = SpreadsheetApp.openById(ID_INVENTORY);

    // 1. Save Line Items
    const lineRows = fullCart.map((item, index) => [
      index + 1,
      summary.orderId,
      item.mainCategory,     // Fixed: Use mainCategory instead of category
      item.sku,              // Fixed: Use sku instead of itemId
      item.itemName,
      item.quantity,
      item.uom,
      item.salePrice,
      item.fullSubtotal,
      ""
    ]);

    const nextLiRow = getFirstEmptyRowInColumn(liSheet, 2);
    liSheet.getRange(nextLiRow, 1, lineRows.length, 10).setValues(lineRows);

    // 2. Save Order Summary
    const ordRow = [[
      "P0",
      summary.orderId,
      summary.customerId,
      summary.customerName,
      new Date(),
      "Received",
      summary.finalTotal,
      "Not Received",
      summary.notes
    ]];

    const nextOrdRow = getFirstEmptyRowInColumn(ordSheet, 2);
    ordSheet.getRange(nextOrdRow, 1, 1, 9).setValues(ordRow);

    // 3. INVENTORY SYNC (Searching by SKU in Column P)
    fullCart.forEach(cartItem => {
      if (VALID_SHEETS.indexOf(cartItem.mainCategory) === -1) return;

      const targetSheet = invSS.getSheetByName(cartItem.mainCategory);
      if (!targetSheet) return;

      const data = targetSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        // Look for SKU in Column P (index 15)
        if (String(data[i][15]) === String(cartItem.sku)) {
          let currentStock = parseFloat(data[i][4]) || 0;
          let reorderPoint = parseFloat(data[i][9]) || 0;
          let newStock = currentStock - cartItem.quantity;

          // Update Stock (Col E - Index 5)
          targetSheet.getRange(i + 1, 5).setValue(newStock);

          // Update Status (Col M - Index 13)
          let status = newStock <= 0 ? "Sold out" : (newStock <= reorderPoint ? "Repurchase needed" : "In stock");
          targetSheet.getRange(i + 1, 13).setValue(status);
          break;
        }
      }
    });

    // 4. EMAIL INVOICE
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
  // Get all values in the specific column (e.g., Column B for Order ID)
  const range = sheet.getRange(1, col, sheet.getMaxRows()).getValues();
  
  // Loop from top to bottom to find the first truly empty cell
  for (let i = 0; i < range.length; i++) {
    if (range[i][0] === "" || range[i][0] === null || range[i][0] === undefined) {
      return i + 1;
    }
  }
  return sheet.getLastRow() + 1;
}