// --- CONFIGURATION ---
var ID_PARENTS = "1xgcQfWYczXmkwpQsbonkRUraAMvlWExNRtm7D_iSJbk";
var ID_INVENTORY = "1YDiJsrkNEj4HxDaNlirGIczAX4h7FExpb3XNs9Xu5co";
var ID_ORDERS_LINE_ITEMS = "1j5ma5hH1vKaoNW0O3JrYL19FZvPLBXMOyN5_0efP0e8";
var ID_ORDERS = "1i3XQ7tfoKKb6RH8CjyP0fryMnbuOthbXnb26-FCa0MU";
var ID_ADMINS = "1iiZtZclKgr7G7ISZFlM1We4LTmMLNkZLp_x4gP2DoOM";
var ID_LEDGER = "17BBdRWeZZCCa7WmhNnIc-P_Vau3n9WkVZaR3XCB8Pck";
var ID_SMS = "1o10_jI39_Pr3QjUoRvz42ZUive08UcKd12aedCWmQTY";

var TAB_PARENTS_MAIN = "main";
var TAB_PARENTS_GUEST = "guest";
var TAB_LINE_ITEMS_MAIN = "main";
var TAB_ORDERS_MAIN = "main";
var TAB_ADMINS_ENABLE_CATEGORY = "enable_maincategory";
var TAB_ADMINS_ACTIVITY_LOGS = "activitiy_logs";
var TAB_ADMINS_VARGA = "varga";
var TAB_LEDGER_MAIN_LEDGER = "main_ledger";
var TAB_SMS_SHEET = "Sheet1";

//var VALID_SHEETS = ["Shridhanya", "Varnam", "Vastram", "GauAmruth", "Tejas", "Madhuram"];

function doGet() {
  // 1. Create a template from the file
  var template = HtmlService.createTemplateFromFile('Index');

  // 2. Evaluate the template to execute <?!= include('Styles'); ?>
  return template.evaluate()
    .setTitle("Vidyagrama Online Order")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getVargas() {
  const ss = SpreadsheetApp.openById(ID_PARENTS);
  const data = ss.getSheetByName(TAB_PARENTS_MAIN).getDataRange().getValues();
  return [...new Set(data.slice(1).map(row => row[1]))].filter(v => v).sort();
}

function getNamesByVarga(varga) {
  const ss = SpreadsheetApp.openById(ID_PARENTS);
  const data = ss.getSheetByName(TAB_PARENTS_MAIN).getDataRange().getValues();
  return data.filter(row => row[1] === varga).map(row => row[2]);
}

function validateLogin(varga, name, mobile) {
  const ss = SpreadsheetApp.openById(ID_PARENTS);
  const data = ss.getSheetByName(TAB_PARENTS_MAIN).getDataRange().getValues();

  // Find the user based on your existing column mapping
  const user = data.find(row =>
    row[1] === varga &&
    row[2] === name &&
    String(row[5]).trim() === String(mobile).trim()
  );

  if (user) {

    const userData = {
      success: true,
      varga: varga,
      email: user[6],
      discount: user[7] || 0,
      name: user[2],
      id: user[0],
      credit: parseFloat(user[8] || 0),
      balance: parseFloat(user[9] || 0)
    };

    // LOG SUCCESS: Useful to see which Varga/Parent is active
    logActivity(
      name,
      "LOGIN_SUCCESS",
      `Varga: ${varga}`,
      "Parents_Main"
    );

    return userData;
  } else {
    return { success: false };
  }
}

function getInventoryData() {
  // Fetch dynamic categories
  const categoryConfig = getCategoryMap();
  const currentValidSheets = categoryConfig.validSheets;

  const adminSS = SpreadsheetApp.openById(ID_ADMINS);
  const adminSheet = adminSS.getSheetByName(TAB_ADMINS_ENABLE_CATEGORY);
  const adminData = adminSheet.getDataRange().getValues();

  // Get current date once outside the loop
  const now = new Date();
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // 1. Get list of currently ACTIVE categories (Normalized to lowercase)
  const activeCategories = adminData.slice(1).reduce((acc, row) => {
    const category = String(row[0]).toLowerCase().trim();
    const status = String(row[1]).toLowerCase().trim();

    // Check if cells are empty
    if (!row[2] || !row[3]) return acc;

    try {
      // 3. Format From/To dates from the sheet into YYYY-MM-DD
      const fromStr = Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      const toStr = Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "yyyy-MM-dd");

      // 4. Compare strings
      if (status === 'enable' && nowStr >= fromStr && nowStr <= toStr) {
        acc.push(category);
      }
    } catch (e) {
      console.log(`Error parsing dates for ${category}: ${e.message}`);
    }

    return acc;
  }, []);

  const ss = SpreadsheetApp.openById(ID_INVENTORY);
  let allItems = [];

  // 2. Normalize valid sheets for comparison
  const normalizedValidSheets = currentValidSheets.map(s => s.toLowerCase().trim());

  // FIXED: Added 'index' to the forEach parameters
  normalizedValidSheets.forEach((sheetName, index) => {
    // Compare lowercase sheet name against our active list
    if (activeCategories.indexOf(sheetName) === -1) return;

    // Use the actual sheet name from the valid list to open the tab
    const originalSheetName = currentValidSheets[index];
    const sheet = ss.getSheetByName(originalSheetName);

    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const items = data.slice(1).map(row => ({
      sku: String(row[15]),
      mainCategory: originalSheetName,
      subCategory: row[1],
      itemName: row[2],
      uom: row[3],
      stock: parseFloat(row[4]) || 0,
      salePrice: parseFloat(row[7]) || 0,
      moq: parseFloat(row[10]) || 0.5,
      imageUrl: (row[16] || "https://via.placeholder.com/150") + "?v=" + new Date().getTime()
    })).filter(item => item.sku && item.sku !== "undefined");

    allItems = allItems.concat(items);
  });

  return allItems;
}

/** Finalize order bulk***/
function finalizeOrderBulk(summary, fullCart, paymentMode, base64Image, txnId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const currentValidSheets = getCategoryMap().validSheets;
    const liSheet = SpreadsheetApp.openById(ID_ORDERS_LINE_ITEMS).getSheetByName(TAB_LINE_ITEMS_MAIN);
    const ordSheet = SpreadsheetApp.openById(ID_ORDERS).getSheetByName(TAB_ORDERS_MAIN);
    const invSS = SpreadsheetApp.openById(ID_INVENTORY);

    // --- STEP 1: CALCULATE DUES & WALLET ADJUSTMENTS ---
    const parentsSS = SpreadsheetApp.openById(ID_PARENTS);
    const parentsSheet = parentsSS.getSheetByName(TAB_PARENTS_MAIN);
    const parentsData = parentsSheet.getDataRange().getValues();

    const userRowIndex = parentsData.findIndex(row => String(row[0]).trim() === String(summary.customerId).trim());

    let initialDues = 0;
    let updatedDues = 0;
    let updatedCredit = 0;
    let creditUsed = 0;
    let totalWalletDeducted = 0;

    if (userRowIndex !== -1) {
      const sheetRowNumber = userRowIndex + 1;

      let currentCredit = parseFloat(parentsData[userRowIndex][8] || 0); // Col I (Credit)
      let currentDues = parseFloat(parentsData[userRowIndex][9] || 0);   // Col J (Dues)
      initialDues = currentDues;

      const grossVal = parseFloat(summary.grossTotal || 0);
      const savingsVal = parseFloat(summary.savingAmount || 0);
      let orderCharge = Math.max(0, grossVal - savingsVal);

      if (paymentMode !== "Gift" && paymentMode !== "Kitchen") {

        // 1. Deduct from available Credit first
        if (currentCredit > 0 && orderCharge > 0) {
          if (currentCredit >= orderCharge) {
            creditUsed = orderCharge;
            currentCredit -= orderCharge;
            orderCharge = 0;
          } else {
            creditUsed = currentCredit;
            orderCharge -= currentCredit;
            currentCredit = 0;
          }
          totalWalletDeducted = creditUsed;
        }

        // 2. Handle remaining charge or previous dues based on payment mode
        if (paymentMode === "Wallet") {
          // Uncovered order balance becomes remaining dues
          currentDues = currentDues + orderCharge;
        } else {
          // Cash/UPI/Auto-Verified clears any remaining previous dues
          if (currentDues > 0) {
            console.log(`💵 [DUES_CLEARED] Cleared ₹${currentDues} previous dues via ${paymentMode}.`);
            currentDues = 0.00;
          }
        }

        updatedDues = parseFloat(currentDues.toFixed(2));
        updatedCredit = parseFloat(currentCredit.toFixed(2));

        // Batch write Col I (Credit) and Col J (Dues) together
        parentsSheet.getRange(sheetRowNumber, 9, 1, 2).setValues([[updatedCredit, updatedDues]]);
        SpreadsheetApp.flush();
      } else {
        updatedDues = currentDues;
        updatedCredit = currentCredit;
      }
    }

    // --- STEP 2: DETERMINE ORDER STATUS & TYPE ---
    const isAutoPaid = (paymentMode === "Cash" || paymentMode === "Gift" || paymentMode === "Kitchen" || paymentMode === "Auto-Verified" || paymentMode === "Wallet" || totalWalletDeducted > 0);
    const orderStatus = "Received";
    const paymentStatus = isAutoPaid ? "Paid" : "Pending Verification";

    let orderType = "Online";
    if (paymentMode === "Cash") orderType = "Cash";
    else if (paymentMode === "Gift") orderType = "Gift";
    else if (paymentMode === "Kitchen") orderType = "Kitchen";
    else if (paymentMode === "Wallet" || totalWalletDeducted > 0) orderType = "Wallet";

    const categoriesInCart = [...new Set(fullCart.map(item => item.mainCategory))];
    let generatedOrderIds = [];

    // --- STEP 3: SAVE ORDERS & LINE ITEMS ---
    categoriesInCart.forEach((cat) => {
      const catItems = fullCart.filter(item => item.mainCategory === cat);
      const catOrderId = generateOrderId(cat);
      generatedOrderIds.push(catOrderId);

      const categoryGrossTotal = catItems.reduce((sum, item) => sum + (parseFloat(item.fullSubtotal) || 0), 0);

      let categoryFinalTotal = categoryGrossTotal;
      if (summary.grossTotal && summary.grossTotal > 0 && paymentMode !== "Gift") {
        const discountRatio = (summary.finalTotal / summary.grossTotal);
        categoryFinalTotal = categoryGrossTotal * discountRatio;
      }

      const lineRows = catItems.map((item, index) => [
        index + 1, catOrderId, item.mainCategory, item.subCategory || "",
        item.sku, item.itemName, item.quantity, item.uom,
        item.salePrice, item.fullSubtotal, ""
      ]);

      const nextLiRow = getFirstEmptyRowInColumn(liSheet, 2);
      liSheet.getRange(nextLiRow, 1, lineRows.length, 11).setValues(lineRows);

      let prefix = "";
      if (paymentMode === "Cash") prefix = "💵 CASH";
      else if (paymentMode === "Gift") prefix = "🎁 GIFT";
      else if (paymentMode === "Kitchen") prefix = "🍳 KITCHEN"; 
      else if (paymentMode === "Auto-Verified") prefix = "⚡ AUTO-VERIFIED";
      else if (paymentMode === "Manual UTR-4") prefix = "🔢 UTR-4 ENTRY";
      else if (paymentMode === "Wallet" || totalWalletDeducted > 0) prefix = "👛 WALLET";

      let notes = `${prefix} | TXN/Ref: ${txnId || "N/A"}`;
      if (totalWalletDeducted > 0) {
        notes += ` | Adjusted ₹${totalWalletDeducted.toFixed(2)} from user credit`;
      }
      if (summary.notes && summary.notes.trim() !== "") {
        notes += ` | Notes: ${summary.notes.trim()}`;
      }

      const ordRow = [[
        "P0", catOrderId, summary.customerId, summary.customerName,
        new Date(), orderStatus, categoryFinalTotal, paymentStatus, notes, orderType
      ]];

      const nextOrdRow = getFirstEmptyRowInColumn(ordSheet, 2);
      ordSheet.getRange(nextOrdRow, 1, 1, 10).setValues(ordRow);
    });

    // --- STEP 4: INVENTORY SYNC ---
    fullCart.forEach(cartItem => {
      if (currentValidSheets.indexOf(cartItem.mainCategory) === -1) return;
      const targetSheet = invSS.getSheetByName(cartItem.mainCategory);
      if (!targetSheet) return;

      const data = targetSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][15]) === String(cartItem.sku)) {
          let currentStock = parseFloat(data[i][4]) || 0;
          let reorderPoint = parseFloat(data[i][9]) || 0;
          let newStock = currentStock - cartItem.quantity;
          let status = newStock <= 0 ? "Sold out" : (newStock <= reorderPoint ? "Repurchase needed" : "In stock");

          targetSheet.getRange(i + 1, 5).setValue(newStock);
          targetSheet.getRange(i + 1, 13).setValue(status);
          break;
        }
      }
    });

    // --- STEP 5: EMAIL NOTIFICATION ---
    summary.allOrderIds = generatedOrderIds.join(", ");
    summary.paymentStatus = paymentStatus;
    summary.prevDues = initialDues;
    summary.creditUsed = totalWalletDeducted;
    summary.totalWalletDeducted = totalWalletDeducted;
    summary.finalAmount = Math.max(0, summary.grossTotal - summary.savingAmount - totalWalletDeducted);

    try {
      sendReceiptEmail(summary, fullCart);
    } catch (emailErr) {
      console.log("Receipt Email Failed: " + emailErr.toString());
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      orderIds: generatedOrderIds,
      mode: paymentMode,
      newBalance: updatedDues,
      newCredit: updatedCredit,
      walletAdjusted: totalWalletDeducted
    };

  } catch (e) {
    console.log("Error in finalizeOrderBulk: " + e.toString());
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Polling function called by frontend during payment overlay countdown
 */
function checkLatestPaymentReceived(customerId, expectedAmount) {
  try {
    // Search your incoming webhook transactions sheet or bank ledger tab
    const bankSS = SpreadsheetApp.openById(ID_ORDERS);
    const bankSheet = bankSS.getSheetByName("Bank_Transactions");
    if (!bankSheet) return false;

    const rows = bankSheet.getLastRow();
    if (rows < 2) return false;

    // Check last 10 incoming transactions for matching amount and status
    const data = bankSheet.getRange(Math.max(2, rows - 10), 1, 10, 5).getValues();

    // Whole rupee value of expected bill amount (e.g., 876.87 -> 876)
    const expectedWhole = Math.floor(parseFloat(expectedAmount) || 0);

    for (let i = data.length - 1; i >= 0; i--) {
      const txAmount = parseFloat(data[i][2]) || 0; // Col C: Amount
      const txStatus = String(data[i][3]).trim();   // Col D: Status (e.g. "UNMATCHED")

      const txWhole = Math.floor(txAmount); // e.g., 876.00 -> 876

      // Match whole rupee amount or absolute difference within 1 Rupee
      const isAmountMatch = (txWhole === expectedWhole) || (Math.abs(txAmount - expectedAmount) <= 1.0);

      if (isAmountMatch && txStatus === "UNMATCHED") {
        return true;
      }
    }
    return false;
  } catch (err) {
    console.log("Error in checkLatestPaymentReceived: " + err.toString());
    return false;
  }
}

/***************Wallet code*************************************/
/**
 * Fetches current credit and balance for a specific customer ID.
 */
function getUserWalletData(customerId) {
  try {
    const parentsSS = SpreadsheetApp.openById(ID_PARENTS);
    const parentsSheet = parentsSS.getSheetByName(TAB_PARENTS_MAIN);
    
    // Always flush to get absolute latest values directly from sheet
    SpreadsheetApp.flush();
    
    const parentsData = parentsSheet.getDataRange().getValues();

    // Row 0 is headers; search matching Customer ID in Column A (index 0)
    const userRow = parentsData.find(row => String(row[0]).trim() === String(customerId).trim());

    if (userRow) {
      // Index 8 = Column I (Credit), Index 9 = Column J (Balance)
      const creditVal = parseFloat(userRow[8]) || 0;
      const balanceVal = parseFloat(userRow[9]) || 0;

      return {
        success: true,
        credit: creditVal,
        balance: balanceVal
      };
    } else {
      return { success: false, error: "User not found" };
    }
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}


function checkEmailQuota() {
  Logger.log("Remaining Daily Email Quota: " + MailApp.getRemainingDailyQuota());
}

function sendReceiptEmail(summary, cart) {
  try {
    const parentSS = SpreadsheetApp.openById(ID_PARENTS);
    const parentData = parentSS.getSheetByName(TAB_PARENTS_MAIN).getDataRange().getValues();
    const user = parentData.find(r => String(r[0]).trim() === String(summary.customerId).trim());

    // User email from sheet or summary
    let userEmail = user ? user[6] : summary.customerEmail;

    const adminEmail = "writetovidyagrama@gmail.com";
    const logoUrl = "https://i.ibb.co/3mk7ddzj/vidyagrama-logo.png";
    const upiId = "9035734752@icici";

    // --- 1. GROUPING LOGIC & OVERALL TOTAL CALCULATION ---
    let overallTotal = 0;
    let itemIndex = 1; // Running Serial Number Counter
    const categories = [...new Set(cart.map(i => i.mainCategory))];
    let tableRows = "";

    categories.forEach(cat => {
      // Extended colspan to 5 for Category header
      tableRows += `
        <tr style="background-color: #fcf8e3;">
          <td colspan="5" style="border: 1px solid #cccccc; padding: 8px; font-weight: bold; color: #8a6d3b; text-transform: uppercase; font-size: 12px;">
            ${cat}
          </td>
        </tr>`;

      const catItems = cart.filter(i => i.mainCategory === cat);
      catItems.forEach(item => {
        let qty = parseFloat(item.quantity) || 0;
        let price = parseFloat(item.salePrice) || 0;
        let unit = (item.uom || '').toString();

        if (unit.toLowerCase() === 'gms' || unit.toLowerCase() === 'gm') {
          qty = qty / 1000;
          unit = 'kg';
        }

        // Use pre-calculated fullSubtotal from cart to prevent floating point drift
        let lineTotal = parseFloat(item.fullSubtotal) || (qty * price);
        overallTotal += lineTotal;

        tableRows += `
          <tr>
            <td align="center" style="border: 1px solid #cccccc; padding: 10px;">${itemIndex++}</td>
            <td style="border: 1px solid #cccccc; padding: 10px;">${item.itemName}</td>
            <td align="right" style="border: 1px solid #cccccc; padding: 10px;">${qty} ${unit}</td>
            <td align="right" style="border: 1px solid #cccccc; padding: 10px;">₹ ${price.toFixed(2)}</td>
            <td align="right" style="border: 1px solid #cccccc; padding: 10px;">₹ ${lineTotal.toFixed(2)}</td>
          </tr>`;
      });
    });

    // --- 2. DISCOUNT, ROUND OFF & WALLET CALCULATIONS ---
    const discountAmount = parseFloat(summary.savingAmount || 0);
    const discountRate = overallTotal > 0 ? ((discountAmount / overallTotal) * 100).toFixed(0) : 0;

    const unroundedItemTotal = Math.max(0, overallTotal - discountAmount);
    
    // Round to nearest whole rupee (₹1.00)
    const roundedItemTotal = Math.round(unroundedItemTotal);
    const roundOffAmt = roundedItemTotal - unroundedItemTotal;

    // Wallet / Balance variables passed from finalizeOrderBulk
    const prevBalance = parseFloat(summary.prevBalance || 0);
    const creditUsed = parseFloat(summary.creditUsed || 0);
    const totalDeducted = parseFloat(summary.totalWalletDeducted || (prevBalance + creditUsed));

    // Final Net Payable
    const finalAmount = Math.max(0, roundedItemTotal - totalDeducted);

    // --- 3. PAYMENT STATUS & NOTES HANDLING ---
    const paymentStatus = summary.paymentStatus || (summary.paymentMode ? `Paid (${summary.paymentMode})` : "Paid");
    const txnRef = summary.txnId || summary.refNo || "N/A";
    const notesText = summary.notes || summary.remarks || "";

    const isPaid = String(paymentStatus).toLowerCase().indexOf("paid") !== -1;

    const paymentStatusHtml = `
      <div style="margin: 15px 0; padding: 12px 15px; border-radius: 6px; background-color: ${isPaid ? '#e8f5e9' : '#fff3e0'}; border: 1px solid ${isPaid ? '#c8e6c9' : '#ffe0b2'};">
        <span style="font-weight: bold; color: ${isPaid ? '#2e7d32' : '#e65100'}; font-size: 14px;">Payment Status:</span>
        <span style="font-weight: bold; color: ${isPaid ? '#1b5e20' : '#bf360c'}; background: ${isPaid ? '#a5d6a7' : '#ffcc80'}; padding: 3px 8px; border-radius: 4px; font-size: 13px; margin-left: 5px;">${paymentStatus.toUpperCase()}</span>
        ${txnRef !== "N/A" ? `<span style="font-size: 12px; color: #555; margin-left: 10px;">(Ref: <strong>${txnRef}</strong>)</span>` : ''}
      </div>`;

    const notesHtml = notesText ? `
      <div style="margin: 15px 0; padding: 12px 15px; border-radius: 6px; background-color: #fff8e1; border: 1px solid #ffe082;">
        <strong style="color: #f57f17; font-size: 13px;">📝 Notes / Remarks:</strong>
        <p style="margin: 4px 0 0 0; font-size: 13px; color: #424242;">${notesText}</p>
      </div>` : '';

    const upiLink = `upi://pay?pa=${upiId}&pn=Vidyakshetra&am=${finalAmount.toFixed(2)}&cu=INR`;
    const qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(upiLink)}`;

    const mobileBannerHtml = finalAmount > 0 && !isPaid ? `
    <div style="margin: 25px 0; border: 2px dashed #2e7d32; padding: 20px; border-radius: 10px; background-color: #f9fdf9;">
      <p style="margin: 0 0 10px 0; font-weight: bold; color: #2e7d32; font-size: 16px;">📱 Payment Instructions:</p>
      <p style="margin: 5px 0; font-size: 14px; color: #333;">
        Please scan the QR code below for the <b>Combined Total</b> of all items.
      </p>
    </div>` : '';

    // --- 4. HTML INVOICE TEMPLATE ---
    const htmlInvoice = `
      <!DOCTYPE html>
      <html>
      <body style="font-family: sans-serif; padding: 20px; color: #333; line-height: 1.5;">
        <table width="100%" style="margin-bottom: 20px; border-bottom: 2px solid #444; padding-bottom: 10px;">
          <tr>
            <td><img src="${logoUrl}" height="70" alt="Logo"></td>
            <td align="right">
              <h1 style="margin:0; font-size: 24px;">TAX INVOICE</h1>
              <p style="margin:5px 0;">Ref: <strong>${summary.allOrderIds || summary.orderId || 'N/A'}</strong></p>
              <p style="margin:5px 0;">Date: ${new Date().toLocaleDateString('en-IN')}</p>
            </td>
          </tr>
        </table>
        
        <p>Namaste <strong>${summary.customerName}</strong>,</p>
        <p>Your order has been received. Here is your consolidated invoice:</p>
        
        ${paymentStatusHtml}
        ${notesHtml}

        <table width="100%" style="border-collapse: collapse; margin-top: 15px;">
          <thead>
            <tr style="background: #f4f4f4;">
              <th align="center" width="5%" style="padding: 10px; border: 1px solid #ccc;">#</th>
              <th align="left" style="padding: 10px; border: 1px solid #ccc;">Description</th>
              <th align="right" style="padding: 10px; border: 1px solid #ccc;">Qty</th>
              <th align="right" style="padding: 10px; border: 1px solid #ccc;">Price</th>
              <th align="right" style="padding: 10px; border: 1px solid #ccc;">Total</th>
            </tr>
          </thead>
          <tbody>${tableRows}</tbody>
          <tfoot>
            <tr>
              <td colspan="4" align="right" style="padding: 10px; border-top: 2px solid #eee;">Subtotal</td>
              <td align="right" style="padding: 10px; border-top: 2px solid #eee;">₹ ${overallTotal.toFixed(2)}</td>
            </tr>

            ${discountAmount > 0 ? `
            <tr>
              <td colspan="4" align="right" style="padding: 10px;">Discount (${discountRate}%)</td>
              <td align="right" style="padding: 10px; color: #1e88e5;">- ₹ ${discountAmount.toFixed(2)}</td>
            </tr>` : ''}

            <!-- ROUND OFF ROW -->
            <tr>
              <td colspan="4" align="right" style="padding: 10px; color: #666;">Round Off</td>
              <td align="right" style="padding: 10px; color: #666;">${roundOffAmt >= 0 ? '+' : ''}₹ ${roundOffAmt.toFixed(2)}</td>
            </tr>

            ${prevBalance > 0 ? `
            <tr>
              <td colspan="4" align="right" style="padding: 10px; color: #00796b;">Wallet Balance Used</td>
              <td align="right" style="padding: 10px; color: #00796b;">- ₹ ${prevBalance.toFixed(2)}</td>
            </tr>` : ''}

            ${creditUsed > 0 ? `
            <tr>
              <td colspan="4" align="right" style="padding: 10px; color: #2e7d32;">Available Credit Applied</td>
              <td align="right" style="padding: 10px; color: #2e7d32;">- ₹ ${creditUsed.toFixed(2)}</td>
            </tr>` : ''}

            <tr style="font-size: 18px;">
              <td colspan="4" align="right" style="padding: 10px; font-weight: bold; border-top: 1px solid #444;">Net Amount Payable</td>
              <td align="right" style="padding: 10px; font-weight: bold; color: ${finalAmount === 0 || isPaid ? '#2e7d32' : '#d32f2f'}; border-top: 1px solid #444;">
                ₹ ${finalAmount.toFixed(2)}
              </td>
            </tr>
          </tfoot>
        </table>

        ${mobileBannerHtml}

        <div style="margin-top: 40px; border-top: 1px solid #eee; padding-top: 20px;">
          <table width="100%">
            <tr>
              <td width="70%" style="vertical-align: top;">
                <p style="font-size: 13px; font-weight: bold; margin-bottom: 5px;">A COMMUNITY ENTERPRISE INSPIRED BY THE VISION OF VIDYAKSHETRA</p>
                <p style="font-size: 11px; color: #666;">Thank you for your support!</p>
              </td>
              ${!isPaid && finalAmount > 0 ? `
              <td width="30%" align="right">
                <p style="font-size: 11px; margin-bottom: 5px; font-weight: bold;">Scan to Pay via UPI</p>
                <img src="${qrCodeUrl}" width="130" height="130" style="border: 1px solid #ccc; padding: 5px;">
              </td>` : ''}
            </tr>
          </table>
        </div>
      </body>
      </html>`;

    // --- 5. MAIL DELIVERY ---
    const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy");

    let toList = [adminEmail];
    if (userEmail && String(userEmail).trim() !== "" && userEmail !== adminEmail) {
      toList.push(userEmail.trim());
    }

    MailApp.sendEmail({
      to: toList.join(","),
      subject: `Tax-Invoice - Vidyagram - ${formattedDate}`,
      htmlBody: htmlInvoice
    });

    console.log("✅ Email sent successfully to: " + toList.join(", "));

  } catch (e) {
    console.log("Email Error: " + e.toString());
  }
}

function getFirstEmptyRowInColumn(sheet, col) {
  const range = sheet.getRange(1, col, sheet.getMaxRows()).getValues();
  for (let i = 0; i < range.length; i++) {
    if (range[i][0] === "" || range[i][0] === null || range[i][0] === undefined) {
      return i + 1;
    }
  }
  return sheet.getLastRow() + 1;
}

//this is sample code to reauth
function REAUTH_DRIVE() {
  // This forces the script to prove it can access the folder
  const folder = DriveApp.getFolderById("1DgR2LyUJfvmGVD9HCaJcILk3Mve2_YYG");
  const sampleBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
  const sampleTxnId = "TEST_123456";
  const sampleCustomer = "Test User";

  // 1. Clean the incoming data (remove headers like "data:image/png;base64,")
  const contentType = sampleBase64.split(',')[0].split(':')[1].split(';')[0];
  const bytes = Utilities.base64Decode(sampleBase64.split(',')[1]);
  const blob = Utilities.newBlob(bytes, contentType, "Screenshot.png");

  const file = folder.createFile(blob);
  Logger.log("Permission Granted for: " + folder.getName());
}

function saveScreenshotToDrive(base64Data, txnId, customerName) {
  // Use your designated folder for screenshots
  const FOLDER_ID = '1DgR2LyUJfvmGVD9HCaJcILk3Mve2_YYG';

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // 1. Clean the incoming data (remove headers like "data:image/png;base64,")
    const contentType = base64Data.split(',')[0].split(':')[1].split(';')[0];
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);

    // 2. Create the blob (similar to your fetch logic)
    const blob = Utilities.newBlob(bytes, contentType);

    // 3. Name the file using Transaction ID and Customer Name for easy searching
    const safeName = customerName.replace(/\s+/g, '_');
    blob.setName("PAYMENT_" + txnId + "_" + safeName + ".png");

    // 4. Save to folder
    const file = folder.createFile(blob);

    // 5. Set sharing so you can view the receipt from the Google Sheet link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return file.getUrl();

  } catch (e) {
    console.error("Screenshot Upload Failed: " + e.message);
    return "Upload Error: " + e.message;
  }
}

function generateOrderId(mainCategory) {
  const ss = SpreadsheetApp.openById(ID_ORDERS);
  const sheet = ss.getSheetByName(TAB_ORDERS_MAIN);

  const catCode = mainCategory.substring(0, 3).toUpperCase();
  const dateStr = Utilities.formatDate(new Date(), "GMT+5:30", "yyyyMM");
  const prefix = `ORD-${catCode}-${dateStr}-`;

  const lastRow = sheet.getLastRow();
  let nextSerial = 1;

  if (lastRow > 1) {
    // Get only the Order ID column (Column B is index 2)
    const existingIds = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();

    const monthlyCatOrders = existingIds
      .filter(id => id && id.toString().startsWith(prefix))
      .map(id => {
        const parts = id.split("-");
        return parseInt(parts[parts.length - 1], 10);
      })
      .sort((a, b) => b - a);

    if (monthlyCatOrders.length > 0) {
      nextSerial = monthlyCatOrders[0] + 1;
    }
  }

  return prefix + ("000" + nextSerial).slice(-3);
}

/* Auto verify transactions maid through UPI payments*/
function autoCheckPayment(userExpectedAmount, userName, last4 = "") {
  cleanupSmsSheet();

  try {
    const ss = SpreadsheetApp.openById("1o10_jI39_Pr3QjUoRvz42ZUive08UcKd12aedCWmQTY");
    const sheet = ss.getSheetByName("Sheet1");
    const data = sheet.getDataRange().getValues();
    const searchAmount = parseFloat(userExpectedAmount);
    const cleanLast4 = last4 ? last4.toString().trim() : "";

    let todayMatches = [];

    // Helper to robustly check if a sheet date is today
    function isToday(rawDate) {
      if (!rawDate) return false;
      const today = new Date();
      if (rawDate instanceof Date) {
        return rawDate.toDateString() === today.toDateString();
      }
      const d = new Date(rawDate);
      if (!isNaN(d.getTime())) {
        return d.toDateString() === today.toDateString();
      }
      // Fallback for string representations like "August 15, 2026..."
      const dateStr = rawDate.toString();
      const day = today.getDate();
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const monthName = months[today.getMonth()];
      return dateStr.includes(day.toString()) && dateStr.includes(monthName);
    }

    // Loop through logs (Newest to Oldest -> Bottom to Top)
    for (let i = data.length - 1; i >= 0; i--) {
      const status = data[i][2] ? data[i][2].toString().trim() : "";
      
      // 1. Skip if already verified
      if (status.toLowerCase() === "verified") continue;

      const rawDate = data[i][0];
      const sender = data[i][1] ? data[i][1].toString().toUpperCase() : "";
      const message = data[i][3] ? data[i][3].toString() : "";

      if (!message) continue;

      // 2. Strict Same-Day Check: Ignore older unverified days
      if (!isToday(rawDate)) continue;

      // 3. Robust ICICI Sender Check
      if (!sender.includes("ICICI") && !sender.includes("ICIC")) continue;

      // 4. Enhanced Amount Regex
      const amtMatch = message.match(/(?:Rs\.?|INR|credited\s+(?:with|by)\s+Rs\.?)\s*([0-9,]+(?:\.[0-9]{1,2})?)/i);

      if (amtMatch) {
        const cleanAmtStr = amtMatch[1].replace(/,/g, '');
        const smsAmount = parseFloat(cleanAmtStr);

        if (!isNaN(smsAmount) && Math.abs(smsAmount - searchAmount) < 0.05) {

          let fullUTR = "UNKNOWN";
          const upiMatch = message.match(/UPI[:\/\s\-]*(\d{10,14})/i);

          if (upiMatch) {
            fullUTR = upiMatch[1];
          } else {
            const digitMatch = message.match(/\b\d{10,14}\b/);
            if (digitMatch) fullUTR = digitMatch[0];
          }

          const isUtrMatch = (cleanLast4 === "") || fullUTR.endsWith(cleanLast4) || message.includes(cleanLast4);

          if (isUtrMatch) {
            todayMatches.push({
              utr: fullUTR,
              sender: sender,
              amount: smsAmount,
              rowIndex: i + 1
            });
          }
        }
      }
    }

    if (todayMatches.length === 0) {
      return { status: "NOT_FOUND", message: "No matching payment found for today." };
    }

    // If exactly 1 match today (or disambiguated by last4), pick the latest one and close it
    if (todayMatches.length === 1 || (cleanLast4 !== "" && todayMatches.length > 0)) {
      const result = todayMatches[0];

      // Mark SMS row as Verified
      sheet.getRange(result.rowIndex, 3).setValue("Verified");

      // Log to Main Ledger
      logToMainLedger({
        date: new Date(),
        mainCategory: "OnlineSales",
        type: "Income",
        subType: "Online Order",
        referenceID: "UPI-AUTO",
        txnID: result.utr,
        entityName: userName,
        amount: result.amount,
        paymentMode: "UPI/Online",
        status: "Cleared",
        notes: "Auto-verified UTR: " + result.utr + " | Row: " + result.rowIndex
      });

      return { status: "SUCCESS", txnId: result.utr, message: "Verified! Received Rs." + result.amount };
    }

    // If multiple unverified payments found on the same day with no last4 provided
    if (todayMatches.length > 1 && cleanLast4 === "") {
      return { status: "DUPLICATES", message: "Multiple payments found today for Rs." + searchAmount + ". Please enter last 4 digits of UTR." };
    }

  } catch (e) {
    return { status: "ERROR", message: e.toString() };
  }
}

/**
 * BUG FIX 2: Maintains only the last 20 transactions in the SMS Log Sheet
 */
function cleanupSmsSheet() {
  const ss = SpreadsheetApp.openById(ID_SMS);
  const sheet = ss.getSheetByName(TAB_SMS_SHEET);
  const lastRow = sheet.getLastRow();
  const maxRowsToKeep = 500; // Keep last 500 sms

  // Row 1 is header, so we check if there are more than 21 rows total
  if (lastRow > maxRowsToKeep + 1) {
    const rowsToDelete = lastRow - (maxRowsToKeep + 1);

    // Delete from Row 2 (the oldest records) downward
    sheet.deleteRows(2, rowsToDelete);
    console.log("Cleanup: Deleted " + rowsToDelete + " old SMS rows.");
  }
}

/*Log transaction details to main ledger */
function logToMainLedger(data) {
  try {
    const ss = SpreadsheetApp.openById(ID_LEDGER);
    const sheet = ss.getSheetByName(TAB_LEDGER_MAIN_LEDGER);

    const dr = (data.type === "Expense") ? data.amount : "";
    const cr = (data.type === "Income") ? data.amount : "";

    const formData = [
      data.date,         // Date
      data.mainCategory, // Category
      data.type,         // Type (Income/Expense)
      data.subType,      // SubType
      data.referenceID,  // Ref ID
      data.txnID || "",  // Txn ID
      data.entityName,   // Entity Name
      dr,                // Debit
      cr,                // Credit
      data.paymentMode,  // Mode
      data.status,       // Status
      data.notes         // Notes
    ];

    // 1. Find the REAL next row based on Column B (Main Category) 
    // instead of sheet.getLastRow()
    var lastRow = sheet.getLastRow();
    var targetRow = 2; // Default to row 2 if sheet is empty


    if (lastRow > 0) {
      var columnValues = sheet.getRange(1, 2, lastRow, 1).getValues(); // Look at Column B
      for (var i = columnValues.length - 1; i >= 0; i--) {
        if (columnValues[i][0] !== "") {
          targetRow = i + 2; // Set target to the row after the last data
          break;
        }
      }
    }

    // 2. Use setValues instead of appendRow to respect the targetRow
    sheet.getRange(targetRow, 1, 1, 12).setValues([formData]);

    return true;
  } catch (e) {
    console.error("Ledger Update Error: " + e.toString());
    return false;
  }
}

/**
 * Enhanced logActivity to accept custom usernames
 */
function logActivity(username, action, details, targetSheet) {
  try {
    const ss = SpreadsheetApp.openById(ID_ADMINS);
    let logSheet = ss.getSheetByName(TAB_ADMINS_ACTIVITY_LOGS);

    if (!logSheet) {
      logSheet = ss.insertSheet(TAB_ADMINS_ACTIVITY_LOGS);
      logSheet.appendRow(["Timestamp", "User", "Action", "Details", "Target"]);
    }

    // 2. AUTO-INSERT ROWS Logic
    const maxRows = logSheet.getMaxRows();
    const lastRow = logSheet.getLastRow();

    // If we are within 5 rows of the bottom, add 100 more rows
    if (maxRows - lastRow < 5) {
      logSheet.insertRowsAfter(maxRows, 100);
    }

    // Use the passed username, or fallback to "System/Guest" if null
    const finalUser = username || "System";

    // Append in your requested order
    logSheet.appendRow([
      new Date(),
      finalUser,    // This will now be vgvdev or vgkrish
      action,      // e.g., "LOGOUT" or "LOGIN"
      details,     // e.g., "User performed manual sign-out"
      targetSheet  // e.g., "Vastram" or "Security"
    ]);
  } catch (e) {
    console.error("Logging failed: " + e.message);
  }
}
/*****************************  Guest login ****************************************************************/
function getEventList() {
  const ss = SpreadsheetApp.openById(ID_ADMINS);
  const sheet = ss.getSheetByName("varga");
  const data = sheet.getDataRange().getValues();

  // Assuming Column A is Varga and Column B is Events
  // map(row => row[1]) picks the Events column
  const events = data.slice(1)
    .map(row => row[1])
    .filter(val => val && val.toString().trim() !== "");

  // Return unique events only
  return [...new Set(events)];
}

/**
 * Saves guest info to the 'guest' sheet and returns a login object.
 */
function registerGuest(event, month, guestName, mobile, futureAssociate) {
  try {
    const ss = SpreadsheetApp.openById(ID_PARENTS);
    let guestSheet = ss.getSheetByName(TAB_PARENTS_GUEST);

    // Get the last row number. If it's 1 (header only), start at 1.
    const lastRow = guestSheet.getLastRow();
    const guestId = (lastRow === 1) ? 1 : lastRow;

    const timestamp = "Registered: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

    // Column 1: Event, Column 2: Name, Column 3: Month, Column 4: Mobile, etc.
    // Based on your request, Month goes to Column 3
    guestSheet.appendRow([
      guestId,
      event,           // Col 2
      month,       // Col 3
      guestName,           // Col 4
      String(mobile),          // Col 5
      futureAssociate ? "Yes" : "No",
      timestamp
    ]);

    // FIX: Changed 'name' to 'guestName' to match your parameter
    logActivity(guestName, "GUEST_SIGNUP", `Event: ${event} | ID: ${guestId} | Month: ${month}`, "Guest_Sheet");

    return {
      success: true,
      name: guestName,
      id: guestId,
      varga: "Guest",
      event: event,
      discount: 0,
      balance: 0,
      credit: 0,
      isGuest: true
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/*********************Get Category from admin sheet************************/

/**
 * NEW: Dynamic fetch all main categories from admin sheet
 * This replaces the static VALID_SHEETS array.
 */
function getCategoryMap(forceRefresh = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "full_category_map";
  let cachedMap = forceRefresh ? null : cache.get(cacheKey);

  if (cachedMap) return JSON.parse(cachedMap);

  try {
    const adminSs = SpreadsheetApp.openById(ID_ADMINS);
    const sheet = adminSs.getSheetByName(TAB_ADMINS_ENABLE_CATEGORY);
    const data = sheet.getDataRange().getValues();

    let categoryMap = {};
    let validSheets = [];
    let shortCodes = {}; // New object for codes

    for (var i = 1; i < data.length; i++) {
      let mainCat = String(data[i][0]).trim();    // Column A
      let subCatsRaw = String(data[i][4] || "");  // Column E
      let shortCode = String(data[i][5]).trim();  // Column F (New!)

      if (mainCat) {
        categoryMap[mainCat] = subCatsRaw.split(',').map(s => s.trim()).filter(String);
        validSheets.push(mainCat);
        shortCodes[mainCat] = shortCode || mainCat.substring(0, 2).toUpperCase();
      }
    }

    const configResult = {
      categoryMap: categoryMap,
      validSheets: validSheets,
      shortCodes: shortCodes, // Include in result
      defaultSheet: validSheets.length > 0 ? validSheets[0] : ""
    };

    cache.put(cacheKey, JSON.stringify(configResult), 1500);
    return configResult;
  } catch (e) {
    return { categoryMap: {}, validSheets: [], shortCodes: {}, error: e.toString() };
  }
}


/**************************************Debug functions*******************/
/**
 * Test function to verify the dynamic category mapping.
 * Run this from the Apps Script editor's function dropdown.
 */
function debugCategoryMap() {
  console.log("--- Starting Category Map Debug ---");

  try {
    // We pass 'true' to force a fresh fetch from the sheet, ignoring the cache
    const config = getCategoryMap(true);

    if (config.error) {
      console.error("Error found in mapping: " + config.error);
      return;
    }

    console.log("1. Valid Sheets (Tabs found): " + JSON.stringify(config.validSheets));
    console.log("2. Short Codes (Column F mapping): " + JSON.stringify(config.shortCodes));

    // Check specific categories
    config.validSheets.forEach(cat => {
      const subs = config.categoryMap[cat] || [];
      const code = config.shortCodes[cat];
      console.log(`> Category: [${cat}] | Code: [${code}] | Sub-Cats: ${subs.length} items`);
      if (subs.length > 0) console.log(`  Detail: ${subs.join(" | ")}`);
    });

    console.log("3. Default Sheet set to: " + config.defaultSheet);
    console.log("--- Debug Complete ---");

  } catch (e) {
    console.error("Critical Failure in Debugger: " + e.toString());
  }
}

/**
 * Test function to debug the Inventory loading logic.
 * Helps identify if a category is missing due to dates, status, or tab naming.
 */
function testSendEmailDirectly() {
  const dummySummary = {
    customerId: "123", // Replace with a valid ID from your sheet to test lookup
    customerName: "Test User",
    orderId: "ORD-9999",
    finalTotal: 100,
    previousBalance: 0,
    creditUsed: 0
  };

  const dummyCart = [
    { mainCategory: "Snacks", itemName: "Test Item", quantity: 1, salePrice: 100, uom: "Packets" }
  ];

  sendReceiptEmail(dummySummary, dummyCart);
}

function debugInventoryLoading() {
  console.log("--- Starting Inventory Debug ---");

  try {
    // 1. Check Category Mapping first
    const config = getCategoryMap(true); // Force refresh
    const validSheets = config.validSheets;
    console.log("Registered Categories in Admin:", validSheets.join(", "));

    // 2. Run the actual inventory fetch
    const startTime = new Date().getTime();
    const items = getInventoryData();
    const endTime = new Date().getTime();

    // 3. Analyze Results
    if (items.length === 0) {
      console.warn("RESULT: No items returned! Check 'enable_maincategory' dates and status.");
    } else {
      console.log(`RESULT: Found ${items.length} total items in ${(endTime - startTime) / 1000}s`);

      // Count items per category to see what's actually loading
      const counts = items.reduce((acc, item) => {
        acc[item.mainCategory] = (acc[item.mainCategory] || 0) + 1;
        return acc;
      }, {});

      console.log("Items loaded per Category:");
      Object.keys(counts).forEach(cat => {
        console.log(` > ${cat}: ${counts[cat]} items`);
      });

      // 4. Sample Item Check (Check if image/SKU mapping is correct)
      const sample = items[0];
      console.log("Sample Item Mapping:", {
        Name: sample.itemName,
        SKU: sample.sku,
        Price: sample.salePrice,
        Img: sample.imageUrl.substring(0, 30) + "..."
      });
    }

  } catch (e) {
    console.error("Critical Failure in Inventory Debugger: " + e.toString());
  }
  console.log("--- Debug Complete ---");
}

function auditInventoryData() {
  const allItems = getInventoryData();
  const ss = SpreadsheetApp.openById(ID_INVENTORY);
  const config = getCategoryMap();

  console.log("--- Inventory Audit Started ---");
  console.log("Total items successfully loaded: " + allItems.length);

  config.validSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues().slice(1);
    console.log(`Checking Sheet: ${sheetName} (Total Rows: ${data.length})`);

    data.forEach((row, i) => {
      const rowNum = i + 2;
      const sku = String(row[15]);
      const stock = parseFloat(row[4]) || 0;
      const name = row[2] || "Unnamed Item";

      let issues = [];
      if (!sku || sku === "undefined") issues.push("Missing/Invalid SKU (Col P)");
      if (stock <= 0) issues.push("Zero/Negative Stock (Col E)");

      if (issues.length > 0) {
        console.warn(`Row ${rowNum} [${name}]: SKIPPED due to: ${issues.join(" & ")}`);
      }
    });
  });
}

/* Test payment status code */
function debug_TestPaymentSystem() {

  const testUTR = "24958648213"; // Matches your sample SMS
  const testAmount = 1.00;
  const testUser = "Test Debug User";

  console.log("--- STARTING DEBUG TEST ---");

  // TEST VERIFICATION LOGIC (checkPaymentInLogs)
  console.log("Step : Running verification for UTR: " + testUTR);
  const result = autoCheckPayment(testAmount, testUser, "8213");

  console.log("Verification Result Status: " + result.status);
  console.log("Verification Result Message: " + result.message);

  // 3. FINAL VALIDATION
  if (result.status === "SUCCESS") {
    console.log("✅ TEST PASSED: Payment verified and Ledger update triggered.");
  } else {
    console.error("❌ TEST FAILED: Logic did not find or match the mock SMS.");
  }

  console.log("--- DEBUG TEST COMPLETE ---");
}

function debugCheckSmsSheet() {
  const ss = SpreadsheetApp.openById("1o10_jI39_Pr3QjUoRvz42ZUive08UcKd12aedCWmQTY");
  const sheet = ss.getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();
  
  console.log("Total Rows in Sheet1:", data.length);
  
  // Inspect the last 5 rows
  for (let i = Math.max(1, data.length - 5); i < data.length; i++) {
    console.log(`--- Row ${i + 1} ---`);
    console.log("Sender (Col B):", data[i][1]);
    console.log("Status (Col C):", data[i][2]);
    console.log("Message (Col D):", data[i][3]);
  }
}