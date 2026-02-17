var ID_ORDER_LINE_ITEMS = "1j5ma5hH1vKaoNW0O3JrYL19FZvPLBXMOyN5_0efP0e8";
var ID_ORDER_SUMMARY = "1i3XQ7tfoKKb6RH8CjyP0fryMnbuOthbXnb26-FCa0MU";

var SUMMARY_SHEET_NAME = "main";
var LINE_ITEMS_SHEET_NAME = "main"; // Check if this should be "Sheet1" or "items"

/** --- FETCH ALL ORDERS --- **/
function getOrdersData() {
  try {
    const ss = SpreadsheetApp.openById(ID_ORDER_SUMMARY);
    const sheet = ss.getSheetByName(SUMMARY_SHEET_NAME) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return [];

    const headers = data[0];
    let orders = [];

    for (let i = 1; i < data.length; i++) {
      let obj = {};
      // Inside your for loop in getOrdersData...
      headers.forEach((h, index) => {
        let val = data[i][index];
        
        // Check if it's a valid Date object from Google Sheets
        if (val instanceof Date) {
          // Check if the date is actually valid before formatting
          if (!isNaN(val.getTime())) {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd-MMM-yyyy");
          } else {
            val = "N/A";
          }
        } else if (val === "" || val === undefined) {
          val = "";
        }

        let key = h ? h.toString().trim() : "col_" + index;
        obj[key] = val;
      });
      orders.push(obj);
    }

    console.log("Found " + orders.length + " rows. Sending to browser...");
    return orders.reverse();

  } catch (e) {
    console.error("Fetch failed: " + e.toString());
    return [];
  }
}

/** --- FETCH ITEMS FOR MODAL/INVOICE --- **/
function getOrderInvoiceData(orderId) {
  const ssSummary = SpreadsheetApp.openById(ID_ORDER_SUMMARY);
  const ssItems = SpreadsheetApp.openById(ID_ORDER_LINE_ITEMS);

  const summarySheet = ssSummary.getSheetByName(SUMMARY_SHEET_NAME);
  const itemsSheet = ssItems.getSheetByName(LINE_ITEMS_SHEET_NAME);

  const summaryData = summarySheet.getDataRange().getValues();
  const itemsData = itemsSheet.getDataRange().getValues();

  const summaryHeaders = summaryData[0];
  const itemsHeaders = itemsData[0];

  const idxSummaryId = summaryHeaders.indexOf("order_id");
  const orderRow = summaryData.find(row => row[idxSummaryId] == orderId);

  if (!orderRow) return { error: "Order not found" };

  let summaryObj = {};
  summaryHeaders.forEach((h, i) => summaryObj[h.toString().trim()] = orderRow[i]);

  const idxOrderIdItems = itemsHeaders.indexOf("order_id");
  const lineItems = itemsData.filter(row => row[idxOrderIdItems] == orderId).map(row => {
    let obj = {};
    itemsHeaders.forEach((h, i) => obj[h.toString().trim()] = row[i]);
    return obj;
  });

  return { summary: summaryObj, items: lineItems };
}

/** --- UNIFIED UPDATE (Status or Payment) --- **/
function updateOrderField(orderId, type, newVal) {
  const ss = SpreadsheetApp.openById(ID_ORDER_SUMMARY);
  const sheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idxOrderID = headers.indexOf("order_id");
  // Determine which column to update based on 'type' from Scripts.html
  const colName = (type === 'status') ? "order_status" : "payment_status";
  const idxTarget = headers.indexOf(colName);

  if (idxOrderID === -1 || idxTarget === -1) {
    return { success: false, message: "Column " + colName + " not found" };
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][idxOrderID] == orderId) {
      sheet.getRange(i + 1, idxTarget + 1).setValue(newVal);
      return { success: true };
    }
  }
  return { success: false, message: "Order ID not found" };
}

/** --- DELETE ORDER (From Both Sheets) --- **/
function deleteOrderData(orderId) {
  try {
    // 1. Delete from Summary Sheet
    const ssSummary = SpreadsheetApp.openById(ID_ORDER_SUMMARY);
    const summarySheet = ssSummary.getSheetByName(SUMMARY_SHEET_NAME);
    const summaryData = summarySheet.getDataRange().getValues();
    const idxSumId = summaryData[0].indexOf("order_id");

    for (let i = 1; i < summaryData.length; i++) {
      if (summaryData[i][idxSumId] == orderId) {
        summarySheet.deleteRow(i + 1);
        break; 
      }
    }

    // 2. Delete from Line Items Sheet
    const ssItems = SpreadsheetApp.openById(ID_ORDER_LINE_ITEMS);
    const itemsSheet = ssItems.getSheetByName(LINE_ITEMS_SHEET_NAME);
    const itemsData = itemsSheet.getDataRange().getValues();
    const idxItemOrderId = itemsData[0].indexOf("order_id");

    // We go backwards when deleting multiple rows to avoid index shifting
    for (let j = itemsData.length - 1; j >= 1; j--) {
      if (itemsData[j][idxItemOrderId] == orderId) {
        itemsSheet.deleteRow(j + 1);
      }
    }

    return { success: true, message: "Order and items deleted successfully" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}