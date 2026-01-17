function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Grocery Shop")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}


var ID_PARENTS = "1xgcQfWYczXmkwpQsbonkRUraAMvlWExNRtm7D_iSJbk";
var ID_INVENTORY = "1YDiJsrkNEj4HxDaNlirGIczAX4h7FExpb3XNs9Xu5co";
var ID_ORDERS_LINE_ITEMS = "1j5ma5hH1vKaoNW0O3JrYL19FZvPLBXMOyN5_0efP0e8";
var ID_ORDERS = "1i3XQ7tfoKKb6RH8CjyP0fryMnbuOthbXnb26-FCa0MU";

// Sheet Tab Names
var TAB_PARENTS = "main";
var TAB_INVENTORY = "main";
var TAB_LINE_ITEMS = "main";
var TAB_ORDERS = "main";

function getVargas() {
  try {
    const ss = SpreadsheetApp.openById(ID_PARENTS);
    const sheet = ss.getSheetByName(TAB_PARENTS); // Confirmed sheet name
    
    if (!sheet) {
      console.error("Sheet 'parents_list' not found!");
      return ["Error: Sheet not found"];
    }

    const data = sheet.getDataRange().getValues();
    
    // Check if there is actual data beyond the header
    if (data.length < 2) {
      console.warn("Sheet is empty or only has headers");
      return [];
    }

    // Column B is index [1]
    const vargas = data.slice(1).map(row => row[1]); 
    
    // Filter out empty cells and get unique values
    const uniqueVargas = [...new Set(vargas)]
      .filter(v => v && v.toString().trim() !== "")
      .sort();
    
    console.log("Success! Vargas found: " + uniqueVargas);
    return uniqueVargas;

  } catch (e) {
    console.error("Critical Error: " + e.message);
    return ["Error: " + e.message];
  }
}

function getNamesByVarga(varga) {
  const ss = SpreadsheetApp.openById(ID_PARENTS);
  const data = ss.getSheetByName(TAB_PARENTS).getDataRange().getValues();
  
  // Filter by Varga (Col B) and return Name (Col C / index 2)
  return data.filter(row => row[1] === varga).map(row => row[2]);
}

function validateLogin(varga, name, mobile) {
  const ss = SpreadsheetApp.openById(ID_PARENTS);
  const data = ss.getSheetByName(TAB_PARENTS).getDataRange().getValues();
  
  // Col B: Varga, Col C: Name, Col F: Mobile (index 5)
  const user = data.find(row => 
    row[1] === varga && 
    row[2] === name && 
    String(row[5]).trim() === String(mobile).trim()
  );
  
  if (user) {
    return { 
      success: true, 
      id: user[0], 
      name: user[2], 
      email: user[6], // Col G
      discount: user[7] || 0 // Col H
    };
  }
  return { success: false };
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
    // Assuming Image URL is in Column P (index 15). Adjust if it's elsewhere!
    imageUrl: row[15] || "https://via.placeholder.com/150" 
  }));
}

// Log line item when "Add to Cart" is clicked
function logLineItem(lineItemObj) {
  const ss = SpreadsheetApp.openById(ID_ORDERS_LINE_ITEMS);
  const sheet = ss.getSheetByName(TAB_LINE_ITEMS);
  sheet.appendRow([
    "LI-" + Date.now(),
    lineItemObj.orderId,
    lineItemObj.category,
    lineItemObj.itemId,
    lineItemObj.itemName,
    1, // quantity
    lineItemObj.uom,
    lineItemObj.price,
    lineItemObj.subtotal,
    "" // Notes
  ]);
}

function finalizeOrder(orderSummary) {
  const ss = SpreadsheetApp.openById(ID_ORDERS);
  const sheet = ss.getSheetByName(TAB_ORDERS);
  sheet.appendRow([
    "Normal",
    orderSummary.orderId,
    orderSummary.customerId,
    orderSummary.customerName,
    new Date(),
    "Confirmed",
    orderSummary.total,
    "Unpaid",
    ""
  ]);
  return true;
}

