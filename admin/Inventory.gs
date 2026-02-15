// --- CONFIGURATION ---
var ID_INVENTORY = "1YDiJsrkNEj4HxDaNlirGIczAX4h7FExpb3XNs9Xu5co";
var VALID_SHEETS = ["dhanyam", "varnam", "vastram", "gavya", "soaps", "snacks"];

/**
 * INVENTORY INTEGRATION
 * Fetches data from the Inventory Spreadsheet across all valid categories
 */
function getInventoryData() {
  try {
    const ss = SpreadsheetApp.openById(ID_INVENTORY);
    let allItems = [];

    VALID_SHEETS.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      // Mapping logic aligned with your Inventory structure
      const items = data.slice(1).map(row => {
        // Safety check: Ensure the row has enough columns to avoid "undefined" errors
        return {
          sku: row[15] ? String(row[15]).trim() : "",      // Column P
          mainCategory: sheetName,
          subCategory: row[1] || "",                      // Column B
          itemName: row[2] || "Unnamed Item",             // Column C
          uom: row[3] || "Unit",                          // Column D
          stock: parseFloat(row[4]) || 0,                 // Column E
          salePrice: parseFloat(row[7]) || 0,              // Column H
          reorderPoint: parseFloat(row[9]) || 0,           // Column J
          status: row[12] || "In stock"                   // Column M
        };
      }).filter(item => item.sku && item.sku !== "" && item.sku !== "undefined");

      allItems = allItems.concat(items);
    });

    return allItems;
  } catch (e) {
    console.log("Inventory Fetch Error: " + e.toString());
    return [];
  }
}

/**
 * Updates stock for a specific SKU in the spreadsheet.
 */
function updateStockValue(sku, newValue) {
  try {
    const ss = SpreadsheetApp.openById(ID_INVENTORY);
    const skuFromUI = (sku || "").toString().trim().toLowerCase();

    if (!skuFromUI) return { success: false, message: "Invalid SKU provided." };

    for (let j = 0; j < VALID_SHEETS.length; j++) {
      const sheetName = VALID_SHEETS[j];
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;

      const data = sheet.getDataRange().getValues();
      const headers = data[0];

      // Dynamic column finding to prevent errors if columns move
      const skuCol = headers.map(h => h.toString().toLowerCase().trim()).indexOf("sku");
      const stockCol = headers.map(h => h.toString().toLowerCase().trim()).indexOf("stock");

      // If headers aren't found in this specific sheet, skip it
      if (skuCol === -1 || stockCol === -1) {
        console.warn("Sheet " + sheetName + " is missing SKU or Stock headers.");
        continue;
      }

      for (let i = 1; i < data.length; i++) {
        const skuInSheet = (data[i][skuCol] || "").toString().trim().toLowerCase();

        if (skuInSheet === skuFromUI) {
          // Write the new value to the specific cell
          sheet.getRange(i + 1, stockCol + 1).setValue(newValue);
          return { success: true, message: "Stock updated in " + sheetName };
        }
      }
    }

    return { success: false, message: "SKU '" + sku + "' not found in any sheet." };
  } catch (e) {
    console.error("Update error: " + e.toString());
    return { success: false, message: "System Error: " + e.message };
  }
}