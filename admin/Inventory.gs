// --- CONFIGURATION ---
// Ensure these IDs match your Inventory Spreadsheet
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
      const items = data.slice(1).map(row => ({
        sku: String(row[15]),          // Column P
        mainCategory: sheetName,
        subCategory: row[1] || "",     // Column B
        itemName: row[2],              // Column C
        uom: row[3],                   // Column D
        stock: parseFloat(row[4]) || 0, // Column E
        salePrice: parseFloat(row[7]) || 0, // Column H
        reorderPoint: parseFloat(row[9]) || 0, // Column J (for low stock alerts)
        status: row[12] || "In stock"   // Column M
      })).filter(item => item.sku && item.sku !== "undefined");

      allItems = allItems.concat(items);
    });

    return allItems;
  } catch (e) {
    console.log("Inventory Fetch Error: " + e.toString());
    return [];
  }
}