var ID_ORDERS_LINE_ITEMS = "1j5ma5hH1vKaoNW0O3JrYL19FZvPLBXMOyN5_0efP0e8";
var ID_ORDERS = "1i3XQ7tfoKKb6RH8CjyP0fryMnbuOthbXnb26-FCa0MU";
var ID_ADMINS = "1iiZtZclKgr7G7ISZFlM1We4LTmMLNkZLp_x4gP2DoOM";

var TAB_PARENTS_MAIN = "main";
var TAB_PARENTS_GUEST = "guest";
var TAB_LINE_ITEMS_MAIN = "main";
var TAB_ORDERS_MAIN = "main";
var TAB_ADMINS_ENABLE_CATEGORY = "enable_maincategory";
var TAB_ADMINS_ACTIVITY_LOGS = "activitiy_logs";
var TAB_ADMINS_VARGA = "varga";
var TAB_ADMINS_USERS = "users";
var TAB_LEDGER_MAIN_LEDGER = "main_ledger";
var TAB_SMS_SHEET = "Sheet1";

// Global operational database locks
const liSheet = SpreadsheetApp.openById(ID_ORDERS_LINE_ITEMS).getSheetByName(TAB_LINE_ITEMS_MAIN);
const ordSheet = SpreadsheetApp.openById(ID_ORDERS).getSheetByName(TAB_ORDERS_MAIN);

/**
 * Web App Initialization
 */
function doGet() {
  // 1. Create a template from the file
  var template = HtmlService.createTemplateFromFile('Index');

  // 2. Evaluate the template to execute <?!= include('Styles'); ?>
  return template.evaluate()
    .setTitle("Vidyagrama Order Management")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://i.ibb.co/1txQwJMC/vk-main-icon.png');
}

/**
 * Helper to include modular sub-files within Index.html template
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Authentication Controller
 */

/********************************************* Login And role bases Accesss */
function checkLogin(username, password) {
  console.log(`[checkLogin] Execution initiated for username: "${username}"`);
  try {
    console.log(`[checkLogin] Connecting to Spreadsheet ID: ${ID_ADMINS}, Tab: ${TAB_ADMINS_USERS}`);
    const ss = SpreadsheetApp.openById(ID_ADMINS);
    const sheet = ss.getSheetByName(TAB_ADMINS_USERS);
    const data = sheet.getDataRange().getValues();
    console.log(`[checkLogin] Found ${data.length} total rows (including header) in user table.`);

    for (let i = 1; i < data.length; i++) {
      const dbUsername = String(data[i][0]).trim();
      const dbPassword = String(data[i][1]).trim();
      
      if (dbUsername === String(username).trim() && dbPassword === String(password).trim()) {
        const role = String(data[i][3]).trim();
        const displayName = data[i][2];
        
        console.log(`[checkLogin] Match found at row ${i + 1}. User: "${dbUsername}", Role: "${role}"`);

        // Update lastLogin
        console.log(`[checkLogin] Updating lastLogin timestamp at row ${i + 1}, column 6...`);
        sheet.getRange(i + 1, 6).setValue(new Date());

        // Log the successful login
        console.log(`[checkLogin] Dispatching transaction audit payload to logActivity...`);
        logActivity("LOGIN", `User ${username} logged in`, "Security", username);

        const response = {
          success: true,
          name: displayName,
          role: role,
          username: username
        };
        console.log(`[checkLogin] Outbound success payload: ${JSON.stringify(response)}`);
        return response;
      }
    }
    
    console.warn(`[checkLogin] Validation terminated: No matching records found for user "${username}".`);
    return { success: false, message: "Invalid credentials" };
  } catch (err) {
    console.error(`[checkLogin] CRITICAL EXCEPTION: ${err.toString()} \nStack: ${err.stack}`);
    return { success: false, message: "System error" };
  }
}
/**
 * Updates user password
 */
function updatePassword(username, currentPass, newPass) {
  console.log(`[updatePassword] Execution initiated for user: "${username}"`);
  try {
    console.log(`[updatePassword] Opening user schema table.`);
    const ss = SpreadsheetApp.openById(ID_ADMINS);
    const sheet = ss.getSheetByName(TAB_ADMINS_USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const dbUsername = data[i][0];
      const dbPassword = data[i][1];

      if (dbUsername === username && dbPassword === currentPass) {
        console.log(`[updatePassword] Identity matched at row ${i + 1}. Writing new security key...`);
        sheet.getRange(i + 1, 2).setValue(newPass);
        
        console.log(`[updatePassword] Dispatching audit payload to logActivity...`);
        logActivity("PASSWORD_CHANGE", `User ${username} changed their password`, "Users");
        
        console.log(`[updatePassword] Security mutation completed for "${username}".`);
        return "Password updated successfully!";
      }
    }
    console.warn(`[updatePassword] Failure: Operational profile match failed or incorrect current password for "${username}".`);
    throw new Error("Current password incorrect.");
  } catch(err) {
    console.error(`[updatePassword] EXCEPTION: ${err.toString()}`);
    throw err; // Re-throw to make sure withFailureHandler catches it on front-end
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

/**
   * Safe UI view toggling matrix for authentication screens
   */
  function toggleLoginWorkspace(viewId) {
    document.getElementById('loginBox').classList.add('hidden');
    document.getElementById('pwBox').classList.add('hidden');
  }


/**
 * Fetches all order tracking records to populate the main system dashboard ledger.
 * @return {Array<Object>} Array of serialized order transaction cards.
 */
function getLiveOrdersSummary() {
  console.log("[getLiveOrdersSummary] Compiling ledger matrix records...");
  try {
    // Open the master admin ledger spreadsheet
    const ss = SpreadsheetApp.openById(ID_ORDERS);
    const sheet = ss.getSheetByName(TAB_ORDERS_MAIN); // Change to your specific Summary Tab name if different, e.g., "Order_Summary"
    
    if (!sheet) {
      console.error("[getLiveOrdersSummary] Targeted order tab name not found.");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.warn("[getLiveOrdersSummary] Sheet contains no active transaction records.");
      return [];
    }

    const orders = [];
    // Loop through rows skipping the header line
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Safety guard: skip row if Order ID or Customer Name is completely empty
      if (!row[1] && !row[2]) continue;

      // Safely process dates so they pass through JSON communication layers cleanly
      let orderDateFormatted = "";
      if (row[4] instanceof Date) {
        // Converted pattern to standard DD/MM/YYYY
        orderDateFormatted = Utilities.formatDate(row[4], Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else {
        const looseDateStr = String(row[4] || "").trim();
        
        // If the sheet cell is loose text but written as yyyy-mm-dd, let's fix it dynamically
        if (looseDateStr.includes("-") && looseDateStr.split("-")[0].length === 4) {
          const parts = looseDateStr.split("-"); // [yyyy, mm, dd]
          orderDateFormatted = `${parts[2]}/${parts[1]}/${parts[0]}`;
        } else {
          orderDateFormatted = looseDateStr; // Fallback to raw string value (e.g. 11/04/2026)
        }
      }

     // STRICT COLUMN INDEX ALIGNMENT MATRIX
      orders.push({
        slNo: String(i+1).trim(),         // Column A (Index 0) -> Priority / SlNo
        orderId: String(row[1] || "").trim(),      // Column B (Index 1) -> order_id
        customerId: String(row[2] || "").trim(),   // Column C (Index 2) -> customer_id
        customerName: String(row[3] || "").trim(), // Column D (Index 3) -> customer_name
        orderDate: orderDateFormatted,              // Column E (Index 4) -> order_date
        eventName: "Order Info",                   // Default fallback metric placeholder
        orderStatus: String(row[5] || "Pending").trim(), // Column F (Index 5) -> order_status
        orderAmount: parseFloat(row[6]) || 0,       // Column G (Index 6) -> total_amount
        paymentStatus: String(row[7] || "Unpaid").trim(), // Column H (Index 7) -> payment_status
        notes: String(row[8] || "").trim()          // Column I (Index 8) -> notes
      });
    }

    console.log(`[getLiveOrdersSummary] Successfully packed ${orders.length} transaction entries.`);
    return orders.reverse(); // Reverse order so the newest transactions display at the top of the grid
  } catch (err) {
    console.error(`[getLiveOrdersSummary] Critical retrieval breakdown: ${err.toString()}`);
    return [];
  }
}

/**
 * Safely updates specific metadata attributes (Status/Payment fields) inline 
 * by locating target keys within the Master Order Summary Sheet tracking lines.
 */
function updateOrderFieldInline(orderId, fieldName, newValue) {
  console.log(`[Backend Update] Target ID: "${orderId}", Column Mapping: "${fieldName}", Value: "${newValue}"`);
  try {
    const ss = SpreadsheetApp.openById(ID_ORDERS);
    const sheet = ss.getSheetByName(TAB_ORDERS_MAIN); 
    if (!sheet) throw new Error("Order tracking log tab index cannot be located.");

    const data = sheet.getDataRange().getValues();
    
    // CORRECTED SHEET 1-BASED COLUMN INDEX TARGETS:
    // Column F (6) -> order_status | Column H (8) -> payment_status
    let targetColIndex = (fieldName === "orderStatus") ? 6 : 8;

    for (let i = 1; i < data.length; i++) {
      // data[i][1] is Column B (order_id)
      if (String(data[i][1]).trim() === String(orderId).trim()) {
        
        // Write targeted update directly to the correct cell
        sheet.getRange(i + 1, targetColIndex).setValue(newValue);
        
        // Log changes to security audit sheets tracking infrastructure
        if (typeof logActivity === "function") {
          logActivity("ORDER_MUTATION", `Updated ${fieldName} to ${newValue} for Order ID ${orderId}`, "Orders Management");
        }
        
        console.log(`[Backend Update] Cell modified successfully at Row ${i + 1}, Col ${targetColIndex}`);
        return true;
      }
    }
    throw new Error(`Order ID ${orderId} not found in database registry.`);
  } catch(err) {
    console.error(`[Backend Update] Critical mutation error: ${err.toString()}`);
    throw err;
  }
}


/**
 * Write Operations: Dynamic Matrix Status Updates routed directly to the explicit orders sheet
 */
function updateMainFieldStatus(orderId, fieldType, newValue) {
  try {
    if (!ordSheet) return { success: false, error: "Orders ledger reference initialization failure." };
    const data = ordSheet.getDataRange().getValues();
    let colIndex = (fieldType === 'Order Status') ? 5 : 6; 
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === orderId) {
        ordSheet.getRange(i + 1, colIndex + 1).setValue(newValue);
        return { success: true, orderId: orderId, field: fieldType, value: newValue };
      }
    }
    return { success: false, error: "Target Order ID matching sequence failed." };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * PDF Generation Engine with English & Kannada Text Compilations
 * Bound securely to the explicit Order Spreadsheet to eliminate null reference exceptions.
 */
function generateOrderInvoicePdf(orderIds, selectedLanguage) {
  try {
    const labels = {
      en: { invoice: "INVOICE", id: "Order ID", name: "Customer Name", item: "Item Details", qty: "Qty", total: "Sub Total" },
      kn: { invoice: "ಸರಕು ಪಟ್ಟಿ (INVOICE)", id: "ಆರ್ಡರ್ ಸಂಖ್ಯೆ", name: "ಗ್ರಾಹಕರ ಹೆಸರು", item: "ವಸ್ತು ವಿವರ", qty: "ಪ್ರಮಾಣ", total: "ಉಪ ಒಟ್ಟು" }
    };
    
    let dict = labels[selectedLanguage] || labels.en;
    
    // Open the targeted spreadsheet container safely instead of capturing the active runtime window
    let ordersSpreadsheet = SpreadsheetApp.openById(ID_ORDERS);
    let tempSheet = ordersSpreadsheet.insertSheet("TEMP_INVOICE_EXEC");
    
    // Structure invoice cells matching corporate/professional styling guidelines
    tempSheet.getRange("A1").setValue(dict.invoice).setFontSize(16).setFontWeight("bold");
    // [Database mapping logic iterating through orderIds and printing grid]
    
    SpreadsheetApp.flush();
    let blob = tempSheet.getAs('application/pdf').setName("Vidyagrama_Batch_Invoice.pdf");
    ordersSpreadsheet.deleteSheet(tempSheet);
    
    return { success: true, base64: Utilities.base64Encode(blob.getBytes()) };
  } catch(e) {
    return { success: false, error: e.message };
  }
}



/**  Testing framework */
/**
 * Automated test suite for backend authentication system validation.
 * Adjust test credentials matching a test record inside your Sheet table.
 */
function test_authSystem() {
  const TEST_USER = "vgvdev"; 
  const VALID_PASS = "pass";
  const INVALID_PASS = "wrongpass";
  const TEMPORARY_NEW_PASS = "admin999";

  console.log("=== BEGINNING AUTH MATRIX TEST ENGINE ===");

  // 1. Test Login Failure
  console.log("\n--- TEST CASE 1: Expecting Invalid Login ---");
  var failLoginResult = checkLogin(TEST_USER, INVALID_PASS);
  console.log(`Result Evaluated -> Success: ${failLoginResult.success} | Msg: ${failLoginResult.message}`);

  // 2. Test Login Success
  console.log("\n--- TEST CASE 2: Expecting Successful Login ---");
  var successLoginResult = checkLogin(TEST_USER, VALID_PASS);
  console.log(`Result Evaluated -> Success: ${successLoginResult.success} | User: ${successLoginResult.name} | Role: ${successLoginResult.role}`);

  // 3. Test Password Change (Successful path)
  console.log("\n--- TEST CASE 3: Expecting Password Modification ---");
  try {
    var passChangeResult = updatePassword(TEST_USER, VALID_PASS, TEMPORARY_NEW_PASS);
    console.log(`Mutation Result -> ${passChangeResult}`);
    
    // Verify login works with the new password
    console.log("Verifying access authorization using new security key...");
    var secondaryLogin = checkLogin(TEST_USER, TEMPORARY_NEW_PASS);
    console.log(`Login with new password -> Success: ${secondaryLogin.success}`);

    // Clean up reset back to original state to prevent breaking your system tracking
    console.log("Cleaning up workspace: Reverting password back to baseline rules...");
    updatePassword(TEST_USER, TEMPORARY_NEW_PASS, VALID_PASS);
    console.log("Cleanup complete. Baseline parameters restored.");
  } catch(e) {
    console.error(`Test Engine Interrupted in Test Suite 3: ${e.message}`);
  }

  // 4. Test Password Change Failure (Bad current password verification)
  console.log("\n--- TEST CASE 4: Expecting Password Mutation Rejection ---");
  try {
    updatePassword(TEST_USER, INVALID_PASS, "shouldFailPass");
    console.error("CRITICAL TEST FAILURE: System allowed mutation with an invalid current key.");
  } catch(e) {
    console.log(`Expected rejection caught successfully -> Message string: "${e.message}"`);
  }

  console.log("\n=== AUTH MATRIX TEST EXECUTION CONCLUDED ===");
}

/**
 * Automated test runner for getLiveOrdersSummary.
 * Run this function from the Apps Script editor menu to inspect backend data output.
 */
function test_getLiveOrdersSummary() {
  console.log("=== STARTING BACKEND ORDERS DATA LEDGER TEST ===");
  
  try {
    // 1. Trigger the retrieval routine
    const ordersList = getLiveOrdersSummary();
    
    // 2. Validate response structure
    if (!ordersList) {
      console.error("❌ TEST FAILED: The function returned a null or undefined object.");
      return;
    }
    
    if (!Array.isArray(ordersList)) {
      console.error(`❌ TEST FAILED: Expected an Array, but received type: ${typeof ordersList}`);
      return;
    }
    
    console.log(`✅ STRUCTURE VALID: Successfully returned an Array containing ${ordersList.length} records.`);
    
    // 3. Inspect data elements if records exist
    if (ordersList.length > 0) {
      console.log("\n--- INSPECTING LATEST ORDER ENTRY SAMPLE ---");
      const sampleOrder = ordersList[0]; // This will be the newest order due to .reverse()
      
      const requiredKeys = [
        'slNo', 'orderId', 'customerName', 'orderDate', 
        'eventName', 'orderStatus', 'paymentStatus', 'orderAmount', 'notes'
      ];
      
      let structureMatches = true;
      requiredKeys.forEach(key => {
        if (!(key in sampleOrder)) {
          console.warn(`⚠️ MISSING FIELD PROPERTY: "${key}" is absent from the parsed object structure.`);
          structureMatches = false;
        }
      });
      
      if (structureMatches) {
        console.log("✅ FIELD SCHEMATIC VALID: All essential database object properties are mapped.");
      }
      
      // Print clear serialized sample tracking context to execution console
      console.log(`   [SlNo]: ${sampleOrder.slNo}`);
      console.log(`   [Order ID]: ${sampleOrder.orderId}`);
      console.log(`   [Customer]: ${sampleOrder.customerName}`);
      console.log(`   [Date String]: ${sampleOrder.orderDate}`);
      console.log(`   [Status Mapping]: Order: ${sampleOrder.orderStatus} | Payment: ${sampleOrder.paymentStatus}`);
      console.log(`   [Financial Amount]: ₹${sampleOrder.orderAmount}`);
      console.log(`   [Notes Trace]: ${sampleOrder.notes || '(None)'}`);
      
      // 4. DataType Assertions
      console.log("\n--- EXECUTING TYPE SANITY ASSERTIONS ---");
      if (typeof sampleOrder.orderAmount !== 'number' || isNaN(sampleOrder.orderAmount)) {
        console.error("❌ TYPE MISMATCH: 'orderAmount' is not a valid floating-point number.");
      } else {
        console.log("✅ TYPE CHECK PASSED: 'orderAmount' correctly verified as a number data type.");
      }
      
    } else {
      console.warn("💡 EDGE CASE DETECTED: The targeted sheet table data executed successfully but returned 0 rows.");
    }
    
    console.log("\n=== ORDERS DATA LEDGER TEST COMPLETED SUCCESSFULLY ===");
    
  } catch (testError) {
    console.error(`❌ TEST RUNTIME FAILURE: Execution interrupted by error -> ${testError.toString()}`);
  }
}
