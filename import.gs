const PRODUCT_RECIPE_SHEET_ID = "1SAs4im0YWpGFLDV0T7xImF5xMgycPmfAWQQFDxtQcTo";
const PRODUCT_RECIPE_TAB_NAME = "Panels";

function _extractAllocatedFromStatus_(status) {
  const m = String(status || "").match(/ALLOCATED\s+(\d+)\s+FROM\s+STOCK/i);
  return m ? (parseInt(m[1], 10) || 0) : 0;
}



function updateManufactureHub() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- CONFIGURATION ---
  const ORDER_SHEET_URL = "https://docs.google.com/spreadsheets/d/1KDDVnIZ5oCruCY4nyKp6XVqwWnL7-N6f-pKAD3xXo_U/edit?gid=0#gid=0"; 
  const ORDER_TAB_NAME = "Sheet1"; 
  const PRODUCT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1SAs4im0YWpGFLDV0T7xImF5xMgycPmfAWQQFDxtQcTo/edit?gid=0#gid=0"; 
  const PRODUCT_TAB_NAME = "Panels"; 

  const HUB_TAB_NAME = "Manufacture Hub";
  const COMP_TAB_NAME = "Components Hub";
  const DELIV_TAB_NAME = "Delivery Hub";
  const WORKSHOP_CUSTOMER = "workshop stock";


  // --- OPEN SHEETS ---
  let orderSpreadsheet, productSpreadsheet;
  try { orderSpreadsheet = SpreadsheetApp.openByUrl(ORDER_SHEET_URL); } catch (e) { Browser.msgBox("Error: Check Order Sheet URL."); return; }
  try { productSpreadsheet = SpreadsheetApp.openByUrl(PRODUCT_SHEET_URL); } catch (e) { Browser.msgBox("Error: Check Product Sheet URL."); return; }

  const orderSheet = orderSpreadsheet.getSheetByName(ORDER_TAB_NAME);
  const productSheet = productSpreadsheet.getSheetByName(PRODUCT_TAB_NAME);
  const hubSheet = ss.getSheetByName(HUB_TAB_NAME);
  const compSheet = ss.getSheetByName(COMP_TAB_NAME);
  const delivSheet = ss.getSheetByName(DELIV_TAB_NAME);

  if (!orderSheet || !productSheet || !hubSheet || !compSheet || !delivSheet) {
  Browser.msgBox("Error: Missing a required tab.");
  return;
}


  // --- READ DATA ---
  const allOrderData = orderSheet.getDataRange().getValues(); // Read whole order sheet
  const productData = productSheet.getDataRange().getValues(); // Read whole product sheet
  const productMap = {};
  for (let p = 1; p < productData.length; p++) {
    const prodRow = productData[p];
    const recipeSku = String(prodRow[0] || "").trim();
    if (!recipeSku) continue;
    if (!productMap[recipeSku]) productMap[recipeSku] = [];
    productMap[recipeSku].push(prodRow);
  }

  let rowsForPanels = [];
  let rowsForComponents = [];
  let rowsForDelivery = [];
  
  // Reporting Arrays
  let successCount = 0;
  let missingSkus = [];
  let processedRows = []; // To store row numbers (1-based) that we need to mark "Imported"

  // --- LOOP THROUGH ORDERS ---
  // Start at i=1 to skip header.
  for (let i = 1; i < allOrderData.length; i++) {
    const orderRow = allOrderData[i];
    
    // Safety: If row is completely empty, skip
    if (!orderRow[0]) continue;

    // Import Status is Column H (index 7)
const importStatus = String(orderRow[7] || "").trim();
const st = importStatus.toUpperCase();

// Skip if already imported
if (st.startsWith("IMPORTED") || st === "IMPORTED" || st === "IMPORTED ") continue;

// ✅ NEW: Require approval before import
if (!st.startsWith("APPROVED")) continue;


// ✅ If the status includes "ALLOCATED X FROM STOCK", prefill stages for those units
const allocMatch = importStatus.match(/ALLOCATED\s+(\d+)\s+FROM\s+STOCK/i);
const allocatedUnits = allocMatch ? (parseInt(allocMatch[1], 10) || 0) : 0;


    const orderId = orderRow[0];        // Col A
    const customer = orderRow[2];       // Col C
    
    // NEW: Define the Address variable (Col D / Index 3)
    const address = orderRow[3];        

    // SHIFTED: Everything else moves right by 1
    const productName = orderRow[4];    // Col E
    const orderSku = String(orderRow[5]).trim(); // Col F
    const orderQty = orderRow[6];       // Col G
    let skuFound = false;

    // --- SEARCH PRODUCT MAP ---
    const productRows = productMap[orderSku] || [];
    if (productRows.length > 0) {
      skuFound = true;
      productRows.forEach((prodRow) => {
        const itemName = prodRow[2];      
        const material = String(prodRow[3]).toLowerCase(); 
        const qtyPerUnit = prodRow[6];    
        const totalQty = qtyPerUnit * orderQty;
        const compSKU = String(orderSku).trim().toUpperCase();       

        // Split Logic
        if (material.includes("component") || material.includes("hardware")) {
          const prefillComps = Math.min(totalQty, (Number(qtyPerUnit) || 1) * allocatedUnits);

rowsForComponents.push([
  orderId,
  customer,
  productName,
  itemName,
  compSKU,
  qtyPerUnit,
  totalQty,
  prefillComps,  // ✅ qtyPacked
  "",
  ""
]);
        } else {
          const prefillPanels = Math.min(totalQty, (Number(qtyPerUnit) || 1) * allocatedUnits);

rowsForPanels.push([
  orderId, customer, productName, orderSku.toUpperCase(),
  itemName, prodRow[3], qtyPerUnit, prodRow[4], prodRow[5], prodRow[7], prodRow[8],
  totalQty,
  prefillPanels,   // qtyCut
  prefillPanels,   // qtyProcessed
  prefillPanels,   // qtyEdgeFinish
  prefillPanels,   // qtyPacked
  "", "", ""
]);
        }
      });
    }

       if (skuFound) {
      const isWorkshop = String(customer).trim().toLowerCase() === WORKSHOP_CUSTOMER;
      const qty = Number(orderQty) || 0;

      // ✅ Only create Delivery Hub rows for NON-workshop customers
      if (!isWorkshop) {
        for (let q = 0; q < qty; q++) {
          rowsForDelivery.push([orderId, customer, address, productName, "", "Pending", "", ""]);
        }
      }

      successCount++;
      processedRows.push(i + 1);
    } else {
      if (orderSku !== "") {
        missingSkus.push(`Row ${i + 1}: ${orderSku}`);
      }
    }

  } // ✅ IMPORTANT: this closes the main "for (let i = 1; ...)" loop


  // --- WRITE TO SHEETS (With Ghost Row Protection) ---
  
  // Helper to find the REAL last row (ignoring blank rows at the bottom)
  const getRealLastRow = (sheet) => {
    const data = sheet.getRange("A1:A").getValues(); // Check Column A
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] !== "") return i + 1;
    }
    return 1; // Default to header
  };

  if (rowsForPanels.length > 0) {
    const nextRow = getRealLastRow(hubSheet) + 1;
    hubSheet.getRange(nextRow, 1, rowsForPanels.length, rowsForPanels[0].length).setValues(rowsForPanels);
  }

  if (rowsForComponents.length > 0) {
    const nextRow = getRealLastRow(compSheet) + 1;
    compSheet.getRange(nextRow, 1, rowsForComponents.length, rowsForComponents[0].length).setValues(rowsForComponents);
  }

  if (rowsForDelivery.length > 0) {
  const nextRow = getRealLastRow(delivSheet) + 1;
  delivSheet
    .getRange(nextRow, 1, rowsForDelivery.length, rowsForDelivery[0].length)
    .setValues(rowsForDelivery);
}


 // --- UPDATE STATUS ---
  // Mark successes as "Imported"
  if (processedRows.length > 0) {
    const statusCol = 8; // Column H is the 8th column
    processedRows.forEach(r => {
      const prev = String(orderSheet.getRange(r, statusCol).getValue() || "").trim();
orderSheet.getRange(r, statusCol).setValue(prev ? `IMPORTED | ${prev}` : "IMPORTED");
    });
  }

  // --- FINAL REPORT ---
  if (missingSkus.length > 0) {
    Browser.msgBox(`⚠️ PARTIAL SUCCESS\n\nImported: ${successCount} orders.\nSKIPPED: ${missingSkus.length} orders because SKU not found in Product Data.\n\nMissing SKUs:\n${missingSkus.join('\n')}`);
  } else if (successCount > 0) {
    Browser.msgBox(`SUCCESS! Imported ${successCount} orders.`);
  } else {
    Browser.msgBox("No new orders found to import.");
  }
}

function importApprovedShopifyRow_(shopifyRowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hubSheet = ss.getSheetByName("Manufacture Hub");
  const compSheet = ss.getSheetByName("Components Hub");
  const delivSheet = ss.getSheetByName("Delivery Hub");
  if (!hubSheet || !compSheet || !delivSheet) throw new Error("Missing Manufacture/Components/Delivery tab(s) in master.");

  // Open Shopify sheet
  const orderSS = SpreadsheetApp.openById(SHOPIFY_ORDERS_SHEET_ID);
  const orderSheet = orderSS.getSheetByName(SHOPIFY_ORDERS_TAB_NAME);
  if (!orderSheet) throw new Error("Shopify sheet tab not found.");

  // Open product recipe sheet
  const prodSS = SpreadsheetApp.openById(PRODUCT_RECIPE_SHEET_ID);
  const productSheet = prodSS.getSheetByName(PRODUCT_RECIPE_TAB_NAME);
  if (!productSheet) throw new Error("Product recipe tab not found.");

  const lastCol = orderSheet.getLastColumn();
  const headers = orderSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => _normTxt(h).toLowerCase());
  const hm = _headerMap_(headers);

  const idxOrderId = hm["order id"];
  const idxCustomer = hm["customer"];
  const idxAddress = hm["shipping address"];
  const idxProductName = hm["product name"];
  const idxSku = hm["product code"];
  const idxQty = hm["quantity ordered"];
  const idxStatus = hm["import status"];

  if ([idxOrderId, idxCustomer, idxAddress, idxProductName, idxSku, idxQty, idxStatus].some(x => x == null)) {
    throw new Error("Shopify sheet headers missing required columns.");
  }

  const row = orderSheet.getRange(shopifyRowIndex, 1, 1, lastCol).getValues()[0];
  const orderId = _normTxt(row[idxOrderId]);
  const customer = _normTxt(row[idxCustomer]);
  const address = _normTxt(row[idxAddress]);
  const productName = _normTxt(row[idxProductName]);
  const orderSku = _normTxt(row[idxSku]);
  const orderQty = Number(row[idxQty]) || 0;
  const status = _normTxt(row[idxStatus]);
  const stUpper = status.toUpperCase().trim();

  if (stUpper.startsWith("IMPORTED")) return "Already imported";
  if (!stUpper.startsWith("APPROVED")) throw new Error("This line is not APPROVED yet.");
  if (!orderId || !orderSku || orderQty <= 0) throw new Error("Missing Order ID / SKU / Qty.");

  // Prefill based on allocation note
  const allocatedUnits = _extractAllocatedFromStatus_(status);

  const productData = productSheet.getDataRange().getValues();
  const productMap = {};
  for (let p = 1; p < productData.length; p++) {
    const prodRow = productData[p];
    const recipeSku = String(prodRow[0] || "").trim();
    if (!recipeSku) continue;
    if (!productMap[recipeSku]) productMap[recipeSku] = [];
    productMap[recipeSku].push(prodRow);
  }
  const rowsForPanels = [];
  const rowsForComponents = [];
  const rowsForDelivery = [];

  let skuFound = false;
  const existingPanelKeys = new Set();
  const existingCompKeys = new Set();
  let existingDeliveryCount = 0;

  // --- DEDUPLICATION TRACKERS (Prevent double import) ---
  const seenPanels = new Set();
  const seenComps = new Set();

  const buildKey = (parts) => parts.map(p => String(p ?? "").trim().toLowerCase()).join("|");

  const hubData = hubSheet.getDataRange().getValues();
  for (let i = 1; i < hubData.length; i++) {
    const row = hubData[i];
    if (_normTxt(row[0]) !== orderId) continue;
    const key = buildKey([row[0], row[2], row[4], row[5], row[7], row[8]]);
    existingPanelKeys.add(key);
  }

  const compData = compSheet.getDataRange().getValues();
  for (let i = 1; i < compData.length; i++) {
    const row = compData[i];
    if (_normTxt(row[0]) !== orderId) continue;
    const key = buildKey([row[0], row[2], row[3], row[4]]);
    existingCompKeys.add(key);
  }

  const delivData = delivSheet.getDataRange().getValues();
  for (let i = 1; i < delivData.length; i++) {
    const row = delivData[i];
    if (_normTxt(row[0]) !== orderId) continue;
    if (_normTxt(row[3]) !== productName) continue;
    existingDeliveryCount++;
  }

  const productRows = productMap[orderSku] || [];
  productRows.forEach((prodRow) => {
    skuFound = true;
    const itemName = prodRow[2];
    const material = String(prodRow[3] || "").toLowerCase();
    const qtyPerUnit = Number(prodRow[6]) || 1;
    const totalQty = qtyPerUnit * orderQty;

    // Unique ID for this part: ItemName + Material + Length + Width
    const partSignature = itemName + "|" + material + "|" + prodRow[7] + "|" + prodRow[8];
    const panelKey = buildKey([orderId, productName, itemName, prodRow[3], prodRow[4], prodRow[5]]);
    const compKey = buildKey([orderId, productName, itemName, orderSku]);

    if (material.includes("component") || material.includes("hardware")) {
      // IGNORE DUPLICATES
      if (seenComps.has(partSignature)) return;
      seenComps.add(partSignature);
      if (existingCompKeys.has(compKey)) return;

      const prefillComps = Math.min(totalQty, qtyPerUnit * allocatedUnits);
      rowsForComponents.push([
        orderId, customer, productName,
        itemName,
        String(orderSku).trim().toUpperCase(),
        qtyPerUnit,
        totalQty,
        prefillComps, // qtyPacked
        "", ""
      ]);
    } else {
      // IGNORE DUPLICATES
      if (seenPanels.has(partSignature)) return;
      seenPanels.add(partSignature);
      if (existingPanelKeys.has(panelKey)) return;

      const prefillPanels = Math.min(totalQty, qtyPerUnit * allocatedUnits);
      rowsForPanels.push([
        orderId, customer, productName, String(orderSku).trim().toUpperCase(),
        itemName, prodRow[3], qtyPerUnit, prodRow[4], prodRow[5], prodRow[7], prodRow[8],
        totalQty,
        prefillPanels, // cut
        prefillPanels, // processed
        prefillPanels, // edge
        prefillPanels, // packed
        "", "", ""
      ]);
    }
  });

  if (!skuFound) throw new Error("SKU not found in product recipe sheet: " + orderSku);

  // Delivery rows (only non-workshop)
  const isWorkshop = String(customer).trim().toLowerCase() === "workshop stock";
  if (!isWorkshop) {
    const remainingDelivery = Math.max(0, orderQty - existingDeliveryCount);
    for (let q = 0; q < remainingDelivery; q++) {
      rowsForDelivery.push([orderId, customer, address, productName, "", "Pending", "", ""]);
    }
  }

  // Ghost-row-safe append
  const getRealLastRow = (sheet) => {
    const colA = sheet.getRange("A1:A").getValues();
    for (let i = colA.length - 1; i >= 0; i--) {
      if (colA[i][0] !== "") return i + 1;
    }
    return 1;
  };

  if (rowsForPanels.length) {
    const nextRow = getRealLastRow(hubSheet) + 1;
    hubSheet.getRange(nextRow, 1, rowsForPanels.length, rowsForPanels[0].length).setValues(rowsForPanels);
  }

  if (rowsForComponents.length) {
    const nextRow = getRealLastRow(compSheet) + 1;
    compSheet.getRange(nextRow, 1, rowsForComponents.length, rowsForComponents[0].length).setValues(rowsForComponents);
  }

  if (rowsForDelivery.length) {
    const nextRow = getRealLastRow(delivSheet) + 1;
    delivSheet.getRange(nextRow, 1, rowsForDelivery.length, rowsForDelivery[0].length).setValues(rowsForDelivery);
  }

  // Mark imported BUT preserve notes
  const cleaned = status.replace(/^IMPORTED\s*\|?/i, "").trim();
  const nextStatus = cleaned ? `IMPORTED | ${cleaned}` : "IMPORTED";
  orderSheet.getRange(shopifyRowIndex, idxStatus + 1).setValue(nextStatus);

  return "SUCCESS";
}

// HELPER: Check for Pending Orders (For Notification Badge)
function hasPendingOrders() {
  try {
    const sh = _getShopifyOrdersSheet();
    const data = sh.getDataRange().getValues();
    
    // Find "Import Status" column index
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const idx = headers.indexOf("import status");
    if (idx === -1) return false;

    // Check rows (skip header)
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][idx]).toUpperCase().trim();
      // If status is NOT "IMPORTED" and NOT "CANCELLED", it is pending
      if (!status.startsWith("IMPORTED") && !status.startsWith("CANCELLED")) {
        return true;
      }
    }
    return false;
  } catch (e) {
    return false;
  }
}
