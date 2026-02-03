const SHOPIFY_ORDERS_SHEET_ID = "1KDDVnIZ5oCruCY4nyKp6XVqwWnL7-N6f-pKAD3xXo_U";
const SHOPIFY_ORDERS_TAB_NAME = "Sheet1";



// 2. FETCH HIERARCHY DATA (Safe Version)


function getDataTree() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hubSheet = ss.getSheetByName("Manufacture Hub");
  const compSheet = ss.getSheetByName("Components Hub");
  const delivSheet = ss.getSheetByName("Delivery Hub"); // Check this tab name exactly!
  const tz = Session.getScriptTimeZone();

  // 1. READ PANELS
  if (!hubSheet) throw new Error("CRITICAL: Manufacture Hub tab not found!");
  const panelData = hubSheet.getDataRange().getValues();
  panelData.shift(); 
  
  // 2. READ COMPONENTS
  let compData = [];
  if (compSheet && compSheet.getLastRow() > 1) {
    compData = compSheet.getDataRange().getValues();
    compData.shift(); 
  }

  // 3. READ DELIVERY (With Safety Check)
  let delivData = [];
  if (delivSheet && delivSheet.getLastRow() > 1) {
    delivData = delivSheet.getDataRange().getValues();
    delivData.shift();
  } else {
    // If missing, we just log it but don't crash the app
    console.log("Warning: Delivery Hub tab missing or empty.");
  }

  let tree = {};

  const norm = v => String(v ?? "").replace(/\u00A0/g, " ").trim();

  // --- PROCESS PANELS ---
  panelData.forEach((row, index) => {
    const orderId = norm(row[0]);
    const customer = row[1];
    let product = norm(row[2]);
if (!orderId) return;
if (!product) product = norm(row[3] || "Unknown Product");


    if (!tree[orderId]) tree[orderId] = { id: orderId, customer: customer, products: {}, delivery: { bucket: {}, rooms: {} } };
    if (!tree[orderId].products[product]) tree[orderId].products[product] = { panels: [], components: [] };

    tree[orderId].products[product].panels.push({
  rowIndex: index + 2,
  panelName: row[4],
  material: row[5],
  qtyPerUnit: Number(row[6]) || 1,       // ✅ NEW (needed for unit maths)
  qtyOrder: Number(row[11]) || 0,        // total panels required for this part
  qtyCut: Number(row[12]) || 0,
  qtyProcessed: Number(row[13]) || 0,
  qtyEdgeFinish: Number(row[14]) || 0,
  qtyPacked: Number(row[15]) || 0
});
  });

  // --- PROCESS COMPONENTS ---
  compData.forEach((row, index) => {
    const orderId = norm(row[0]);
    const product = norm(row[2]);

    if (!tree[orderId]) return;
    if (!tree[orderId].products[product]) tree[orderId].products[product] = { panels: [], components: [] };

    tree[orderId].products[product].components.push({
  rowIndex: index + 2,
  compName: row[3],                       // Component
  sku: row[4],                            // SKU
  qtyPerUnit: Number(row[5]) || 1,        // Qty Per Unit
  qtyRequired: Number(row[6]) || 0,       // Qty Required
  qtyPacked: Number(row[7]) || 0,         // Qty packed
  lastUser: String(row[8] || ""),
  lastUpdated: (row[9] instanceof Date)
  ? Utilities.formatDate(row[9], tz, "dd/MM/yyyy HH:mm")
  : String(row[9] || "")
});
  });

 
  // --- PROCESS DELIVERY ---
delivData.forEach((row) => {
  const orderId = norm(row[0]);
  const product = norm(row[3]);
  const room = norm(row[4]); // ✅ was String(...)
  const status = norm(row[5]) || "Pending"; // ✅ normalise

  if (!tree[orderId]) return;

  if (room === "") {
    if (!tree[orderId].delivery.bucket[product]) {
      tree[orderId].delivery.bucket[product] = { qty: 0 };
    }
    tree[orderId].delivery.bucket[product].qty++;
  } else {
    if (!tree[orderId].delivery.rooms[room]) {
      tree[orderId].delivery.rooms[room] = [];
    }

    let existingItem = tree[orderId].delivery.rooms[room]
      .find(i => norm(i.name) === product && norm(i.status) === status); // ✅ normalise

    if (existingItem) existingItem.qty++;
    else tree[orderId].delivery.rooms[room].push({ name: product, qty: 1, status: status });
  }
});


  return tree;
}

// 3. UPDATE QUANTITIES (Robust Dynamic Version)
function updateQty(rowIndex, colName, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manufacture Hub");
  const logSheet = ss.getSheetByName("Activity Log"); 

  // 1. DYNAMIC COLUMN MAPPING (Finds column by name, not number)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[String(h || "").trim().toLowerCase()] = i + 1);

  // Map app keys to spreadsheet header names
  let headerName = "";
  if (colName === 'cut') headerName = "qty cut";
  if (colName === 'processed') headerName = "qty processed";
  if (colName === 'edgeFinish') headerName = "qty edge finish";
  if (colName === 'packed') headerName = "qty packed";

  const colIndex = map[headerName];

  if (colIndex > 0) {
    const timestamp = new Date();
    let userEmail = Session.getActiveUser().getEmail();
    if (userEmail === "") userEmail = "Workshop App User"; 

    // 2. UPDATE THE QUANTITY
    sheet.getRange(rowIndex, colIndex).setValue(value);

    // 3. UPDATE METADATA (Only if columns exist)
    if (map["last action"]) sheet.getRange(rowIndex, map["last action"]).setValue(colName);
    if (map["last user"]) sheet.getRange(rowIndex, map["last user"]).setValue(userEmail);
    if (map["last updated"]) sheet.getRange(rowIndex, map["last updated"]).setValue(timestamp);

    // 4. APPEND TO ACTIVITY LOG
    if (logSheet) {
      const rowData = sheet.getRange(rowIndex, 1, 1, 5).getValues()[0];
      const orderId = rowData[0];   
      const panelName = rowData[4]; 
      logSheet.appendRow([timestamp, userEmail, orderId, panelName, colName, value]);
    }
    return "Success";
  }
  return "Error: Column not found";
}


// 4. PROCESS PANEL UPDATES (Robust Dynamic + Safe Lock + Array Return + Delta Fix)
function processBatch(updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manufacture Hub");
  if (!sheet) throw new Error("Sheet 'Manufacture Hub' missing");

  let userEmail = Session.getActiveUser().getEmail() || "Workshop App";
  const timestamp = new Date();

  // 1. DYNAMIC COLUMN MAPPING
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[String(h || "").trim().toLowerCase()] = i + 1);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 

  try {
    // 2. APPLY UPDATES
    const results = []; // Return list for frontend
    updates.sort((a, b) => (a.ts || 0) - (b.ts || 0));

    updates.forEach(item => {
      const rowIndex = item.rowIndex; 
      const colName = item.colName;
      
      // --- FIX: Calculate Value from Base + Delta if 'value' is missing ---
      let value = item.value;
      if (value === undefined && item.delta !== undefined) {
         value = (Number(item.base) || 0) + (Number(item.delta) || 0);
      }
      // --------------------------------------------------------------------

      // Map app keys to spreadsheet header names
      let headerName = "";
      if (colName === 'cut') headerName = "qty cut";
      if (colName === 'processed') headerName = "qty processed";
      if (colName === 'edgeFinish') headerName = "qty edge finish";
      if (colName === 'packed') headerName = "qty packed";

      const colIndex = map[headerName];
      
      if (colIndex > 0) {
        // Update Quantity
        sheet.getRange(rowIndex, colIndex).setValue(value);
        
        // Update Metadata
        if (map["last action"]) sheet.getRange(rowIndex, map["last action"]).setValue(colName);
        if (map["last user"]) sheet.getRange(rowIndex, map["last user"]).setValue(userEmail);
        if (map["last updated"]) sheet.getRange(rowIndex, map["last updated"]).setValue(timestamp);

        // Add to results array for the frontend
        results.push({
            rowIndex: rowIndex,
            colName: colName,
            value: value
        });

        // 3. TRIGGER SYNC (If Packed)
        if (colName === 'packed') {
           const rowData = sheet.getRange(rowIndex, 1, 1, 3).getValues()[0];
           checkProductCompletion(rowData[0], rowData[2]);
        }
      }
    });

    return results; // Return the ARRAY

  } catch (e) {
    console.error(e);
    throw new Error("Save failed: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// HELPER: Calculate Min Sets
function calculateSets(rows) {
  let minSets = Infinity;
  if (rows.length === 0) return 0;

  rows.forEach(r => {
    const sets = Math.floor(r.qtyPacked / r.qtyPerUnit);
    if (sets < minSets) minSets = sets;
  });
  
  if (minSets === Infinity) return 0;
  return minSets;
}

function processComponentBatch(updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Components Hub");
  if (!sheet) throw new Error("Sheet 'Components Hub' missing");
  if (!Array.isArray(updates)) return [];

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  const cache = CacheService.getScriptCache();
  const userEmail = Session.getActiveUser().getEmail() || "Workshop App";
  const results = [];

  try {
    updates.sort((a, b) => (a.ts || 0) - (b.ts || 0));

    updates.forEach(u => {
      const rowIndex = Number(u.rowIndex);
      if (!rowIndex) return;

      const opId = String(u.opId || "");
      if (opId && cache.get("op:" + opId)) {
        const cur = Number(sheet.getRange(rowIndex, 8).getValue()) || 0;
        results.push({ rowIndex, value: cur, duplicate: true });
        return;
      }

      const required = Number(sheet.getRange(rowIndex, 7).getValue()) || 0; // Col G
      const cur = Number(sheet.getRange(rowIndex, 8).getValue()) || 0;      // Col H

      const mode = String(u.mode || "inc");
      let next = cur;

      if (mode === "set") {
        const base = Number(u.base);
        const val  = Number(u.value);

        // Conflict check (prevents overwriting someone else's edit)
        if (!isNaN(base) && base !== cur) {
          results.push({ rowIndex, conflict: true, current: cur });
          return;
        }

        next = isNaN(val) ? cur : val;
      } else {
        const delta = Number(u.delta) || 0;
        next = cur + delta;
      }

      if (next < 0) next = 0;
      if (required > 0 && next > required) next = required;

      sheet.getRange(rowIndex, 8).setValue(next);     // Packed
      sheet.getRange(rowIndex, 9).setValue(userEmail); // Last User
      sheet.getRange(rowIndex,10).setValue(new Date()); // Last Updated

      if (opId) cache.put("op:" + opId, "1", 21600); // 6h
      results.push({ rowIndex, value: next });
    });

    SpreadsheetApp.flush();
    return results;
  } finally {
    lock.releaseLock();
  }
}



// --- FURNITURE STOCK LOGIC ---

function updateFinishedGoodsStock(sku, productName, qtyChange, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Finished Goods");

  // Create if missing (correct headers)
  if (!sheet) {
    sheet = ss.insertSheet("Finished Goods");
    sheet.appendRow(["Customer Name", "SKU", "Product Name", "Qty Available", "Manufactured Total"]);
  }

  const data = sheet.getDataRange().getValues();

  const targetSku = String(sku).trim().toLowerCase();
  let foundRow = -1;

  // SKU is column B now
  for (let i = 1; i < data.length; i++) {
    const rowSku = String(data[i][1]).trim().toLowerCase();
    if (rowSku === targetSku) {
      foundRow = i + 1; // 1-based row index
      break;
    }
  }

  const n = (v) => Number(v) || 0;
  const delta = n(qtyChange);

  if (foundRow > 0) {
    // Qty Available is Col D
    const currentQty = n(sheet.getRange(foundRow, 4).getValue());
    sheet.getRange(foundRow, 4).setValue(currentQty + delta);

    // keep name fresh (Col C)
    sheet.getRange(foundRow, 3).setValue(productName);
  } else {
    // New row: treat as workshop stock
    sheet.appendRow(["Workshop Stock", sku, productName, delta, ""]);
  }

  // Log
  logStockTransaction(`Finished Goods: ${sku}`, delta, reason);

  return { success: true, sku, newQtyChange: delta };
}


// --- APP VIEW FETCHERS ---

function getFurnitureStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Finished Goods");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const stock = [];

  // Skip Header (Start at 1)
  for (let i = 1; i < data.length; i++) {
    // Only return if Qty > 0
    const qty = Number(data[i][3]) || 0; // Col D
    if (qty > 0) {
      stock.push({
        rowIndex: i + 1,
        customer: data[i][0], // Col A
        sku: data[i][1],      // Col B
        product: data[i][2],  // Col C
        qty: qty              // Col D
      });
    }
  }
  return stock;
}

function adjustFurnitureStock(rowIndex, change, reason) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Finished Goods");
  if (!sheet) throw new Error("Finished Goods tab missing");

  const current = sheet.getRange(rowIndex, 4).getValue(); // Col D = Qty
  const newVal = (Number(current) || 0) + Number(change);

  if (newVal < 0) throw new Error("Stock cannot be negative");

  sheet.getRange(rowIndex, 4).setValue(newVal);

  const prodName = sheet.getRange(rowIndex, 3).getValue(); // Col C = Product Name
  logStockTransaction("Finished Good: " + prodName, change, reason);

  return "Success";

    } finally {
    lock.releaseLock();
  }

}

function _normTxt(v) {
  return String(v ?? "").replace(/\u00A0/g, " ").trim();
}


function _getShopifyOrdersSheet() {
  const ss = SpreadsheetApp.openById(SHOPIFY_ORDERS_SHEET_ID);
  const sh = ss.getSheetByName(SHOPIFY_ORDERS_TAB_NAME);
  if (!sh) throw new Error(`Shopify Orders tab not found: "${SHOPIFY_ORDERS_TAB_NAME}"`);
  return sh;
}

function _headerMap_(headers) {
  const m = {};
  headers.forEach((h, i) => m[_normTxt(h).toLowerCase()] = i);
  return m;
}

function _getFinishedGoodsRowMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Finished Goods");
  if (!sh || sh.getLastRow() < 2) return {};

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => _normTxt(h).toLowerCase());
  const hm = _headerMap_(headers);

  // Expected headers:
  // Customer Name | SKU | Product Name | Qty Available | Manufactured Total
  const idxCustomer = hm["customer name"];
  const idxSku = hm["sku"];
  const idxAvail = hm["qty available"];

  if (idxCustomer == null || idxSku == null || idxAvail == null) return {};

  const map = {}; // skuLower -> { rowIndex, avail, sku, customer }
  for (let r = 1; r < data.length; r++) {
    const customer = _normTxt(data[r][idxCustomer]);
    const sku = _normTxt(data[r][idxSku]);
    if (!sku) continue;

    // only Workshop Stock rows count as allocatable supply
    if (customer.toLowerCase() !== "workshop stock") continue;

    const avail = Number(data[r][idxAvail]) || 0;
    map[sku.toLowerCase()] = {
      rowIndex: r + 1, // sheet row index
      sku,
      avail
    };
  }
  return map;
}

function _extractAllocatedFromStatus_(status) {
  const m = String(status || "").match(/ALLOCATED\s+(\d+)\s+FROM\s+STOCK/i);
  return m ? (parseInt(m[1], 10) || 0) : 0;
}


function getPendingShopifyOrdersWithStock() {
  const sh = _getShopifyOrdersSheet();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => _normTxt(h).toLowerCase());
  const hm = _headerMap_(headers);

  const idxOrderId = hm["order id"];
  const idxCustomer = hm["customer"];
  const idxProductName = hm["product name"];
  const idxSku = hm["product code"];
  const idxQty = hm["quantity ordered"];
  const idxStatus = hm["import status"];

  if ([idxOrderId, idxCustomer, idxProductName, idxSku, idxQty, idxStatus].some(x => x == null)) {
    throw new Error("Shopify Orders headers don't match expected names. Check spelling/case.");
  }

  const fg = _getFinishedGoodsRowMap_();

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const orderId = _normTxt(values[r][idxOrderId]);
    const cust = _normTxt(values[r][idxCustomer]);
    const prodName = _normTxt(values[r][idxProductName]);
    const sku = _normTxt(values[r][idxSku]);
    const qtyOrdered = Number(values[r][idxQty]) || 0;
    const status = _normTxt(values[r][idxStatus]);

    if (!orderId || !sku || qtyOrdered <= 0) continue;

    const st = status.trim().toUpperCase();
    if (st.startsWith("IMPORTED")) continue;
    if (st.startsWith("CANCELLED")) continue;

    const allocated = _extractAllocatedFromStatus_(status);
    const remainingToAllocate = Math.max(0, qtyOrdered - allocated);

    const fgRow = fg[sku.toLowerCase()];
    const avail = fgRow ? fgRow.avail : 0;

    const recommend = Math.min(avail, remainingToAllocate);

    out.push({
      rowIndex: r + 1,
      orderId,
      customer: cust,
      productName: prodName,
      sku,
      qtyOrdered,
      importStatus: status,
      allocatedFromStock: allocated,
      remainingToAllocate,
      availableInStock: avail,
      recommendedAllocate: recommend
    });
  }

  return out;
}



function approveShopifyOrder(shopifyRowIndex) {
  const sh = _getShopifyOrdersSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => _normTxt(h).toLowerCase());
  const hm = _headerMap_(headers);

  const idxStatus = hm["import status"];
  if (idxStatus == null) throw new Error("Import Status column not found in Shopify sheet");

  const current = _normTxt(sh.getRange(shopifyRowIndex, idxStatus + 1).getValue());
  const upper = current.toUpperCase();

  // don’t approve if already imported
  if (upper.startsWith("IMPORTED")) return "Already imported";

  // if blank, set APPROVED; if has notes, prefix APPROVED
  const next = current ? `APPROVED | ${current}` : "APPROVED";
  sh.getRange(shopifyRowIndex, idxStatus + 1).setValue(next);

  return "SUCCESS";
}


// 5. UPDATE COMPONENT STATUS
function updateComponentStatus(rowIndex, isPacked) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Components Hub");
  
  // Column 7 (G) is Status
  const status = isPacked ? "Packed" : "Pending";
  sheet.getRange(rowIndex, 7).setValue(status);
  
  return "Success";

    } finally {
    lock.releaseLock();
  }

}

// 7. DELIVERY: ASSIGN ITEMS TO ROOM (Planning Phase)
function assignToRoom(orderId, productName, qtyToAssign, roomName) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Delivery Hub");
  if (!sheet) return "Error: Delivery Hub missing";

  const normStatus = (s) => {
    s = String(s || "").trim();
    return s === "" ? "Pending" : s;
  };

  const data = sheet.getDataRange().getValues();
  const user = Session.getActiveUser().getEmail() || "Workshop App";
  const ts = new Date();

  let assigned = 0;

  // Structure: A ID | B Cust | C Addr | D Prod | E Room | F Status | G User | H Updated
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const normTxt = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim();

if (normTxt(row[0]) !== normTxt(orderId)) continue;
if (normTxt(row[3]) !== normTxt(productName)) continue;
if (normTxt(row[4]) !== "") continue; // already assigned


    const st = normStatus(row[5]);
    if (st !== "Pending") continue;

    if (assigned < qtyToAssign) {
      sheet.getRange(i + 1, 5).setValue(roomName);     // Room (E)
      sheet.getRange(i + 1, 6).setValue("Pending");    // Status (F)
      sheet.getRange(i + 1, 7).setValue(user);         // Last User (G)
      sheet.getRange(i + 1, 8).setValue(ts);           // Last Updated (H)
      assigned++;
    }

    if (assigned >= qtyToAssign) break;
  }

  if (assigned === 0) return "Error: No unassigned items found for " + productName;
  return "SUCCESS";

    } finally {
    lock.releaseLock();
  }

}


// 8. DELIVERY: UPDATE STATUS (DIAGNOSTIC VERSION)
function updateDeliveryStatus(orderId, roomName, productName, oldStatus, newStatus, qtyToUpdate) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const normTxt = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Delivery Hub");
  if (!sheet) return "Error: Delivery Hub missing";

  const normStatus = (s) => {
    s = String(s || "").trim();
    return s === "" ? "Pending" : s;
  };

  oldStatus = normStatus(oldStatus);
  newStatus = normStatus(newStatus);
  qtyToUpdate = Math.max(1, Number(qtyToUpdate) || 1);

  // SAFETY: only consumes factory allowance if moving from Pending -> Delivered/Fitted
  const movingIntoSite = (newStatus === "Delivered" || newStatus === "Fitted");
  const alreadyOnSite = (oldStatus === "Delivered" || oldStatus === "Fitted");
  const delta = (movingIntoSite && !alreadyOnSite) ? qtyToUpdate : 0;

  if (delta > 0) {
    const maxAllowed = getMaxReadyFromFactory(orderId, productName);

    const data = sheet.getDataRange().getValues();
    let currentlyUsed = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      

if (normTxt(row[0]) === normTxt(orderId) && normTxt(row[3]) === normTxt(productName)) {

        const st = normStatus(row[5]);
        if (st === "Delivered" || st === "Fitted") currentlyUsed++;
      }
    }

    if (currentlyUsed + delta > maxAllowed) {
      return "LIMIT_REACHED:" + maxAllowed + ":" + currentlyUsed;
    }
  }

  // APPLY: update only qtyToUpdate rows matching oldStatus in that room/product
  const data = sheet.getDataRange().getValues();
  const user = Session.getActiveUser().getEmail() || "Workshop App";
  const ts = new Date();

  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (normTxt(row[0]) !== normTxt(orderId)) continue;
    if (normTxt(row[4]) !== normTxt(roomName)) continue;
    if (normTxt(row[3]) !== normTxt(productName)) continue;


    const st = normStatus(row[5]);
    if (st !== oldStatus) continue;

    sheet.getRange(i + 1, 6).setValue(newStatus); // Status (F)
    sheet.getRange(i + 1, 7).setValue(user);      // Last User (G)
    sheet.getRange(i + 1, 8).setValue(ts);        // Last Updated (H)

    updated++;
    if (updated >= qtyToUpdate) break;
  }

  if (updated === 0) return "Error: No matching items found to update.";
  return "SUCCESS";

    } finally {
    lock.releaseLock();
  }

}


// 9. HELPER: CALCULATE MAX COMPLETED UNITS (Strict Match)
function getMaxReadyFromFactory(orderId, productName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hubSheet = ss.getSheetByName("Manufacture Hub");
  const compSheet = ss.getSheetByName("Components Hub");
  
  const normTxt = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim();

  const targetId = normTxt(orderId);
  const targetProd = normTxt(productName);

  // 1. Get Panel Data
  const pData = hubSheet.getDataRange().getValues();
  let minPanelsReady = Infinity; 
  let hasPanels = false;

  for (let i = 1; i < pData.length; i++) {
    const row = pData[i];
    // Strict Match
    if (normTxt(row[0]) === targetId && normTxt(row[2]) === targetProd) {
      hasPanels = true;
      const qtyPerUnit = Number(row[6]) || 1;
      const qtyPacked = Number(row[15]) || 0; // Col P
      const possibleUnits = Math.floor(qtyPacked / qtyPerUnit);
      if (possibleUnits < minPanelsReady) minPanelsReady = possibleUnits;
    }
  }

  // 2. Get Component Data
  // 2. Get Component Data (Qty-based)
const cData = compSheet.getDataRange().getValues();
let minCompsReady = Infinity;
let hasComps = false;

for (let i = 1; i < cData.length; i++) {
  const row = cData[i];

  if (normTxt(row[0]) === targetId && normTxt(row[2]) === targetProd) {
    hasComps = true;

    const perUnit = Number(row[5]) || 1;   // Col F
    const packed = Number(row[7]) || 0;    // Col H
    const possibleUnits = Math.floor(packed / perUnit);

    if (possibleUnits < minCompsReady) minCompsReady = possibleUnits;
  }
}

if (minCompsReady === Infinity) minCompsReady = 0;


  // 3. Final Calculation
  let limit = 9999;
  
  if (hasPanels) limit = Math.min(limit, minPanelsReady);
  if (hasComps) limit = Math.min(limit, minCompsReady);
  
  // CRITICAL FIX: If we found NO panels and NO components, we assume 
  // the Product Name didn't match anything. Therefore, we have 0 stock.
  if (!hasPanels && !hasComps) return 0;

  if (limit === 9999 && (hasPanels || hasComps)) return 9999; // All good
  
  return limit;
}

// 10. DELIVERY: UNASSIGN ITEMS (Return to Bucket)
function unassignFromRoom(orderId, roomName, productName, qtyToUnassign) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const normTxt = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Delivery Hub");
  if (!sheet) return "Error: Delivery Hub missing";

  const data = sheet.getDataRange().getValues();
  const user = Session.getActiveUser().getEmail() || "Workshop App";
  const ts = new Date();

  let done = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (normTxt(row[0]) !== normTxt(orderId)) continue;
    if (normTxt(row[4]) !== normTxt(roomName)) continue;
    if (normTxt(row[3]) !== normTxt(productName)) continue;


    if (done < qtyToUnassign) {
      sheet.getRange(i + 1, 5).setValue("");          // clear Room (E)
      sheet.getRange(i + 1, 6).setValue("Pending");   // reset Status (F)
      done++;
    }

    if (done >= qtyToUnassign) break;
  }

  return "Returned " + done + " items to bucket";

    } finally {
    lock.releaseLock();
  }

}


// --- CNC EXPORT ---
function exportCncXml(orderId, productName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hubSheet = ss.getSheetByName("Manufacture Hub");
  
  // 1. Get Data
  const data = hubSheet.getDataRange().getValues();
  let xmlContent = '<?xml version="1.0" encoding="UTF-8"?>\n<Job>\n';
  let count = 0;
  
  // 2. Loop through rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Check match: Order ID (Col A) and Product Name (Col C)
    if (String(row[0]) === String(orderId) && String(row[2]) === String(productName)) {
      
      const length = row[7]; // Col H
      const width = row[8];  // Col I
      const qty = row[11];   // Col L (Qty per Order)
      const barcodeId = row[19]; // Col T (The new Barcode ID)
      
      const grain = "Y"; 

      // 3. Build XML Item
      xmlContent += '  <Panel>\n';
      xmlContent += `    <Length>${length}</Length>\n`;
      xmlContent += `    <Width>${width}</Width>\n`;
      xmlContent += `    <Qty>${qty}</Qty>\n`;
      xmlContent += `    <Grain>${grain}</Grain>\n`;
      xmlContent += `    <Desc_1>${barcodeId}</Desc_1>\n`;
      xmlContent += '  </Panel>\n';
      
      count++;
    }
  }
  xmlContent += '</Job>';

  if (count === 0) return "Error: No panels found for " + productName;

  // 4. Save to Drive
  const fileName = `CNC_${orderId}_${productName}.xml`.replace(/ /g, "_");
  
  // --- THE FIX IS HERE: Use "text/xml" string instead of MimeType.XML ---
  const file = DriveApp.createFile(fileName, xmlContent, "text/xml");
  
  return "Saved to Drive: " + fileName;
}

// --- INVENTORY MANAGEMENT (CORRECTED) ---

// 1. FETCH INVENTORY DATA (Simplified for Single Material Column)
function getInventoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const compSheet = ss.getSheetByName("Component Stock");
  const woodSheet = ss.getSheetByName("Wood Stock");
  const edgeSheet = ss.getSheetByName("Edge Band Stock");
  
  const result = { components: [], wood: [], edge: [] };

  // A. COMPONENTS
  if (compSheet && compSheet.getLastRow() > 1) {
    const data = compSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      result.components.push({
        rowIndex: i + 1,
        sku: data[i][0],
        name: data[i][1],
        category: data[i][2],
        supplier: data[i][3],
        stock: Number(data[i][4]) || 0,
        min: Number(data[i][5]) || 0,
        image: data[i][6]
      });
    }
  }

  // B. WOOD STOCK (New 5-Column Layout)
  if (woodSheet && woodSheet.getLastRow() > 1) {
    const data = woodSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // DEBUG: Ensure we are reading the right column
      // row[0] = Material, row[1] = Type, row[2] = Len, row[3] = Wid, row[4] = Qty
      result.wood.push({
        rowIndex: i + 1,
        material: String(data[i][0]), // Col A
        type: data[i][1],             // Col B
        length: data[i][2],           // Col C
        width: data[i][3],            // Col D
        size: `${data[i][2]} x ${data[i][3]}`, 
        onOrder: Number(data[i][4]) || 0, // Col E = Qty on Order
        qty: Number(data[i][5]) || 0      // Col F = Qty in Stock
      });
    }
  }
    // C. EDGE BAND STOCK (3-Column Layout)
  if (edgeSheet && edgeSheet.getLastRow() > 1) {
    const data = edgeSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      result.edge.push({
        rowIndex: i + 1,
        material: String(data[i][0] || "").trim(),
        thickness: String(data[i][1] || "").trim(),
        onOrder: Number(data[i][2]) || 0, // Col C = Rolls on Order
        rolls: Number(data[i][3]) || 0    // Col D = Rolls In Stock
      });
    }
  }
  return result;
}

// ADJUST WOOD STOCK (With Logging)
function adjustWoodStock(rowIndex, change, reason) {
    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Wood Stock");
  
  // 1. Update Value
  const currentVal = sheet.getRange(rowIndex, 6).getValue(); // Col F = Qty in Stock
  const newVal = (Number(currentVal) || 0) + Number(change);
  
  if (newVal < 0) throw new Error("Result cannot be negative");
  
  sheet.getRange(rowIndex, 6).setValue(newVal);

  // 2. Log History
  const materialName = sheet.getRange(rowIndex, 1).getValue(); // Col A
  logStockTransaction(materialName, change, reason);

  return newVal;

    } finally {
    lock.releaseLock();
  }

}

function adjustEdgeStock(rowIndex, change, reason) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Edge Band Stock");
  if (!sheet) throw new Error("Edge Band Stock tab missing");

  const currentVal = sheet.getRange(rowIndex, 4).getValue(); // Col D = Rolls In Stock
  const newVal = (Number(currentVal) || 0) + Number(change);

  if (newVal < 0) throw new Error("Stock cannot be negative");

  sheet.getRange(rowIndex, 4).setValue(newVal);

  const mat = sheet.getRange(rowIndex, 1).getValue(); // Col A
  const thk = sheet.getRange(rowIndex, 2).getValue(); // Col B
  logStockTransaction(`Edge: ${mat} | ${thk}mm`, change, reason);

  return newVal;

    } finally {
    lock.releaseLock();
  }

}

function adjustComponentStock(rowIndex, change, reason) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Component Stock");
  if (!sheet) throw new Error("Component Stock tab missing");

  const currentVal = sheet.getRange(rowIndex, 5).getValue(); // Col E = Stock
  const newVal = (Number(currentVal) || 0) + Number(change);

  if (newVal < 0) throw new Error("Stock cannot be negative");

  sheet.getRange(rowIndex, 5).setValue(newVal);

  const sku = sheet.getRange(rowIndex, 1).getValue();  // Col A
  const name = sheet.getRange(rowIndex, 2).getValue(); // Col B
  logStockTransaction(`Component: ${sku} | ${name}`, change, reason);

  return newVal;

    } finally {
    lock.releaseLock();
  }

}



// 3. ALLOCATE WOOD (Updated for Summary Log)
function allocateWood(rowIndex, qtyUsed, projectId, productName, offcutData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("Wood Stock");
  
  if (!stockSheet) throw new Error("Wood Stock tab missing");

  // 1. DEDUCT STOCK (Physical Inventory)
  const currentQty = stockSheet.getRange(rowIndex, 6).getValue(); // Col F = Qty in Stock
  const newQty = (Number(currentQty) || 0) - Number(qtyUsed);
  if (newQty < 0) throw new Error("Insufficient stock.");
  stockSheet.getRange(rowIndex, 6).setValue(newQty);

  // 2. LOG HISTORY (The "Book" Icon - Transactional)
  const materialName = stockSheet.getRange(rowIndex, 1).getValue(); // Col A
  let historyReason = `CNC Job: #${projectId} (${productName})`;
  if(offcutData) historyReason += " [Offcut Generated]";
  
  // This goes to 'Stock History' tab
  logStockTransaction(materialName, -qtyUsed, historyReason);

  // 3. UPDATE PROJECT SUMMARY (The "Wood Usage Log" Tab)
  // This aggregates the totals nicely
  updateProjectUsageSummary(projectId, productName, materialName, qtyUsed);

  // 4. CREATE OFFCUT (Physical Inventory)
  if (offcutData) {
    stockSheet.appendRow([
      materialName,        // A: Material
      "Offcut",            // B: Type
      offcutData.length,   // C: Length
      offcutData.width,    // D: Width
      0,                   // E: Qty on Order
      offcutData.qty       // f: Qty
    ]);
  }
  
  return "Success";
}

function adjustWoodOnOrder(rowIndex, change) {
    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Wood Stock");
  if (!sheet) throw new Error("Wood Stock tab missing");

  const currentVal = sheet.getRange(rowIndex, 5).getValue(); // Col E = Qty on Order
  const newVal = (Number(currentVal) || 0) + Number(change);

  if (newVal < 0) throw new Error("On Order cannot be negative");

  sheet.getRange(rowIndex, 5).setValue(newVal);
  return newVal;

    } finally {
    lock.releaseLock();
  }

}

function adjustEdgeOnOrder(rowIndex, change) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Edge Band Stock");
  if (!sheet) throw new Error("Edge Band Stock tab missing");

  const currentVal = sheet.getRange(rowIndex, 3).getValue(); // Col C = Rolls on Order
  const newVal = (Number(currentVal) || 0) + Number(change);

  if (newVal < 0) throw new Error("On Order cannot be negative");

  sheet.getRange(rowIndex, 3).setValue(newVal);
  return newVal;

    } finally {
    lock.releaseLock();
  }

}

function receiveWoodFromOrder(rowIndex, qtyReceived) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Wood Stock");
  if (!sheet) throw new Error("Wood Stock tab missing");

  const onOrder = Number(sheet.getRange(rowIndex, 5).getValue()) || 0; // Col E
  const inStock = Number(sheet.getRange(rowIndex, 6).getValue()) || 0; // Col F

  const qty = Number(qtyReceived) || 0;
  if (qty <= 0) throw new Error("Receive qty must be > 0");
  if (qty > onOrder) throw new Error("Cannot receive more than On Order");

  sheet.getRange(rowIndex, 5).setValue(onOrder - qty);
  sheet.getRange(rowIndex, 6).setValue(inStock + qty);

  const materialName = sheet.getRange(rowIndex, 1).getValue(); // Col A
  logStockTransaction(materialName, qty, "Restock / Delivery (from Order)");

  return "Success";

    } finally {
    lock.releaseLock();
  }

}

function receiveEdgeFromOrder(rowIndex, rollsReceived) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Edge Band Stock");
  if (!sheet) throw new Error("Edge Band Stock tab missing");

  const onOrder = Number(sheet.getRange(rowIndex, 3).getValue()) || 0; // Col C
  const inStock = Number(sheet.getRange(rowIndex, 4).getValue()) || 0; // Col D

  const qty = Number(rollsReceived) || 0;
  if (qty <= 0) throw new Error("Receive qty must be > 0");
  if (qty > onOrder) throw new Error("Cannot receive more than On Order");

  sheet.getRange(rowIndex, 3).setValue(onOrder - qty);
  sheet.getRange(rowIndex, 4).setValue(inStock + qty);

  const mat = sheet.getRange(rowIndex, 1).getValue(); // Col A
  const thk = sheet.getRange(rowIndex, 2).getValue(); // Col B
  logStockTransaction(`Edge: ${mat} | ${thk}mm`, qty, "Restock / Delivery (from Order)");

  return "Success";

    } finally {
    lock.releaseLock();
  }

}


// HELPER: Centralized Stock Logging
function logStockTransaction(material, change, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName("Stock History");
  if (!historySheet) return;

  // ✅ Ensure header has "Source" column (A-E)
  if (historySheet.getLastRow() === 0) {
    historySheet.appendRow(["Timestamp", "User", "Source", "Material", "Change", "Reason"]);
  } else {
    const headers = historySheet.getRange(1, 1, 1, historySheet.getLastColumn()).getValues()[0];
    if (headers.length < 6 || String(headers[2]).toLowerCase() !== "source") {
      historySheet.insertColumnAfter(2); // Insert new Col C
      historySheet.getRange(1, 1, 1, 6).setValues([["Timestamp", "User", "Source", "Material", "Change", "Reason"]]);
    }
  }

  const user = Session.getActiveUser().getEmail() || "Workshop App";
  const timestamp = new Date();

  // ✅ Detect source from material prefix
  let source = "Other";
  const m = String(material || "").toLowerCase();
  if (m.startsWith("edge:")) source = "Edgebanding";
   else if (m.startsWith("component:")) source = "Components";
   else if (m.startsWith("finished goods:") || m.startsWith("finished good:")) source = "Finished Goods";
   else source = "Wood";


  historySheet.appendRow([timestamp, user, source, material, change, reason]);
}


// HELPER: Update Project Summary (Find match and increment, or create new)
function updateProjectUsageSummary(projectId, productName, materialName, qtyToAdd) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Wood Usage Log");
  
  // Safety: Create tab if missing
  if (!sheet) {
    sheet = ss.insertSheet("Wood Usage Log");
    sheet.appendRow(["Project", "Product", "Material", "Total Qty"]);
  }

  const data = sheet.getDataRange().getValues();
  let foundRowIndex = -1;

  // Search for existing entry (Skip header row)
  // Matching: Project (Col A) + Product (Col B) + Material (Col C)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(projectId) && 
        String(data[i][1]) === String(productName) && 
        String(data[i][2]) === String(materialName)) {
      foundRowIndex = i + 1; // Store the 1-based Row Index
      break;
    }
  }

  if (foundRowIndex > 0) {
    // MATCH FOUND: Update existing quantity
    const currentQty = Number(sheet.getRange(foundRowIndex, 4).getValue()) || 0; // Col D
    sheet.getRange(foundRowIndex, 4).setValue(currentQty + qtyToAdd);
  } else {
    // NO MATCH: Create new summary row
    sheet.appendRow([projectId, productName, materialName, qtyToAdd]);
  }
}



// 6. GET HISTORY FOR A MATERIAL (Robust Version)
function getMaterialHistory(materialName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Stock History");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const history = [];
  
  // Normalize search term
  const searchMat = String(materialName).trim().toLowerCase();

  // Loop backwards (newest first)
  // Skip header (i > 0)
  for (let i = data.length - 1; i > 0; i--) {
    const rowMat = String(data[i][3]).trim().toLowerCase(); // Col D (Material)
    
    if (rowMat === searchMat) {
      // Format Date Server-Side to prevent Frontend Crash
      let dateDisplay = "Unknown Date";
      try {
        const rawDate = data[i][0];
        if (rawDate instanceof Date) {
          dateDisplay = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        } else {
          dateDisplay = String(rawDate);
        }
      } catch (e) {
        dateDisplay = "Date Error";
      }

      history.push({
        dateStr: dateDisplay, // <--- Sending pre-formatted string
        user: data[i][1],
        change: Number(data[i][4]) || 0,   // Col E (Change)
        reason: data[i][5],                // Col F (Reason)
        source: data[i][2]                 // Col C (Source)
      });

      // Limit to last 20 entries
      if (history.length >= 20) break;
    }
  }
  return history;
}

// --- AGGREGATED FINISHED GOODS LOGIC (Customer + SKU) ---

function checkProductCompletion(orderId, productName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hubSheet = ss.getSheetByName("Manufacture Hub");
  const data = hubSheet.getDataRange().getValues();

  // 1) Find metadata for this order/product (customer + SKU)
  let targetCustomer = "";
  let targetSKU = "";
  let realProdName = productName;

  const targetIdStr = String(orderId).trim().toLowerCase();
  const targetProdStr = String(productName).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][0]).trim().toLowerCase() === targetIdStr &&
      String(data[i][2]).trim().toLowerCase() === targetProdStr
    ) {
      targetCustomer = data[i][1]; // Col B
      targetSKU = data[i][3];      // Col D
      realProdName = data[i][2];   // Col C
      break;
    }
  }

  if (targetCustomer === "" || targetSKU === "") return; // Not found

  // ✅ Only maintain Finished Goods for Workshop Stock
  if (String(targetCustomer).trim().toLowerCase() !== "workshop stock") return;

  // 2) Aggregate packed panels for this Customer + SKU
  let panelTotals = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (
      String(row[1]).trim().toLowerCase() === String(targetCustomer).trim().toLowerCase() &&
      String(row[3]).trim().toLowerCase() === String(targetSKU).trim().toLowerCase()
    ) {
      const pName = row[4];
      const qtyPerUnit = Number(row[6]) || 1;
      const qtyPacked = Number(row[15]) || 0;

      if (!panelTotals[pName]) panelTotals[pName] = { packed: 0, required: qtyPerUnit };
      panelTotals[pName].packed += qtyPacked;
    }
  }

  // 3) Limiting factor = finished units
  let totalFinished = Infinity;
  let hasPanels = false;

  for (const key in panelTotals) {
    hasPanels = true;
    const p = panelTotals[key];
    const sets = Math.floor(p.packed / p.required);
    if (sets < totalFinished) totalFinished = sets;
  }

  if (!hasPanels || totalFinished === Infinity) totalFinished = 0;

  // 4) Sync to Finished Goods
  syncToFinishedGoods(targetCustomer, targetSKU, realProdName, totalFinished);
}

function syncToFinishedGoods(customer, sku, productName, newManufacturedTotal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Finished Goods");

  // Create if missing
  if (!sheet) {
    sheet = ss.insertSheet("Finished Goods");
    sheet.appendRow(["Customer Name", "SKU", "Product Name", "Qty Available", "Manufactured Total"]);
  }

  const data = sheet.getDataRange().getValues();

  const targetCust = String(customer).trim().toLowerCase();
  const targetSku  = String(sku).trim().toLowerCase();

  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    const rowCust = String(data[i][0]).trim().toLowerCase();
    const rowSku  = String(data[i][1]).trim().toLowerCase();
    if (rowCust === targetCust && rowSku === targetSku) {
      foundRow = i + 1; // 1-based row index
      break;
    }
  }

  // Helper to safely read numbers
  const n = (v) => Number(v) || 0;

  if (foundRow > 0) {
    // Existing row
    const curAvail = n(sheet.getRange(foundRow, 4).getValue());  // Col D
    const curMade  = n(sheet.getRange(foundRow, 5).getValue());  // Col E

    const delta = Math.max(0, n(newManufacturedTotal) - curMade);

    // Update Manufactured Total (Col E)
    sheet.getRange(foundRow, 5).setValue(n(newManufacturedTotal));

    // Increase Available by delta (Col D) but NEVER overwrite user allocations
    if (delta > 0) sheet.getRange(foundRow, 4).setValue(curAvail + delta);

    // Keep product name fresh (Col C)
    sheet.getRange(foundRow, 3).setValue(productName);

  } else {
    // New row: set Available = Manufactured Total
    const made = n(newManufacturedTotal);
    if (made > 0) {
      sheet.appendRow([customer, sku, productName, made, made]);
    }
  }
}

function allocateFinishedGoodsToOrder(shopifyRowIndex, qtyToAllocate) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  qtyToAllocate = Number(qtyToAllocate) || 0;
  if (qtyToAllocate <= 0) return "Nothing to allocate";

  const sh = _getShopifyOrdersSheet();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => _normTxt(h).toLowerCase());
  const hm = _headerMap_(headers);

  const idxSku = hm["product code"];
  const idxOrderId = hm["order id"];
  const idxCustomer = hm["customer"];
  const idxProductName = hm["product name"];
  const idxQty = hm["quantity ordered"];
  const idxStatus = hm["import status"];

  if ([idxSku, idxOrderId, idxCustomer, idxProductName, idxQty, idxStatus].some(x => x == null)) {
    throw new Error("Shopify sheet headers missing required columns.");
  }

  const row = sh.getRange(shopifyRowIndex, 1, 1, lastCol).getValues()[0];

  const sku = _normTxt(row[idxSku]);
  const orderId = _normTxt(row[idxOrderId]);
  const customer = _normTxt(row[idxCustomer]);
  const productName = _normTxt(row[idxProductName]);
  const qtyOrdered = Number(row[idxQty]) || 0;

  let status = _normTxt(row[idxStatus]);
  const stUpper = status.toUpperCase().trim();
  if (stUpper.startsWith("IMPORTED")) throw new Error("This line is already imported.");

  if (!sku || !orderId) throw new Error("Missing SKU or Order ID in Shopify row");
  if (qtyOrdered <= 0) throw new Error("Quantity Ordered must be > 0");

  const alreadyAllocated = _extractAllocatedFromStatus_(status);
  const remainingToAllocate = Math.max(0, qtyOrdered - alreadyAllocated);
  if (remainingToAllocate <= 0) throw new Error("This line is already fully allocated from stock.");
  if (qtyToAllocate > remainingToAllocate) throw new Error(`You can only allocate up to ${remainingToAllocate} more.`);

  // Finished Goods lookup (Workshop Stock only)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fgSheet = ss.getSheetByName("Finished Goods");
  if (!fgSheet) throw new Error("Finished Goods tab missing");

  const fgData = fgSheet.getDataRange().getValues();
  const fgHeaders = fgData[0].map(h => _normTxt(h).toLowerCase());
  const fhm = _headerMap_(fgHeaders);

  const idxFgCustomer = fhm["customer name"];
  const idxFgSku = fhm["sku"];
  const idxAvail = fhm["qty available"];
  if ([idxFgCustomer, idxFgSku, idxAvail].some(x => x == null)) {
    throw new Error("Finished Goods headers must include: Customer Name, SKU, Qty Available");
  }

  let fgRowIndex = -1;
  for (let r = 1; r < fgData.length; r++) {
    const c = _normTxt(fgData[r][idxFgCustomer]).toLowerCase();
    const s = _normTxt(fgData[r][idxFgSku]).toLowerCase();
    if (c === "workshop stock" && s === sku.toLowerCase()) {
      fgRowIndex = r + 1;
      break;
    }
  }
  if (fgRowIndex < 0) throw new Error(`No Workshop Stock found in Finished Goods for SKU: ${sku}`);

  const currentAvail = Number(fgSheet.getRange(fgRowIndex, idxAvail + 1).getValue()) || 0;
  if (qtyToAllocate > currentAvail) throw new Error(`Not enough stock. Available: ${currentAvail}`);

  // Deduct stock
  fgSheet.getRange(fgRowIndex, idxAvail + 1).setValue(currentAvail - qtyToAllocate);

  // Log stock history
  const reason = `Allocation to ${orderId} (${customer})`;
  logStockTransaction(`Finished Goods: ${sku} ${productName}`.trim(), -qtyToAllocate, reason);

  // Update Import Status (replace or add ALLOCATED note)
  const newTotalAllocated = alreadyAllocated + qtyToAllocate;

  let base = status.replace(/(\|\s*)?ALLOCATED\s+\d+\s+FROM\s+STOCK/ig, "").trim();
  base = base.replace(/\|\s*\|/g, "|").replace(/^\|\s*/, "").replace(/\s*\|$/, "").trim();

  const allocNote = `ALLOCATED ${newTotalAllocated} FROM STOCK`;
  const nextStatus = base ? `${base} | ${allocNote}` : allocNote;

  sh.getRange(shopifyRowIndex, idxStatus + 1).setValue(nextStatus);

  return "SUCCESS";

    } finally {
    lock.releaseLock();
  }

}

function approveAndImportShopifyRow(shopifyRowIndex) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  // 1) Approve line
  const res = approveShopifyOrder(shopifyRowIndex);
  if (res === "Already imported") return res;
  if (res !== "SUCCESS") return res;

  // 2) Import that ONE line immediately (function lives in Import.gs)
  return importApprovedShopifyRow_(shopifyRowIndex);

    } finally {
    lock.releaseLock();
  }

}

function ping() {
  return { ok: true, ts: new Date().toISOString() };
}

// HELPER: Dynamic Column Finder
// caches headers to avoid repeated reads in the same execution
function getColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    map[String(h).trim().toLowerCase()] = i + 1; // Store 1-based index
  });
  return map;
}
