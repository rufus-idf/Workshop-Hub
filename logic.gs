const SHOPIFY_ORDERS_SHEET_ID = "1KDDVnIZ5oCruCY4nyKp6XVqwWnL7-N6f-pKAD3xXo_U";
const SHOPIFY_ORDERS_TAB_NAME = "Sheet1";
const FURNITURE_STOCK_TAB_NAME = "Furniture Stock";
const LEGACY_FINISHED_GOODS_TAB_NAME = "Finished Goods";

function squeezeSpaces_(v) {
  return String(v ?? "").replace(/\u00A0/g, " ").trim().replace(/\s+/g, " ");
}

function canonicalRoomName_(v) {
  const cleaned = squeezeSpaces_(v);
  if (!cleaned) return "";

  const m = cleaned.match(/^room\s*([a-z0-9-]+)$/i);
  if (m) return "Room " + String(m[1]).toUpperCase();

  return cleaned.replace(/\b\w/g, function(ch) { return ch.toUpperCase(); });
}

function roomKey_(v) {
  return squeezeSpaces_(v).toLowerCase().replace(/[^a-z0-9]/g, "");
}

// Helper: get Furniture Stock sheet (auto-migrates legacy "Finished Goods" -> "Furniture Stock")
function _getFurnitureStockSheet_(createIfMissing) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Prefer new name
  let sh = ss.getSheetByName(FURNITURE_STOCK_TAB_NAME);

  // Migrate legacy name if needed
  if (!sh) {
    const legacy = ss.getSheetByName(LEGACY_FINISHED_GOODS_TAB_NAME);
    if (legacy) {
      try {
        legacy.setName(FURNITURE_STOCK_TAB_NAME);
        sh = legacy;
      } catch (e) {
        // If rename fails (e.g. name already taken), just keep using legacy
        sh = legacy;
      }
    }
  }

  if (!sh && createIfMissing) {
    sh = ss.insertSheet(FURNITURE_STOCK_TAB_NAME);
    sh.appendRow(["Customer Name", "SKU", "Product Name", "Qty Available", "Manufactured Total"]);
  }

  return sh;
}




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
  const room = canonicalRoomName_(row[4]); // ✅ canonical room
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
  const panelInfoCols = getPanelInfoColumnMap_(map);

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
    const previousValue = Number(sheet.getRange(rowIndex, colIndex).getValue()) || 0;
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

    const delta = Number(value) - previousValue;
    if (delta !== 0) {
      const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
      logPanelHistoryEntry_(rowData, panelInfoCols, buildPanelHistoryPayload_(colName, delta, userEmail, timestamp));
    }
    return "Success";
  }
  return "Error: Column not found";
}

function getPanelInfoColumnMap_(sheetHeaderMap) {
  return {
    orderId: sheetHeaderMap["order id"] || 1,
    productName: sheetHeaderMap["product name"] || sheetHeaderMap["product"] || 3,
    sku: sheetHeaderMap["sku"] || sheetHeaderMap["product sku"] || sheetHeaderMap["product code"] || 4,
    panelName: sheetHeaderMap["panel name"] || 5,
    barcodeId: sheetHeaderMap["barcode id"] || sheetHeaderMap["barcode"] || 20
  };
}

function buildPanelHistoryPayload_(colName, delta, userEmail, timestamp) {
  const changeLabels = {
    cut: "Cut",
    processed: "Processed",
    edgeFinish: "Edge Finish",
    packed: "Packed"
  };

  return {
    changeType: changeLabels[colName] || colName,
    quantity: delta,
    reason: "Manufacturing",
    user: userEmail || "Workshop App User",
    timestamp: timestamp || new Date()
  };
}

function buildDamagePayload_(qty, reason, userEmail, timestamp) {
  return {
    changeType: "Damaged",
    quantity: -Math.abs(Number(qty) || 0),
    reason: reason || "Damaged",
    user: userEmail || "Workshop App User",
    timestamp: timestamp || new Date()
  };
}

function logPanelHistoryEntry_(panelRow, panelInfoCols, payload) {
  const sheet = getPanelHistorySheet_();
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((h, i) => headerMap[String(h || "").trim().toLowerCase()] = i);

  const row = new Array(headers.length).fill("");
  const setVal = (key, value) => {
    const idx = headerMap[key];
    if (idx !== undefined) row[idx] = value;
  };

  const orderId = panelRow[panelInfoCols.orderId - 1];
  const productName = panelRow[panelInfoCols.productName - 1];
  const sku = panelInfoCols.sku ? panelRow[panelInfoCols.sku - 1] : "";
  const panelName = panelRow[panelInfoCols.panelName - 1];
  const barcodeId = panelInfoCols.barcodeId ? panelRow[panelInfoCols.barcodeId - 1] : "";

  setVal("event id", Utilities.getUuid());
  setVal("barcode id", barcodeId);
  setVal("order id", orderId);
  setVal("product", productName);
  setVal("product name", productName);
  setVal("sku", sku);
  setVal("product sku", sku);
  setVal("panel name", panelName);
  setVal("quantity", payload.quantity);
  setVal("change", payload.changeType);
  setVal("change type", payload.changeType);
  setVal("reason", payload.reason);
  setVal("user", payload.user);
  setVal("timestamp", payload.timestamp);

  sheet.appendRow(row);
}

function getPanelHistorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Panel History");
  if (!sheet) {
    sheet = ss.insertSheet("Panel History");
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Event ID",
      "Barcode ID",
      "Order ID",
      "Product",
      "Panel Name",
      "Quantity",
      "Change",
      "Reason",
      "User",
      "Timestamp"
    ]);
  }

  return sheet;
}

function markPanelDamaged(rowIndex, qty, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manufacture Hub");
  if (!sheet) throw new Error("Sheet 'Manufacture Hub' missing");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const qtyNum = Math.floor(Number(qty) || 0);
    if (!rowIndex || qtyNum <= 0) throw new Error("Invalid damaged quantity.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach((h, i) => map[String(h || "").trim().toLowerCase()] = i + 1);
    const panelInfoCols = getPanelInfoColumnMap_(map);

    const colQtyOrder = map["qty order"];
    const colQtyCut = map["qty cut"];
    const colQtyProcessed = map["qty processed"];
    const colQtyEdge = map["qty edge finish"];
    const colQtyPacked = map["qty packed"];
    const colLastAction = map["last action"];
    const colLastUser = map["last user"];
    const colLastUpdated = map["last updated"];

    if (!colQtyOrder || !colQtyCut || !colQtyProcessed || !colQtyEdge || !colQtyPacked) {
      throw new Error("Manufacture Hub missing required quantity columns.");
    }

    const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const cut = Number(rowData[colQtyCut - 1]) || 0;
    const processed = Number(rowData[colQtyProcessed - 1]) || 0;
    const edge = Number(rowData[colQtyEdge - 1]) || 0;
    let packed = Number(rowData[colQtyPacked - 1]) || 0;

    let remaining = qtyNum;

    const reducePacked = Math.min(remaining, packed);
    packed -= reducePacked;
    remaining -= reducePacked;

    const maxEdgeReduce = Math.max(0, edge - packed);
    const reduceEdge = Math.min(remaining, maxEdgeReduce);
    const edgeNext = edge - reduceEdge;
    remaining -= reduceEdge;

    const maxProcReduce = Math.max(0, processed - edgeNext);
    const reduceProc = Math.min(remaining, maxProcReduce);
    const procNext = processed - reduceProc;
    remaining -= reduceProc;

    const maxCutReduce = Math.max(0, cut - procNext);
    const reduceCut = Math.min(remaining, maxCutReduce);
    const cutNext = cut - reduceCut;
    remaining -= reduceCut;

    if (remaining > 0) {
      throw new Error("Not enough panels available in production stages to mark damaged.");
    }

    const timestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail() || "Workshop App User";

    sheet.getRange(rowIndex, colQtyPacked).setValue(packed);
    sheet.getRange(rowIndex, colQtyEdge).setValue(edgeNext);
    sheet.getRange(rowIndex, colQtyProcessed).setValue(procNext);
    sheet.getRange(rowIndex, colQtyCut).setValue(cutNext);

    if (colLastAction) sheet.getRange(rowIndex, colLastAction).setValue("Damaged");
    if (colLastUser) sheet.getRange(rowIndex, colLastUser).setValue(userEmail);
    if (colLastUpdated) sheet.getRange(rowIndex, colLastUpdated).setValue(timestamp);

    logPanelHistoryEntry_(rowData, panelInfoCols, buildDamagePayload_(qtyNum, reason, userEmail, timestamp));

    const newRow = rowData.slice();
    newRow[colQtyOrder - 1] = qtyNum;
    newRow[colQtyCut - 1] = 0;
    newRow[colQtyProcessed - 1] = 0;
    newRow[colQtyEdge - 1] = 0;
    newRow[colQtyPacked - 1] = 0;
    if (colLastAction) newRow[colLastAction - 1] = "";
    if (colLastUser) newRow[colLastUser - 1] = "";
    if (colLastUpdated) newRow[colLastUpdated - 1] = "";

    sheet.appendRow(newRow);

    const targetOrderId = rowData[panelInfoCols.orderId - 1];
    const targetProduct = rowData[panelInfoCols.productName - 1];
    const updatedData = sheet.getDataRange().getValues();
    if (targetOrderId && targetProduct) {
      syncProductCompletions_(updatedData, [{ orderId: targetOrderId, productName: targetProduct }]);
    }

    SpreadsheetApp.flush();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
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
  const panelInfoCols = getPanelInfoColumnMap_(map);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 

  try {
    // 2. APPLY UPDATES
    const results = []; // Return list for frontend
    const packedTargets = new Map();
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
        const previousValue = Number(sheet.getRange(rowIndex, colIndex).getValue()) || 0;
        // Update Quantity
        sheet.getRange(rowIndex, colIndex).setValue(value);
        
        // Update Metadata
        if (map["last action"]) sheet.getRange(rowIndex, map["last action"]).setValue(colName);
        if (map["last user"]) sheet.getRange(rowIndex, map["last user"]).setValue(userEmail);
        if (map["last updated"]) sheet.getRange(rowIndex, map["last updated"]).setValue(timestamp);

        const delta = Number(value) - previousValue;
        if (delta !== 0) {
          const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
          logPanelHistoryEntry_(rowData, panelInfoCols, buildPanelHistoryPayload_(colName, delta, userEmail, timestamp));
        }

        // Add to results array for the frontend
        results.push({
            rowIndex: rowIndex,
            colName: colName,
            value: value
        });

        // 3. TRIGGER SYNC (If Packed)
        if (colName === 'packed') {
          const rowData = sheet.getRange(rowIndex, 1, 1, 3).getValues()[0];
          const orderId = rowData[0];
          const productName = rowData[2];
          const key = `${String(orderId)}||${String(productName)}`;
          packedTargets.set(key, { orderId, productName });
        }
      }
    });

    if (packedTargets.size > 0) {
      const data = sheet.getDataRange().getValues();
      syncProductCompletions_(data, Array.from(packedTargets.values()));
    }

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

  const stockSheet = ss.getSheetByName("Component Stock");
  const hubSheet = ss.getSheetByName("Manufacture Hub");

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  const cache = CacheService.getScriptCache();
  const userEmail = Session.getActiveUser().getEmail() || "Workshop App";
  const results = [];

  try {
    const stockMap = {};
    if (stockSheet && stockSheet.getLastRow() > 1) {
      const stockData = stockSheet.getDataRange().getValues();
      for (let i = 1; i < stockData.length; i++) {
        const name = _normTxt(stockData[i][1]);
        const qty = Number(stockData[i][4]) || 0;
        if (name) stockMap[`name:${name.toLowerCase()}`] = { rowIndex: i + 1, qty };
      }
    }

    const productSkuMap = {};
    if (hubSheet && hubSheet.getLastRow() > 1) {
      const hubData = hubSheet.getDataRange().getValues();
      for (let i = 1; i < hubData.length; i++) {
        const orderId = _normTxt(hubData[i][0]);
        const productName = _normTxt(hubData[i][2]);
        const sku = _normTxt(hubData[i][3]);
        if (orderId && productName && sku) {
          productSkuMap[`${orderId}||${productName}`.toLowerCase()] = sku;
        }
      }
    }

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

      const delta = next - cur;

      if (delta > 0) {
        if (!stockSheet) {
          results.push({ rowIndex, outOfStock: true, current: cur, available: 0, message: "Component Stock tab missing" });
          return;
        }

        const compRow = sheet.getRange(rowIndex, 1, 1, 5).getValues()[0];
        const orderId = _normTxt(compRow[0]);
        const productName = _normTxt(compRow[2]);
        const compName = _normTxt(compRow[3]);
        const nameKey = compName ? `name:${compName.toLowerCase()}` : "";
        const stockEntry = nameKey ? stockMap[nameKey] : null;

        const available = stockEntry ? Number(stockEntry.qty) || 0 : 0;
        if (!stockEntry || available < delta) {
          results.push({ rowIndex, outOfStock: true, current: cur, available, message: "Insufficient component stock" });
          return;
        }

        const newStockQty = available - delta;
        stockSheet.getRange(stockEntry.rowIndex, 5).setValue(newStockQty);
        stockEntry.qty = newStockQty;

        if (nameKey && stockMap[nameKey]) stockMap[nameKey].qty = newStockQty;

        const productSku = productSkuMap[`${orderId}||${productName}`.toLowerCase()] || "";
        const reason = `Allocated to ${productSku || productName || "Product"}`;
        logStockTransaction(compName || "Component", -delta, reason, orderId || "Project");
      }

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

function updateFurnitureStock(sku, productName, qtyChange, reason) {
  const sheet = _getFurnitureStockSheet_(true);

  const data = sheet.getDataRange().getValues();
  const targetSku = String(sku).trim().toLowerCase();
  let foundRow = -1;

  for (let i = 1; i < data.length; i++) {
    const rowSku = String(data[i][1]).trim().toLowerCase(); // Col B
    if (rowSku === targetSku) { foundRow = i + 1; break; }
  }

  const n = v => Number(v) || 0;
  const delta = n(qtyChange);

  if (foundRow > 0) {
    const currentQty = n(sheet.getRange(foundRow, 4).getValue()); // Col D
    sheet.getRange(foundRow, 4).setValue(currentQty + delta);
    sheet.getRange(foundRow, 3).setValue(productName);
  } else {
    sheet.appendRow(["Workshop Stock", sku, productName, delta, ""]);
  }

  // Better history shape (Material = SKU, Source = Furniture Stock)
  logStockTransaction(String(sku).trim(), delta, reason, "Furniture Stock");

  return { success: true, sku, newQtyChange: delta };
}



// --- APP VIEW FETCHERS ---

function getFurnitureStock() {
  const sheet = _getFurnitureStockSheet_(false);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const stock = [];

  for (let i = 1; i < data.length; i++) {
    const qty = Number(data[i][3]) || 0; // Col D
    if (qty > 0) {
      stock.push({
        rowIndex: i + 1,
        customer: data[i][0],
        sku: data[i][1],
        product: data[i][2],
        qty: qty
      });
    }
  }
  return stock;
}


function adjustFurnitureStock(rowIndex, change, reason) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = _getFurnitureStockSheet_(false);
    if (!sheet) throw new Error("Furniture Stock tab missing");

    const current = sheet.getRange(rowIndex, 4).getValue(); // Col D
    const newVal = (Number(current) || 0) + Number(change);
    if (newVal < 0) throw new Error("Stock cannot be negative");

    sheet.getRange(rowIndex, 4).setValue(newVal);

    const sku = String(sheet.getRange(rowIndex, 2).getValue() || "").trim(); // Col B
    const prodName = sheet.getRange(rowIndex, 3).getValue(); // Col C

    logStockTransaction(sku || String(prodName || "").trim(), change, reason, "Furniture Stock");
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
  const sheet = _getFurnitureStockSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => _normTxt(h).toLowerCase());
  const hm = _headerMap_(headers);

  const idxCustomer = hm["customer name"];
  const idxSku = hm["sku"];
  const idxAvail = hm["qty available"];
  if (idxCustomer == null || idxSku == null || idxAvail == null) return {};

  const map = {};
  for (let r = 1; r < data.length; r++) {
    const customer = _normTxt(data[r][idxCustomer]);
    const sku = _normTxt(data[r][idxSku]);
    if (!sku) continue;
    if (customer.toLowerCase() !== "workshop stock") continue;

    const avail = Number(data[r][idxAvail]) || 0;
    map[sku.toLowerCase()] = { rowIndex: r + 1, sku, avail };
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

function _classifyMaterial_(material) {
  const mat = _normTxt(material).toLowerCase();
  if (!mat) return "unknown";
  if (mat.includes("component") || mat.includes("hardware")) return "component";
  if (mat.endsWith("mdf")) return "mdf";
  if (mat.endsWith("ply")) return "ply";
  return "unknown";
}

function getSmartOrderSummary(orderId) {
  const orderSheet = _getShopifyOrdersSheet();
  const orderValues = orderSheet.getDataRange().getValues();

  let targetOrderId = _normTxt(orderId);
  if (!targetOrderId) {
    if (orderValues.length < 2) throw new Error("Order ID is required. Shopify Orders has no data rows.");

    const headers = orderValues[0].map(h => _normTxt(h).toLowerCase());
    const hm = _headerMap_(headers);
    const idxOrderId = hm["order id"];
    if (idxOrderId == null) {
      throw new Error("Shopify Orders headers don't match expected names. Missing 'Order ID'.");
    }

    for (let r = 1; r < orderValues.length; r++) {
      const candidateOrderId = _normTxt(orderValues[r][idxOrderId]);
      if (candidateOrderId) {
        targetOrderId = candidateOrderId;
        break;
      }
    }

    if (!targetOrderId) throw new Error("Order ID is required. No non-empty Order ID values were found.");
  }
  if (orderValues.length < 2) return { orderId: targetOrderId, products: [], totals: {}, totalToBuild: 0 };

  const headers = orderValues[0].map(h => _normTxt(h).toLowerCase());
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

  const orderLines = [];
  let customerName = "";

  for (let r = 1; r < orderValues.length; r++) {
    const rowOrderId = _normTxt(orderValues[r][idxOrderId]);
    if (rowOrderId !== targetOrderId) continue;

    const qtyOrdered = Number(orderValues[r][idxQty]) || 0;
    const status = _normTxt(orderValues[r][idxStatus]);
    const allocated = _extractAllocatedFromStatus_(status);
    const toBuild = Math.max(0, qtyOrdered - allocated);

    if (toBuild <= 0) continue;

    orderLines.push({
      productName: _normTxt(orderValues[r][idxProductName]),
      sku: _normTxt(orderValues[r][idxSku]).toUpperCase(),
      qtyOrdered,
      allocated,
      toBuild
    });

    if (!customerName) customerName = _normTxt(orderValues[r][idxCustomer]);
  }

  if (orderLines.length === 0) {
    return { orderId: targetOrderId, customer: customerName, products: [], totals: {}, totalToBuild: 0 };
  }

  const prodSS = SpreadsheetApp.openById(PRODUCT_RECIPE_SHEET_ID);
  const productSheet = prodSS.getSheetByName(PRODUCT_RECIPE_TAB_NAME);
  if (!productSheet) throw new Error("Product recipe tab not found.");
  const productData = productSheet.getDataRange().getValues();
  const productMap = {};

  for (let p = 1; p < productData.length; p++) {
    const sku = _normTxt(productData[p][0]).toUpperCase();
    if (!sku) continue;
    if (!productMap[sku]) productMap[sku] = [];
    productMap[sku].push(productData[p]);
  }

  const inventory = getInventoryData();
  const compStockMap = {};
  const woodStockList = [];
  const edgeStockList = [];

  (inventory.components || []).forEach(item => {
    const nameKey = _normTxt(item.name).toLowerCase();
    if (!nameKey) return;
    compStockMap[nameKey] = (compStockMap[nameKey] || 0) + (Number(item.stock) || 0);
  });

  (inventory.wood || []).forEach(item => {
    const mat = _normTxt(item.material).toLowerCase();
    if (!mat) return;
    woodStockList.push({ material: mat, qty: Number(item.qty) || 0 });
  });

  (inventory.edge || []).forEach(item => {
    const mat = _normTxt(item.material).toLowerCase();
    if (!mat) return;
    edgeStockList.push({ material: mat, rolls: Number(item.rolls) || 0 });
  });

  const sumMatchingStock = (material, list, field) => {
    const key = _normTxt(material).toLowerCase();
    if (!key) return 0;
    return list.reduce((sum, item) => {
      if (!item.material) return sum;
      if (item.material.includes(key) || key.includes(item.material)) {
        return sum + (Number(item[field]) || 0);
      }
      return sum;
    }, 0);
  };

  const MDF_SHEET_AREA = 2.8 * 2.07;
  const PLY_SHEET_AREA = 3.05 * 1.22;
  const EDGE_ROLL_METERS = 75;

  const totalsCompMap = {};
  const totalsWoodMap = {};
  const totalsEdgeMap = {};

  const products = orderLines.map(line => {
    const recipeRows = productMap[line.sku] || [];
    const compMap = {};
    const woodMap = {};
    const edgeMap = {};

    recipeRows.forEach(row => {
      const panelName = _normTxt(row[2]);
      const material = _normTxt(row[3]);
      const qtyPerUnit = Number(row[6]) || 0;
      const totalUnits = qtyPerUnit * line.toBuild;
      if (!totalUnits) return;

      // Recipe column H / I values are already per finished product row
      // (panel area/perimeter values include the panel quantity in column G).
      // Only scale by how many units of the product we need to build.
      const areaPerProduct = Number(row[7]) || 0;
      const perimeterPerProduct = Number(row[8]) || 0;
      const totalArea = areaPerProduct * line.toBuild;
      const totalPerimeter = perimeterPerProduct * line.toBuild;
      const matType = _classifyMaterial_(material);
      const materialKey = material.toLowerCase();

      if (matType === "component") {
        if (!compMap[panelName]) compMap[panelName] = { name: panelName, qty: 0 };
        compMap[panelName].qty += totalUnits;
        if (!totalsCompMap[panelName]) totalsCompMap[panelName] = { name: panelName, qty: 0 };
        totalsCompMap[panelName].qty += totalUnits;
      } else {
        if (!woodMap[materialKey]) {
          woodMap[materialKey] = { material, type: matType, area: 0 };
        }
        woodMap[materialKey].area += totalArea;

        if (!totalsWoodMap[materialKey]) {
          totalsWoodMap[materialKey] = { material, type: matType, area: 0 };
        }
        totalsWoodMap[materialKey].area += totalArea;

        if (matType === "mdf") {
          if (!edgeMap[materialKey]) {
            edgeMap[materialKey] = { material, meters: 0 };
          }
          edgeMap[materialKey].meters += totalPerimeter;

          if (!totalsEdgeMap[materialKey]) {
            totalsEdgeMap[materialKey] = { material, meters: 0 };
          }
          totalsEdgeMap[materialKey].meters += totalPerimeter;
        }
      }
    });

    const components = Object.values(compMap).map(item => {
      const stock = compStockMap[item.name.toLowerCase()] || 0;
      return {
        name: item.name,
        qty: item.qty,
        stock,
        short: Math.max(0, item.qty - stock)
      };
    });

    const wood = Object.values(woodMap).map(item => {
      const sheetArea = item.type === "mdf" ? MDF_SHEET_AREA : item.type === "ply" ? PLY_SHEET_AREA : 0;
      const sheetsNeeded = sheetArea ? Math.ceil(item.area / sheetArea) : 0;
      const stockSheets = sumMatchingStock(item.material, woodStockList, "qty");
      return {
        material: item.material,
        type: item.type,
        area: item.area,
        sheetArea,
        sheetsNeeded,
        stockSheets,
        shortSheets: Math.max(0, sheetsNeeded - stockSheets)
      };
    });

    const edge = Object.values(edgeMap).map(item => {
      const rollsNeeded = Math.ceil(item.meters / EDGE_ROLL_METERS);
      const stockRolls = sumMatchingStock(item.material, edgeStockList, "rolls");
      return {
        material: item.material,
        meters: item.meters,
        rollsNeeded,
        stockRolls,
        shortRolls: Math.max(0, rollsNeeded - stockRolls)
      };
    });

    return {
      productName: line.productName,
      sku: line.sku,
      qtyRequired: line.qtyOrdered,
      qtyAllocated: line.allocated,
      toBuild: line.toBuild,
      components,
      wood,
      edge
    };
  });

  const totals = {
    components: Object.values(totalsCompMap).map(item => {
      const stock = compStockMap[item.name.toLowerCase()] || 0;
      return {
        name: item.name,
        qty: item.qty,
        stock,
        short: Math.max(0, item.qty - stock)
      };
    }),
    wood: Object.values(totalsWoodMap).map(item => {
      const sheetArea = item.type === "mdf" ? MDF_SHEET_AREA : item.type === "ply" ? PLY_SHEET_AREA : 0;
      const sheetsNeeded = sheetArea ? Math.ceil(item.area / sheetArea) : 0;
      const stockSheets = sumMatchingStock(item.material, woodStockList, "qty");
      return {
        material: item.material,
        type: item.type,
        area: item.area,
        sheetArea,
        sheetsNeeded,
        stockSheets,
        shortSheets: Math.max(0, sheetsNeeded - stockSheets)
      };
    }),
    edge: Object.values(totalsEdgeMap).map(item => {
      const rollsNeeded = Math.ceil(item.meters / EDGE_ROLL_METERS);
      const stockRolls = sumMatchingStock(item.material, edgeStockList, "rolls");
      return {
        material: item.material,
        meters: item.meters,
        rollsNeeded,
        stockRolls,
        shortRolls: Math.max(0, rollsNeeded - stockRolls)
      };
    })
  };

  const totalToBuild = orderLines.reduce((sum, line) => sum + (Number(line.toBuild) || 0), 0);

  return {
    orderId: targetOrderId,
    customer: customerName,
    products,
    totals,
    totalToBuild,
    sheetSizes: {
      mdf: "2800 x 2070mm",
      ply: "3050 x 1220mm"
    },
    edgeRollMeters: EDGE_ROLL_METERS
  };
}



function exportSmartOrderToSheets(orderId) {
  const summary = getSmartOrderSummary(orderId);
  if (!summary || !summary.products || summary.products.length === 0) {
    throw new Error("No Smart Order items to export for this order.");
  }

  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
  const fileName = `Smart Order - ${summary.orderId} - ${timestamp}`;

  const out = SpreadsheetApp.create(fileName);
  const summarySheet = out.getSheets()[0];
  summarySheet.setName("Summary");

  const generatedOn = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");
  let row = 1;

  summarySheet.getRange(row++, 1).setValue("Smart Order Export").setFontWeight("bold").setFontSize(14);
  summarySheet.getRange(row++, 1).setValue(`Order ID: ${summary.orderId}`);
  summarySheet.getRange(row++, 1).setValue(`Customer: ${summary.customer || "Unknown"}`);
  summarySheet.getRange(row++, 1).setValue(`Total units to build: ${summary.totalToBuild || 0}`);
  summarySheet.getRange(row++, 1).setValue(`Generated: ${generatedOn}`);
  row += 1;

  row = _writeSmartOrderTableSection_(summarySheet, row, "ORDER PRODUCTS", ["Product Name", "Product SKU", "Quantity Required", "Quantity Allocated", "Quantity to Build"],
    (summary.products || []).map(prod => [
      prod.productName || "",
      prod.sku || "",
      Number(prod.qtyRequired) || 0,
      Number(prod.qtyAllocated) || 0,
      Number(prod.toBuild) || 0
    ])
  );

  row = _writeSmartOrderTableSection_(summarySheet, row, "TOTAL WOOD REQUIREMENTS", ["Material", "Type", "Sheet Size", "Area (m²)", "Sheets Req", "In Stock", "Needed"],
    (summary.totals && summary.totals.wood ? summary.totals.wood : []).map(item => [
      item.material,
      item.type,
      item.type === "mdf" ? (summary.sheetSizes && summary.sheetSizes.mdf ? summary.sheetSizes.mdf : "2800 x 2070mm")
        : item.type === "ply" ? (summary.sheetSizes && summary.sheetSizes.ply ? summary.sheetSizes.ply : "3050 x 1220mm")
        : "",
      Number(item.area) || 0,
      Number(item.sheetsNeeded) || 0,
      Number(item.stockSheets) || 0,
      Number(item.shortSheets) || 0
    ])
  );

  row = _writeSmartOrderTableSection_(summarySheet, row, "TOTAL EDGE REQUIREMENTS", ["Material", "Meters", `Rolls (${summary.edgeRollMeters || 75}m)`, "In Stock", "Needed"],
    (summary.totals && summary.totals.edge ? summary.totals.edge : []).map(item => [
      item.material,
      Number(item.meters) || 0,
      Number(item.rollsNeeded) || 0,
      Number(item.stockRolls) || 0,
      Number(item.shortRolls) || 0
    ])
  );

  _writeSmartOrderTableSection_(summarySheet, row, "TOTAL COMPONENT REQUIREMENTS", ["Component", "Qty Req", "In Stock", "Needed"],
    (summary.totals && summary.totals.components ? summary.totals.components : []).map(item => [
      item.name,
      Number(item.qty) || 0,
      Number(item.stock) || 0,
      Number(item.short) || 0
    ])
  );

  (summary.products || []).forEach((prod, idx) => {
    const sheetName = _smartOrderSheetName_(`${idx + 1}. ${prod.productName || prod.sku || "Product"}`);
    const sh = out.insertSheet(sheetName);
    let r = 1;

    sh.getRange(r++, 1).setValue("Smart Order Product Breakdown").setFontWeight("bold").setFontSize(13);
    sh.getRange(r++, 1).setValue(`Order ID: ${summary.orderId}`);
    sh.getRange(r++, 1).setValue(`Customer: ${summary.customer || "Unknown"}`);
    sh.getRange(r++, 1).setValue(`Product: ${prod.productName || prod.sku || "Unknown"}`);
    sh.getRange(r++, 1).setValue(`SKU: ${prod.sku || ""}`);
    sh.getRange(r++, 1).setValue(`Units to build: ${Number(prod.toBuild) || 0}`);
    sh.getRange(r++, 1).setValue(`Generated: ${generatedOn}`);
    r += 1;

    r = _writeSmartOrderTableSection_(sh, r, "WOOD", ["Material", "Type", "Sheet Size", "Area (m²)", "Sheets Req", "In Stock", "Needed"],
      (prod.wood || []).map(item => [
        item.material,
        item.type,
        item.type === "mdf" ? (summary.sheetSizes && summary.sheetSizes.mdf ? summary.sheetSizes.mdf : "2800 x 2070mm")
          : item.type === "ply" ? (summary.sheetSizes && summary.sheetSizes.ply ? summary.sheetSizes.ply : "3050 x 1220mm")
          : "",
        Number(item.area) || 0,
        Number(item.sheetsNeeded) || 0,
        Number(item.stockSheets) || 0,
        Number(item.shortSheets) || 0
      ])
    );

    r = _writeSmartOrderTableSection_(sh, r, "EDGEBAND", ["Material", "Meters", `Rolls (${summary.edgeRollMeters || 75}m)`, "In Stock", "Needed"],
      (prod.edge || []).map(item => [
        item.material,
        Number(item.meters) || 0,
        Number(item.rollsNeeded) || 0,
        Number(item.stockRolls) || 0,
        Number(item.shortRolls) || 0
      ])
    );

    _writeSmartOrderTableSection_(sh, r, "COMPONENTS", ["Component", "Qty Req", "In Stock", "Needed"],
      (prod.components || []).map(item => [
        item.name,
        Number(item.qty) || 0,
        Number(item.stock) || 0,
        Number(item.short) || 0
      ])
    );

    sh.autoResizeColumns(1, 7);
  });

  summarySheet.autoResizeColumns(1, 7);

  return {
    spreadsheetId: out.getId(),
    url: out.getUrl(),
    name: out.getName()
  };
}

function exportRoomListToSheets(orderId) {
  const targetOrderId = String(orderId || "").trim();
  if (!targetOrderId) throw new Error("Missing order ID for room list export.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const deliverySheet = ss.getSheetByName("Delivery Hub");
  if (!deliverySheet) throw new Error("Delivery Hub tab is missing.");

  const data = deliverySheet.getDataRange().getValues();
  if (!data || data.length <= 1) throw new Error("No delivery data found.");

  const norm = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim();
  const rows = [];
  let customerName = "";

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (norm(row[0]) !== targetOrderId) continue;

    const customer = String(row[1] || "").trim();
    if (!customerName && customer) customerName = customer;

    const product = String(row[3] || "").trim();
    const room = canonicalRoomName_(row[4]);
    const status = norm(row[5]) || "Pending";

    if (!room) continue;
    rows.push({ room, product, status });
  }

  if (rows.length === 0) {
    throw new Error("No room assignments found for this order.");
  }

  rows.sort((a, b) => {
    const roomCmp = a.room.localeCompare(b.room);
    if (roomCmp !== 0) return roomCmp;
    const productCmp = a.product.localeCompare(b.product);
    if (productCmp !== 0) return productCmp;
    return a.status.localeCompare(b.status);
  });

  const grouped = new Map();
  rows.forEach(item => {
    const key = `${item.room}|||${item.product}|||${item.status}`;
    grouped.set(key, (grouped.get(key) || 0) + 1);
  });

  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
  const generatedOn = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");
  const out = SpreadsheetApp.create(`Room List - ${targetOrderId} - ${timestamp}`);
  const sh = out.getSheets()[0];
  sh.setName("Room List");

  let r = 1;
  sh.getRange(r++, 1).setValue("Room List Export").setFontWeight("bold").setFontSize(14);
  sh.getRange(r++, 1).setValue(`Order ID: ${targetOrderId}`);
  sh.getRange(r++, 1).setValue(`Customer: ${customerName || "Unknown"}`);
  sh.getRange(r++, 1).setValue(`Generated: ${generatedOn}`);
  r += 1;

  const headerRow = r;
  sh.getRange(r, 1, 1, 4).setValues([["Room", "Product", "Qty", "Status"]]);
  sh.getRange(r, 1, 1, 4).setFontWeight("bold").setBackground("#f1f3f4");
  r++;

  const tableRows = Array.from(grouped.entries()).map(([key, qty]) => {
    const [room, product, status] = key.split("|||");
    return [room, product, qty, status];
  });

  sh.getRange(r, 1, tableRows.length, 4).setValues(tableRows);

  const lastRow = r + tableRows.length - 1;
  sh.getRange(headerRow, 1, lastRow - headerRow + 1, 4).createFilter();
  sh.autoResizeColumns(1, 4);

  return {
    spreadsheetId: out.getId(),
    url: out.getUrl(),
    rowsExported: tableRows.length
  };
}

function _writeSmartOrderTableSection_(sheet, startRow, title, headers, rows) {
  let row = startRow;
  sheet.getRange(row, 1).setValue(title).setFontWeight("bold");
  row += 1;

  if (!rows || rows.length === 0) {
    sheet.getRange(row, 1).setValue("No items.").setFontStyle("italic").setFontColor("#666666");
    return row + 2;
  }

  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f1f3f4");
  row += 1;

  sheet.getRange(row, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(row, 1, rows.length, headers.length).setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);

  return row + rows.length + 2;
}

function _smartOrderSheetName_(name) {
  const cleaned = String(name || "Product")
    .replace(/[\\/?*\[\]:]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned.substring(0, 99) || "Product";
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
  roomName = canonicalRoomName_(roomName);
  if (!roomName) return "Error: Room name is required";

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
  roomName = canonicalRoomName_(roomName);

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
    if (roomKey_(row[4]) !== roomKey_(roomName)) continue;
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
  roomName = canonicalRoomName_(roomName);

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
    if (roomKey_(row[4]) !== roomKey_(roomName)) continue;
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
  const materialLabel = String(name || "").trim();
  logStockTransaction(materialLabel || `Component: ${sku} | ${name}`, change, reason);

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
function logStockTransaction(material, change, reason, sourceOverride) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName("Stock History");
  if (!historySheet) return;

  // Ensure header has columns: Timestamp, User, Source, Material, Change, Reason
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

  // Source: explicit override wins, otherwise infer from material prefix
  let source = String(sourceOverride || "").trim();
  if (!source) {
    source = "Other";
    const m = String(material || "").toLowerCase();
    if (m.startsWith("edge:")) source = "Edgebanding";
    else if (m.startsWith("component:")) source = "Components";
    else if (m.startsWith("furniture stock:") || m.startsWith("finished goods:") || m.startsWith("finished good:")) source = "Furniture Stock";
    else source = "Wood";
  }

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
    const isComponentMatch = rowMat.startsWith("component:") && rowMat.includes(`| ${searchMat}`);

    if (rowMat === searchMat || isComponentMatch) {
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

function syncProductCompletions_(data, targets) {
  if (!Array.isArray(data) || data.length < 2) return;
  if (!Array.isArray(targets) || targets.length === 0) return;

  const norm = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim().toLowerCase();
  const orderProductMap = {};
  const panelTotalsByCustomerSku = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const orderId = norm(row[0]);
    const customer = norm(row[1]);
    const productName = norm(row[2]);
    const sku = norm(row[3]);
    const panelName = String(row[4] || "");
    const qtyPerUnit = Number(row[6]) || 1;
    const qtyPacked = Number(row[15]) || 0;

    if (orderId && productName) {
      orderProductMap[`${orderId}||${productName}`] = {
        customer,
        sku,
        productName: row[2]
      };
    }

    if (customer && sku && panelName) {
      const key = `${customer}||${sku}`;
      if (!panelTotalsByCustomerSku[key]) panelTotalsByCustomerSku[key] = {};
      if (!panelTotalsByCustomerSku[key][panelName]) {
        panelTotalsByCustomerSku[key][panelName] = { packed: 0, required: qtyPerUnit };
      }
      panelTotalsByCustomerSku[key][panelName].packed += qtyPacked;
    }
  }

  targets.forEach(({ orderId, productName }) => {
    const key = `${norm(orderId)}||${norm(productName)}`;
    const meta = orderProductMap[key];
    if (!meta || meta.customer !== "workshop stock" || !meta.sku) return;

    const panelTotals = panelTotalsByCustomerSku[`${meta.customer}||${meta.sku}`];
    if (!panelTotals) return;

    let totalFinished = Infinity;
    let hasPanels = false;
    Object.keys(panelTotals).forEach((panelKey) => {
      hasPanels = true;
      const p = panelTotals[panelKey];
      const sets = Math.floor((Number(p.packed) || 0) / (Number(p.required) || 1));
      if (sets < totalFinished) totalFinished = sets;
    });

    if (!hasPanels || totalFinished === Infinity) totalFinished = 0;
    syncToFinishedGoods(meta.customer, meta.sku, meta.productName, totalFinished);
  });
}

function syncToFinishedGoods(customer, sku, productName, newManufacturedTotal) {
  const sheet = _getFurnitureStockSheet_(true);

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

  const n = (v) => Number(v) || 0;
  const made = n(newManufacturedTotal);

  // Only log "Completed Manufacture" for Workshop Stock rows
  const shouldLogCompletion = (targetCust === "workshop stock");
  const cleanSku = String(sku || "").trim();

  if (foundRow > 0) {
    const curAvail = n(sheet.getRange(foundRow, 4).getValue()); // Col D
    const curMade  = n(sheet.getRange(foundRow, 5).getValue()); // Col E

    const delta = Math.max(0, made - curMade);

    // Update Manufactured Total (Col E)
    sheet.getRange(foundRow, 5).setValue(made);

    // Increase Available by delta (Col D)
    if (delta > 0) {
      sheet.getRange(foundRow, 4).setValue(curAvail + delta);

      // Stock history entry for completions
      if (shouldLogCompletion && cleanSku) {
        logStockTransaction(cleanSku, delta, "Completed Manufacture", "Workshop Stock");
      }
    }

    // Keep product name fresh (Col C)
    sheet.getRange(foundRow, 3).setValue(productName);

  } else {
    // New row: set Available = Manufactured Total
    if (made > 0) {
      sheet.appendRow([customer, sku, productName, made, made]);

      if (shouldLogCompletion && cleanSku) {
        logStockTransaction(cleanSku, made, "Completed Manufacture", "Workshop Stock");
      }
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
  const fgSheet = _getFurnitureStockSheet_(false);
   if (!fgSheet) throw new Error("Furniture Stock tab missing");


  const fgData = fgSheet.getDataRange().getValues();
  const fgHeaders = fgData[0].map(h => _normTxt(h).toLowerCase());
  const fhm = _headerMap_(fgHeaders);

  const idxFgCustomer = fhm["customer name"];
  const idxFgSku = fhm["sku"];
  const idxAvail = fhm["qty available"];
  if ([idxFgCustomer, idxFgSku, idxAvail].some(x => x == null)) {
    throw new Error("Furniture Stock headers must include: Customer Name, SKU, Qty Available");
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
  if (fgRowIndex < 0) throw new Error(`No Workshop Stock found in Furniture Stock for SKU: ${sku}`);

  const currentAvail = Number(fgSheet.getRange(fgRowIndex, idxAvail + 1).getValue()) || 0;
  if (qtyToAllocate > currentAvail) throw new Error(`Not enough stock. Available: ${currentAvail}`);

  // Deduct stock
  fgSheet.getRange(fgRowIndex, idxAvail + 1).setValue(currentAvail - qtyToAllocate);

  // Log stock history
  const reason = `Allocation to ${orderId} (${customer})`;
  logStockTransaction(`Furniture Stock: ${sku} ${productName}`.trim(), -qtyToAllocate, reason);

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

function authorizeSmartOrderAccess() {
  const orderSheet = _getShopifyOrdersSheet();
  const values = orderSheet.getDataRange().getValues();

  if (values.length < 2) {
    SpreadsheetApp.openById(PRODUCT_RECIPE_SHEET_ID).getSheetByName(PRODUCT_RECIPE_TAB_NAME);
    return {
      ok: true,
      message: "Authorization completed. Shopify Orders has no data rows yet."
    };
  }

  const headers = values[0].map(h => _normTxt(h).toLowerCase());
  const hm = _headerMap_(headers);
  const idxOrderId = hm["order id"];
  if (idxOrderId == null) {
    throw new Error("Shopify Orders headers don't match expected names. Missing 'Order ID'.");
  }

  let sampleOrderId = "";
  for (let r = 1; r < values.length; r++) {
    sampleOrderId = _normTxt(values[r][idxOrderId]);
    if (sampleOrderId) break;
  }

  if (!sampleOrderId) {
    SpreadsheetApp.openById(PRODUCT_RECIPE_SHEET_ID).getSheetByName(PRODUCT_RECIPE_TAB_NAME);
    return {
      ok: true,
      message: "Authorization completed. Add an Order ID row before testing Smart Order output."
    };
  }

  const summary = getSmartOrderSummary(sampleOrderId);
  return {
    ok: true,
    orderId: sampleOrderId,
    totalToBuild: Number(summary.totalToBuild) || 0,
    message: "Authorization completed."
  };
}
