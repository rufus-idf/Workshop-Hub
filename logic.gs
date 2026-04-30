const DISCORD_WEBHOOK_URL = "https://discord.com/api/webhooks/1498336458258513971/i7T7hLbe_EWlxKDUMvwyqoezT7ixZ9FF0_BYWdkOVWeysGMhWj8J2CA2S8V9VzYSRIyG";
const SHOPIFY_ORDERS_SHEET_ID = "1KDDVnIZ5oCruCY4nyKp6XVqwWnL7-N6f-pKAD3xXo_U";
const SHOPIFY_ORDERS_TAB_NAME = "Sheet1";
const FURNITURE_STOCK_TAB_NAME = "Furniture Stock";
const LEGACY_FINISHED_GOODS_TAB_NAME = "Finished Goods";
const OFFCUT_SHEET_ID = "1-qS6gWekGtEhjczboAyAShJHHamK0ZuVlR7CFbubxxo";
const OFFCUT_INVENTORY_TAB_NAME = "offcut_inventory";
const OFFCUT_SHAPES_TAB_NAME = "offcut_shapes";
const OFFCUT_TEXTURE_LIBRARY_TAB_NAME = "texture_library";

function squeezeSpaces_(v) {
  return String(v ?? "").replace(/\u00A0/g, " ").trim().replace(/\s+/g, " ");
}

function normalizeEdgeMaterialName_(value) {
  return squeezeSpaces_(value)
    .toLowerCase()
    .replace(/\b(edgebanding|edge\s*band|edgeband|edge|mdf)\b/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function getEdgeBandColumnMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => squeezeSpaces_(h).toLowerCase());

  const findIndex = (names, fallback) => {
    for (const name of names) {
      const idx = headers.indexOf(name);
      if (idx !== -1) return idx + 1;
    }
    return fallback;
  };

  return {
    material: findIndex(["material"], 1),
    thickness: findIndex(["thickness (mm)", "thickness"], 2),
    rollLength: findIndex(["roll length", "roll length (m)"], 3),
    onOrder: findIndex(["rolls on order", "on order"], 4),
    rolls: findIndex(["rolls in stock", "in stock"], 5),
    price: findIndex(["price", "price (£)", "cost per roll", "cost"], null)
  };
}

function getHeaderIndexMap_(headers) {
  const map = {};
  (headers || []).forEach((header, index) => {
    const key = squeezeSpaces_(header).toLowerCase();
    if (key) map[key] = index;
  });
  return map;
}

function getRowValueByHeaders_(row, headerMap, names) {
  const headerNames = Array.isArray(names) ? names : [names];
  for (let i = 0; i < headerNames.length; i++) {
    const key = squeezeSpaces_(headerNames[i]).toLowerCase();
    if (Object.prototype.hasOwnProperty.call(headerMap, key)) {
      return row[headerMap[key]];
    }
  }
  return "";
}

function parseJsonArraySafe_(value, fallback) {
  if (Array.isArray(value)) return value;
  const raw = String(value || "").trim();
  if (!raw) return fallback;
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : fallback;
  } catch (err) {
    return fallback;
  }
}

function getOffcutTextureLibrary_(ss) {
  const sheet = ss.getSheetByName(OFFCUT_TEXTURE_LIBRARY_TAB_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const data = sheet.getDataRange().getValues();
  const map = getHeaderIndexMap_(data[0]);
  const library = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const material = squeezeSpaces_(getRowValueByHeaders_(row, map, "material"));
    if (!material) continue;

    const activeRaw = String(getRowValueByHeaders_(row, map, "active") || "").trim().toLowerCase();
    const isActive = activeRaw === "true" || activeRaw === "yes" || activeRaw === "y" || activeRaw === "1";
    if (!isActive) continue;

    const textureUrl = squeezeSpaces_(getRowValueByHeaders_(row, map, "texture_url"));
    const solidColorHex = squeezeSpaces_(getRowValueByHeaders_(row, map, "solid_color_hex"));
    const textureType = squeezeSpaces_(getRowValueByHeaders_(row, map, "texture_type")).toUpperCase();

    library[material.toLowerCase()] = {
      textureType: textureType,
      textureUrl: textureUrl,
      solidColorHex: /^#?[0-9a-f]{3,8}$/i.test(solidColorHex) ? (solidColorHex.startsWith("#") ? solidColorHex : `#${solidColorHex}`) : "",
      hasTexture: /^https?:\/\//i.test(textureUrl)
    };
  }

  return library;
}

function getExternalOffcutInventory_() {
  try {
    const ss = SpreadsheetApp.openById(OFFCUT_SHEET_ID);
    const inventorySheet = ss.getSheetByName(OFFCUT_INVENTORY_TAB_NAME);
    const shapesSheet = ss.getSheetByName(OFFCUT_SHAPES_TAB_NAME);
    const textureLibrary = getOffcutTextureLibrary_(ss);

    if (!inventorySheet || inventorySheet.getLastRow() < 2) return [];

    const inventoryData = inventorySheet.getDataRange().getValues();
    const inventoryMap = getHeaderIndexMap_(inventoryData[0]);

    const shapesByOffcutId = {};
    const shapesByRef = {};
    if (shapesSheet && shapesSheet.getLastRow() > 1) {
      const shapesData = shapesSheet.getDataRange().getValues();
      const shapesMap = getHeaderIndexMap_(shapesData[0]);

      for (let i = 1; i < shapesData.length; i++) {
        const row = shapesData[i];
        const offcutId = squeezeSpaces_(getRowValueByHeaders_(row, shapesMap, "offcut_id"));
        const shapeRef = squeezeSpaces_(getRowValueByHeaders_(row, shapesMap, "shape_ref"));
        const shapeRecord = {
          shapeRef: shapeRef,
          coordUnit: squeezeSpaces_(getRowValueByHeaders_(row, shapesMap, "coord_unit")),
          bboxXmm: Number(getRowValueByHeaders_(row, shapesMap, "bbox_x_mm")) || 0,
          bboxYmm: Number(getRowValueByHeaders_(row, shapesMap, "bbox_y_mm")) || 0,
          verticesJson: parseJsonArraySafe_(getRowValueByHeaders_(row, shapesMap, "vertices_json"), []),
          holesJson: parseJsonArraySafe_(getRowValueByHeaders_(row, shapesMap, "holes_json"), []),
          version: getRowValueByHeaders_(row, shapesMap, "version")
        };

        if (offcutId) shapesByOffcutId[offcutId] = shapeRecord;
        if (shapeRef) shapesByRef[shapeRef] = shapeRecord;
      }
    }

    const offcuts = [];
    for (let i = 1; i < inventoryData.length; i++) {
      const row = inventoryData[i];
      const offcutId = squeezeSpaces_(getRowValueByHeaders_(row, inventoryMap, "offcut_id"));
      const status = squeezeSpaces_(getRowValueByHeaders_(row, inventoryMap, "status")).toUpperCase();
      const qty = Number(getRowValueByHeaders_(row, inventoryMap, "qty")) || 0;
      if (!offcutId || status !== "IN_STOCK" || qty <= 0) continue;

      const bboxW = Number(getRowValueByHeaders_(row, inventoryMap, "bbox_w_mm")) || 0;
      const bboxH = Number(getRowValueByHeaders_(row, inventoryMap, "bbox_h_mm")) || 0;
      const lengthMm = Math.max(bboxW, bboxH);
      const widthMm = Math.min(bboxW, bboxH);
      const areaMm2 = Number(getRowValueByHeaders_(row, inventoryMap, "area_mm2")) || 0;
      const shapeRef = squeezeSpaces_(getRowValueByHeaders_(row, inventoryMap, "shape_ref"));
      const shape = shapesByOffcutId[offcutId] || shapesByRef[shapeRef] || null;
      const material = squeezeSpaces_(getRowValueByHeaders_(row, inventoryMap, "material"));
      const textureEntry = textureLibrary[material.toLowerCase()] || null;

      offcuts.push({
        rowIndex: `offcut:${offcutId}`,
        externalOffcutId: offcutId,
        material: material,
        type: "Offcut",
        shapeType: squeezeSpaces_(getRowValueByHeaders_(row, inventoryMap, "shape_type")) || "POLYGON",
        length: lengthMm,
        width: widthMm,
        size: `${lengthMm} x ${widthMm}`,
        areaMm2: areaMm2,
        areaM2: areaMm2 / 1000000,
        onOrder: 0,
        qty: qty,
        status: status,
        thicknessMm: Number(getRowValueByHeaders_(row, inventoryMap, "thickness_mm")) || 0,
        grade: squeezeSpaces_(getRowValueByHeaders_(row, inventoryMap, "grade")),
        location: squeezeSpaces_(getRowValueByHeaders_(row, inventoryMap, "location")),
        shapeRef: shapeRef,
        textureType: textureEntry ? textureEntry.textureType : "",
        textureUrl: textureEntry && textureEntry.hasTexture ? textureEntry.textureUrl : "",
        solidColorHex: textureEntry ? textureEntry.solidColorHex : "",
        coordUnit: shape ? shape.coordUnit : "",
        verticesJson: shape ? shape.verticesJson : [],
        holesJson: shape ? shape.holesJson : []
      });
    }

    return offcuts;
  } catch (err) {
    console.warn("Unable to load external offcut inventory.", err);
    return [];
  }
}

function canonicalRoomName_(v) {
  // Google Sheets auto-converts values like "3/1" to Date objects — recover them as D/M strings
  if (v instanceof Date && !isNaN(v.getTime())) {
    v = (v.getDate()) + '/' + (v.getMonth() + 1);
  }
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
const ORDER_MERGE_PROP_KEY = "WORKSHOP_HUB_ORDER_MERGES_V1";

function _getOrderIdArray_(orderIdsOrId) {
  const raw = Array.isArray(orderIdsOrId) ? orderIdsOrId : [orderIdsOrId];
  const seen = {};
  return raw
    .map(v => squeezeSpaces_(v))
    .filter(v => {
      if (!v || seen[v]) return false;
      seen[v] = true;
      return true;
    });
}

function _loadOrderMergeMap_() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(ORDER_MERGE_PROP_KEY);
  if (!raw) return {};

  try {
    return JSON.parse(raw) || {};
  } catch (err) {
    console.warn("Invalid order merge map. Resetting.", err);
    return {};
  }
}

function _resolveOrderMergeRoot_(orderId, mergeMap) {
  let current = squeezeSpaces_(orderId);
  const seen = {};

  while (current && mergeMap[current] && !seen[current]) {
    seen[current] = true;
    current = squeezeSpaces_(mergeMap[current]);
  }

  return current;
}

function _normalizeOrderMergeMap_(mergeMap) {
  const rawMap = mergeMap || {};
  const cleaned = {};

  Object.keys(rawMap).forEach(sourceId => {
    const source = squeezeSpaces_(sourceId);
    const target = squeezeSpaces_(rawMap[sourceId]);
    if (!source || !target || source === target) return;
    cleaned[source] = target;
  });

  Object.keys(cleaned).forEach(sourceId => {
    const root = _resolveOrderMergeRoot_(cleaned[sourceId], cleaned);
    if (!root || root === sourceId) delete cleaned[sourceId];
    else cleaned[sourceId] = root;
  });

  return cleaned;
}

function _saveOrderMergeMap_(mergeMap) {
  const cleaned = _normalizeOrderMergeMap_(mergeMap);
  PropertiesService.getDocumentProperties().setProperty(ORDER_MERGE_PROP_KEY, JSON.stringify(cleaned));
  return cleaned;
}

function _getShopifyOrderNotesMap_() {
  const sh = _getShopifyOrdersSheet();
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return {};

  const headers = values[0].map(h => _normTxt(h).toLowerCase());
  const hm = _headerMap_(headers);
  const idxOrderId = hm["order id"];
  const idxNotes = hm["notes"];
  if (idxOrderId == null || idxNotes == null) return {};

  const notesMap = {};
  for (let i = 1; i < values.length; i++) {
    const orderId = _normTxt(values[i][idxOrderId]);
    const note = squeezeSpaces_(values[i][idxNotes]);
    if (!orderId || !note) continue;
    if (!notesMap[orderId]) notesMap[orderId] = [];
    if (notesMap[orderId].indexOf(note) === -1) notesMap[orderId].push(note);
  }

  return notesMap;
}

function _addOrderNoteEntries_(orderObj, orderId, notes) {
  if (!orderObj) return;
  const safeOrderId = squeezeSpaces_(orderId);
  const entries = Array.isArray(notes) ? notes : [notes];
  if (!Array.isArray(orderObj.noteEntries)) orderObj.noteEntries = [];

  entries.map(v => squeezeSpaces_(v)).filter(Boolean).forEach(note => {
    const exists = orderObj.noteEntries.some(entry => entry.orderId === safeOrderId && entry.note === note);
    if (!exists) orderObj.noteEntries.push({ orderId: safeOrderId, note: note });
  });
}

function _formatOrderNotes_(noteEntries) {
  const entries = Array.isArray(noteEntries) ? noteEntries : [];
  if (!entries.length) return "";
  if (entries.length === 1) return entries[0].note;
  return entries.map(entry => `Order ${entry.orderId}: ${entry.note}`).join("\n\n");
}

function _ensureTreeOrder_(tree, orderId, customer) {
  const safeOrderId = squeezeSpaces_(orderId);
  if (!safeOrderId) return null;

  if (!tree[safeOrderId]) {
    tree[safeOrderId] = {
      id: safeOrderId,
      customer: customer,
      products: {},
      delivery: { bucket: {}, rooms: {} },
      noteEntries: [],
      notes: "",
      memberOrderIds: [safeOrderId],
      displayOrderIds: [safeOrderId],
      mergeTargetOrderId: safeOrderId,
      isMerged: false
    };
  }

  if (!tree[safeOrderId].customer && customer) tree[safeOrderId].customer = customer;
  return tree[safeOrderId];
}

function _ensureTreeProduct_(orderObj, productName, sourceOrderId) {
  const safeProduct = squeezeSpaces_(productName) || "Unknown Product";
  if (!orderObj.products[safeProduct]) {
    orderObj.products[safeProduct] = { panels: [], components: [], sourceOrderIds: [] };
  }

  const prod = orderObj.products[safeProduct];
  const safeSource = squeezeSpaces_(sourceOrderId);
  if (safeSource && prod.sourceOrderIds.indexOf(safeSource) === -1) prod.sourceOrderIds.push(safeSource);
  return { key: safeProduct, value: prod };
}

function _addUniqueOrderIds_(target, ids) {
  const nextIds = _getOrderIdArray_(ids);
  target = Array.isArray(target) ? target : [];
  nextIds.forEach(id => {
    if (target.indexOf(id) === -1) target.push(id);
  });
  return target;
}

function _getDeliveryTotalsForOrderProduct_(orderObj, prodName) {
  const safeProd = squeezeSpaces_(prodName);
  const del = (orderObj && orderObj.delivery) ? orderObj.delivery : { bucket: {}, rooms: {} };
  let total = 0;
  let delivered = 0;
  let fitted = 0;

  if (del.bucket && del.bucket[safeProd]) total += Number(del.bucket[safeProd].qty) || 0;

  Object.keys(del.rooms || {}).forEach(room => {
    (del.rooms[room] || []).forEach(item => {
      if (squeezeSpaces_(item.name) !== safeProd) return;
      const qty = Number(item.qty) || 0;
      total += qty;
      if (item.status === "Delivered" || item.status === "Fitted") delivered += qty;
      if (item.status === "Fitted") fitted += qty;
    });
  });

  return { total: total, delivered: delivered, fitted: fitted };
}

function _getRequiredUnitsForProduct_(prodData, deliveryTotal) {
  if ((Number(deliveryTotal) || 0) > 0) return Number(deliveryTotal) || 0;

  let maxUnits = 0;
  (prodData && prodData.panels ? prodData.panels : []).forEach(panel => {
    const per = Number(panel.qtyPerUnit) || 1;
    const qtyOrder = Number(panel.qtyOrder) || 0;
    const units = per ? Math.round(qtyOrder / per) : 0;
    if (units > maxUnits) maxUnits = units;
  });

  return maxUnits;
}

function _getManufacturedUnitsForProduct_(prodData) {
  const panelTotals = {};
  let panelUnits = Infinity;

  (prodData && prodData.panels ? prodData.panels : []).forEach(panel => {
    const key = String(panel.panelName || "");
    const per = Number(panel.qtyPerUnit) || 1;
    const packed = Number(panel.qtyPacked) || 0;
    if (!panelTotals[key]) panelTotals[key] = { packed: 0, per: per };
    panelTotals[key].packed += packed;
  });

  Object.keys(panelTotals).forEach(key => {
    const entry = panelTotals[key];
    const sets = Math.floor((Number(entry.packed) || 0) / (Number(entry.per) || 1));
    if (sets < panelUnits) panelUnits = sets;
  });
  if (panelUnits === Infinity) panelUnits = 0;

  let compUnits = Infinity;
  (prodData && prodData.components ? prodData.components : []).forEach(comp => {
    const per = Number(comp.qtyPerUnit) || 1;
    const packed = Number(comp.qtyPacked) || 0;
    const sets = Math.floor(packed / per);
    if (sets < compUnits) compUnits = sets;
  });
  if (compUnits === Infinity) compUnits = 999999;

  return Math.min(panelUnits, compUnits);
}

function _decorateOrderTree_(tree) {
  Object.keys(tree).forEach(orderId => {
    const orderObj = tree[orderId];
    orderObj.id = squeezeSpaces_(orderObj.id || orderId);
    orderObj.mergeTargetOrderId = squeezeSpaces_(orderObj.mergeTargetOrderId || orderObj.id || orderId);
    orderObj.memberOrderIds = _addUniqueOrderIds_([], orderObj.memberOrderIds && orderObj.memberOrderIds.length ? orderObj.memberOrderIds : [orderObj.id]);

    const preferred = orderObj.mergeTargetOrderId || orderObj.id;
    orderObj.displayOrderIds = _addUniqueOrderIds_([], [preferred].concat(orderObj.memberOrderIds));
    orderObj.isMerged = orderObj.displayOrderIds.length > 1;
    orderObj.notes = _formatOrderNotes_(orderObj.noteEntries);

    let orderRequired = 0;
    let orderManufactured = 0;
    let orderFitted = 0;

    Object.keys(orderObj.products || {}).forEach(prodName => {
      const prodData = orderObj.products[prodName] || { panels: [], components: [] };
      prodData.sourceOrderIds = _addUniqueOrderIds_([], prodData.sourceOrderIds && prodData.sourceOrderIds.length ? prodData.sourceOrderIds : [orderObj.id]);
      const delivery = _getDeliveryTotalsForOrderProduct_(orderObj, prodName);
      const required = _getRequiredUnitsForProduct_(prodData, delivery.total);
      const manufactured = Math.min(_getManufacturedUnitsForProduct_(prodData), required);
      orderRequired += required;
      orderManufactured += manufactured;
      orderFitted += Math.min(delivery.fitted || 0, required);
    });

    const isWorkshop = squeezeSpaces_(orderObj.customer).toLowerCase() === "workshop stock";
    orderObj.isWorkshop = isWorkshop;
    orderObj.isComplete = orderRequired > 0 ? (isWorkshop ? orderManufactured >= orderRequired : orderFitted >= orderRequired) : false;
  });

  return tree;
}

function _mergeOrderObjects_(targetOrder, sourceOrder) {
  if (!targetOrder || !sourceOrder) return targetOrder;

  targetOrder.memberOrderIds = _addUniqueOrderIds_(targetOrder.memberOrderIds, sourceOrder.memberOrderIds || [sourceOrder.id]);
  targetOrder.displayOrderIds = _addUniqueOrderIds_(targetOrder.displayOrderIds, sourceOrder.displayOrderIds || [sourceOrder.id]);
  targetOrder.noteEntries = (targetOrder.noteEntries || []).concat(sourceOrder.noteEntries || []);

  Object.keys(sourceOrder.products || {}).forEach(prodName => {
    const ensured = _ensureTreeProduct_(targetOrder, prodName, sourceOrder.id);
    const targetProd = ensured.value;
    const sourceProd = sourceOrder.products[prodName] || { panels: [], components: [], sourceOrderIds: [] };
    targetProd.sourceOrderIds = _addUniqueOrderIds_(targetProd.sourceOrderIds, sourceProd.sourceOrderIds || [sourceOrder.id]);
    targetProd.panels = targetProd.panels.concat(sourceProd.panels || []);
    targetProd.components = targetProd.components.concat(sourceProd.components || []);
  });

  Object.keys((sourceOrder.delivery && sourceOrder.delivery.bucket) || {}).forEach(prodName => {
    const src = sourceOrder.delivery.bucket[prodName];
    if (!targetOrder.delivery.bucket[prodName]) targetOrder.delivery.bucket[prodName] = { qty: 0, sourceOrderIds: [] };
    targetOrder.delivery.bucket[prodName].qty += Number(src.qty) || 0;
    targetOrder.delivery.bucket[prodName].sourceOrderIds = _addUniqueOrderIds_(targetOrder.delivery.bucket[prodName].sourceOrderIds, src.sourceOrderIds || [sourceOrder.id]);
  });

  Object.keys((sourceOrder.delivery && sourceOrder.delivery.rooms) || {}).forEach(roomName => {
    if (!targetOrder.delivery.rooms[roomName]) targetOrder.delivery.rooms[roomName] = [];

    (sourceOrder.delivery.rooms[roomName] || []).forEach(item => {
      const existing = targetOrder.delivery.rooms[roomName].find(entry =>
        squeezeSpaces_(entry.name) === squeezeSpaces_(item.name) &&
        squeezeSpaces_(entry.status) === squeezeSpaces_(item.status)
      );

      if (existing) {
        existing.qty = (Number(existing.qty) || 0) + (Number(item.qty) || 0);
        existing.sourceOrderIds = _addUniqueOrderIds_(existing.sourceOrderIds, item.sourceOrderIds || [sourceOrder.id]);
      } else {
        targetOrder.delivery.rooms[roomName].push({
          name: item.name,
          qty: Number(item.qty) || 0,
          status: item.status,
          sourceOrderIds: _addUniqueOrderIds_([], item.sourceOrderIds || [sourceOrder.id])
        });
      }
    });
  });

  return targetOrder;
}

function _applyCustomOrderMergesToTree_(tree) {
  const storedMap = _loadOrderMergeMap_();
  const mergeMap = _normalizeOrderMergeMap_(storedMap);
  if (JSON.stringify(storedMap || {}) !== JSON.stringify(mergeMap)) _saveOrderMergeMap_(mergeMap);

  Object.keys(mergeMap).forEach(sourceId => {
    const rootId = _resolveOrderMergeRoot_(mergeMap[sourceId], mergeMap);
    if (!sourceId || !rootId || sourceId === rootId) return;

    const sourceOrder = tree[sourceId];
    const targetOrder = tree[rootId];
    if (!sourceOrder || !targetOrder) return;
    if (squeezeSpaces_(sourceOrder.customer).toLowerCase() !== squeezeSpaces_(targetOrder.customer).toLowerCase()) return;
    if (squeezeSpaces_(sourceOrder.customer).toLowerCase() === "workshop stock") return;

    _mergeOrderObjects_(targetOrder, sourceOrder);
    delete tree[sourceId];
  });

  return tree;
}

function _buildDataTree_(applyCustomMerges) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hubSheet = ss.getSheetByName("Manufacture Hub");
  const compSheet = ss.getSheetByName("Components Hub");
  const delivSheet = ss.getSheetByName("Delivery Hub");
  const tz = Session.getScriptTimeZone();

  if (!hubSheet) throw new Error("CRITICAL: Manufacture Hub tab not found!");
  const panelData = hubSheet.getDataRange().getValues();
  panelData.shift();

  let compData = [];
  if (compSheet && compSheet.getLastRow() > 1) {
    compData = compSheet.getDataRange().getValues();
    compData.shift();
  }

  let delivData = [];
  if (delivSheet && delivSheet.getLastRow() > 1) {
    delivData = delivSheet.getDataRange().getValues();
    delivData.shift();
  } else {
    console.log("Warning: Delivery Hub tab missing or empty.");
  }

  const tree = {};
  const notesMap = _getShopifyOrderNotesMap_();
  const norm = v => String(v ?? "").replace(/ /g, " ").trim();

  panelData.forEach((row, index) => {
    const orderId = norm(row[0]);
    const customer = row[1];
    let product = norm(row[2]);
    if (!orderId) return;
    if (!product) product = norm(row[3] || "Unknown Product");

    const orderObj = _ensureTreeOrder_(tree, orderId, customer);
    _addOrderNoteEntries_(orderObj, orderId, notesMap[orderId]);
    const ensured = _ensureTreeProduct_(orderObj, product, orderId);

    if (!ensured.value.sku) ensured.value.sku = norm(row[3]);
    ensured.value.panels.push({
      rowIndex: index + 2,
      sourceOrderId: orderId,
      panelName: row[4],
      material: row[5],
      qtyPerUnit: Number(row[6]) || 1,
      qtyOrder: Number(row[11]) || 0,
      qtyCut: Number(row[12]) || 0,
      qtyProcessed: Number(row[13]) || 0,
      qtyEdgeFinish: Number(row[14]) || 0,
      qtyPacked: Number(row[15]) || 0
    });
  });

  compData.forEach((row, index) => {
    const orderId = norm(row[0]);
    let product = norm(row[2]);
    if (!orderId) return;
    if (!product) product = norm(row[3] || "Unknown Product");

    const orderObj = _ensureTreeOrder_(tree, orderId, row[1]);
    _addOrderNoteEntries_(orderObj, orderId, notesMap[orderId]);
    const ensured = _ensureTreeProduct_(orderObj, product, orderId);

    ensured.value.components.push({
      rowIndex: index + 2,
      sourceOrderId: orderId,
      compName: row[3],
      sku: row[4],
      qtyPerUnit: Number(row[5]) || 1,
      qtyRequired: Number(row[6]) || 0,
      qtyPacked: Number(row[7]) || 0,
      lastUser: String(row[8] || ""),
      lastUpdated: (row[9] instanceof Date)
        ? Utilities.formatDate(row[9], tz, "dd/MM/yyyy HH:mm")
        : String(row[9] || "")
    });
  });

  delivData.forEach(row => {
    const orderId = norm(row[0]);
    const product = norm(row[3]);
    const room = canonicalRoomName_(row[4]);
    const status = norm(row[5]) || "Pending";
    if (!orderId) return;

    const orderObj = _ensureTreeOrder_(tree, orderId, row[1]);
    _addOrderNoteEntries_(orderObj, orderId, notesMap[orderId]);
    _ensureTreeProduct_(orderObj, product, orderId);

    if (room === "") {
      if (!orderObj.delivery.bucket[product]) orderObj.delivery.bucket[product] = { qty: 0, sourceOrderIds: [] };
      orderObj.delivery.bucket[product].qty++;
      orderObj.delivery.bucket[product].sourceOrderIds = _addUniqueOrderIds_(orderObj.delivery.bucket[product].sourceOrderIds, [orderId]);
      return;
    }

    if (!orderObj.delivery.rooms[room]) orderObj.delivery.rooms[room] = [];
    const existingItem = orderObj.delivery.rooms[room].find(item => norm(item.name) === product && norm(item.status) === status);

    if (existingItem) {
      existingItem.qty++;
      existingItem.sourceOrderIds = _addUniqueOrderIds_(existingItem.sourceOrderIds, [orderId]);
    } else {
      orderObj.delivery.rooms[room].push({
        name: product,
        qty: 1,
        status: status,
        sourceOrderIds: [orderId]
      });
    }
  });

  if (applyCustomMerges) _applyCustomOrderMergesToTree_(tree);
  return _decorateOrderTree_(tree);
}

function getDataTree() {
  return _buildDataTree_(true);
}

// 3. UPDATE QUANTITIES (Robust Dynamic Version)
function updateQty(rowIndex, colName, value, callerUser) {
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
    let userEmail = (callerUser && String(callerUser).trim()) ? String(callerUser).trim() : (Session.getActiveUser().getEmail() || "Workshop App User");

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
    processed: "CNC'd",
    edgeFinish: "Edge Finished",
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

function markPanelDamaged(rowIndex, qty, reason, callerUser) {
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

    const colQtyOrder = map["qty order"] || map["qty per order"];
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
    const qtyOrder = Number(rowData[colQtyOrder - 1]) || 0;

    const timestamp = new Date();
    const userEmail = (callerUser && String(callerUser).trim()) ? String(callerUser).trim() : (Session.getActiveUser().getEmail() || "Workshop App User");

    sheet.getRange(rowIndex, colQtyOrder).setValue(qtyOrder + qtyNum);

    if (colLastAction) sheet.getRange(rowIndex, colLastAction).setValue("Damaged");
    if (colLastUser) sheet.getRange(rowIndex, colLastUser).setValue(userEmail);
    if (colLastUpdated) sheet.getRange(rowIndex, colLastUpdated).setValue(timestamp);

    logPanelHistoryEntry_(rowData, panelInfoCols, buildDamagePayload_(qtyNum, reason, userEmail, timestamp));

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
function processBatch(updates, callerUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manufacture Hub");
  if (!sheet) throw new Error("Sheet 'Manufacture Hub' missing");

  let userEmail = (callerUser && String(callerUser).trim()) ? String(callerUser).trim() : (Session.getActiveUser().getEmail() || "Workshop App");
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
        const currentVal = Number(sheet.getRange(rowIndex, colIndex).getValue()) || 0;

        // Conflict detection: base is what the client last saw; if the sheet has moved on, someone else saved first
        if (item.base !== undefined && currentVal !== (Number(item.base) || 0)) {
          results.push({ rowIndex, colName, value: currentVal, conflict: true });
          return; // skip this update
        }

        const previousValue = currentVal;
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

function processComponentBatch(updates, callerUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Components Hub");
  if (!sheet) throw new Error("Sheet 'Components Hub' missing");
  if (!Array.isArray(updates)) return [];

  const stockSheet = ss.getSheetByName("Component Stock");
  const hubSheet = ss.getSheetByName("Manufacture Hub");

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  const cache = CacheService.getScriptCache();
  const userEmail = (callerUser && String(callerUser).trim()) ? String(callerUser).trim() : (Session.getActiveUser().getEmail() || "Workshop App");
  const results = [];
  const packedTargets = new Map();

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
        const reason = `Allocated to ${productSku || productName || "Product"}${orderId ? ` (Order ${orderId})` : ""}`;
        logStockTransaction(compName || "Component", -delta, reason, "Component Stock");
      }

      sheet.getRange(rowIndex, 8).setValue(next);     // Packed
      sheet.getRange(rowIndex, 9).setValue(userEmail); // Last User
      sheet.getRange(rowIndex,10).setValue(new Date()); // Last Updated

      if (opId) cache.put("op:" + opId, "1", 21600); // 6h
      results.push({ rowIndex, value: next });

      // Track for furniture stock sync (any change to packed count)
      if (next !== cur) {
        const infoRow = sheet.getRange(rowIndex, 1, 1, 3).getValues()[0];
        const oid = _normTxt(infoRow[0]);
        const pname = _normTxt(infoRow[2]);
        if (oid && pname) packedTargets.set(`${oid}||${pname}`, { orderId: oid, productName: pname });
      }
    });

    // Trigger furniture stock sync for any workshop stock products with changed component packs
    if (packedTargets.size > 0 && hubSheet && hubSheet.getLastRow() > 1) {
      syncProductCompletions_(hubSheet.getDataRange().getValues(), Array.from(packedTargets.values()));
    }

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
  const idxNotes = hm["notes"];

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
    const notes = idxNotes == null ? "" : squeezeSpaces_(values[r][idxNotes]);

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
      notes,
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

/**
 * Reads the "Yield Rules" tab from the product recipe spreadsheet.
 * Returns rules indexed by SKU (uppercase):
 *   { "SKU": { wood: [{ material, sheetsPerUnit }], edge: [{ material, metersPerUnit }] }, ... }
 *
 * Edge material names must match the format "{Stock Material} {Thickness}mm"
 * e.g. "Bardolino Edgebanding 0.8mm", "Bardolino Edgebanding 2mm"
 */
function loadYieldRules_(prodSS) {
  const sheet = prodSS.getSheetByName('Yield Rules');
  if (!sheet) { console.warn('[YieldRules] No "Yield Rules" tab found in product sheet.'); return {}; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const headers = data[0].map(h => squeezeSpaces_(String(h)).toLowerCase());
  const skuCol  = headers.indexOf('product sku');
  const typeCol = headers.indexOf('type');
  const matCol  = headers.indexOf('material');
  const qtyCol  = headers.indexOf('qty per unit');

  if ([skuCol, typeCol, matCol, qtyCol].some(i => i < 0)) {
    throw new Error('Yield Rules tab is missing required columns: Product SKU, Type, Material, Qty Per Unit');
  }

  const rules = {};
  for (let i = 1; i < data.length; i++) {
    const sku      = squeezeSpaces_(String(data[i][skuCol] || '')).toUpperCase();
    const type     = squeezeSpaces_(String(data[i][typeCol] || '')).toLowerCase();
    const material = squeezeSpaces_(String(data[i][matCol]  || ''));
    const qty      = Number(data[i][qtyCol]) || 0;
    if (!sku || !material || !qty) continue;

    if (!rules[sku]) rules[sku] = { wood: [], edge: [] };
    if      (type === 'wood') rules[sku].wood.push({ material, sheetsPerUnit: qty });
    else if (type === 'edge') rules[sku].edge.push({ material, metersPerUnit: qty });
  }
  return rules;
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
  const idxNotes = hm["notes"];

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

  // Components tab: col A=SKU, col C=Component Name, col D=Item Code, col E=Qty Per Unit
  const componentMap = {};
  const compProductSheet = prodSS.getSheetByName("Components");
  if (compProductSheet) {
    const compData = compProductSheet.getDataRange().getValues();
    for (let p = 1; p < compData.length; p++) {
      const sku = _normTxt(compData[p][0]).toUpperCase();
      if (!sku) continue;
      if (!componentMap[sku]) componentMap[sku] = [];
      componentMap[sku].push(compData[p]);
    }
  }

  const yieldRules = loadYieldRules_(prodSS);

  const inventory = getInventoryData();
  const compStockMap = {};
  const woodStockList = [];
  const edgeStockList = [];

  (inventory.components || []).forEach(item => {
    const nameKey = _normTxt(item.name).toLowerCase();
    const codeKey = _normTxt(item.itemCode || '').toLowerCase();
    const s = Number(item.stock) || 0;
    const pr = Number(item.price) || 0;
    if (nameKey) {
      const ex = compStockMap[nameKey] || { stock: 0, price: 0 };
      compStockMap[nameKey] = { stock: ex.stock + s, price: pr || ex.price };
    }
    if (codeKey && codeKey !== nameKey) {
      const ex = compStockMap[codeKey] || { stock: 0, price: 0 };
      compStockMap[codeKey] = { stock: ex.stock + s, price: pr || ex.price };
    }
  });

  (inventory.wood || []).forEach(item => {
    const mat = _normTxt(item.material).toLowerCase();
    if (!mat) return;
    woodStockList.push({ material: mat, qty: Number(item.qty) || 0 });
  });

  (inventory.edge || []).forEach(item => {
    const mat = _normTxt(item.material).toLowerCase();
    if (!mat) return;
    const rollLength = Number(item.rollLength) || 0;
    const rolls = Number(item.rolls) || 0;
    edgeStockList.push({
      material: mat,
      thickness: squeezeSpaces_(String(item.thickness || '')),
      normalized: normalizeEdgeMaterialName_(mat),
      rollLength,
      rolls,
      stockMeters: rollLength * rolls,
      price: Number(item.price) || 0
    });
  });

  // Yield-rule edge lookup: "bardolino edgebanding 0.8mm" → total meters in stock.
  // Sums across all roll sizes of the same material + thickness.
  const yieldEdgeStockMap = {};
  edgeStockList.forEach(item => {
    if (!item.thickness) return;
    const key = `${item.material} ${item.thickness}mm`.toLowerCase();
    yieldEdgeStockMap[key] = (yieldEdgeStockMap[key] || 0) + item.stockMeters;
  });

  // Price per sheet for wood: materialKey (lowercase) → price per sheet.
  // Uses the first non-zero price found per material.
  const woodPriceMap = {};
  (inventory.wood || []).forEach(item => {
    const mat = _normTxt(item.material).toLowerCase();
    if (mat && !woodPriceMap[mat] && (Number(item.price) || 0) > 0) {
      woodPriceMap[mat] = Number(item.price);
    }
  });

  // Price per meter for yield-rule edgeband: "bardolino edgebanding 0.8mm" → £/m.
  // Derived from price per roll ÷ roll length. Uses first priced row per key.
  const yieldEdgePricePerMeterMap = {};
  edgeStockList.forEach(item => {
    if (!item.thickness || !item.rollLength || !item.price) return;
    const key = `${item.material} ${item.thickness}mm`.toLowerCase();
    if (!yieldEdgePricePerMeterMap[key]) {
      yieldEdgePricePerMeterMap[key] = item.price / item.rollLength;
    }
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
  const sumMatchingEdgeStock = (material, list) => {
    const key = normalizeEdgeMaterialName_(material);
    if (!key) return 0;
    return list.reduce((sum, item) => {
      if (!item.normalized) return sum;
      if (item.normalized === key || item.normalized.includes(key) || key.includes(item.normalized)) {
        return sum + (Number(item.stockMeters) || 0);
      }
      return sum;
    }, 0);
  };
  const totalsCompMap = {};
  const totalsWoodMap = {};      // recipe-based: material key → { material, type, area }
  const totalsEdgeMap = {};      // recipe-based: wood-material key → { material, meters }
  const totalsWoodYieldMap = {}; // yield-based:  material key → { material, rawSheets }
  const totalsEdgeYieldMap = {}; // yield-based:  edgeband key → { material, meters }

  const products = orderLines.map(line => {
    const recipeRows = productMap[line.sku] || [];
    const hasYield = !!yieldRules[line.sku];
    const compMap = {};
    const woodMap = {}; // per-product wood accumulator
    const edgeMap = {}; // per-product edge accumulator

    // ---- Components: from Components tab (col C=name, col D=itemCode, col E=qtyPerUnit) ----
    (componentMap[line.sku] || []).forEach(row => {
      const compName   = _normTxt(row[2]);
      const itemCode   = _normTxt(row[3]);
      const qtyPerUnit = Number(row[4]) || 0;
      const totalUnits = qtyPerUnit * line.toBuild;
      if (!totalUnits || !compName) return;

      if (!compMap[compName]) compMap[compName] = { name: compName, itemCode, qty: 0 };
      compMap[compName].qty += totalUnits;
      if (!totalsCompMap[compName]) totalsCompMap[compName] = { name: compName, itemCode, qty: 0 };
      totalsCompMap[compName].qty += totalUnits;
    });

    // ---- Wood & Edge: yield rules take priority over recipe ----
    if (hasYield) {
      const rules = yieldRules[line.sku];

      rules.wood.forEach(w => {
        const matKey = w.material.toLowerCase();
        if (!woodMap[matKey]) woodMap[matKey] = { material: w.material, rawSheets: 0 };
        woodMap[matKey].rawSheets += w.sheetsPerUnit * line.toBuild;

        if (!totalsWoodYieldMap[matKey]) totalsWoodYieldMap[matKey] = { material: w.material, rawSheets: 0 };
        totalsWoodYieldMap[matKey].rawSheets += w.sheetsPerUnit * line.toBuild;
      });

      rules.edge.forEach(e => {
        const matKey = e.material.toLowerCase();
        if (!edgeMap[matKey]) edgeMap[matKey] = { material: e.material, meters: 0 };
        edgeMap[matKey].meters += e.metersPerUnit * line.toBuild;

        if (!totalsEdgeYieldMap[matKey]) totalsEdgeYieldMap[matKey] = { material: e.material, meters: 0 };
        totalsEdgeYieldMap[matKey].meters += e.metersPerUnit * line.toBuild;
      });

    } else {
      // Fall back to recipe-based area/perimeter calculation
      recipeRows.forEach(row => {
        const material  = _normTxt(row[3]);
        const qtyPerUnit = Number(row[7]) || 0;
        const totalUnits = qtyPerUnit * line.toBuild;
        if (!totalUnits) return;

        const areaPerProduct      = Number(row[8]) || 0;
        const perimeterPerProduct = Number(row[9]) || 0;
        const totalArea      = areaPerProduct * line.toBuild;
        const totalPerimeter = perimeterPerProduct * line.toBuild;
        const matType    = _classifyMaterial_(material);
        const materialKey = material.toLowerCase();
        if (matType === "component") return;

        if (!woodMap[materialKey]) woodMap[materialKey] = { material, type: matType, area: 0 };
        woodMap[materialKey].area += totalArea;
        if (!totalsWoodMap[materialKey]) totalsWoodMap[materialKey] = { material, type: matType, area: 0 };
        totalsWoodMap[materialKey].area += totalArea;

        if (matType === "mdf") {
          if (!edgeMap[materialKey]) edgeMap[materialKey] = { material, meters: 0 };
          edgeMap[materialKey].meters += totalPerimeter;
          if (!totalsEdgeMap[materialKey]) totalsEdgeMap[materialKey] = { material, meters: 0 };
          totalsEdgeMap[materialKey].meters += totalPerimeter;
        }
      });
    }

    // ---- Build result arrays for this product ----
    const components = Object.values(compMap).map(item => {
      const byCode  = item.itemCode ? (compStockMap[item.itemCode.toLowerCase()] || null) : null;
      const byName  = compStockMap[item.name.toLowerCase()] || null;
      const entry   = byCode || byName || { stock: 0, price: 0 };
      const short = Math.max(0, item.qty - entry.stock);
      return { name: item.name, qty: item.qty, stock: entry.stock, short, price: entry.price, totalPrice: entry.price * short };
    });

    const wood = hasYield
      ? Object.values(woodMap).map(item => {
          const matType      = _classifyMaterial_(item.material);
          const sheetArea    = matType === "mdf" ? MDF_SHEET_AREA : matType === "ply" ? PLY_SHEET_AREA : 0;
          const sheetsNeeded = Math.ceil(item.rawSheets);
          const stockSheets  = sumMatchingStock(item.material, woodStockList, "qty");
          const shortSheets  = Math.max(0, sheetsNeeded - stockSheets);
          const price        = woodPriceMap[item.material.toLowerCase()] || 0;
          return {
            material: item.material,
            type: matType,
            area: parseFloat((item.rawSheets * sheetArea).toFixed(3)),
            sheetArea,
            sheetsNeeded,
            stockSheets,
            shortSheets,
            price,
            totalCost: parseFloat((price * shortSheets).toFixed(2))
          };
        })
      : Object.values(woodMap).map(item => {
          const sheetArea    = item.type === "mdf" ? MDF_SHEET_AREA : item.type === "ply" ? PLY_SHEET_AREA : 0;
          const sheetsNeeded = sheetArea ? Math.ceil(item.area / sheetArea) : 0;
          const stockSheets  = sumMatchingStock(item.material, woodStockList, "qty");
          const shortSheets  = Math.max(0, sheetsNeeded - stockSheets);
          const price        = woodPriceMap[item.material.toLowerCase()] || 0;
          return { material: item.material, type: item.type, area: item.area, sheetArea, sheetsNeeded, stockSheets, shortSheets, price, totalCost: parseFloat((price * shortSheets).toFixed(2)) };
        });

    const edge = hasYield
      ? Object.values(edgeMap).map(item => {
          const matKey      = item.material.toLowerCase();
          const stockMeters = yieldEdgeStockMap[matKey] || 0;
          const shortMeters = parseFloat(Math.max(0, item.meters - stockMeters).toFixed(2));
          const pricePerM   = yieldEdgePricePerMeterMap[matKey] || 0;
          return {
            material: item.material,
            meters:      parseFloat(item.meters.toFixed(2)),
            stockMeters: parseFloat(stockMeters.toFixed(2)),
            shortMeters,
            price:     parseFloat(pricePerM.toFixed(4)),
            totalCost: parseFloat((pricePerM * shortMeters).toFixed(2))
          };
        })
      : Object.values(edgeMap).map(item => {
          const stockMeters = sumMatchingEdgeStock(item.material, edgeStockList);
          const shortMeters = Math.max(0, item.meters - stockMeters);
          const matched     = edgeStockList.find(e => e.normalized && (e.normalized === normalizeEdgeMaterialName_(item.material) || e.normalized.includes(normalizeEdgeMaterialName_(item.material))));
          const pricePerM   = (matched && matched.price && matched.rollLength) ? matched.price / matched.rollLength : 0;
          return { material: item.material, meters: item.meters, stockMeters, shortMeters, price: parseFloat(pricePerM.toFixed(4)), totalCost: parseFloat((pricePerM * shortMeters).toFixed(2)) };
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

  // Build totals: merge yield-rule entries and recipe-fallback entries.
  // Yield-rule entries are always listed first (more accurate).
  const totalsWoodArray = [
    ...Object.values(totalsWoodYieldMap).map(item => {
      const matType      = _classifyMaterial_(item.material);
      const sheetArea    = matType === "mdf" ? MDF_SHEET_AREA : matType === "ply" ? PLY_SHEET_AREA : 0;
      const sheetsNeeded = Math.ceil(item.rawSheets);
      const stockSheets  = sumMatchingStock(item.material, woodStockList, "qty");
      const shortSheets  = Math.max(0, sheetsNeeded - stockSheets);
      const price        = woodPriceMap[item.material.toLowerCase()] || 0;
      return {
        material: item.material,
        type: matType,
        area: parseFloat((item.rawSheets * sheetArea).toFixed(3)),
        sheetArea,
        sheetsNeeded,
        stockSheets,
        shortSheets,
        price,
        totalCost: parseFloat((price * shortSheets).toFixed(2))
      };
    }),
    ...Object.values(totalsWoodMap).map(item => {
      const sheetArea    = item.type === "mdf" ? MDF_SHEET_AREA : item.type === "ply" ? PLY_SHEET_AREA : 0;
      const sheetsNeeded = sheetArea ? Math.ceil(item.area / sheetArea) : 0;
      const stockSheets  = sumMatchingStock(item.material, woodStockList, "qty");
      const shortSheets  = Math.max(0, sheetsNeeded - stockSheets);
      const price        = woodPriceMap[item.material.toLowerCase()] || 0;
      return { material: item.material, type: item.type, area: item.area, sheetArea, sheetsNeeded, stockSheets, shortSheets, price, totalCost: parseFloat((price * shortSheets).toFixed(2)) };
    })
  ];

  const totalsEdgeArray = [
    ...Object.values(totalsEdgeYieldMap).map(item => {
      const matKey      = item.material.toLowerCase();
      const stockMeters = yieldEdgeStockMap[matKey] || 0;
      const shortMeters = parseFloat(Math.max(0, item.meters - stockMeters).toFixed(2));
      const pricePerM   = yieldEdgePricePerMeterMap[matKey] || 0;
      return {
        material: item.material,
        meters:      parseFloat(item.meters.toFixed(2)),
        stockMeters: parseFloat(stockMeters.toFixed(2)),
        shortMeters,
        price:     parseFloat(pricePerM.toFixed(4)),
        totalCost: parseFloat((pricePerM * shortMeters).toFixed(2))
      };
    }),
    ...Object.values(totalsEdgeMap).map(item => {
      const stockMeters = sumMatchingEdgeStock(item.material, edgeStockList);
      const shortMeters = Math.max(0, item.meters - stockMeters);
      const matched     = edgeStockList.find(e => e.normalized && (e.normalized === normalizeEdgeMaterialName_(item.material) || e.normalized.includes(normalizeEdgeMaterialName_(item.material))));
      const pricePerM   = (matched && matched.price && matched.rollLength) ? matched.price / matched.rollLength : 0;
      return { material: item.material, meters: item.meters, stockMeters, shortMeters, price: parseFloat(pricePerM.toFixed(4)), totalCost: parseFloat((pricePerM * shortMeters).toFixed(2)) };
    })
  ];

  const totals = {
    components: Object.values(totalsCompMap).map(item => {
      const byCode  = item.itemCode ? (compStockMap[item.itemCode.toLowerCase()] || null) : null;
      const byName  = compStockMap[item.name.toLowerCase()] || null;
      const entry   = byCode || byName || { stock: 0, price: 0 };
      const short = Math.max(0, item.qty - entry.stock);
      return { name: item.name, qty: item.qty, stock: entry.stock, short, price: entry.price, totalPrice: entry.price * short };
    }),
    wood: totalsWoodArray,
    edge: totalsEdgeArray
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
    }
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

  row = _writeSmartOrderTableSection_(summarySheet, row, "TOTAL WOOD REQUIREMENTS", ["Material", "Type", "Sheet Size", "Area (m²)", "Sheets Req", "In Stock", "Needed", "Sheet Price (£)", "To Order (£)"],
    (summary.totals && summary.totals.wood ? summary.totals.wood : []).map(item => [
      item.material,
      item.type,
      item.type === "mdf" ? (summary.sheetSizes && summary.sheetSizes.mdf ? summary.sheetSizes.mdf : "2800 x 2070mm")
        : item.type === "ply" ? (summary.sheetSizes && summary.sheetSizes.ply ? summary.sheetSizes.ply : "3050 x 1220mm")
        : "",
      Number(item.area) || 0,
      Number(item.sheetsNeeded) || 0,
      Number(item.stockSheets) || 0,
      Number(item.shortSheets) || 0,
      Number(item.price) || 0,
      Number(item.totalCost) || 0
    ])
  );

  row = _writeSmartOrderTableSection_(summarySheet, row, "TOTAL EDGE REQUIREMENTS", ["Material", "Meters", "In Stock (m)", "Needed (m)", "Price/m (£)", "To Order (£)"],
    (summary.totals && summary.totals.edge ? summary.totals.edge : []).map(item => [
      item.material,
      Number(item.meters) || 0,
      Number(item.stockMeters) || 0,
      Number(item.shortMeters) || 0,
      Number(item.price) || 0,
      Number(item.totalCost) || 0
    ])
  );

  _writeSmartOrderTableSection_(summarySheet, row, "TOTAL COMPONENT REQUIREMENTS", ["Component", "Qty Req", "In Stock", "Needed", "Unit Price (£)", "Total Cost (£)"],
    (summary.totals && summary.totals.components ? summary.totals.components : []).map(item => [
      item.name,
      Number(item.qty) || 0,
      Number(item.stock) || 0,
      Number(item.short) || 0,
      Number(item.price) || 0,
      Number(item.totalPrice) || 0
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

    r = _writeSmartOrderTableSection_(sh, r, "WOOD", ["Material", "Type", "Sheet Size", "Area (m²)", "Sheets Req", "In Stock", "Needed", "Sheet Price (£)", "To Order (£)"],
      (prod.wood || []).map(item => [
        item.material,
        item.type,
        item.type === "mdf" ? (summary.sheetSizes && summary.sheetSizes.mdf ? summary.sheetSizes.mdf : "2800 x 2070mm")
          : item.type === "ply" ? (summary.sheetSizes && summary.sheetSizes.ply ? summary.sheetSizes.ply : "3050 x 1220mm")
          : "",
        Number(item.area) || 0,
        Number(item.sheetsNeeded) || 0,
        Number(item.stockSheets) || 0,
        Number(item.shortSheets) || 0,
        Number(item.price) || 0,
        Number(item.totalCost) || 0
      ])
    );

    r = _writeSmartOrderTableSection_(sh, r, "EDGEBAND", ["Material", "Meters", "In Stock (m)", "Needed (m)", "Price/m (£)", "To Order (£)"],
      (prod.edge || []).map(item => [
        item.material,
        Number(item.meters) || 0,
        Number(item.stockMeters) || 0,
        Number(item.shortMeters) || 0,
        Number(item.price) || 0,
        Number(item.totalCost) || 0
      ])
    );

    _writeSmartOrderTableSection_(sh, r, "COMPONENTS", ["Component", "Qty Req", "In Stock", "Needed", "Unit Price (£)", "Total Cost (£)"],
      (prod.components || []).map(item => [
        item.name,
        Number(item.qty) || 0,
        Number(item.stock) || 0,
        Number(item.short) || 0,
        Number(item.price) || 0,
        Number(item.totalPrice) || 0
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

function exportRoomListToSheets(orderIdsOrId) {
  const targetOrderIds = _getOrderIdArray_(orderIdsOrId);
  if (!targetOrderIds.length) throw new Error("Missing order ID for room list export.");
  const orderIdSet = {};
  targetOrderIds.forEach(id => orderIdSet[id] = true);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const deliverySheet = ss.getSheetByName("Delivery Hub");
  if (!deliverySheet) throw new Error("Delivery Hub tab is missing.");

  const data = deliverySheet.getDataRange().getValues();
  if (!data || data.length <= 1) throw new Error("No delivery data found.");

  const norm = (v) => String(v ?? "").replace(/ /g, " ").trim();
  const rows = [];
  let customerName = "";

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!orderIdSet[norm(row[0])]) continue;

    const customer = String(row[1] || "").trim();
    if (!customerName && customer) customerName = customer;

    const product = String(row[3] || "").trim();
    const room = canonicalRoomName_(row[4]);
    const status = norm(row[5]) || "Pending";

    if (!room) continue;
    rows.push({ room: room, product: product, status: status });
  }

  if (rows.length === 0) throw new Error("No room assignments found for this order.");

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

  const orderLabel = targetOrderIds.join(" & ");
  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
  const generatedOn = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");
  const out = SpreadsheetApp.create(`Room List - ${orderLabel} - ${timestamp}`);
  const sh = out.getSheets()[0];
  sh.setName("Room List");

  let r = 1;
  sh.getRange(r++, 1).setValue("Room List Export").setFontWeight("bold").setFontSize(14);
  sh.getRange(r++, 1).setValue(`Order ID: ${orderLabel}`);
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
function getShopifyMergeOptions(shopifyRowIndex) {
  const sh = _getShopifyOrdersSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => _normTxt(h).toLowerCase());
  const hm = _headerMap_(headers);

  const idxOrderId = hm["order id"];
  const idxCustomer = hm["customer"];
  if (idxOrderId == null || idxCustomer == null) throw new Error("Shopify Orders headers missing Order ID / Customer.");

  const row = sh.getRange(shopifyRowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
  const orderId = _normTxt(row[idxOrderId]);
  const customer = squeezeSpaces_(row[idxCustomer]);
  const customerNorm = customer.toLowerCase();
  const mergeMap = _normalizeOrderMergeMap_(_loadOrderMergeMap_());
  const existingMergeTargetId = _resolveOrderMergeRoot_(orderId, mergeMap);
  const tree = getDataTree();

  const candidates = Object.keys(tree)
    .map(key => tree[key])
    .filter(orderObj => {
      if (!orderObj || orderObj.isWorkshop || orderObj.isComplete) return false;
      if (squeezeSpaces_(orderObj.customer).toLowerCase() !== customerNorm) return false;
      const memberIds = _getOrderIdArray_(orderObj.memberOrderIds || [orderObj.id]);
      return memberIds.indexOf(orderId) === -1;
    })
    .map(orderObj => ({
      targetOrderId: orderObj.mergeTargetOrderId || orderObj.id,
      displayOrderIds: _getOrderIdArray_(orderObj.displayOrderIds || [orderObj.id]),
      customer: orderObj.customer
    }))
    .sort((a, b) => b.displayOrderIds.join(" & ").localeCompare(a.displayOrderIds.join(" & ")));

  return {
    orderId: orderId,
    customer: customer,
    existingMergeTargetId: (existingMergeTargetId && existingMergeTargetId !== orderId) ? existingMergeTargetId : "",
    candidates: candidates
  };
}

function mergeImportedOrderIntoExisting_(sourceOrderId, targetOrderId, expectedCustomer) {
  const sourceId = squeezeSpaces_(sourceOrderId);
  const targetId = squeezeSpaces_(targetOrderId);
  const customerNorm = squeezeSpaces_(expectedCustomer).toLowerCase();
  if (!sourceId || !targetId || sourceId === targetId) return targetId;

  const currentMap = _normalizeOrderMergeMap_(_loadOrderMergeMap_());
  const sourceRoot = _resolveOrderMergeRoot_(sourceId, currentMap) || sourceId;
  const targetRoot = _resolveOrderMergeRoot_(targetId, currentMap) || targetId;
  if (!sourceRoot || !targetRoot || sourceRoot === targetRoot) return targetRoot;

  const tree = getDataTree();
  const sourceOrder = tree[sourceRoot] || _buildDataTree_(false)[sourceRoot];
  const targetOrder = tree[targetRoot] || _buildDataTree_(false)[targetRoot];
  if (!sourceOrder) throw new Error(`Imported order ${sourceId} could not be found for merging.`);
  if (!targetOrder) throw new Error(`Merge target ${targetId} is no longer available.`);
  if (targetOrder.isWorkshop || targetOrder.isComplete) throw new Error(`Order ${targetRoot} is not available for merge.`);

  const sourceCustomer = squeezeSpaces_(sourceOrder.customer).toLowerCase();
  const targetCustomer = squeezeSpaces_(targetOrder.customer).toLowerCase();
  if (customerNorm && targetCustomer !== customerNorm) throw new Error(`Order ${targetRoot} belongs to a different customer.`);
  if (sourceCustomer !== targetCustomer) throw new Error("Only orders with the same customer name can be merged.");

  const nextMap = _normalizeOrderMergeMap_(currentMap);
  Object.keys(nextMap).forEach(key => {
    const root = _resolveOrderMergeRoot_(key, nextMap);
    if (root === sourceRoot) nextMap[key] = targetRoot;
  });
  nextMap[sourceRoot] = targetRoot;
  _saveOrderMergeMap_(nextMap);
  return targetRoot;
}

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

function _getAllowedOrderIdSet_(orderIdsOrId) {
  const ids = _getOrderIdArray_(orderIdsOrId);
  const set = {};
  ids.forEach(id => set[id] = true);
  return { ids: ids, set: set };
}

// 7. DELIVERY: ASSIGN ITEMS TO ROOM (Planning Phase)
function assignToRoom(orderIdsOrId, productName, qtyToAssign, roomName, callerUser) {
  roomName = canonicalRoomName_(roomName);
  if (!roomName) return "Error: Room name is required";
  const allowedOrders = _getAllowedOrderIdSet_(orderIdsOrId);
  if (!allowedOrders.ids.length) return "Error: Order ID is required";

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
  const user = (callerUser && String(callerUser).trim()) ? String(callerUser).trim() : (Session.getActiveUser().getEmail() || "Workshop App");
  const ts = new Date();

  let assigned = 0;

  // Structure: A ID | B Cust | C Addr | D Prod | E Room | F Status | G User | H Updated
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const normTxt = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim();

if (!allowedOrders.set[normTxt(row[0])]) continue;
if (normTxt(row[3]) !== normTxt(productName)) continue;
if (normTxt(row[4]) !== "") continue; // already assigned


    const st = normStatus(row[5]);
    if (st !== "Pending") continue;

    if (assigned < qtyToAssign) {
      sheet.getRange(i + 1, 5).setNumberFormat('@').setValue(roomName);  // Room (E) — forced text to prevent date auto-conversion
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


function bulkAssignToRooms(assignments, orderIdsOrId, callerUser) {
  if (!Array.isArray(assignments) || assignments.length === 0) return { success: false, error: "No assignments provided" };

  const allowedOrders = _getAllowedOrderIdSet_(orderIdsOrId);
  if (!allowedOrders.ids.length) return { success: false, error: "Order ID is required" };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Delivery Hub");
    if (!sheet) return { success: false, error: "Delivery Hub missing" };

    const data = sheet.getDataRange().getValues();
    const user = (callerUser && String(callerUser).trim()) ? String(callerUser).trim() : (Session.getActiveUser().getEmail() || "Workshop App");
    const ts = new Date();
    const normTxt = v => String(v ?? "").replace(/ /g, " ").trim();

    // Build a map: productName -> [{assignmentIndex, roomName, remaining}]
    // so we can process all assignments in a single pass through the sheet
    const assignMap = {};
    assignments.forEach((a, idx) => {
      const prod = normTxt(a.productName);
      const room = canonicalRoomName_(a.roomName);
      const qty  = Number(a.qty) || 0;
      if (!prod || !room || qty <= 0) return;
      if (!assignMap[prod]) assignMap[prod] = [];
      assignMap[prod].push({ room, remaining: qty });
    });

    const results = {};
    Object.keys(assignMap).forEach(p => { results[p] = { assigned: 0 }; });

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!allowedOrders.set[normTxt(row[0])]) continue;
      if (normTxt(row[4]) !== "") continue; // already assigned
      const st = String(row[5] || "").trim() || "Pending";
      if (st !== "Pending") continue;

      const prod = normTxt(row[3]);
      const queue = assignMap[prod];
      if (!queue) continue;

      // Find the first assignment for this product that still has remaining qty
      const slot = queue.find(s => s.remaining > 0);
      if (!slot) continue;

      sheet.getRange(i + 1, 5).setNumberFormat('@').setValue(slot.room);
      sheet.getRange(i + 1, 6, 1, 3).setValues([["Pending", user, ts]]);
      data[i][4] = slot.room; // mark as assigned in local copy

      slot.remaining--;
      results[prod].assigned++;

      // If this slot is exhausted, move to next slot for this product
      if (slot.remaining <= 0) queue.shift();
    }

    return { success: true, results };
  } finally {
    lock.releaseLock();
  }
}

function getDeliveryHistorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Delivery History");
  if (!sheet) sheet = ss.insertSheet("Delivery History");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "User", "Source", "Order ID", "Item", "Change", "Reason"]);
  }

  return sheet;
}

function logDeliveryFittingHistory_(orderId, productName, newStatus, qty, roomName, userEmail, timestamp) {
  if (newStatus !== "Delivered" && newStatus !== "Fitted") return;

  const sheet = getDeliveryHistorySheet_();
  const safeQty = Math.max(1, Number(qty) || 1);
  const safeOrder = String(orderId || "").trim();
  const safeItem = String(productName || "").trim() || "Unknown Product";
  const safeUser = String(userEmail || "").trim() || "Workshop App";
  const safeRoom = String(roomName || "").trim();
  const ts = timestamp instanceof Date ? timestamp : new Date();

  const reason = safeRoom ? `Room: ${safeRoom}` : "";
  sheet.appendRow([ts, safeUser, "Delivery and Fitting", safeOrder, safeItem, `${newStatus} (+${safeQty})`, reason]);
}

// 8. DELIVERY: UPDATE STATUS (DIAGNOSTIC VERSION)
function updateDeliveryStatus(orderIdsOrId, roomName, productName, oldStatus, newStatus, qtyToUpdate, callerUser) {
  roomName = canonicalRoomName_(roomName);
  const allowedOrders = _getAllowedOrderIdSet_(orderIdsOrId);

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
    const maxAllowed = getMaxReadyFromFactory(allowedOrders.ids, productName);

    const data = sheet.getDataRange().getValues();
    let currentlyUsed = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      

if (allowedOrders.set[normTxt(row[0])] && normTxt(row[3]) === normTxt(productName)) {

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
  const user = (callerUser && String(callerUser).trim()) ? String(callerUser).trim() : (Session.getActiveUser().getEmail() || "Workshop App");
  const ts = new Date();

  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (!allowedOrders.set[normTxt(row[0])]) continue;
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

  logDeliveryFittingHistory_(allowedOrders.ids.join(" & "), productName, newStatus, updated, roomName, user, ts);
  return "SUCCESS";

    } finally {
    lock.releaseLock();
  }

}


// 9. HELPER: CALCULATE MAX COMPLETED UNITS (Strict Match)
function getMaxReadyFromFactory(orderIdsOrId, productName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hubSheet = ss.getSheetByName("Manufacture Hub");
  const compSheet = ss.getSheetByName("Components Hub");
  
  const normTxt = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim();

  const allowedOrders = _getAllowedOrderIdSet_(orderIdsOrId);
  const targetProd = normTxt(productName);

  // 1. Get Panel Data
  const pData = hubSheet.getDataRange().getValues();
  let minPanelsReady = Infinity; 
  let hasPanels = false;

  for (let i = 1; i < pData.length; i++) {
    const row = pData[i];
    // Strict Match
    if (allowedOrders.set[normTxt(row[0])] && normTxt(row[2]) === targetProd) {
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

  if (allowedOrders.set[normTxt(row[0])] && normTxt(row[2]) === targetProd) {
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
function unassignFromRoom(orderIdsOrId, roomName, productName, qtyToUnassign) {
  roomName = canonicalRoomName_(roomName);
  const allowedOrders = _getAllowedOrderIdSet_(orderIdsOrId);

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

    if (!allowedOrders.set[normTxt(row[0])]) continue;
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
function exportCncXml(orderIdsOrId, productName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hubSheet = ss.getSheetByName("Manufacture Hub");
  const allowedOrders = _getAllowedOrderIdSet_(orderIdsOrId);
  
  // 1. Get Data
  const data = hubSheet.getDataRange().getValues();
  let xmlContent = '<?xml version="1.0" encoding="UTF-8"?>\n<Job>\n';
  let count = 0;
  
  // 2. Loop through rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Check match: Order ID (Col A) and Product Name (Col C)
    if (allowedOrders.set[String(row[0]).trim()] && String(row[2]) === String(productName)) {
      
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
  const fileName = `CNC_${_getOrderIdArray_(orderIdsOrId).join("_")}_${productName}.xml`.replace(/ /g, "_");
  
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
        itemCode: data[i][0],
        name: data[i][1],
        category: data[i][2],
        supplier: data[i][3],
        stock: Number(data[i][4]) || 0,    // Col E = Current Stock
        onOrder: Number(data[i][5]) || 0,  // Col F = On Order
        price: Number(data[i][6]) || 0,    // Col G = Price
        link: String(data[i][7] || '').trim() // Col H = Link
      });
    }
  }

  // B. WOOD STOCK (New 5-Column Layout)
  if (woodSheet && woodSheet.getLastRow() > 1) {
    const data = woodSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // row[0] = Material, row[1] = Colour Code, row[2] = Length, row[3] = Width, row[4] = Qty on Order, row[5] = Qty, row[6] = Price
      result.wood.push({
        rowIndex: i + 1,
        material: String(data[i][0]),                  // Col A
        colourCode: String(data[i][1] || "").trim(),   // Col B
        length: data[i][2],                            // Col C
        width: data[i][3],                             // Col D
        size: `${data[i][2]} x ${data[i][3]}`,
        onOrder: Number(data[i][4]) || 0,              // Col E = Qty on Order
        qty: Number(data[i][5]) || 0,                  // Col F = Qty in Stock
        price: Number(data[i][6]) || 0                 // Col G = Price per sheet
      });
    }
  }

  result.wood = result.wood.concat(getExternalOffcutInventory_());

  // C. EDGE BAND STOCK (Header-driven)
  if (edgeSheet && edgeSheet.getLastRow() > 1) {
    const data = edgeSheet.getDataRange().getValues();
    const map = getEdgeBandColumnMap_(edgeSheet);
    for (let i = 1; i < data.length; i++) {
      result.edge.push({
        rowIndex: i + 1,
        material: String(data[i][map.material - 1] || "").trim(),
        thickness: String(data[i][map.thickness - 1] || "").trim(),
        rollLength: Number(data[i][map.rollLength - 1]) || 0,
        onOrder: Number(data[i][map.onOrder - 1]) || 0,
        rolls: Number(data[i][map.rolls - 1]) || 0,
        price: map.price ? (Number(data[i][map.price - 1]) || 0) : 0
      });
    }
  }
  return result;
}

function getEdgeInventoryOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const edgeSheet = ss.getSheetByName("Edge Band Stock");
  const result = [];
  if (!edgeSheet || edgeSheet.getLastRow() < 2) return result;
  const data = edgeSheet.getDataRange().getValues();
  const map = getEdgeBandColumnMap_(edgeSheet);
  for (let i = 1; i < data.length; i++) {
    result.push({
      rowIndex: i + 1,
      material:   String(data[i][map.material   - 1] || "").trim(),
      thickness:  String(data[i][map.thickness  - 1] || "").trim(),
      rollLength: Number(data[i][map.rollLength - 1]) || 0,
      onOrder:    Number(data[i][map.onOrder    - 1]) || 0,
      rolls:      Number(data[i][map.rolls      - 1]) || 0,
      price:      map.price ? (Number(data[i][map.price - 1]) || 0) : 0
    });
  }
  return result;
}

// ADD NEW STOCK ITEMS
function addWoodItem(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Wood Stock");
  if (!sheet) throw new Error("Wood Stock tab missing");
  sheet.appendRow([
    String(data.material    || '').trim(),  // Col A
    String(data.colourCode  || '').trim(),  // Col B
    Number(data.length)     || 0,           // Col C
    Number(data.width)      || 0,           // Col D
    Number(data.onOrder)    || 0,           // Col E = Qty on Order
    Number(data.qty)        || 0,           // Col F = Qty in Stock
    Number(data.price)      || 0            // Col G = Price
  ]);
  return sheet.getLastRow();
}

function addComponentItem(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Component Stock");
  if (!sheet) throw new Error("Component Stock tab missing");
  sheet.appendRow([
    String(data.itemCode || '').trim(),  // Col A
    String(data.name     || '').trim(),  // Col B
    String(data.category || '').trim(),  // Col C
    String(data.supplier || '').trim(),  // Col D
    Number(data.stock)   || 0,           // Col E = Current Stock
    Number(data.onOrder) || 0,           // Col F = On Order
    Number(data.price)   || 0,           // Col G = Price
    String(data.link     || '').trim()   // Col H = Link
  ]);
  return sheet.getLastRow();
}

function addEdgeBandItem(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Edge Band Stock");
  if (!sheet) throw new Error("Edge Band Stock tab missing");
  const colMap = getEdgeBandColumnMap_(sheet);
  const lastCol = Math.max(
    colMap.material, colMap.thickness, colMap.rollLength,
    colMap.onOrder, colMap.rolls, colMap.price || 1
  );
  const newRow = new Array(lastCol).fill('');
  newRow[colMap.material   - 1] = String(data.material  || '').trim();
  newRow[colMap.thickness  - 1] = Number(data.thickness)  || 0;
  newRow[colMap.rollLength - 1] = Number(data.rollLength) || 0;
  newRow[colMap.onOrder    - 1] = Number(data.onOrder)    || 0;
  newRow[colMap.rolls      - 1] = Number(data.rolls)      || 0;
  if (colMap.price) newRow[colMap.price - 1] = Number(data.price) || 0;
  sheet.appendRow(newRow);
  return sheet.getLastRow();
}

// ADJUST WOOD STOCK (With Logging)
function adjustWoodStock(rowIndex, change, reason, callerUser) {
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
  logStockTransaction(materialName, change, reason, null, callerUser);

  return newVal;

    } finally {
    lock.releaseLock();
  }

}

function adjustEdgeStock(rowIndex, change, reason, callerUser) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Edge Band Stock");
  if (!sheet) throw new Error("Edge Band Stock tab missing");

  const colMap = getEdgeBandColumnMap_(sheet);
  const currentVal = sheet.getRange(rowIndex, colMap.rolls).getValue();
  const newVal = (Number(currentVal) || 0) + Number(change);

  if (newVal < 0) throw new Error("Stock cannot be negative");

  sheet.getRange(rowIndex, colMap.rolls).setValue(newVal);

  const mat = sheet.getRange(rowIndex, colMap.material).getValue();
  const thk = sheet.getRange(rowIndex, colMap.thickness).getValue();
  const rollLength = sheet.getRange(rowIndex, colMap.rollLength).getValue();
  logStockTransaction(`Edge: ${mat} | ${thk}mm | ${rollLength}m`, change, reason, null, callerUser);

  return newVal;

    } finally {
    lock.releaseLock();
  }

}

function adjustComponentStock(rowIndex, change, reason, callerUser) {

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
  logStockTransaction(materialLabel || `Component: ${sku} | ${name}`, change, reason, null, callerUser);

  return newVal;

    } finally {
    lock.releaseLock();
  }

}



function adjustComponentOnOrder(rowIndex, change) {
    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Component Stock");
  if (!sheet) throw new Error("Component Stock tab missing");

  const currentVal = sheet.getRange(rowIndex, 6).getValue(); // Col F = On Order
  const newVal = (Number(currentVal) || 0) + Number(change);

  if (newVal < 0) throw new Error("On Order cannot be negative");

  sheet.getRange(rowIndex, 6).setValue(newVal);
  return newVal;

    } finally {
    lock.releaseLock();
  }
}

function receiveComponentFromOrder(rowIndex, qtyReceived, callerUser) {
    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Component Stock");
  if (!sheet) throw new Error("Component Stock tab missing");

  const onOrder = Number(sheet.getRange(rowIndex, 6).getValue()) || 0; // Col F = On Order
  const inStock = Number(sheet.getRange(rowIndex, 5).getValue()) || 0; // Col E = Current Stock

  const qty = Number(qtyReceived) || 0;
  if (qty <= 0) throw new Error("Receive qty must be > 0");
  if (qty > onOrder) throw new Error("Cannot receive more than On Order");

  sheet.getRange(rowIndex, 6).setValue(onOrder - qty);
  sheet.getRange(rowIndex, 5).setValue(inStock + qty);

  const name = sheet.getRange(rowIndex, 2).getValue(); // Col B = Name
  logStockTransaction(name, qty, "Restock / Delivery (from Order)", null, callerUser);

  return "Success";

    } finally {
    lock.releaseLock();
  }
}

function allocateExternalOffcut_(itemKey, qtyUsed, projectId, productName, callerUser) {
  const offcutId = String(itemKey || '').replace(/^offcut:/i, '').trim();
  if (!offcutId) throw new Error("Invalid offcut id.");

  const extSs = SpreadsheetApp.openById(OFFCUT_SHEET_ID);
  const inventorySheet = extSs.getSheetByName(OFFCUT_INVENTORY_TAB_NAME);
  if (!inventorySheet || inventorySheet.getLastRow() < 2) throw new Error("External offcut inventory tab missing.");

  const data = inventorySheet.getDataRange().getValues();
  const map = getHeaderIndexMap_(data[0]);

  let rowIndex = -1;
  let row = null;
  for (let i = 1; i < data.length; i++) {
    const candidateId = squeezeSpaces_(getRowValueByHeaders_(data[i], map, "offcut_id"));
    if (candidateId === offcutId) {
      rowIndex = i + 1;
      row = data[i];
      break;
    }
  }

  if (!rowIndex || !row) throw new Error("Offcut not found in external inventory.");

  const qtyColIndex = (map["qty"] ?? -1) + 1;
  const statusColIndex = (map["status"] ?? -1) + 1;
  if (qtyColIndex <= 0) throw new Error("External offcut inventory is missing a qty column.");

  const currentQty = Number(getRowValueByHeaders_(row, map, "qty")) || 0;
  if (qtyUsed > currentQty) throw new Error("Insufficient offcut stock.");

  const materialName = squeezeSpaces_(getRowValueByHeaders_(row, map, "material"));
  const bboxW = Number(getRowValueByHeaders_(row, map, "bbox_w_mm")) || 0;
  const bboxH = Number(getRowValueByHeaders_(row, map, "bbox_h_mm")) || 0;
  const lengthMm = Math.max(bboxW, bboxH);
  const widthMm = Math.min(bboxW, bboxH);
  const newQty = currentQty - qtyUsed;

  inventorySheet.getRange(rowIndex, qtyColIndex).setValue(newQty);
  if (statusColIndex > 0) {
    inventorySheet.getRange(rowIndex, statusColIndex).setValue(newQty > 0 ? "IN_STOCK" : "DEPLETED");
  }

  const historyMaterialName = `Offcut ${offcutId}: ${materialName} - ${lengthMm} x ${widthMm}`;
  const historyReason = `CNC Job: #${projectId} (${productName})`;
  logStockTransaction(historyMaterialName, -qtyUsed, historyReason, null, callerUser);
  updateProjectUsageSummary(projectId, productName, materialName, qtyUsed);

  return "Success";
}

function allocateWoodSheet_(rowIndex, qtyUsed, projectId, productName, callerUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("Wood Stock");
  if (!stockSheet) throw new Error("Wood Stock tab missing");

  const currentQty = stockSheet.getRange(rowIndex, 6).getValue();
  const newQty = (Number(currentQty) || 0) - Number(qtyUsed);
  if (newQty < 0) throw new Error("Insufficient stock.");
  stockSheet.getRange(rowIndex, 6).setValue(newQty);

  const materialName = stockSheet.getRange(rowIndex, 1).getValue();
  const historyReason = `CNC Job: #${projectId} (${productName})`;
  logStockTransaction(materialName, -qtyUsed, historyReason, null, callerUser);
  updateProjectUsageSummary(projectId, productName, materialName, qtyUsed);

  return "Success";
}

// 3. ALLOCATE WOOD / OFFCUTS
function allocateWood(itemKey, qtyUsed, projectId, productName, callerUser) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const qty = Math.floor(Number(qtyUsed) || 0);
    if (qty <= 0) throw new Error("Quantity must be at least 1.");

    const key = String(itemKey || "").trim();
    if (!key) throw new Error("Missing stock item key.");

    if (/^offcut:/i.test(key)) {
      return allocateExternalOffcut_(key, qty, projectId, productName, callerUser);
    }

    const rowIndex = Number(key);
    if (!rowIndex) throw new Error("Invalid wood stock row.");
    return allocateWoodSheet_(rowIndex, qty, projectId, productName, callerUser);
  } finally {
    lock.releaseLock();
  }
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

  const colMap = getEdgeBandColumnMap_(sheet);
  const currentVal = sheet.getRange(rowIndex, colMap.onOrder).getValue();
  const newVal = (Number(currentVal) || 0) + Number(change);

  if (newVal < 0) throw new Error("On Order cannot be negative");

  sheet.getRange(rowIndex, colMap.onOrder).setValue(newVal);
  return newVal;

    } finally {
    lock.releaseLock();
  }

}

function receiveWoodFromOrder(rowIndex, qtyReceived, callerUser) {

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
  logStockTransaction(materialName, qty, "Restock / Delivery (from Order)", null, callerUser);

  return "Success";

    } finally {
    lock.releaseLock();
  }

}

function receiveEdgeFromOrder(rowIndex, rollsReceived, callerUser) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Edge Band Stock");
  if (!sheet) throw new Error("Edge Band Stock tab missing");

  const colMap = getEdgeBandColumnMap_(sheet);
  const onOrder = Number(sheet.getRange(rowIndex, colMap.onOrder).getValue()) || 0;
  const inStock = Number(sheet.getRange(rowIndex, colMap.rolls).getValue()) || 0;

  const qty = Number(rollsReceived) || 0;
  if (qty <= 0) throw new Error("Receive qty must be > 0");
  if (qty > onOrder) throw new Error("Cannot receive more than On Order");

  sheet.getRange(rowIndex, colMap.onOrder).setValue(onOrder - qty);
  sheet.getRange(rowIndex, colMap.rolls).setValue(inStock + qty);

  const mat = sheet.getRange(rowIndex, colMap.material).getValue();
  const thk = sheet.getRange(rowIndex, colMap.thickness).getValue();
  const rollLength = sheet.getRange(rowIndex, colMap.rollLength).getValue();
  logStockTransaction(`Edge: ${mat} | ${thk}mm | ${rollLength}m`, qty, "Restock / Delivery (from Order)", null, callerUser);

  return "Success";

    } finally {
    lock.releaseLock();
  }

}


// ACTION LIST
function getProductCatalogue() {
  const prodSS = SpreadsheetApp.openById(PRODUCT_RECIPE_SHEET_ID);

  const panelSheet = prodSS.getSheetByName(PRODUCT_RECIPE_TAB_NAME);
  if (!panelSheet) return [];
  const panelData = panelSheet.getDataRange().getValues();

  const compSheet = prodSS.getSheetByName("Components");
  const compData = compSheet ? compSheet.getDataRange().getValues() : [];

  const catalogue = {};

  for (let i = 1; i < panelData.length; i++) {
    const row = panelData[i];
    const sku = String(row[0] || "").trim().toUpperCase();
    if (!sku) continue;
    if (!catalogue[sku]) catalogue[sku] = { sku, productName: String(row[1] || "").trim() || sku, panels: [], components: [] };
    const panelName   = String(row[2] || "").trim();
    const material    = String(row[3] || "").trim();
    const familyCode  = String(row[4] || "").trim();
    const qty         = Number(row[7]) || 0;
    if (panelName) catalogue[sku].panels.push({ name: panelName, material, qty, familyCode });
  }

  for (let i = 1; i < compData.length; i++) {
    const row = compData[i];
    const sku = String(row[0] || "").trim().toUpperCase();
    if (!sku) continue;
    if (!catalogue[sku]) catalogue[sku] = { sku, productName: String(row[1] || "").trim() || sku, panels: [], components: [] };
    const name     = String(row[2] || "").trim();
    const itemCode = String(row[3] || "").trim();
    const qty      = Number(row[4]) || 0;
    const supplier = String(row[5] || "").trim();
    const link     = String(row[6] || "").trim();
    if (name) catalogue[sku].components.push({ name, itemCode, qty, supplier, link });
  }

  return Object.values(catalogue).sort((a, b) => a.sku.localeCompare(b.sku));
}

// ── Delivery Schedule ────────────────────────────────────────────────────────

function _getDeliveryScheduleSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Delivery Schedule");
  if (!sheet) {
    sheet = ss.insertSheet("Delivery Schedule");
    sheet.appendRow(["Delivery No","Date","Order ID","Customer","Rooms","Notes","Created By","Created At"]);
  }
  return sheet;
}

function _nextDeliveryNumber_() {
  const sheet = _getDeliveryScheduleSheet_();
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const m = String(data[i][0] || "").match(/D-(\d+)/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  return "D-" + String(max + 1).padStart(3, "0");
}

function saveDeliverySchedule(orderId, customer, dateStr, rooms, notes, callerUser) {
  const sheet = _getDeliveryScheduleSheet_();
  const delivNo = _nextDeliveryNumber_();
  const user = (callerUser && String(callerUser).trim()) ? String(callerUser).trim() : (Session.getActiveUser().getEmail() || "Workshop App");
  sheet.appendRow([delivNo, dateStr, orderId, customer, rooms.join(","), notes || "", user, new Date()]);
  return delivNo;
}

function getDeliverySchedules(orderId) {
  const sheet = _getDeliveryScheduleSheet_();
  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[2] || "").trim() !== String(orderId || "").trim()) continue;
    results.push({
      delivNo:   String(row[0] || ""),
      date:      row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(row[1] || ""),
      orderId:   String(row[2] || ""),
      customer:  String(row[3] || ""),
      rooms:     String(row[4] || "").split(",").map(r => r.trim()).filter(Boolean),
      notes:     String(row[5] || ""),
      createdBy: String(row[6] || ""),
      rowIndex:  i + 1
    });
  }
  return results.reverse();
}

function getOrderAddress(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Delivery Hub");
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  const norm = v => String(v ?? "").replace(/ /g, " ").trim();
  for (let i = 1; i < data.length; i++) {
    if (norm(data[i][0]) === norm(orderId) && norm(data[i][2])) return norm(data[i][2]);
  }
  return "";
}

function deleteDeliverySchedule(rowIndex) {
  const sheet = _getDeliveryScheduleSheet_();
  sheet.deleteRow(rowIndex);
  return true;
}

// ── Room Notes ────────────────────────────────────────────────────────────────

function saveRoomNote(orderId, roomName, note, callerUser) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  let sheet  = ss.getSheetByName("Room Notes");
  if (!sheet) {
    sheet = ss.insertSheet("Room Notes");
    sheet.appendRow(["Order ID", "Room Name", "Note", "Updated By", "Updated At"]);
  }
  const user    = callerUser && String(callerUser).trim() ? String(callerUser).trim() : "Workshop App";
  const normTxt = v => String(v ?? "").replace(/ /g, " ").trim();
  const ordId   = normTxt(orderId);
  const room    = normTxt(roomName);
  const data    = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normTxt(data[i][0]) === ordId && normTxt(data[i][1]) === room) {
      sheet.getRange(i + 1, 3).setValue(note || "");
      sheet.getRange(i + 1, 4).setValue(user);
      sheet.getRange(i + 1, 5).setValue(new Date());
      return true;
    }
  }
  sheet.appendRow([ordId, room, note || "", user, new Date()]);
  return true;
}

function getRoomNotes(orderId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Room Notes");
  if (!sheet) return {};
  const normTxt = v => String(v ?? "").replace(/ /g, " ").trim();
  const ordId   = normTxt(orderId);
  const data    = sheet.getDataRange().getValues();
  const result  = {};
  for (let i = 1; i < data.length; i++) {
    if (normTxt(data[i][0]) === ordId) {
      result[normTxt(data[i][1])] = String(data[i][2] || "");
    }
  }
  return result;
}

function batchMarkDelivered(orderId, items, callerUser) {
  // items = [{roomName, prodName, qty, sourceOrderIds: [...]}]
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const normTxt = v => String(v ?? "").replace(/ /g, " ").trim();
    const normSt  = s => { s = String(s || "").trim(); return s === "" ? "Pending" : s; };
    const user = callerUser && String(callerUser).trim() ? String(callerUser).trim() : "Workshop App";
    const ts   = new Date();

    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const delivSheet = ss.getSheetByName("Delivery Hub");
    if (!delivSheet) return { error: "Delivery Hub sheet missing" };

    const data    = delivSheet.getDataRange().getValues();
    const results = [];

    for (const item of items) {
      const allowedOrders = _getAllowedOrderIdSet_(item.sourceOrderIds);
      const prodNorm  = normTxt(item.prodName);
      const roomNorm  = canonicalRoomName_(item.roomName);
      const qtyToMark = Math.max(1, Number(item.qty) || 1);

      // Safety: count already on site across all allowed orders for this product
      const maxAllowed = getMaxReadyFromFactory(allowedOrders.ids, item.prodName);
      let onSite = 0;
      for (let i = 1; i < data.length; i++) {
        if (!allowedOrders.set[normTxt(data[i][0])]) continue;
        if (normTxt(data[i][3]) !== prodNorm) continue;
        const st = normSt(data[i][5]);
        if (st === "Delivered" || st === "Fitted") onSite++;
      }
      if (onSite + qtyToMark > maxAllowed) {
        results.push({ prodName: item.prodName, roomName: item.roomName, result: "BLOCKED", maxAllowed, onSite });
        continue;
      }

      // Apply: find Pending rows in this room for this product and mark Delivered
      let updated = 0;
      for (let i = 1; i < data.length && updated < qtyToMark; i++) {
        if (!allowedOrders.set[normTxt(data[i][0])]) continue;
        if (roomKey_(data[i][4]) !== roomKey_(roomNorm)) continue;
        if (normTxt(data[i][3]) !== prodNorm) continue;
        if (normSt(data[i][5]) !== "Pending") continue;
        delivSheet.getRange(i + 1, 6).setValue("Delivered");
        delivSheet.getRange(i + 1, 7).setValue(user);
        delivSheet.getRange(i + 1, 8).setValue(ts);
        data[i][5] = "Delivered"; // keep in-memory copy in sync for subsequent safety checks
        updated++;
      }

      if (updated > 0) {
        logDeliveryFittingHistory_(allowedOrders.ids.join(" & "), item.prodName, "Delivered", updated, item.roomName, user, ts);
        results.push({ prodName: item.prodName, roomName: item.roomName, result: "SUCCESS", updated });
      } else {
        results.push({ prodName: item.prodName, roomName: item.roomName, result: "NOT_FOUND", updated: 0 });
      }
    }

    return { results };
  } finally {
    lock.releaseLock();
  }
}

// ── End Delivery Schedule ─────────────────────────────────────────────────────

// ── Material Weights ──────────────────────────────────────────────────────────

function _createMaterialWeightsSheet_(prodSS) {
  const sheet = prodSS.insertSheet('Material Weights');
  sheet.getRange(1, 1, 1, 4).setValues([['Material', 'Sheet Width (mm)', 'Sheet Length (mm)', 'Full Sheet Weight (kg)']]);
  sheet.getRange(2, 1, 2, 4).setValues([
    ['MDF', 2800, 2070, 80],
    ['Ply', 3050, 1220, 50]
  ]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f4f6');
  sheet.autoResizeColumns(1, 4);
}

function getMaterialWeights_(prodSS) {
  const sheet = prodSS.getSheetByName('Material Weights');
  if (!sheet || sheet.getLastRow() < 2) return { mdf: 13.803, ply: 13.441 };
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const matKey = String(data[i][0] || '').trim().toLowerCase();
    const w  = Number(data[i][1]) || 0;
    const l  = Number(data[i][2]) || 0;
    const kg = Number(data[i][3]) || 0;
    if (!matKey || !w || !l || !kg) continue;
    map[matKey] = parseFloat((kg / ((w * l) / 1000000)).toFixed(4));
  }
  return map;
}

function getProductWeightsBySku(skus) {
  if (!skus || !skus.length) return {};
  const skuSet = new Set(skus.map(s => String(s).trim().toUpperCase()).filter(Boolean));
  if (!skuSet.size) return {};

  const prodSS = SpreadsheetApp.openById(PRODUCT_RECIPE_SHEET_ID);

  if (!prodSS.getSheetByName('Material Weights')) _createMaterialWeightsSheet_(prodSS);

  const recipeSheet = prodSS.getSheetByName(PRODUCT_RECIPE_TAB_NAME);
  if (!recipeSheet || recipeSheet.getLastRow() < 2) return {};

  const weightMap = getMaterialWeights_(prodSS);
  const data = recipeSheet.getDataRange().getValues();
  const skuWeights = {};

  for (let i = 1; i < data.length; i++) {
    const sku = String(data[i][0] || '').trim().toUpperCase();
    if (!sku || !skuSet.has(sku)) continue;

    const material    = String(data[i][3] || '').trim().toLowerCase();
    const areaPerUnit = Number(data[i][8]) || 0; // Col I = area per unit (m²)
    if (!areaPerUnit) continue;

    let kgPerM2 = 0;
    for (const [key, val] of Object.entries(weightMap)) {
      if (material.includes(key)) { kgPerM2 = val; break; }
    }
    if (!kgPerM2) continue;

    skuWeights[sku] = (skuWeights[sku] || 0) + areaPerUnit * kgPerM2;
  }

  Object.keys(skuWeights).forEach(k => { skuWeights[k] = parseFloat(skuWeights[k].toFixed(1)); });
  return skuWeights;
}

// ── End Material Weights ──────────────────────────────────────────────────────

// ── Surplus Stock ─────────────────────────────────────────────────────────────

function _ensureSurplusSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Surplus Stock');
  if (!sheet) {
    sheet = ss.insertSheet('Surplus Stock');
    sheet.getRange(1, 1, 1, 10).setValues([[
      'ID', 'Panel Name', 'Family Code', 'Source SKU', 'Source Order ID',
      'Qty Declared', 'Qty Available', 'Declared By', 'Date Declared', 'Notes'
    ]]);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 10);
  }
  return sheet;
}

function getSurplusStock() {
  const sheet = _ensureSurplusSheet_();
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const qtyAvail = Number(data[i][6]) || 0;
    result.push({
      rowIndex:    i + 1,
      id:          Number(data[i][0]) || (i + 1),
      panelName:   String(data[i][1] || '').trim(),
      familyCode:  String(data[i][2] || '').trim(),
      sourceSku:   String(data[i][3] || '').trim(),
      sourceOrder: String(data[i][4] || '').trim(),
      qtyDeclared: Number(data[i][5]) || 0,
      qtyAvailable: qtyAvail,
      declaredBy:  String(data[i][7] || '').trim(),
      date:        data[i][8] instanceof Date ? Utilities.formatDate(data[i][8], tz, 'dd/MM/yyyy') : String(data[i][8] || ''),
      notes:       String(data[i][9] || '').trim()
    });
  }
  return result.filter(r => r.qtyAvailable > 0);
}

function addSurplusEntry(panelName, familyCode, sourceSku, sourceOrderId, qty, callerUser, notes) {
  if (!qty || qty <= 0) throw new Error('Quantity must be greater than zero');
  const sheet = _ensureSurplusSheet_();
  const lastRow = sheet.getLastRow();
  let nextId = 1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    nextId = Math.max(...ids.map(v => Number(v) || 0)) + 1;
  }
  const user = String(callerUser || '').trim() || 'Workshop App';
  sheet.appendRow([
    nextId,
    String(panelName  || '').trim(),
    String(familyCode || '').trim().toUpperCase(),
    String(sourceSku  || '').trim().toUpperCase(),
    String(sourceOrderId || '').trim(),
    qty, qty,
    user, new Date(),
    String(notes || '').trim()
  ]);
  return nextId;
}

function allocateSurplusToOrder(surplusRowIndex, qtyToAllocate, targetOrderId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const norm = v => String(v ?? '').replace(/\s+/g, ' ').trim();

    // 1. Decrement surplus stock
    const surplusSheet = _ensureSurplusSheet_();
    const current = Number(surplusSheet.getRange(surplusRowIndex, 7).getValue()) || 0;
    if (qtyToAllocate > current) throw new Error(`Only ${current} available in surplus`);
    surplusSheet.getRange(surplusRowIndex, 7).setValue(current - qtyToAllocate);

    // 2. Mark matching panels as packed in the target order's Manufacture Hub rows
    if (!targetOrderId) return true;
    const surplusRow  = surplusSheet.getRange(surplusRowIndex, 1, 1, 10).getValues()[0];
    const familyCode  = norm(surplusRow[2]).toUpperCase();
    const surplusSku  = norm(surplusRow[3]).toUpperCase();
    const surplusPanelName = norm(surplusRow[1]).toLowerCase();

    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const hubSheet = ss.getSheetByName('Manufacture Hub');
    if (!hubSheet) return true;

    const hubData = hubSheet.getDataRange().getValues();
    const matchingPanelNames = new Set();

    if (familyCode) {
      // Primary: match by family code via recipe
      const prodSS      = SpreadsheetApp.openById(PRODUCT_RECIPE_SHEET_ID);
      const recipeSheet = prodSS.getSheetByName(PRODUCT_RECIPE_TAB_NAME);
      if (recipeSheet) {
        const orderSkus = new Set();
        hubData.forEach(row => { if (norm(row[0]) === norm(targetOrderId)) orderSkus.add(norm(row[3]).toUpperCase()); });
        const recipeData = recipeSheet.getDataRange().getValues();
        for (let i = 1; i < recipeData.length; i++) {
          if (!orderSkus.has(norm(recipeData[i][0]).toUpperCase())) continue;
          if (norm(recipeData[i][4]).toUpperCase() !== familyCode) continue;
          matchingPanelNames.add(norm(recipeData[i][2]).toLowerCase());
        }
      }
    } else {
      // Fallback: no family code — match by same panel name within same SKU
      for (let i = 1; i < hubData.length; i++) {
        if (norm(hubData[i][0]) !== norm(targetOrderId)) continue;
        if (norm(hubData[i][3]).toUpperCase() !== surplusSku) continue;
        if (norm(hubData[i][4]).toLowerCase() === surplusPanelName) {
          matchingPanelNames.add(surplusPanelName);
          break;
        }
      }
    }
    if (!matchingPanelNames.size) return true;

    // Update qtyCut/qtyProcessed/qtyEdgeFinish/qtyPacked for matching hub rows
    for (let i = 1; i < hubData.length; i++) {
      if (norm(hubData[i][0]) !== norm(targetOrderId)) continue;
      if (!matchingPanelNames.has(norm(hubData[i][4]).toLowerCase())) continue;
      const qtyOrder = Number(hubData[i][11]) || 0;
      const addQty   = Math.min(qtyToAllocate, qtyOrder);
      // cols 13-16 = qtyCut, qtyProcessed, qtyEdgeFinish, qtyPacked (1-indexed)
      [13, 14, 15, 16].forEach(col => {
        const cur = Number(hubSheet.getRange(i + 1, col).getValue()) || 0;
        hubSheet.getRange(i + 1, col).setValue(Math.min(qtyOrder, cur + addQty));
      });
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}

function getSurplusSummaryBySkus(skus) {
  if (!skus || !skus.length) return {};
  const surplus = getSurplusStock();
  const result  = {};
  skus.forEach(sku => {
    const upper = String(sku).trim().toUpperCase();
    const total = surplus
      .filter(s => s.sourceSku.toUpperCase() === upper)
      .reduce((sum, s) => sum + (s.qtyAvailable || 0), 0);
    if (total > 0) result[upper] = total;
  });
  return result;
}

// Search surplus by SKU directly — used by import modal before order is in Manufacture Hub
function getSurplusForSku(sku) {
  if (!sku) return [];
  const normSku = String(sku).trim().toUpperCase();
  const surplus = getSurplusStock();
  const matches = [];

  // Primary: same source SKU + any panel name
  surplus.forEach(s => {
    if (s.sourceSku.toUpperCase() === normSku) {
      matches.push({ ...s, orderPanelName: s.panelName, orderProductName: normSku });
    }
  });

  // Secondary: family code match via recipe (for panels from a different SKU but same family)
  if (surplus.some(s => s.sourceSku.toUpperCase() !== normSku)) {
    try {
      const prodSS   = SpreadsheetApp.openById(PRODUCT_RECIPE_SHEET_ID);
      const recSheet = prodSS.getSheetByName(PRODUCT_RECIPE_TAB_NAME);
      if (recSheet) {
        const recData = recSheet.getDataRange().getValues();
        const norm    = v => String(v ?? '').replace(/\s+/g, ' ').trim();
        // Get family codes for this SKU
        const skuFamilies = new Set();
        for (let i = 1; i < recData.length; i++) {
          if (norm(recData[i][0]).toUpperCase() !== normSku) continue;
          const fc = norm(recData[i][4]).toUpperCase();
          if (fc) skuFamilies.add(fc);
        }
        const alreadyMatched = new Set(matches.map(m => m.rowIndex));
        surplus.forEach(s => {
          if (alreadyMatched.has(s.rowIndex)) return;
          const fc = s.familyCode.toUpperCase();
          if (fc && skuFamilies.has(fc)) {
            matches.push({ ...s, orderPanelName: s.panelName, orderProductName: normSku });
          }
        });
      }
    } catch(e) { /* recipe lookup optional */ }
  }

  return matches;
}

function getAvailableSurplusForOrder(orderId) {
  // Get all panel family codes for this order's products from the recipe
  const prodSS   = SpreadsheetApp.openById(PRODUCT_RECIPE_SHEET_ID);
  const recSheet = prodSS.getSheetByName(PRODUCT_RECIPE_TAB_NAME);
  if (!recSheet) return [];

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const hubSheet = ss.getSheetByName('Manufacture Hub');
  if (!hubSheet) return [];

  const norm = v => String(v ?? '').replace(/\s+/g, ' ').trim();

  // Get SKUs in this order
  const hubData = hubSheet.getDataRange().getValues();
  const orderSkus = new Set();
  const skuProdMap = {}; // sku → productName
  hubData.forEach(row => {
    if (norm(row[0]) !== norm(orderId)) return;
    const sku = norm(row[3]).toUpperCase();
    if (sku) { orderSkus.add(sku); skuProdMap[sku] = norm(row[2]); }
  });
  if (!orderSkus.size) return [];

  // Get family codes used in this order's panels
  const recData = recSheet.getDataRange().getValues();
  const orderFamilies = new Map(); // familyCode → { panelName, sku, productName }
  // Also build panel name map for SKU-based fallback
  const orderPanelsBySku = new Map(); // sku → [panelName, ...]
  for (let i = 1; i < recData.length; i++) {
    const sku    = norm(recData[i][0]).toUpperCase();
    const pName  = norm(recData[i][2]);
    const family = norm(recData[i][4]).toUpperCase();
    if (!orderSkus.has(sku)) continue;
    if (family) {
      if (!orderFamilies.has(family)) {
        orderFamilies.set(family, { panelName: pName, sku, productName: skuProdMap[sku] || sku });
      }
    }
    // Always track panels by SKU for fallback
    if (!orderPanelsBySku.has(sku)) orderPanelsBySku.set(sku, []);
    orderPanelsBySku.get(sku).push(pName);
  }

  // Match surplus entries by family code (primary) OR by SKU when no family code (fallback)
  const surplus = getSurplusStock();
  const matches = [];
  const seen = new Set(); // avoid duplicates

  surplus.forEach(s => {
    const fc  = s.familyCode.toUpperCase();
    const key = `${s.rowIndex}`;
    if (seen.has(key)) return;

    if (fc && orderFamilies.has(fc)) {
      // Primary: family code match
      const orderPanel = orderFamilies.get(fc);
      seen.add(key);
      matches.push({ ...s, orderPanelName: orderPanel.panelName, orderProductName: orderPanel.productName });
    } else if (!fc && orderSkus.has(s.sourceSku.toUpperCase())) {
      // Fallback: no family code but same SKU — panels are identical products
      const sku = s.sourceSku.toUpperCase();
      const panelNames = orderPanelsBySku.get(sku) || [];
      const orderPanelName = panelNames.find(n => n.toLowerCase() === s.panelName.toLowerCase()) || s.panelName;
      seen.add(key);
      matches.push({ ...s, orderPanelName, orderProductName: skuProdMap[sku] || sku });
    }
  });
  return matches;
}

function declareSurplusFromManufacturing(orderId, panelName, familyCode, sourceSku, qty, alreadyProcessed, callerUser) {
  // alreadyProcessed = true: panels exist physically → add to surplus AND reduce order panel qty
  // alreadyProcessed = false: panels not yet cut → just reduce order required qty, nothing added to surplus
  const norm = v => String(v ?? '').replace(/\s+/g, ' ').trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hubSheet = ss.getSheetByName('Manufacture Hub');
  if (!hubSheet) throw new Error('Manufacture Hub tab missing');

  const data = hubSheet.getDataRange().getValues();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (norm(data[i][0]) !== norm(orderId)) continue;
      if (norm(data[i][4]).toLowerCase() !== norm(panelName).toLowerCase()) continue;
      const currentQtyOrder = Number(data[i][11]) || 0;
      const newQtyOrder = Math.max(0, currentQtyOrder - qty);
      hubSheet.getRange(i + 1, 12).setValue(newQtyOrder); // Col L = qtyOrder
      if (!alreadyProcessed) {
        // Also zero out any already-logged packed/cut if they exceed new qty
        const cols = [13, 14, 15, 16]; // qtyCut, qtyProcessed, qtyEdge, qtyPacked
        cols.forEach(col => {
          const cur = Number(hubSheet.getRange(i + 1, col).getValue()) || 0;
          if (cur > newQtyOrder) hubSheet.getRange(i + 1, col).setValue(newQtyOrder);
        });
      }
      found = true;
      break;
    }
    if (!found) throw new Error(`Panel "${panelName}" not found in order ${orderId}`);

    if (alreadyProcessed) {
      addSurplusEntry(panelName, familyCode, sourceSku, orderId, qty, callerUser,
        `Declared surplus during manufacturing of order ${orderId}`);
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}

function markPanelAsPackedFromSurplus(orderId, panelName, qty) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const norm = v => String(v ?? '').replace(/\s+/g, ' ').trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hubSheet = ss.getSheetByName('Manufacture Hub');
    if (!hubSheet) throw new Error('Manufacture Hub tab missing');

    const data = hubSheet.getDataRange().getValues();
    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      if (norm(data[i][0]) !== norm(orderId)) continue;
      if (norm(data[i][4]).toLowerCase() !== norm(panelName).toLowerCase()) continue;
      const qtyOrder = Number(data[i][11]) || 0;
      const setQty   = Math.min(qty, qtyOrder);
      // Set cut, processed, edge, packed all to setQty
      [13, 14, 15, 16].forEach(col => {
        hubSheet.getRange(i + 1, col).setValue(setQty);
      });
      updated++;
    }
    if (!updated) throw new Error(`Panel "${panelName}" not found in order ${orderId}`);
    return updated;
  } finally {
    lock.releaseLock();
  }
}

// ── End Surplus Stock ─────────────────────────────────────────────────────────

// ── Discord Notifications ─────────────────────────────────────────────────────

function postToDiscord_(embed) {
  try {
    const res = UrlFetchApp.fetch(DISCORD_WEBHOOK_URL + '?wait=false', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'User-Agent': 'WorkshopHubBot/1.0 (Google Apps Script)' },
      payload: JSON.stringify({ embeds: [embed] }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code !== 200 && code !== 204) {
      console.warn('Discord webhook returned ' + code + ': ' + res.getContentText());
    }
  } catch (e) {
    console.warn('Discord notification failed:', e.message);
  }
}

function notifyTaskCreated_(task, dueDate, assignedTo, assignedBy, priority, taskId) {
  const priorityColours = { 'High': 15612550, 'Medium': 16097803, 'Low': 3818378 };
  const colour = priorityColours[priority] || 3818378;

  const fields = [
    { name: '📝 Task',        value: task || '—',        inline: false },
    { name: '👤 Assigned To', value: assignedTo || '—',  inline: true  },
    { name: '📌 Assigned By', value: assignedBy || '—',  inline: true  },
    { name: '⚡ Priority',    value: priority || '—',    inline: true  }
  ];
  if (dueDate) fields.push({ name: '📅 Due Date', value: dueDate, inline: true });

  postToDiscord_({
    title: '📋 New Task Created',
    color: colour,
    fields,
    footer: { text: `Workshop Hub · Task #${taskId}` },
    timestamp: new Date().toISOString()
  });
}

function notifyTaskCompleted_(task, assignedTo, assignedBy, newStatus, taskId) {
  const colour = newStatus === 'Complete' ? 3066993 : 9807270; // green : grey

  postToDiscord_({
    title: newStatus === 'Complete' ? '✅ Task Completed' : `🔄 Task Updated — ${newStatus}`,
    color: colour,
    fields: [
      { name: '📝 Task',        value: task || '—',       inline: false },
      { name: '👤 Assigned To', value: assignedTo || '—', inline: true  },
      { name: '📌 Assigned By', value: assignedBy || '—', inline: true  },
      { name: '📊 New Status',  value: newStatus || '—',  inline: true  }
    ],
    footer: { text: `Workshop Hub · Task #${taskId}` },
    timestamp: new Date().toISOString()
  });
}

// ── End Discord Notifications ─────────────────────────────────────────────────

function getActionList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Action List");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const result = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[3];
    const dueVal  = row[6];
    result.push({
      rowIndex:   i + 1,
      id:         Number(row[0]) || (i + 1),
      assignedTo: String(row[1] || '').trim(),
      assignedBy: String(row[2] || '').trim(),
      date:       dateVal instanceof Date ? Utilities.formatDate(dateVal, tz, "dd/MM/yyyy") : String(dateVal || '').trim(),
      task:       String(row[4] || '').trim(),
      priority:   String(row[5] || '').trim(),
      dueDate:    dueVal instanceof Date ? Utilities.formatDate(dueVal, tz, "dd/MM/yyyy") : String(dueVal || '').trim(),
      status:     String(row[7] || 'Not Started').trim()
    });
  }

  return result;
}

function addActionItem(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Action List");
  if (!sheet) throw new Error("Action List tab missing");

  const lastRow = sheet.getLastRow();
  let nextId = 1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const maxId = Math.max(...ids.map(v => Number(v) || 0));
    nextId = maxId + 1;
  }

  let dueDate = '';
  if (data.dueDate) {
    const parsed = new Date(data.dueDate);
    if (!isNaN(parsed.getTime())) dueDate = parsed;
  }

  const assignedTo = String(data.assignedTo || '').trim();
  const assignedBy = String(data.assignedBy || '').trim();
  const task       = String(data.task       || '').trim();
  const priority   = String(data.priority   || '').trim();

  sheet.appendRow([nextId, assignedTo, assignedBy, new Date(), task, priority, dueDate, 'Not Started']);

  const dueDateDisplay = dueDate instanceof Date
    ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "dd/MM/yyyy")
    : '';
  notifyTaskCreated_(task, dueDateDisplay, assignedTo, assignedBy, priority, nextId);

  return nextId;
}

function updateActionStatus(rowIndex, status) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Action List");
    if (!sheet) throw new Error("Action List tab missing");
    sheet.getRange(rowIndex, 8).setValue(status); // Col H = Status

    // Read row details for the Discord notification
    const row = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
    const taskId    = Number(row[0]) || rowIndex;
    const assignedTo = String(row[1] || '').trim();
    const assignedBy = String(row[2] || '').trim();
    const task       = String(row[4] || '').trim();
    notifyTaskCompleted_(task, assignedTo, assignedBy, status, taskId);

    return "Success";
  } finally {
    lock.releaseLock();
  }
}

// --- PRESENCE (Active Users) ---
const _PRESENCE_USERS = ["Ben","Adam","Ian","Rufus","Theo","David","Eli","Tom","Allan","Peter"];

function heartbeatPresence(userName) {
  if (!userName) return;
  CacheService.getScriptCache().put("whub_p_" + String(userName).trim(), "1", 90);
}

function removePresence(userName) {
  if (!userName) return;
  CacheService.getScriptCache().remove("whub_p_" + String(userName).trim());
}

function getActiveUsers() {
  const keys = _PRESENCE_USERS.map(n => "whub_p_" + n);
  const vals = CacheService.getScriptCache().getAll(keys);
  return _PRESENCE_USERS.filter(n => vals["whub_p_" + n] != null);
}

// HELPER: Centralized Stock Logging
function logStockTransaction(material, change, reason, sourceOverride, callerUser) {
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

  const user = (callerUser && String(callerUser).trim()) ? String(callerUser).trim() : (Session.getActiveUser().getEmail() || "Workshop App");
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


function getTrackingHistory(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const maxRows = Math.max(1, Number(limit) || 1500);
  const entries = [];

  const normalizeTrackingOrderId_ = (value) => {
    const txt = String(value || '').trim();
    if (!txt) return '';
    return txt.startsWith('#') ? txt : `#${txt}`;
  };

  const formatTrackingItemLabel_ = (materialValue) => {
    const raw = String(materialValue || '').trim();
    if (!raw) return 'Stock Item';

    const offcutMatch = raw.match(/^offcut\s*:\s*(.+?)\s*-\s*(\d+)\s*x\s*(\d+)$/i);
    if (offcutMatch) {
      return `Offcut: ${offcutMatch[1].trim()} - ${offcutMatch[2]} x ${offcutMatch[3]}`;
    }

    return raw;
  };

  const pullByHeaders = (sheetName, mapFn) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const map = {};
    headers.forEach((h, i) => map[h] = i);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const entry = mapFn(row, map);
      if (entry) entries.push(entry);
    }
  };

  pullByHeaders('Panel History', (row, map) => {
    const orderId = row[map['order id']] || '';
    const timestamp = row[map['timestamp']] || '';
    const user = row[map['user']] || '';
    const panel = row[map['panel name']] || '';
    const product = row[map['product']] || row[map['product name']] || '';
    const change = row[map['change']] || row[map['change type']] || '';
    const qty = Number(row[map['quantity']]) || 0;
    const reason = row[map['reason']] || '';

    if (!timestamp && !orderId && !user && !panel && !change) return null;

    return {
      source: 'Panel History',
      orderId: normalizeTrackingOrderId_(orderId),
      user: String(user || '').trim() || 'Unknown',
      timestamp: timestamp,
      item: String(panel || product || 'Panel').trim(),
      change: qty ? `${change || 'Updated'} (${qty > 0 ? '+' : ''}${qty})` : String(change || 'Updated'),
      reason: String(reason || '').trim()
    };
  });

  pullByHeaders('Stock History', (row, map) => {
    const timestamp = row[map['timestamp']] || '';
    const user = row[map['user']] || '';
    const sourceRaw = String(row[map['source']] || 'Stock History').trim();
    const material = row[map['material']] || '';
    const changeVal = row[map['change']];
    const reason = row[map['reason']] || '';

    if (!timestamp && !user && !material && (changeVal === '' || changeVal === null || changeVal === undefined)) return null;

    const sourceLooksLikeOrderId = /^#?[a-z0-9-]+$/i.test(sourceRaw) && /^#?\d+/i.test(sourceRaw);
    const sourceOrderMatch = sourceRaw.match(/^(#?[a-z0-9-]+)$/i);
    const reasonOrderMatch = String(reason || '').match(/#\s*[a-z0-9-]+/i)
      || String(reason || '').match(/\b(?:order|batch)\s*#?\s*([a-z0-9-]+)/i);
    const materialOrderMatch = String(material || '').match(/#\s*[a-z0-9-]+/i);

    const normalizedSource = sourceLooksLikeOrderId
      ? (String(material || '').toLowerCase().includes('component') ? 'Component Stock' : 'Stock History')
      : sourceRaw;

    const reasonOrderId = reasonOrderMatch
      ? String(reasonOrderMatch[0]).replace(/.*?#/, '#').replace(/\s+/g, '')
      : '';
    const materialOrderId = materialOrderMatch
      ? String(materialOrderMatch[0]).replace(/\s+/g, '')
      : '';
    const sourceOrderId = (sourceLooksLikeOrderId && sourceOrderMatch)
      ? String(sourceOrderMatch[1] || '').replace(/\s+/g, '')
      : '';

    const resolvedOrderId = reasonOrderId || materialOrderId || sourceOrderId;

    const numericChange = Number(changeVal);
    const changeText = Number.isFinite(numericChange)
      ? `${numericChange > 0 ? '+' : ''}${numericChange}`
      : String(changeVal || 'Updated');

    return {
      source: normalizedSource,
      orderId: normalizeTrackingOrderId_(resolvedOrderId),
      user: String(user || '').trim() || 'Unknown',
      timestamp: timestamp,
      item: formatTrackingItemLabel_(material),
      change: changeText,
      reason: String(reason || '').trim()
    };
  });

  pullByHeaders('Delivery History', (row, map) => {
    const timestamp = row[map['timestamp']] || '';
    const user = row[map['user']] || '';
    const source = row[map['source']] || 'Delivery and Fitting';
    const orderId = row[map['order id']] || '';
    const item = row[map['item']] || row[map['product']] || '';
    const change = row[map['change']] || '';
    const reason = row[map['reason']] || '';

    if (!timestamp && !user && !orderId && !item && !change) return null;

    return {
      source: String(source || 'Delivery and Fitting').trim(),
      orderId: normalizeTrackingOrderId_(orderId),
      user: String(user || '').trim() || 'Unknown',
      timestamp: timestamp,
      item: String(item || 'Product').trim(),
      change: String(change || 'Updated').trim(),
      reason: String(reason || '').trim()
    };
  });

  entries.sort((a, b) => {
    const aTime = (a.timestamp instanceof Date) ? a.timestamp.getTime() : new Date(a.timestamp).getTime();
    const bTime = (b.timestamp instanceof Date) ? b.timestamp.getTime() : new Date(b.timestamp).getTime();
    return (Number.isFinite(bTime) ? bTime : 0) - (Number.isFinite(aTime) ? aTime : 0);
  });

  return entries.slice(0, maxRows).map(e => {
    const rawTime = e.timestamp;
    let timeMs = null;
    let dateStr = '';

    if (rawTime instanceof Date) {
      timeMs = rawTime.getTime();
      dateStr = Utilities.formatDate(rawTime, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    } else {
      const parsed = new Date(rawTime);
      if (!isNaN(parsed.getTime())) {
        timeMs = parsed.getTime();
        dateStr = Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      } else {
        dateStr = String(rawTime || '');
      }
    }

    const dateKey = timeMs ? Utilities.formatDate(new Date(timeMs), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';

    return {
      source: e.source,
      orderId: e.orderId,
      user: e.user,
      dateStr: dateStr,
      dateKey: dateKey,
      timeMs: timeMs,
      item: e.item,
      change: e.change,
      reason: e.reason
    };
  });
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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const norm = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim().toLowerCase();
  const orderProductMap = {};
  const panelTotalsByCustomerSku = {};

  // Build panel totals from Manufacture Hub
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

  // Build component totals from Components Hub
  // Cols: A=orderId, B=customer, C=productName, D=compName, E=itemCode, F=qtyPerUnit, G=totalQty, H=qtyPacked
  const compTotalsByCustomerSku = {};
  const compSheet = ss.getSheetByName("Components Hub");
  if (compSheet && compSheet.getLastRow() > 1) {
    const compData = compSheet.getDataRange().getValues();
    for (let i = 1; i < compData.length; i++) {
      const row = compData[i];
      const orderId = norm(row[0]);
      const customer = norm(row[1]);
      const productName = norm(row[2]);
      const compName = String(row[3] || "");
      const qtyPerUnit = Number(row[5]) || 1;
      const qtyPacked = Number(row[7]) || 0;

      // Resolve SKU from the order+product map built above
      const meta = orderProductMap[`${orderId}||${productName}`];
      const sku = meta ? meta.sku : "";
      if (!customer || !sku || !compName) continue;

      const key = `${customer}||${sku}`;
      if (!compTotalsByCustomerSku[key]) compTotalsByCustomerSku[key] = {};
      if (!compTotalsByCustomerSku[key][compName]) {
        compTotalsByCustomerSku[key][compName] = { packed: 0, required: qtyPerUnit };
      }
      compTotalsByCustomerSku[key][compName].packed += qtyPacked;
    }
  }

  targets.forEach(({ orderId, productName }) => {
    const key = `${norm(orderId)}||${norm(productName)}`;
    const meta = orderProductMap[key];
    if (!meta || meta.customer !== "workshop stock" || !meta.sku) return;

    const customerSkuKey = `${meta.customer}||${meta.sku}`;

    // Panel limiting factor
    const panelTotals = panelTotalsByCustomerSku[customerSkuKey];
    let panelFinished = Infinity;
    let hasPanels = false;
    if (panelTotals) {
      Object.keys(panelTotals).forEach(panelKey => {
        hasPanels = true;
        const p = panelTotals[panelKey];
        const sets = Math.floor((Number(p.packed) || 0) / (Number(p.required) || 1));
        if (sets < panelFinished) panelFinished = sets;
      });
    }
    if (!hasPanels || panelFinished === Infinity) panelFinished = 0;

    // Component limiting factor
    const compTotals = compTotalsByCustomerSku[customerSkuKey];
    let compFinished = Infinity;
    let hasComps = false;
    if (compTotals) {
      Object.keys(compTotals).forEach(compKey => {
        hasComps = true;
        const c = compTotals[compKey];
        const sets = Math.floor((Number(c.packed) || 0) / (Number(c.required) || 1));
        if (sets < compFinished) compFinished = sets;
      });
    }
    if (!hasComps) compFinished = Infinity; // no components → not a limiting factor

    // Overall: limited by whichever is lower (components only limit if they exist)
    const totalFinished = hasPanels
      ? Math.min(panelFinished, compFinished === Infinity ? panelFinished : compFinished)
      : 0;

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

function approveAndImportShopifyRow(shopifyRowIndex, mergeTargetOrderId) {

    const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

  // 1) Approve line
  const res = approveShopifyOrder(shopifyRowIndex);
  if (res === "Already imported") return res;
  if (res !== "SUCCESS") return res;

  // 2) Import that ONE line immediately (function lives in Import.gs)
  return importApprovedShopifyRow_(shopifyRowIndex, mergeTargetOrderId);

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
