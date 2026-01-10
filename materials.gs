// =============================================================================
// INVENTORY ENGINE (inventory.gs)
// Handles FIFO, Stock Movements, and Database Updates
// =============================================================================

const MATERIALS_SS_ID = '1V6RihfeEAlt78-eRgeO3b3xDL_tBCr2BppUIL_T5anw'; // Materials Spreadsheet

/**
 * CORE: Process a Stock Transaction
 */
function processInventoryTransaction(txData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    const ss = SpreadsheetApp.openById(MATERIALS_SS_ID);
    
    let result;

    // --- LOGIC CHANGE: OPENING BALANCE (Delta Approach) ---
    // We treat the input qty as the TARGET Balance, not just an addition.
    if (txData.type === 'OPENING_BALANCE') {
        const delta = getOpeningBalanceDelta(ss, txData.refId, txData.whId, txData.qty);
        
        if (delta.diff === 0) {
            return { success: true, message: "Stock is already at this level. No changes made." };
        }
        
        // Update the transaction data to reflect the CHANGE, not the total
        txData.qty = delta.absDiff;
        // Append the explanatory note
        txData.notes = `${delta.notePrefix} ${txData.notes || ''}`;
        
        if (delta.diff > 0) {
            // We need to ADD stock to reach the target
            // We use the Unit Cost provided by the user for these new items
            result = handleStockIn(ss, txData);
        } else {
            // We need to REMOVE stock to reach the target
            // We ignore user cost and use FIFO (removing oldest layers)
            result = handleStockOut(ss, txData);
        }

    } else {
        // --- STANDARD TRANSACTIONS (Adjustments, Transfers) ---
        // Validate Availability for Outbound
        if (!validateStockTransaction(ss, txData)) {
            throw new Error("خطأ في التحقق: الرصيد غير كافٍ أو البيانات غير صالحة");
        }

        switch (txData.type) {
          case 'ADJUSTMENT_IN':
            result = handleStockIn(ss, txData);
            break;
          case 'ADJUSTMENT_OUT':
            result = handleStockOut(ss, txData);
            break;
          case 'TRANSFER':
            result = handleTransfer(ss, txData);
            break;
          default:
            throw new Error("Unknown Transaction Type: " + txData.type);
        }
    }
    
    return { success: true, txId: result.txId };

  } catch (e) {
    Logger.log("Inventory Error: " + e.toString());
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * HELPER: Calculate the difference between Current Stock and Target Opening Balance
 */
function getOpeningBalanceDelta(ss, refId, whId, targetQty) {
    const sheet = ss.getSheetByName('Current_Stock');
    const lastRow = sheet.getLastRow();
    let currentQty = 0;

    if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // Cols A, B, C
        for(let i=0; i<data.length; i++) {
            if(data[i][0] == refId && data[i][1] == whId) {
                currentQty = Number(data[i][2]);
                break;
            }
        }
    }

    const diff = targetQty - currentQty;
    
    if (diff > 0) {
        return { 
            diff: diff, 
            absDiff: diff, 
            notePrefix: `[Auto-Adjust] Added ${diff} to reach Opening Balance of ${targetQty}.` 
        };
    } else {
        return { 
            diff: diff, 
            absDiff: Math.abs(diff), 
            notePrefix: `[Auto-Adjust] Removed ${Math.abs(diff)} to reach Opening Balance of ${targetQty}.` 
        };
    }
}

/**
 * Helper: Find the last recorded Unit Cost for a specific Ref ID
 * Logic: Scans backwards to find the most recent cost, REGARDLESS of warehouse.
 */
function getLastRecordedCost(ss, refId) {
  const sheet = ss.getSheetByName('Transactions');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  
  // Fetch Columns A to G (Indexes 0 to 6)
  // Col D (Index 3) = Ref ID
  // Col G (Index 6) = Unit Cost
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  // Iterate backwards (Newest to Oldest)
  for (let i = data.length - 1; i >= 0; i--) {
    // Check if Ref ID matches
    if (String(data[i][3]) === String(refId)) {
      const cost = Number(data[i][6]); 
      // Ensure cost is a valid positive number
      if (cost > 0) return cost;
    }
  }
  return 0; // Return 0 if no history found
}

function handleStockIn(ss, data) {
  const layersSheet = ss.getSheetByName('FIFO_Layers');
  const txSheet = ss.getSheetByName('Transactions');
  
  // --- COST LOGIC ---
  let finalCost = 0;

  // 1. Check if we have a valid manual input
  if (data.cost !== null && data.cost !== undefined && data.cost !== '') {
     finalCost = Number(data.cost);
  } 
  // Fallback: Check unitCost
  else if (data.unitCost !== null && data.unitCost !== undefined && data.unitCost !== '' && Number(data.unitCost) !== 0) {
     finalCost = Number(data.unitCost);
  }
  
  // 2. Auto-Detect if 0
  if (finalCost === 0) {
     const detectedCost = getLastRecordedCost(ss, data.refId);
     if (detectedCost === 0) {
       throw new Error("عفواً، لا يوجد سعر مسجل لهذا الصنف من قبل. يرجى إدخال سعر الوحدة يدوياً.");
     }
     finalCost = detectedCost;
  }
  
  data.unitCost = finalCost; 
  // ------------------

  const timestamp = data.forcedDate || new Date();
  const txId = data.forcedTxId || ('TX-' + timestamp.getTime());
  const layerId = 'LAY-' + timestamp.getTime() + (data.forcedTxId ? '-IN' : '');
  
  // FIFO Layers
  layersSheet.appendRow([layerId, data.refId, data.whId, timestamp, data.unitCost, data.qty, data.qty, txId]);
  
  const totalValue = data.qty * data.unitCost;
  
  // Transactions
  txSheet.appendRow([txId, timestamp, data.type, data.refId, data.whId, data.qty, data.unitCost, totalValue, layerId, data.notes]);
  
  // Update Stock Sheet
  updateCurrentStockCache(ss, data.refId, data.whId, data.qty, totalValue);
  
  // --- FIXED RETURN OBJECT ---
  return { success: true, message: 'تم حفظ العملية بنجاح', txId: txId };
}

function handleStockOut(ss, data) {
  const layersSheet = ss.getSheetByName('FIFO_Layers');
  const txSheet = ss.getSheetByName('Transactions');
  
  const reqQty = Number(data.qty);
  const refId = data.refId;
  const whId = data.whId;
  
  // 1. Fetch FIFO Layers
  const range = layersSheet.getDataRange();
  const values = range.getValues();
  
  let relevantLayers = [];
  let specificLayerIndex = -1;

  // --- NEW: SPECIFIC LAYER LOGIC ---
  if (data.specificLayerId) {
      // Find the EXACT layer requested
      for (let i = 1; i < values.length; i++) {
          if (String(values[i][0]) === String(data.specificLayerId)) {
              specificLayerIndex = i;
              relevantLayers.push({
                  rowIndex: i + 1, 
                  date: new Date(values[i][3]),
                  cost: Number(values[i][4]),
                  remQty: Number(values[i][6]),
                  layerId: values[i][0]
              });
              break;
          }
      }
      if (relevantLayers.length === 0) throw new Error("Specified Layer ID not found or invalid.");
  } 
  else {
      // --- STANDARD AUTO-FIFO LOGIC ---
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][1]) === String(refId) && String(values[i][2]) === String(whId) && Number(values[i][6]) > 0) {
          relevantLayers.push({
            rowIndex: i + 1, 
            date: new Date(values[i][3]),
            cost: Number(values[i][4]),
            remQty: Number(values[i][6]),
            layerId: values[i][0]
          });
        }
      }
      // Sort FIFO
      relevantLayers.sort((a, b) => a.date - b.date);
  }
  
  // 2. Validate Stock
  const totalAvailable = relevantLayers.reduce((sum, l) => sum + l.remQty, 0);
  if (totalAvailable < reqQty) {
    throw new Error(`Insufficient Stock in selected layer(s). Available: ${totalAvailable}, Requested: ${reqQty}`);
  }
  
  // 3. Consume Layers
  let qtyToDeduct = reqQty;
  let totalCostValue = 0;
  let consumedLayersIds = [];
  
  for (let layer of relevantLayers) {
    if (qtyToDeduct <= 0) break;

    let deductFromThis = Math.min(qtyToDeduct, layer.remQty);
    
    qtyToDeduct -= deductFromThis;
    totalCostValue += (deductFromThis * layer.cost);
    consumedLayersIds.push(layer.layerId);

    // Update Sheet
    const newRemQty = layer.remQty - deductFromThis;
    layersSheet.getRange(layer.rowIndex, 7).setValue(newRemQty);
  }
  
  // 4. Calculate Weighted Average Cost
  const avgUnitCost = totalCostValue / reqQty;

  // 5. Record Transaction
  const timestamp = data.forcedDate || new Date();
  const txId = data.forcedTxId || ('TX-' + timestamp.getTime() + '-OUT');

  txSheet.appendRow([
    txId, timestamp, data.txType || data.type, refId, whId, -reqQty, 
    avgUnitCost, totalCostValue, consumedLayersIds.join(','), data.notes
  ]);

  updateCurrentStockCache(ss, refId, whId, -reqQty, -totalCostValue);
  
  return { 
      success: true, 
      txId: txId, 
      avgCost: avgUnitCost, 
      totalValue: totalCostValue 
  };
}

/**
 * LOGIC: Warehouse Transfer (Out from Source -> In to Dest)
 */
function handleTransfer(ss, data) {
  const timestamp = new Date();
  const baseId = 'TRF-' + timestamp.getTime(); 
  
  // Step 1: Stock OUT from Source (Specific Layer or Auto)
  const outResult = handleStockOut(ss, {
    type: 'TRANSFER_OUT',
    refId: data.refId,
    whId: data.whId, 
    qty: data.qty,
    notes: `Transfer to ${data.relatedWhId}`,
    specificLayerId: data.specificLayerId, // <--- PASSED HERE
    forcedTxId: baseId + '-OUT', 
    forcedDate: timestamp
  });

  // Calculate specific unit cost from the OUT result
  const transferredUnitCost = outResult.avgCost || (outResult.totalValue / data.qty);

  // Step 2: Stock IN to Destination
  const inResult = handleStockIn(ss, {
    type: 'TRANSFER_IN',
    refId: data.refId,
    whId: data.relatedWhId,
    qty: data.qty,
    unitCost: transferredUnitCost,
    cost: transferredUnitCost, 
    notes: `Transfer from ${data.whId}`,
    forcedTxId: baseId + '-IN', 
    forcedDate: timestamp
  });

  return { txId: baseId };
}

/**
 * Updates the 'Current_Stock' sheet directly.
 * Handles both adding stock (positive values) and removing stock (negative values).
 */
function updateCurrentStockCache(ss, refId, whId, qtyChange, valueChange) {
  const stockSheet = ss.getSheetByName('Current_Stock');
  
  // Optimistic Check: Try to find the row
  const range = stockSheet.getDataRange();
  const values = range.getValues();
  let foundIndex = -1;

  // Loop starting from row 1 (skipping header)
  // Col 0 = RefID, Col 1 = WhID
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(refId) && String(values[i][1]) === String(whId)) {
      foundIndex = i + 1; // Convert 0-based array index to 1-based Sheet Row
      break;
    }
  }

  if (foundIndex > 0) {
    // --- UPDATE EXISTING ROW ---
    // Get current values
    const currentQty = Number(values[foundIndex - 1][2]);   // Col C
    const currentValue = Number(values[foundIndex - 1][4]); // Col E
    
    const newQty = currentQty + Number(qtyChange);
    const newValue = currentValue + Number(valueChange);
    
    // Calculate new Average Cost (prevent division by zero)
    const newAvg = newQty !== 0 ? (newValue / newQty) : 0;
    
    // Update Cols: C(3)=Qty, D(4)=Avg, E(5)=Value, F(6)=Date
    stockSheet.getRange(foundIndex, 3).setValue(newQty);
    stockSheet.getRange(foundIndex, 4).setValue(newAvg);
    stockSheet.getRange(foundIndex, 5).setValue(newValue);
    stockSheet.getRange(foundIndex, 6).setValue(new Date());

  } else {
    // --- CREATE NEW ROW ---
    // Calculate Avg Cost
    const avg = Number(qtyChange) !== 0 ? (Number(valueChange) / Number(qtyChange)) : 0;
    
    // Append: [RefID, WhID, Qty, AvgCost, TotalValue, Date]
    stockSheet.appendRow([
      refId, 
      whId, 
      Number(qtyChange), 
      avg, 
      Number(valueChange), 
      new Date()
    ]);
  }
}

/**
 * HELPER: Fetch Active Layers for FIFO
 */
function getActiveLayers(sheet, refId, whId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const layers = [];
  
  for (let i = 0; i < data.length; i++) {
    // Col A=ID, B=Ref, C=WH, G=Remaining
    if (data[i][1] == refId && data[i][2] == whId && data[i][6] > 0) {
      layers.push({
        id: data[i][0],
        cost: data[i][4],
        remaining: data[i][6],
        date: data[i][3],
        rowIndex: i + 2
      });
    }
  }
  
  // Sort by Date ASC (Oldest First)
  return layers.sort((a, b) => new Date(a.date) - new Date(b.date));
}

/**
 * HELPER: Validation
 */
function validateStockTransaction(ss, data) {
  if (data.qty <= 0) return false;
  
  // For Out/Transfer, check availability
  if (['ADJUSTMENT_OUT', 'TRANSFER'].includes(data.type)) {
    const sheet = ss.getSheetByName('Current_Stock');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    
    const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const stock = rows.find(r => r[0] == data.refId && r[1] == data.whId);
    
    if (!stock || stock[2] < data.qty) {
      return false; // Not enough stock
    }
  }
  return true;
}

/**
 * NEW: Query Item History (Include ID)
 */
function getInventoryHistory(refId) {
  const ss = SpreadsheetApp.openById(MATERIALS_SS_ID);
  const sheet = ss.getSheetByName('Transactions');
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return [];
  // Cols: A=ID (Index 0), B=Date (1), C=Type (2), D=Ref (3), E=WH (4), F=Qty_Change (5), G=Cost (6), J=Notes (9)
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const history = [];
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][3] == refId) {
      history.push({
        id: data[i][0], 
        date: data[i][1],
        type: data[i][2],
        whId: data[i][4],
        qtyChange: Number(data[i][5]),
        cost: Number(data[i][6]) || 0, // [Fix 1] Include Cost
        notes: data[i][9]
      });
    }
  }
  
  // [Fix 4] Removed backend sorting to save processing. Frontend handles it.
  return history; 
}
// =============================================================================
// MATERIALS MODULE API (materials.gs)
// =============================================================================

/**
 * FETCH DASHBOARD DATA (Optimized with Caching)
 */
function getMaterialsInitialData() {
  try {
    const ss = SpreadsheetApp.openById(MATERIALS_SS_ID);
    
    // 1. Fetch Warehouses
    const whSheet = ss.getSheetByName('Warehouses');
    const whData = whSheet.getRange(2, 1, whSheet.getLastRow() - 1, 6).getValues();
    const warehouses = whData.map(r => ({ 
      id: r[0], name: r[2], color: r[5] || '#64748b' 
    }));
    
    // 2. Fetch Material Structure & Stock
    const matData = getCachedMaterialStructure(ss);
    
    // --- UPDATED: Fetch Reorder Points for Groups ---
    const allMatSheet = ss.getSheetByName('All Materials');
    const allMatData = allMatSheet.getDataRange().getValues();
    const reorderMap = {};
    
    // Map Group Name (Column E, Index 4) -> Reorder Point (Column L, Index 11)
    allMatData.forEach(row => {
      // row[4] is Column E (Group Name)
      // row[11] is Column L (Reorder Point)
      if(row[4]) reorderMap[row[4]] = Number(row[11]) || 0; 
    });
    
    // Attach Reorder Point to the Group Objects
    matData.grouped.forEach(group => {
      group.reorder = reorderMap[group.name] || 0;
    });
    // ------------------------------------------------

    const stockSheet = ss.getSheetByName('Current_Stock');
    const stockRows = stockSheet.getLastRow() > 1 ? stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 6).getValues() : [];
    
    const dashboard = stockRows.map(row => {
      const refId = row[0];
      const details = matData.flatMap[refId] || { name: 'Unknown', reorder: 0 };
      return {
        refId: refId,
        whId: row[1],
        qty: Number(row[2]),
        value: Number(row[4]),
        name: details.name, 
        reorderPoint: Number(details.reorder),
        status: (Number(row[2]) <= details.reorder && details.reorder > 0) ? 'LOW' : 'OK'
      };
    });
    
    return JSON.stringify({
      success: true,
      warehouses: warehouses,
      stock: dashboard,
      materialsGrouped: matData.grouped
    });
    
  } catch (e) {
    return JSON.stringify({ success: false, message: e.toString() });
  }
}

/**
 * NEW: Fetch Item History & Balance at Date
 */
function getMaterialDetails(refId) {
  try {
    const ss = SpreadsheetApp.openById(MATERIALS_SS_ID); // Open SS to get stock
    
    // [Fix 3] Fetch actual live stock for sanity check
    const stockSheet = ss.getSheetByName('Current_Stock');
    let liveStock = 0;
    const stockData = stockSheet.getDataRange().getValues();
    // Sum qty for this refId across ALL warehouses
    for (let i = 1; i < stockData.length; i++) {
      if (String(stockData[i][0]) === String(refId)) {
        liveStock += (Number(stockData[i][2]) || 0);
      }
    }

    const history = getInventoryHistory(refId);
    return JSON.stringify({ success: true, history: history, liveStock: liveStock });
  } catch (e) {
    return JSON.stringify({ success: false, message: e.toString() });
  }
}

function submitBatchTransfer(data) {
  const ss = SpreadsheetApp.openById(MATERIALS_SS_ID);
  
  const successfulIndices = [];
  const failures = []; 
  const timestamp = new Date().getTime(); // Common timestamp for the whole batch

  if (!data.items || data.items.length === 0) {
    return { success: false, message: "No items to transfer" };
  }

  data.items.forEach((item, index) => {
    try {
      const qty = parseFloat(item.qty);
      if (qty <= 0) throw new Error("Quantity must be positive");

      // GENERATE ID: TRF + Timestamp + Index (padded) -> e.g. TRF-1678899000-01
      // We use concatenation to ensure the regex (\d+) matches the whole numeric part.
      // Example ID: TRF-167889900001
      const uniqueSuffix = String(index).padStart(3, '0');
      const baseId = `TRF-${timestamp}${uniqueSuffix}`;

      // --- STEP 1: REMOVE FROM SOURCE ---
      const outResult = handleStockOut(ss, {
         txType: 'TRANSFER_OUT',
         refId: item.refId,
         whId: data.sourceWh,
         qty: qty,
         notes: `Transfer to ${data.destWh}`,
         // FORCE THE ID:
         forcedTxId: `${baseId}-OUT`,
         forcedDate: new Date(timestamp)
      });

      // --- STEP 2: ADD TO DESTINATION ---
      handleStockIn(ss, {
         txType: 'TRANSFER_IN',
         type: 'TRANSFER_IN',
         refId: item.refId,
         whId: data.destWh,
         qty: qty,
         unitCost: outResult.avgCost, 
         cost: outResult.avgCost,     
         notes: `Transfer from ${data.sourceWh}`,
         // FORCE THE ID (Must match the baseId of the OUT transaction):
         forcedTxId: `${baseId}-IN`,
         forcedDate: new Date(timestamp)
      });

      successfulIndices.push(index);
    } catch (e) {
      failures.push({
        index: index,
        refId: item.refId,
        message: e.message
      });
    }
  });

  const isCompleteSuccess = failures.length === 0;
  const isPartial = successfulIndices.length > 0 && failures.length > 0;
  
  let msg = "تم تنفيذ التحويل بنجاح";
  if (!isCompleteSuccess) {
    msg = isPartial 
      ? `تم تحويل ${successfulIndices.length} أصناف، وفشل ${failures.length} أصناف.`
      : `فشلت العملية لجميع الأصناف (${failures.length}).`;
  }

  return {
    success: isCompleteSuccess,
    partial: isPartial,
    processedIndices: successfulIndices, 
    failures: failures, 
    message: msg
  };
}

function submitMaterialTransaction(form) {
  const ss = SpreadsheetApp.openById(MATERIALS_SS_ID);
  
  const txData = {
    txType: form.txType,
    type: form.txType,
    refId: form.refId,
    whId: form.whId,
    qty: parseFloat(form.qty),
    cost: form.cost, // Raw input
    unitCost: form.cost ? parseFloat(form.cost) : 0,
    notes: form.notes
  };

  try {
    // --- ROUTING LOGIC ---
    
    if (txData.txType === 'OPENING_BALANCE') {
      // 1. OPENING BALANCE (Smart Adjust: Add or Deduct)
      return handleOpeningBalance(ss, txData);
      
    } else if (txData.txType === 'ADJUSTMENT_OUT') {
      // 2. STOCK OUT (FIFO)
      return handleStockOut(ss, txData);
      
    } else {
      // 3. STOCK IN (Manual/Auto Cost)
      // Covers: ADJUSTMENT_IN
      return handleStockIn(ss, txData);
    }
    
  } catch (e) {
    Logger.log("Error in submitMaterialTransaction: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * OPTIMIZATION: Cache the Master Structure for 25 minutes
 */
function getCachedMaterialStructure(ss) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('MATERIALS_STRUCTURE_V1');
  
  if (cachedData) {
    return JSON.parse(cachedData);
  }
  
  // If not in cache, fetch and store
  const data = getMasterMaterialStructure(ss);
  
  // Cache for 1500 seconds (25 mins)
  // We must ensure the JSON string is not too large (100KB limit). 
  // If it's too large, CacheService fails silently or throws. 
  try {
    cache.put('MATERIALS_STRUCTURE_V1', JSON.stringify(data), 1500);
  } catch (e) {
    Logger.log("Cache put failed (likely too large): " + e.message);
  }
  
  return data;
}

function getMasterMaterialStructure(ss) {
  // ... (Same logic as before) ...
  const sheet = ss.getSheetByName('All Materials');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 5) return { flatMap: {}, grouped: [] };
  
  const data = sheet.getRange(5, 1, lastRow - 4, lastCol).getValues();
  const flatMap = {};
  const groupsObj = {}; 
  
  data.forEach(row => {
    const genericName = row[4]; 
    const reorderPoint = row[11];
    if (!genericName) return;
    if (!groupsObj[genericName]) groupsObj[genericName] = [];

    for (let i = 0; i < 9; i++) {
      const refId = row[50 + i]; // AY
      if (refId && refId.toString().trim() !== '') {
        const fullName = row[59 + i]; // BH
        const supplier = row[14 + i];
        const mark = row[23 + i];
        const size = row[32 + i];
        const pack = row[41 + i];
        
        const varLabel = `${supplier || ''} ${mark || ''} [${size || ''} ${pack || ''}]`.trim();

        flatMap[refId] = {
          name: fullName || genericName,
          reorder: reorderPoint || 0
        };
        groupsObj[genericName].push({
          id: refId,
          label: varLabel || "Standard"
        });
      }
    }
  });

  const grouped = Object.keys(groupsObj).map(key => ({
    name: key,
    variations: groupsObj[key]
  }));

  return { flatMap: flatMap, grouped: grouped };
}

/**
 * API: Fetch last cost for frontend display
 */
function api_getLastCost(refId) {
  try {
    const ss = SpreadsheetApp.openById(MATERIALS_SS_ID);
    // Uses the helper we created in inventory.gs
    const cost = getLastRecordedCost(ss, refId); 
    return { success: true, cost: cost };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_getItemHistory(refId) {
  try {
    const ss = SpreadsheetApp.openById(MATERIALS_SS_ID);
    const sheet = ss.getSheetByName('Transactions');
    const data = sheet.getDataRange().getValues();
    
    // Data Structure:
    // [0]TxID, [1]Date, [2]Type, [3]RefID, [4]WhID, [5]Qty, [6]Cost, [7]Total, [8]Layers, [9]Notes
    
    const history = data
      .slice(1) // Skip header
      .filter(row => String(row[3]) === String(refId)) // Filter by Ref ID
      .map(row => {
        // Format Date
        let dateStr = '';
        try { dateStr = new Date(row[1]).toLocaleDateString('en-GB'); } catch(e){}

        return {
          txId: row[0], // <--- ENSURE THIS IS MAPPED
          date: dateStr,
          type: row[2],
          whId: row[4],
          qty: Number(row[5]),
          cost: Number(row[6]),
          notes: row[9] || ''
        };
      })
      .reverse(); // Newest first
      
    return { success: true, history: history };
    
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getCurrentStockQty(ss, refId, whId) {
  const sheet = ss.getSheetByName('Current_Stock');
  const data = sheet.getDataRange().getValues();
  
  // Loop rows (skip header)
  // Col A (0) = RefID, Col B (1) = WhID, Col C (2) = Qty
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(refId) && String(data[i][1]) === String(whId)) {
      return Number(data[i][2]) || 0;
    }
  }
  return 0; // Not found = 0
}

function handleOpeningBalance(ss, data) {
    // 1. Get REAL Stock from FIFO Layers (Source of Truth)
    // We ignore the cache here to prevent the "Insufficient Stock" error
    const currentQty = getFifoStockQty(ss, data.refId, data.whId);
    const targetQty = data.qty; // User's input is the TARGET

    // 2. Calculate Difference
    const diff = targetQty - currentQty;

    // If matches exactly, do nothing
    if (Math.abs(diff) < 0.000001) {
        return { success: true, message: 'الرصيد مطابق بالفعل (لم يتم إجراء أي تغيير)', txId: null };
    }

    // 3. Prepare Sub-Transaction
    const subData = { 
        ...data,
        qty: Math.abs(diff), // Functions expect positive quantity
        notes: data.notes + ` (Auto-Adjust: Target ${targetQty}, Was ${currentQty})`
    };

    let result;
    if (diff > 0) {
        // Case A: Shortage -> ADD Stock
        result = handleStockIn(ss, subData);
    } else {
        // Case B: Surplus -> REMOVE Stock
        // Since we calculated diff based on FIFO, we are guaranteed to have enough stock to remove.
        result = handleStockOut(ss, subData);
    }
    
    // 4. CRITICAL: Force Sync Cache to Target
    // Since previous cache might have been wrong (phantom stock), 
    // we force the Current_Stock sheet to match the Target Qty exactly.
    // We assume the cost used is either the manual input or the last detected cost.
    const usedCost = subData.unitCost || subData.cost || 0; // Attempt to grab cost from subData handling
    forceSetStockCache(ss, data.refId, data.whId, targetQty, usedCost);

    return result;
}

/**
 * Helper: Sums up actual available quantity in FIFO Layers
 */
function getFifoStockQty(ss, refId, whId) {
  const sheet = ss.getSheetByName('FIFO_Layers');
  const data = sheet.getDataRange().getValues();
  let total = 0;
  
  // Col B (1) = RefID, Col C (2) = WhID, Col G (6) = Remaining Qty
  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === String(refId) && String(data[i][2]) === String(whId)) {
      total += Number(data[i][6]) || 0;
    }
  }
  return total;
}

/**
 * Force Overwrite the 'Current_Stock' sheet with a specific quantity.
 * Used for Opening Balances to fix sync issues.
 */
function forceSetStockCache(ss, refId, whId, exactQty, lastCost) {
  const stockSheet = ss.getSheetByName('Current_Stock');
  const range = stockSheet.getDataRange();
  const values = range.getValues();
  
  // Calculate Value
  const totalValue = exactQty * lastCost;
  
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(refId) && String(values[i][1]) === String(whId)) {
      // Found row - Overwrite it
      stockSheet.getRange(i + 1, 3).setValue(exactQty);     // Qty
      stockSheet.getRange(i + 1, 5).setValue(totalValue);   // Value
      stockSheet.getRange(i + 1, 6).setValue(new Date());   // Date
      return;
    }
  }
  
  // If not found, create new
  stockSheet.appendRow([refId, whId, exactQty, lastCost, totalValue, new Date()]);
}
/**
 * BATCH ADJUSTMENT (High-Performance Version v2)
 * Reads all data once, processes in memory, writes once.
 * [UPDATED]: Adds "Target vs Was" details to Transaction Notes.
 */
function processStockTakeBatch(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
    
    const ss = SpreadsheetApp.openById(MATERIALS_SS_ID);
    const stockSheet = ss.getSheetByName('Current_Stock');
    const layersSheet = ss.getSheetByName('FIFO_Layers');
    const txSheet = ss.getSheetByName('Transactions');
    
    // 1. BULK READ
    const stockData = stockSheet.getDataRange().getValues(); 
    const layersData = layersSheet.getDataRange().getValues(); 
    
    const newTxRows = [];
    const newLayerRows = [];
    const processedDate = new Date(payload.date);
    const timestamp = processedDate.getTime();
    
    let successCount = 0;
    let failCount = 0;
    const errors = [];

    // Map Stock for fast lookup
    const stockMap = new Map();
    for(let i=1; i<stockData.length; i++) {
        stockMap.set(`${stockData[i][0]}_${stockData[i][1]}`, i);
    }

    // 2. PROCESS LOOP
    payload.rows.forEach((row, idx) => {
        try {
            const refId = String(row.refId);
            const whId = String(row.whId);
            const targetQty = Number(row.qty);
            const mapKey = `${refId}_${whId}`;
            
            // A. Get Current System Stock
            let currentQty = 0;
            let currentValue = 0;
            let stockRowIdx = -1;
            
            if (stockMap.has(mapKey)) {
                stockRowIdx = stockMap.get(mapKey);
                currentQty = Number(stockData[stockRowIdx][2]);
                currentValue = Number(stockData[stockRowIdx][4]);
            }

            // B. Calculate Delta
            const diff = targetQty - currentQty;
            
            if (Math.abs(diff) < 0.000001) return; 

            // [FIX] Prepare Detailed Note
            const noteSuffix = ` (Auto-Adjust: Target ${targetQty}, Was ${currentQty})`;
            const userNote = row.notes || '';
            const finalTxNote = `Batch Stock Take: ${userNote}${noteSuffix}`;

            // C. Prepare Transaction Data
            const txId = `TX-${timestamp}-${idx}`;
            let unitCost = 0;
            let totalValChange = 0;

            if (diff > 0) {
                // --- CASE: ADD STOCK (IN) ---
                unitCost = Number(row.cost);
                if (!unitCost || unitCost === 0) {
                    unitCost = (currentQty > 0) ? (currentValue / currentQty) : 0;
                }

                totalValChange = diff * unitCost;

                const layerId = `LAY-${timestamp}-${idx}`;
                newLayerRows.push([
                    layerId, refId, whId, processedDate, unitCost, diff, diff, txId
                ]);

                newTxRows.push([
                    txId, processedDate, 'OPENING_BALANCE', refId, whId, diff, 
                    unitCost, totalValChange, layerId, finalTxNote // <--- Updated Note
                ]);

            } else {
                // --- CASE: REMOVE STOCK (OUT) ---
                const qtyToRemove = Math.abs(diff);
                let remainingToRemove = qtyToRemove;
                let costAccumulator = 0;
                const consumedLayerIds = [];

                for (let i = 1; i < layersData.length; i++) {
                    if (remainingToRemove <= 0) break;
                    
                    if (String(layersData[i][1]) === refId && String(layersData[i][2]) === whId && Number(layersData[i][6]) > 0) {
                        const availableInLayer = Number(layersData[i][6]);
                        const deduct = Math.min(remainingToRemove, availableInLayer);
                        
                        layersData[i][6] = availableInLayer - deduct; 
                        remainingToRemove -= deduct;
                        costAccumulator += (deduct * Number(layersData[i][4]));
                        consumedLayerIds.push(layersData[i][0]);
                    }
                }

                unitCost = (qtyToRemove > 0) ? (costAccumulator / qtyToRemove) : 0;
                totalValChange = -costAccumulator;

                newTxRows.push([
                    txId, processedDate, 'OPENING_BALANCE', refId, whId, diff, 
                    unitCost, -costAccumulator, consumedLayerIds.join(','), finalTxNote // <--- Updated Note
                ]);
            }

            // D. Update Current Stock Cache
            const newTotalValue = currentValue + totalValChange;
            const newAvgCost = (targetQty > 0) ? (newTotalValue / targetQty) : 0;

            if (stockRowIdx !== -1) {
                stockData[stockRowIdx][2] = targetQty;
                stockData[stockRowIdx][3] = newAvgCost;
                stockData[stockRowIdx][4] = newTotalValue;
                stockData[stockRowIdx][5] = new Date();
            } else {
                const newRow = [refId, whId, targetQty, newAvgCost, newTotalValue, new Date()];
                stockData.push(newRow);
                stockMap.set(mapKey, stockData.length - 1); 
            }

            successCount++;

        } catch (e) {
            failCount++;
            errors.push(`Row ${idx+1}: ${e.message}`);
        }
    });

    // 3. BULK WRITE
    if (newTxRows.length > 0) {
        txSheet.getRange(txSheet.getLastRow() + 1, 1, newTxRows.length, newTxRows[0].length).setValues(newTxRows);
    }

    if (newLayerRows.length > 0) {
        layersSheet.getRange(layersSheet.getLastRow() + 1, 1, newLayerRows.length, newLayerRows[0].length).setValues(newLayerRows);
    }

    if (layersData.length > 1) {
        const remainingQtyCol = layersData.map(r => [r[6]]); 
        layersSheet.getRange(1, 7, remainingQtyCol.length, 1).setValues(remainingQtyCol);
    }

    if (stockData.length > 0) {
        stockSheet.getRange(1, 1, stockData.length, stockData[0].length).setValues(stockData);
    }

    return { success: successCount, failed: failCount, errors: errors };

  } catch (e) {
    return { success: 0, failed: 1, errors: [e.message] };
  } finally {
    lock.releaseLock();
  }
}

/**
 * API: Fetch Available FIFO Layers for a specific item/warehouse
 * Returns JSON String to avoid serialization issues
 */
function getOpenLayers(refId, whId) {
  const ss = SpreadsheetApp.openById(MATERIALS_SS_ID);
  const sheet = ss.getSheetByName('FIFO_Layers');
  
  // Guard: Return empty JSON array if sheet missing
  if (!sheet) return '[]';

  const data = sheet.getDataRange().getValues();
  const layers = [];
  
  // Iterate (Skip Header)
  // Col A(0)=ID, B(1)=Ref, C(2)=WH, D(3)=Date, E(4)=Cost, G(6)=RemQty
  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === String(refId) && String(data[i][2]) === String(whId)) {
      const rem = Number(data[i][6]);
      if (rem > 0) {
        layers.push({
          id: data[i][0],
          date: data[i][3], // Date Object
          cost: Number(data[i][4]),
          qty: rem
        });
      }
    }
  }
  
  // Sort Oldest First and STRINGIFY
  layers.sort((a, b) => new Date(a.date) - new Date(b.date));
  return JSON.stringify(layers);
}
