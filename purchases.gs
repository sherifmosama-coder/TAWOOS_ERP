// =============================================================================
// PURCHASES MODULE BACKEND (purchases.gs)
// =============================================================================

const PURCHASES_SPREADSHEET_ID = '15riHuotxGDPJV5Sw0GJoXqatbANj8poTWpUBRcdkbe4';

/**
 * FETCH INITIAL DATA
 */
function getPurchasesInitialData() {
  try {
    const ss = SpreadsheetApp.openById(PURCHASES_SPREADSHEET_ID);
    
    // 1. Permissions Check
    const userEmail = Session.getActiveUser().getEmail();
    let isAdmin = false;
    try {
        const permissions = getPermissions();
        isAdmin = permissions[userEmail] === 'admin';
    } catch(e) { isAdmin = false; }

    // 2. Fetch Operations
    const indexSheet = ss.getSheetByName('Index');
    let operations = [];
    if (indexSheet) {
      const lastRow = indexSheet.getLastRow();
      if (lastRow >= 2) {
         operations = indexSheet.getRange(2, 3, lastRow - 1, 1).getValues()
           .map(r => r[0].toString().trim()).filter(p => p !== '');
      }
    }

    // 3. Fetch Warehouses
    const matData = JSON.parse(getMaterialsInitialData());
    
    // 4. Fetch Materials Structure & Units
    const matStruct = getPurchasesMaterialStructure(SpreadsheetApp.openById(MATERIALS_SS_ID));

    // 5. Fetch History
    const historySheet = ss.getSheetByName('استلام و ارتجاع');
    let history = [];
    if (historySheet && historySheet.getLastRow() >= 2) {
      const data = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 22).getValues();
      history = data.map(r => ({
        id: r[1] || r[0],
        date: r[2] instanceof Date ? r[2].toISOString().slice(0,10) : r[2],
        op: r[3],
        supplier: r[8],
        item: r[5],
        total: Number(r[21]) || 0
      })).reverse();
    }

    return { 
      success: true, 
      isAdmin: isAdmin,
      warehouses: matData.warehouses || [],
      operations: operations,
      materialsGrouped: matStruct.groups, 
      units: matStruct.units, // Universal Unit List
      history: history 
    };

  } catch (e) {
    Logger.log('getPurchasesInitialData Error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * PARSER: Reads 'All Materials' for 3-Step Flow + Units + Taxes
 */
function getPurchasesMaterialStructure(ss) {
  const sheet = ss.getSheetByName('All Materials');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 5) return { groups: [], units: [] };
  
  const data = sheet.getRange(5, 1, lastRow - 4, lastCol).getValues();
  const groupsObj = {};
  const unitsSet = new Set();
  
  data.forEach(row => {
    const genericName = row[4]; // Col E: Basic Name
    const isService = String(row[8]).toUpperCase() === 'X'; // Col I
    const smallUnit = row[10]; // Col K: Small Unit
    const vatRate = row[12];   // Col M: VAT Rate
    const whtRate = row[13];   // Col N: WHT Rate
    
    if (smallUnit) unitsSet.add(smallUnit);

    if (!genericName) return;
    if (!groupsObj[genericName]) {
        groupsObj[genericName] = {
            name: genericName,
            isService: isService,
            variations: []
        };
    }

    // Loop through 9 variations
    for (let i = 0; i < 9; i++) {
      const refId = row[50 + i]; // Col AY + i
      if (refId && refId.toString().trim() !== '') {
        const specificName = row[59 + i]; // Col BH
        const supplier = row[14 + i];     // Col O
        const pack = row[41 + i];         // Col AP
        const packQty = row[32 + i];      // Col AG
        const largeUnitName = row[41 + i]; // Col AP
        const specialSign = row[23 + i];   // Col X + i (NEW: Special Sign)

        if (largeUnitName) unitsSet.add(largeUnitName);

        groupsObj[genericName].variations.push({
          id: refId,
          specificName: specificName || genericName,
          supplier: supplier || 'Unknown',
          specialSign: specialSign || '', // NEW
          pack: packQty || 1,
          smallUnit: smallUnit || 'قطعة',
          largeUnit: largeUnitName || 'كرتونة',
          vatRate: (typeof vatRate === 'number') ? vatRate : 0.14,
          whtRate: (typeof whtRate === 'number') ? whtRate : 0.01
        });
      }
    }
  });

  return { 
      groups: Object.values(groupsObj), 
      units: Array.from(unitsSet).sort() 
  };
}

/**
 * SAVE TRANSACTION
 */
function savePurchaseTransaction(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(PURCHASES_SPREADSHEET_ID);
    let sheet = ss.getSheetByName('استلام و ارتجاع');
    
    // --- 1. GENERATE ID ---
    const isReceipt = form.operation === 'استلام';
    const idCol = isReceipt ? 2 : 1; 
    const lastRow = sheet.getLastRow();
    
    let maxSeq = 0;
    if (lastRow >= 2) {
        const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat();
        ids.forEach(val => {
            const str = String(val);
            const num = parseInt(str.slice(-3)); 
            if (!isNaN(num) && num > maxSeq) maxSeq = num;
        });
    }
    const nextSeq = String(maxSeq + 1).padStart(3, '0');
    const transactionId = isReceipt ? `2026${nextSeq}` : `Re-2026${nextSeq}`;

    // --- 2. PREPARE DATA & CALCULATE SHIPPING ---
    const rowsToAdd = [];
    const structure = getPurchasesMaterialStructure(SpreadsheetApp.openById(MATERIALS_SS_ID)).groups;
    
    // Calculate Total Qty of Transaction to distribute shipping
    const grandTotalQty = form.items.reduce((sum, item) => sum + parseFloat(item.totalQty), 0);
    const totalShipping = parseFloat(form.transportationFee) || 0;

    form.items.forEach(item => {
        // Validation & Service Check
        const basicGroup = structure.find(g => g.name === item.basicName);
        const isService = basicGroup ? basicGroup.isService : false;

        // Shipping Share Calculation (Weighted)
        const rowQty = parseFloat(item.totalQty);
        const shippingShare = grandTotalQty > 0 ? (rowQty / grandTotalQty) * totalShipping : 0;
        
        // Inventory Cost (Loaded Cost = Price + Unit Shipping Share)
        const unitShipping = rowQty > 0 ? (shippingShare / rowQty) : 0;
        const loadedCost = parseFloat(item.price) + unitShipping;

        // Shifted Columns (H is Special Sign, I is Supplier, etc.)
        const row = [
            !isReceipt ? transactionId : '',    // A
            isReceipt ? transactionId : '',     // B
            form.date,                          // C
            form.operation,                     // D
            form.poNumber || '',                // E
            item.basicName,                     // F
            item.specificName,                  // G
            item.specialSign,                   // H (New: Special Sign)
            form.supplier,                      // I (Shifted)
            item.warehouseName,                 // J (Shifted)
            item.qtyMajor,                      // K (Shifted)
            item.unitMajor,                     // L
            item.qtyMinor,                      // M (Pack Size)
            item.unitMinor,                     // N
            item.totalQty,                      // O
            item.price,                         // P (Invoice Price)
            item.isTax ? 'نعم' : 'لا',          // Q
            item.isDisc1 ? 'نعم' : 'لا',        // R
            item.preTaxTotal,                   // S
            item.taxValue,                      // T
            item.discountValue,                 // U
            item.lineTotal,                     // V
            shippingShare.toFixed(2),           // W (Calculated Shipping Share)
            form.notes                          // X
        ];
        rowsToAdd.push(row);

        // --- 3. INVENTORY UPDATE (Using Loaded Cost) ---
        if (!isService) {
            const txType = isReceipt ? 'ADJUSTMENT_IN' : 'ADJUSTMENT_OUT'; 
            processInventoryTransaction({
                type: txType,
                refId: item.refId, 
                whId: item.warehouseId,
                qty: parseFloat(item.totalQty),
                unitCost: loadedCost, // NEW: Price + Shipping per unit
                cost: loadedCost,     // NEW: Price + Shipping per unit
                notes: `${form.operation} - ${form.supplier}`,
                forcedTxId: transactionId,
                forcedDate: new Date(form.date)
            });
        }
    });

    if (rowsToAdd.length > 0) {
        sheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }

    return { success: true, message: 'تم الحفظ بنجاح', transactionId: transactionId };

  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
