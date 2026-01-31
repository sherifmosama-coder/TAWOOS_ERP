// =============================================================================
// PURCHASES MODULE BACKEND (purchases.gs)
// =============================================================================

const PURCHASES_SPREADSHEET_ID = '15riHuotxGDPJV5Sw0GJoXqatbANj8poTWpUBRcdkbe4';

/**
 * FETCH INITIAL DATA
 */
function getPurchasesInitialData(clientEmail) {
  try {
    const ss = SpreadsheetApp.openById(PURCHASES_SPREADSHEET_ID);
    
    // 1. Permissions Check
    const userEmail = clientEmail || ''; 
    const accessLevel = resolveModulePermission(userEmail, 'operations', 'purchases');
    
    // Allow both Admin and Editor to have "Edit" rights
    const canEdit = (accessLevel === 'editor' || accessLevel === 'admin');
    const isAdmin = (accessLevel === 'admin');

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

    // 5. Fetch History & Calculate Next IDs
const historySheet = ss.getSheetByName('استلام و ارتجاع');
let history = [];
let maxR = 0; // Receipt
let maxRe = 0; // Return
let maxC = 0; // Cost
let maxSa = 0; // Sample

if (historySheet && historySheet.getLastRow() >= 2) {
  // Expand fetch range to Column AF (Index 32)
  const data = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 32).getValues();

  // History for Frontend (Rich Object)
  history = data.map(r => ({
    id: r[1] || r[0], // Col A or B
    date: r[2] instanceof Date ? r[2].toISOString().slice(0,10) : r[2], // Col C
    op: r[3],           // Col D
    po: r[4],           // Col E
    item: r[5],         // Col F (Basic)
    specific: r[6],     // Col G (Specific)
    supplier: r[8],     // Col I
    warehouse: r[9],    // Col J (Warehouse)
    qty: Number(r[14]) || 0, // Col O (Total Units)
    price: Number(r[15]) || 0, // Col P
    vatVal: Number(r[19]) || 0, // Col T
    total: Number(r[21]) || 0,  // Col V (Line Total)
    note: r[23],        // Col X (Item Note)
    user: r[24],        // Col Y (User Email)
    ts: r[25] instanceof Date ? r[25].toISOString() : r[25], // Col Z (Timestamp)
    invNum: r[26],      // Col AA
    invDate: r[27] instanceof Date ? r[27].toISOString().slice(0,10) : r[27], // Col AB
    whtNum: r[28],      // Col AC
    linkInv: r[29],     // Col AD (Supplier Invoice)
    linkWht: r[30],     // Col AE (WHT Note)
    linkRec: r[31]      // Col AF (Receipt/Return Note)
  })).reverse();

  // Calculate Next IDs
  // Col A (0): Returns (RE-), Cost (C-), Sample (Sa-)
  // Col B (1): Receipts (2026...)
  data.forEach(row => {
    // Receipts (Col B)
    if (row[1]) {
        const str = String(row[1]);
        if(str.startsWith('2026')) {
            const num = parseInt(str.slice(-3)); 
            if (!isNaN(num) && num > maxR) maxR = num;
        }
    }
    // Others (Col A)
    if (row[0]) {
        const str = String(row[0]);
        if (str.startsWith('RE-2026')) {
            const num = parseInt(str.slice(-3)); 
            if (!isNaN(num) && num > maxRe) maxRe = num;
        } else if (str.startsWith('C-2026')) {
            const num = parseInt(str.slice(-3)); 
            if (!isNaN(num) && num > maxC) maxC = num;
        } else if (str.startsWith('Sa-2026')) {
            const num = parseInt(str.slice(-3)); 
            if (!isNaN(num) && num > maxSa) maxSa = num;
        }
    }
  });
}

return { 
  success: true, 
  isAdmin: isAdmin,
  canEdit: canEdit, 
  userRole: accessLevel, 
  warehouses: matData.warehouses || [],
  operations: operations,
  materialsGrouped: matStruct.groups, 
  units: matStruct.units,
  history: history,
  nextReceiptId: `2026${String(maxR + 1).padStart(3, '0')}`,
  nextReturnId: `RE-2026${String(maxRe + 1).padStart(3, '0')}`,
  nextCostId: `C-2026${String(maxC + 1).padStart(3, '0')}`,
  nextSampleId: `Sa-2026${String(maxSa + 1).padStart(3, '0')}`
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
    const genericName = row[4]; 
    const isService = String(row[8]).toUpperCase() === 'X'; 
    const smallUnit = row[10]; 
    const vatRate = row[12];   
    const whtRate = row[13];   
    
    if (smallUnit) unitsSet.add(smallUnit);

    if (!genericName) return;
    if (!groupsObj[genericName]) {
        groupsObj[genericName] = {
            name: genericName,
            isService: isService,
            variations: []
        };
    }

    for (let i = 0; i < 9; i++) {
      const refId = row[50 + i]; 
      if (refId && refId.toString().trim() !== '') {
        const specificName = row[59 + i]; 
        const supplier = row[14 + i];     
        const pack = row[41 + i];         
        const packQty = row[32 + i];      
        const largeUnitName = row[41 + i]; 
        const specialSign = row[23 + i];   

        if (largeUnitName) unitsSet.add(largeUnitName);

        groupsObj[genericName].variations.push({
          id: refId,
          specificName: specificName || genericName,
          supplier: supplier || 'Unknown',
          specialSign: specialSign || '',
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
 * SAVE TRANSACTION (UPDATED WITH PDF GENERATION)
 */
function savePurchaseTransaction(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // --- SECURITY GUARD ---
    const userEmail = form.userEmail; 
    const role = resolveModulePermission(userEmail, 'operations', 'purchases');
    
    if (role !== 'editor' && role !== 'admin') {
      return { success: false, message: 'Security Alert: You do not have permission to edit Purchases.' };
    }

    const ss = SpreadsheetApp.openById(PURCHASES_SPREADSHEET_ID);
    let sheet = ss.getSheetByName('استلام و ارتجاع');
    
    // --- 1. GENERATE ID ---
    const op = form.operation;
    const isReceipt = (op === 'استلام');
    let prefix = '';
    let idCol = 1; // Col A (0)

    if (isReceipt) {
        prefix = '2026';
        idCol = 2; // Col B (1)
    } else if (op === 'ارتجاع' || op === 'return') {
        prefix = 'RE-2026';
    } else if (op === 'تعديل تكلفة' || op.includes('Cost')) {
        prefix = 'C-2026';
    } else if (op === 'عينة' || op.includes('Sample')) {
        prefix = 'Sa-2026';
    } else {
        prefix = 'RE-2026'; 
    }

    const lastRow = sheet.getLastRow();
    let maxSeq = 0;

    if (lastRow >= 2) {
        const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat();
        ids.forEach(val => {
            const str = String(val);
            if (str.startsWith(prefix)) {
                const num = parseInt(str.slice(-3)); 
                if (!isNaN(num) && num > maxSeq) maxSeq = num;
            }
        });
    }
    const transactionId = `${prefix}${String(maxSeq + 1).padStart(3, '0')}`;

    // --- 2. GENERATE PDF NOTE ---
    let pdfData = { url: '', html: '' };
    try {
      pdfData = createPurchaseNote(transactionId, form);
    } catch (err) {
      Logger.log('PDF Generation Failed: ' + err.toString());
      // Continue saving even if PDF fails, but log it
    }

    // --- 3. PREPARE DATA ---
    const rowsToAdd = [];
    const structure = getPurchasesMaterialStructure(SpreadsheetApp.openById(MATERIALS_SS_ID)).groups;
    
    const grandTotalQty = form.items.reduce((sum, item) => sum + parseFloat(item.totalQty), 0);
    const totalShipping = parseFloat(form.transportationFee) || 0;

    form.items.forEach(item => {
        const basicGroup = structure.find(g => g.name === item.basicName);
        const isService = basicGroup ? basicGroup.isService : false;

        const rowQty = parseFloat(item.totalQty);
        const shippingShare = grandTotalQty > 0 ? (rowQty / grandTotalQty) * totalShipping : 0;
        const unitShipping = rowQty > 0 ? (shippingShare / rowQty) : 0;
        const loadedCost = parseFloat(item.price) + unitShipping;

        const row = [
            !isReceipt ? transactionId : '',    // A (Return ID)
            isReceipt ? transactionId : '',     // B (Receipt ID)
            form.date,                          // C
            form.operation,                     // D
            form.poNumber || '',                // E
            item.basicName,                     // F
            item.specificName,                  // G
            item.specialSign,                   // H
            form.supplier,                      // I
            item.warehouseName,                 // J
            item.qtyMajor,                      // K
            item.unitMajor,                     // L
            item.qtyMinor,                      // M
            item.unitMinor,                     // N
            item.totalQty,                      // O
            item.price,                         // P
            item.isTax ? 'TRUE' : 'FALSE',      // Q
            item.isDisc1 ? 'TRUE' : 'FALSE',    // R
            item.preTaxTotal,                   // S
            item.taxValue,                      // T
            item.discountValue,                 // U
            item.lineTotal,                     // V
            shippingShare.toFixed(2),           // W
            item.note,                          // X
            userEmail,                          // Y
            Utilities.formatDate(new Date(), "GMT+2", "yyyy-MM-dd HH:mm:ss"), // Z
            form.invoiceNumber || '',           // AA
            form.invoiceDate || '',             // AB
            form.whtNoteNumber || '',           // AC
            '',                                 // AD (Invoice Link - User can add later)
            '',                                 // AE (WHT Link - User can add later)
            pdfData.url                         // AF (Auto-Generated Note Link)
        ];
        rowsToAdd.push(row);

        // --- 4. INVENTORY UPDATE ---
        if (!isService) {
            const txType = isReceipt ? 'ADJUSTMENT_IN' : 'ADJUSTMENT_OUT'; 
            processInventoryTransaction({
                type: txType,
                refId: item.refId, 
                whId: item.warehouseId,
                qty: parseFloat(item.totalQty),
                unitCost: loadedCost, 
                cost: loadedCost,     
                notes: `${form.operation} - ${form.supplier}`,
                forcedTxId: transactionId,
                forcedDate: new Date(form.date)
            });
        }
    });

    if (rowsToAdd.length > 0) {
        sheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }

    // Return PDF URL and HTML for direct printing
    return { 
        success: true, 
        message: 'تم الحفظ بنجاح', 
        transactionId: transactionId, 
        pdfUrl: pdfData.url,
        printHtml: pdfData.html
    };

  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * FETCH EDIT HISTORY (CELL NOTES)
 */
function getTransactionEditHistory(transactionId) {
  const ss = SpreadsheetApp.openById(PURCHASES_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('استلام و ارتجاع');
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return [];

  // Get IDs (Col A & B) and Notes (All Columns)
  // We fetch values to find the row, and notes to get the history
  const range = sheet.getRange(2, 1, lastRow - 1, 32); // A to AF
  const values = range.getValues();
  
  // Filter row indices that match the ID
  const matchIndices = [];
  values.forEach((r, i) => {
    // Check Receipt ID (Col B) or Return ID (Col A)
    if (r[1] == transactionId || r[0] == transactionId) {
      matchIndices.push(i);
    }
  });

  if (matchIndices.length === 0) return [];

  // Now fetch notes ONLY for those rows to optimize
  // We can't fetch disjoint ranges easily, so we fetch the whole block of notes if matches are found
  // or just fetch specific rows. For simplicity and speed in Apps Script, fetching the specific rows is better.
  
  const historyLog = [];
  const headers = ['Ref', 'ID', 'Date', 'Op', 'PO', 'Item', 'Specific', 'Sign', 'Supplier', 'Warehouse', 
                   'Qty', 'Unit', 'Pack', 'Unit', 'Total', 'Price', 'Tax', 'Disc', 'Pre', 'TaxVal', 
                   'DiscVal', 'Total', 'Ship', 'Note', 'User', 'Time', 'Inv', 'InvDate', 'WHT', 'L1', 'L2', 'L3'];

  matchIndices.forEach(idx => {
    // Row Index + 2 (header) = Actual Row Number
    const rowNum = idx + 2;
    const rowNotes = sheet.getRange(rowNum, 1, 1, 32).getNotes()[0];
    const item = values[idx][5]; // Item Name

    rowNotes.forEach((note, colIdx) => {
      if (note && note.trim() !== '') {
        historyLog.push({
          item: item,
          field: headers[colIdx] || `Col ${colIdx+1}`,
          content: note
        });
      }
    });
  });

  return historyLog;
}

/**
 * GET/CREATE FOLDER
 */
function getPurchasesFolder() {
  const folders = DriveApp.getFoldersByName("Purchases Operations");
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder("Purchases Operations");
  }
}

/**
 * GENERATE MODERN SPLIT PDF (Grayscale, Rounded, SVGs)
 * V5: One-Page A4 Fit, Basic Sans-Serif Font, Smart Formatting
 */
function createPurchaseNote(transactionId, form) {
  const folder = getPurchasesFolder();
  const isReceipt = (form.operation === 'استلام');
  
  // 1. TITLES & SIGNATURE LOGIC
  const title = isReceipt ? 'إذن استلام مخزني' : 'إذن ارتجاع بضاعة';
  const subTitleSupplier = 'نسخة المورد';
  const subTitleInternal = 'نسخة الأرشيف';
  
  const signSupplierLeft = isReceipt ? 'توقيع المستلم (المخزن)' : 'توقيع المستلم (المورد)';
  const signSupplierRight = isReceipt ? 'اعتماد المورد' : 'توقيع المسلم (المخزن)';

  // 2. EMBED LOGO (Base64)
  const logoUrl = 'https://lh3.googleusercontent.com/d/1KuZm8n-1MFpWNTUIVbnHBONCVnkWZh7z';
  let logoBase64 = '';
  try {
    const imageBlob = UrlFetchApp.fetch(logoUrl).getBlob();
    const b64 = Utilities.base64Encode(imageBlob.getBytes());
    logoBase64 = `data:${imageBlob.getContentType()};base64,${b64}`;
  } catch (e) { logoBase64 = ''; }

  // 3. HELPERS
  const fmt = (num) => {
    const n = Number(num);
    return n.toLocaleString('en-US', {
      minimumFractionDigits: Number.isInteger(n) ? 0 : 2,
      maximumFractionDigits: 2
    });
  };

  const hasWht = form.items.some(i => i.discountValue > 0);
  const hasNotes = form.items.some(i => i.note); 

  // 4. ICONS
  const icons = {
    id: `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 9h16M4 15h16M10 3L8 21M16 3l-2 18"/></svg>`,
    date: `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`,
    user: `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>`,
    check: `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>`,
    wht: `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>`,
    clock: `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>`,
    note: `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#6b7280" stroke-width="2"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path></svg>`
  };

  // 5. BUILD ROWS
  let supplierRows = '';
  let internalRows = '';
  const internalColSpan = 8 + (hasWht ? 1 : 0);

  form.items.forEach((item, i) => {
    // --- SUPPLIER ROWS ---
    supplierRows += `
      <tr>
        <td style="text-align:center;">${i + 1}</td>
        <td style="font-family:monospace; font-weight:bold;">${item.refId || '-'}</td>
        <td style="font-weight:bold;">${item.basicName}</td>
        ${hasNotes ? `<td style="font-size:0.7rem; color:#4b5563;">${item.note || '-'}</td>` : ''}
        <td style="text-align:center; font-weight:bold;">${fmt(item.totalQty)}</td>
        <td style="text-align:center;">${item.unitMinor}</td>
      </tr>`;

    // --- INTERNAL ROWS ---
    const taxInfo = item.isTax ? `<span style="font-size:0.7rem;">${fmt(item.taxValue)}</span>` : '-';
    const whtInfo = hasWht ? `<span style="font-size:0.7rem; color:#dc2626;">${item.discountValue > 0 ? '('+fmt(item.discountValue)+')' : '-'}</span>` : '';
    const unitDisplay = `
        <div style="font-weight:bold;">${item.unitMinor}</div>
        ${item.qtyMajor > 0 ? `<div style="font-size:0.6rem; color:#6b7280;">[ ${fmt(item.qtyMajor)} ${item.unitMajor} ]</div>` : ''}
    `;
    const borderStyle = item.note ? 'border-bottom:none;' : '';

    internalRows += `
      <tr>
        <td style="text-align:center; ${borderStyle}">${i + 1}</td>
        <td style="${borderStyle}">
            <div style="font-weight:bold;">${item.basicName}</div>
            <div style="font-size:0.7rem; color:#6b7280;">${item.specificName}</div>
        </td>
        <td style="text-align:center; ${borderStyle}">${item.warehouseName}</td>
        <td style="text-align:center; ${borderStyle}">${unitDisplay}</td>
        <td style="text-align:center; font-weight:bold; ${borderStyle}">${fmt(item.totalQty)}</td>
        <td style="text-align:right; ${borderStyle}">${fmt(item.price)}</td>
        <td style="text-align:center; ${borderStyle}">${taxInfo}</td>
        ${hasWht ? `<td style="text-align:center; ${borderStyle}">${whtInfo}</td>` : ''}
        <td style="text-align:right; ${borderStyle}">${fmt(item.lineTotal)}</td>
      </tr>`;
    
    if (item.note) {
      internalRows += `
      <tr>
        <td colspan="${internalColSpan}" style="padding: 0 10px 6px 10px; background:white; border-bottom:1px solid #f3f4f6;">
            <div style="display:flex; align-items:center; gap:4px; font-size:0.7rem; color:#4b5563; background:#f9fafb; padding:2px 8px; border-radius:4px; width:fit-content;">
                ${icons.note}
                <span style="font-style:italic;">${item.note}</span>
            </div>
        </td>
      </tr>`;
    }
  });

  // 6. METADATA
  const metaDate = `<div class="chip">${icons.date} ${form.date}</div>`;
  const metaSupplier = `<div class="chip">${icons.user} ${form.supplier}</div>`;
  const metaPO = form.poNumber ? `<div class="chip">${icons.check} PO: ${form.poNumber}</div>` : '';
  const userName = form.userEmail.split('@')[0];
  const timestamp = Utilities.formatDate(new Date(), "GMT+2", "yyyy-MM-dd HH:mm");
  
  const idChip = `
    <div class="chip" style="font-family:monospace; font-size:0.9rem; background:#f3f4f6; color:#374151; font-weight:800;">
        ${icons.id} ${transactionId}
    </div>`;

  // 7. HTML STRUCTURE
  // Note: We use 'Arial' as the primary font because it is "Basic" and renders Arabic correctly on most systems (including PDF generation).
  const htmlContent = `
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
      <meta charset="UTF-8">
      <style>
        /* Force A4 Size */
        @page { size: A4; margin: 0; }
        
        body { 
            font-family: 'Arial', 'Tahoma', sans-serif; /* Basic Fonts */
            margin: 0; padding: 0; 
            color: #1f2937; background: white; 
            -webkit-print-color-adjust: exact; 
            width: 210mm; height: 297mm; /* Exact A4 dimensions */
            overflow: hidden; /* Prevent spillover */
        }

        .page-container { 
            width: 100%; height: 100%; 
            box-sizing: border-box; 
            display: flex; flex-direction: column; 
        }
        
        /* Two Halves - strictly 50% max */
        .half-section { 
            height: 50%; 
            padding: 15px 25px; /* Reduced Padding */
            box-sizing: border-box; 
            display: flex; flex-direction: column; 
            overflow: hidden;
        }
        
        .separator { 
            height: 0; border-top: 1px dashed #d1d5db; 
            margin: 0; position: relative; width: 100%; 
        }
        .separator::after { 
            content: '✂'; position: absolute; left: 50%; top: -12px; 
            background: white; padding: 0 5px; color: #9ca3af; 
            font-size: 1rem; transform: translateX(-50%); 
        }

        /* Scaled Down Header */
        .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; border-bottom: 1px solid #e5e7eb; padding-bottom: 8px; }
        .brand-box { display: flex; align-items: center; gap: 8px; }
        .brand-logo { width: 40px; height: 40px; border-radius: 6px; object-fit: cover; filter: grayscale(100%); opacity: 0.8; }
        .doc-title h2 { margin: 0; font-size: 1.2rem; font-weight: 700; }
        .doc-title span { font-size: 0.75rem; font-weight: 600; background: #f3f4f6; padding: 2px 6px; border-radius: 4px; color: #4b5563; }
        
        .meta-grid { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 10px; }
        .chip { display: inline-flex; align-items: center; gap: 4px; background: #f9fafb; border: 1px solid #e5e7eb; padding: 2px 8px; border-radius: 50px; font-size: 0.7rem; font-weight: 600; color: #374151; }

        /* Compact Table */
        table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 0.75rem; margin-bottom: auto; }
        th { background: #f3f4f6; color: #374151; padding: 6px; text-align: right; font-weight: 700; border-top: 1px solid #e5e7eb; border-bottom: 1px solid #e5e7eb; }
        th:first-child { border-top-right-radius: 6px; border-bottom-right-radius: 6px; border-right: 1px solid #e5e7eb; }
        th:last-child { border-top-left-radius: 6px; border-bottom-left-radius: 6px; border-left: 1px solid #e5e7eb; }
        td { padding: 6px; border-bottom: 1px solid #f3f4f6; vertical-align: middle; }
        tr:last-child td { border-bottom: none; }

        /* Footer */
        .footer-box { margin-top: 10px; background: #f9fafb; border-radius: 8px; padding: 10px; display: flex; justify-content: space-between; border: 1px solid #f3f4f6; }
        .sign-area { text-align: center; width: 30%; }
        .sign-line { margin-top: 25px; border-top: 1px solid #d1d5db; width: 80%; margin-left: auto; margin-right: auto; }
        .sign-label { font-size: 0.65rem; font-weight: 700; color: #6b7280; margin-top: 4px; }
        
        .user-footer { font-size: 0.65rem; color: #9ca3af; text-align: left; margin-top: 4px; display: flex; gap: 15px; justify-content: flex-end; font-family: monospace; }
      </style>
    </head>
    <body>
      <div class="page-container">
        
        <div class="half-section">
          <div class="header">
            <div class="brand-box">
                <img src="${logoBase64}" class="brand-logo">
                <div class="doc-title">
                    <h2>${title}</h2>
                    <span>${subTitleSupplier}</span>
                </div>
            </div>
            ${idChip}
          </div>

          <div class="meta-grid">
            ${metaDate} ${metaSupplier} ${metaPO}
          </div>

          <table>
            <thead>
                <tr>
                    <th width="5%">#</th>
                    <th width="15%">كود</th>
                    <th>بيان الصنف</th>
                    ${hasNotes ? '<th>ملاحظات</th>' : ''}
                    <th width="10%">الكمية</th>
                    <th width="10%">الوحدة</th>
                </tr>
            </thead>
            <tbody>${supplierRows}</tbody>
          </table>

          <div class="footer-box">
             <div class="sign-area"><div class="sign-line"></div><div class="sign-label">${signSupplierLeft}</div></div>
             <div class="sign-area"><div class="sign-line"></div><div class="sign-label">${signSupplierRight}</div></div>
          </div>
          
          <div class="user-footer">
             <span>${icons.clock} Printed: ${timestamp}</span>
          </div>
        </div>

        <div class="separator"></div>

        <div class="half-section">
          <div class="header">
            <div class="brand-box">
                <img src="${logoBase64}" class="brand-logo">
                <div class="doc-title">
                    <h2>${title}</h2>
                    <span>${subTitleInternal}</span>
                </div>
            </div>
            ${idChip}
          </div>

          <div class="meta-grid">
            ${metaDate} ${metaSupplier}
            ${form.invoiceNumber ? `<div class="chip">${icons.check} Inv: ${form.invoiceNumber}</div>` : ''}
            ${form.whtNoteNumber ? `<div class="chip">${icons.wht} WHT: ${form.whtNoteNumber}</div>` : ''}
          </div>

          <table>
            <thead>
                <tr>
                    <th width="4%">#</th>
                    <th>الصنف</th>
                    <th width="10%">المخزن</th>
                    <th width="10%">الوحدة</th>
                    <th width="10%">الكمية</th>
                    <th width="10%">السعر</th>
                    <th width="8%">ضريبة</th>
                    ${hasWht ? '<th width="8%">أ.ت.ص</th>' : ''}
                    <th width="12%">الإجمالي</th>
                </tr>
            </thead>
            <tbody>${internalRows}</tbody>
          </table>

          <div class="footer-box" style="margin-top:auto;">
             <div class="sign-area"><div class="sign-line"></div><div class="sign-label">أمين المخزن</div></div>
             <div class="sign-area"><div class="sign-line"></div><div class="sign-label">توكيد الجودة (QA)</div></div>
             <div class="sign-area"><div class="sign-line"></div><div class="sign-label">مدير التشغيل</div></div>
          </div>

          <div class="user-footer">
             <span>${icons.user} ${userName}</span>
             <span>${icons.clock} ${timestamp}</span>
          </div>
        </div>

      </div>
    </body>
    </html>
  `;

  const blob = Utilities.newBlob(htmlContent, MimeType.HTML, `${transactionId}_Note.html`).getAs(MimeType.PDF);
  blob.setName(`${transactionId}_${form.supplier}.pdf`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return { url: file.getUrl(), html: htmlContent };
}
