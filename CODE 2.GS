/**
 * =============================================================================
 * RETURN STOCK MODULE - SECURE BACKEND
 * =============================================================================
 */

const SPREADSHEET_ID = '19gyEZlBxFiTBKpAt3wjwC1lNkglG9Oaovcz2d4ZnGcg'; 
const SHEET_NAME = 'صرف و ارتجاع';
const FOLDER_ID = '1lyfSYYZAIQTW6i8q59DXqoLjQ5e563Oc';

/**
 * SECURITY CORE: Determine User Role based on Email
 * You should customize the list of emails or connect this to your Permissions Sheet.
 */
function getReturnModulePermission(userEmail) {
  if (!userEmail) {
    return 'viewer'; // No email = viewer
  }
  
  // Use the global permission system
  const userPerms = getUserPermissions(userEmail);
  
  if (!userPerms.success || !userPerms.permissions) {
    return 'viewer'; // Safe default
  }
  
  // Find the orders.return permission
  const returnPerm = userPerms.permissions.find(p => p.tabId === 'orders.return');
  
  if (!returnPerm || !returnPerm.permission) {
    return 'viewer'; // No permission = viewer
  }
  
  // Map permission values to return module roles
  const permValue = returnPerm.permission.toLowerCase();
  
  if (permValue === 'admin') return 'admin';
  if (permValue === 'editor') return 'editor';
  if (permValue === 'viewer') return 'viewer';
  
  // Default to viewer for any other value
  return 'viewer';
}

/**
 * Fetches Data AND Permission Level in one secure package.
 */
/**
 * Fetches Data, Permission Level, AND Order Details for the Return Module.
 */
function getReturnStockRecords(userEmail, startDate, endDate) {
  try {
    const role = getReturnModulePermission(userEmail); // Secure Role
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet '${SHEET_NAME}' not found`);

    // --- 1. FETCH RETURN DATA ---
    const data = sheet.getDataRange().getValues();
    const headers = data[9]; // Row 10 is headers
    const returnRecords = [];
    
    // Cairo Time Setup
    const now = new Date();
    const cairoOffset = 2 * 60; 
    const cairoTime = now.getTime() + (now.getTimezoneOffset() * 60000) + (cairoOffset * 60000);
    const todayStr = Utilities.formatDate(new Date(cairoTime), 'Africa/Cairo', 'yyyy-MM-dd');

    // --- 2. FETCH ORDERS DATA (For linking Clients/Summaries) ---
    // We try to find the 'Orders' sheet. Adjust name if yours is different (e.g. 'الطلبات')
    let ordersSheet = ss.getSheetByName('Orders'); 
    if (!ordersSheet) ordersSheet = ss.getSheetByName('الطلبات'); // Fallback
    
    let ordersMap = {}; // Map<DismissNumber, Array<OrderDetails>>
    
    if (ordersSheet) {
      const ordersData = ordersSheet.getDataRange().getValues();
      // Assume Headers on Row 1, Data starts Row 2
      // Col I (Client) = Index 8
      // Col S (Summary) = Index 18
      // Col Y (Dismiss #) = Index 24
      // Col AP (Type) = Index 41
      
      for (let k = 1; k < ordersData.length; k++) {
        const oRow = ordersData[k];
        const dismissRef = String(oRow[24] || '').trim(); // Col Y
        
        if (dismissRef) {
          if (!ordersMap[dismissRef]) ordersMap[dismissRef] = [];
          
          ordersMap[dismissRef].push({
            client: oRow[8],   // Col I
            summary: oRow[18], // Col S
            type: oRow[41]     // Col AP
          });
        }
      }
    }

    // --- 3. PROCESS RECORDS ---
    for (let i = 12; i < data.length; i++) {
      const row = data[i];
      const dismissNumber = String(row[0]);
      const pdfUrl = row[7];
      
      if (!dismissNumber || !pdfUrl) continue;
      
      const statusVal = row[9];
      const summaryText = row[10];  
      const returnDateVal = row[12];
      
      // Filter Logic (Same as before)
      const isSaved = (statusVal !== "" && statusVal !== null);
      let isWorkedOnToday = false;
      if (returnDateVal instanceof Date) {
        const rDate = Utilities.formatDate(returnDateVal, 'Africa/Cairo', 'yyyy-MM-dd');
        if (rDate === todayStr) isWorkedOnToday = true;
      }
      
      if (!startDate && !endDate) {
        if (statusVal === true && !isWorkedOnToday) continue;
      } else {
        const start = startDate || '2020-01-01';
        const end = endDate || todayStr;
        const loadingDateStr = (row[4] instanceof Date) ? Utilities.formatDate(row[4], 'Africa/Cairo', 'yyyy-MM-dd') : '';
        const returnDateStr = (row[12] instanceof Date) ? Utilities.formatDate(row[12], 'Africa/Cairo', 'yyyy-MM-dd') : '';
        const loadingInRange = loadingDateStr && loadingDateStr >= start && loadingDateStr <= end;
        const returnInRange = returnDateStr && returnDateStr >= start && returnDateStr <= end;
        if (!loadingInRange && !returnInRange) continue;
      }
      
      // Delivery Info
      const deliveryInfo = {
        truck: String(row[2] || '').split('-')[0].trim() || '-',
        driver: String(row[2] || '').split('-')[1]?.trim() || '',
      };

      // Products - UPDATED FOR MATRIX
      const products = [];
      for (let j = 18; j < row.length; j += 7) {
        if (j + 6 >= row.length) break;
        
        // Fetch all requested offsets
        const valSupply = Number(row[j + 0]) || 0;     // Supply Qty
        const valClientRet = Number(row[j + 1]) || 0;  // Return Qty (Client?)
        const valLoaded = Number(row[j + 2]) || 0;     // Loaded Qty
        const valExtras = Number(row[j + 4]) || 0;     // Extra Qty
        const valExpRet = Number(row[j + 5]) || 0;     // Expected Return
        const valActRet = row[j + 6];                  // Returned Qty (Actual)

        // Visibility Logic: If any activity exists
        if (valSupply > 0 || (valActRet !== "" && valActRet !== null) || valExpRet > 0) {
            
            let actualToSend = null;
            if (isSaved) {
                actualToSend = (valActRet !== "" && valActRet !== null) ? Number(valActRet) : 0;
            }

            products.push({
                name: headers[j],
                blockIndex: j,
                // Full Matrix Data
                matrix: {
                  supply: valSupply,
                  clientRet: valClientRet,
                  loaded: valLoaded,
                  extras: valExtras,
                  expectedRet: valExpRet,
                  actualRet: actualToSend
                },
                // Legacy support for frontend card
                expectedQty: valExpRet, // Mapping expected return to main expected
                actualQty: actualToSend,
                details: { client: valClientRet, extras: valExtras }
            });
        }
      }
      
      if (products.length > 0) {
          returnRecords.push({
            dismissNumber: dismissNumber,
            loadingDate: (row[4] instanceof Date) ? Utilities.formatDate(row[4], 'Africa/Cairo', 'yyyy-MM-dd') : "",
            returnDate: (row[12] instanceof Date) ? Utilities.formatDate(row[12], 'Africa/Cairo', 'yyyy-MM-dd') : "",
            status: statusVal,
            summary: summaryText,
            pdfUrl: pdfUrl,
            returnNumber: row[8] || '', 
            returnNoteUrl: row[11] || '',
            orderIds: row[3] || '',
            delivery: deliveryInfo,
            products: products,
            // Attach Linked Orders
            linkedOrders: ordersMap[dismissNumber] || [] 
          });
      }
    }
    
    returnRecords.sort((a, b) => b.loadingDate < a.loadingDate ? -1 : 1);
    return JSON.stringify({ success: true, role: role, records: returnRecords });
  } catch (error) {
    return JSON.stringify({ success: false, message: error.toString(), records: [] });
  }
}

/**
 * Saves Transaction - SECURE CHECK
 */
function saveReturnTransaction(recordData) {
  try {
    // 1. RE-VERIFY PERMISSION ON SERVER (Prevents hacking)
    const role = getReturnModulePermission(recordData.userEmail);
    if (role === 'viewer') {
        throw new Error("Permission Denied: You are a Viewer.");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    for (let i = 12; i < data.length; i++) {
      if (data[i][0].toString() === recordData.dismissNumber.toString()) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Dismiss record not found.");
    
    // Check if trying to edit a completed record without admin rights
    const currentStatus = data[rowIndex-1][9]; // Col J is index 9
    if (currentStatus === true && role !== 'admin') {
         throw new Error("Permission Denied: Only Admins can edit completed records.");
    }

    // 2. Perform Save
    recordData.products.forEach(p => {
      sheet.getRange(rowIndex, p.blockIndex + 7).setValue(p.receivedQty);
    });

    sheet.getRange(rowIndex, 11).setValue(recordData.summary);
    sheet.getRange(rowIndex, 13).setValue(new Date(recordData.returnDate));

    const statusCell = sheet.getRange(rowIndex, 10);
    const timestamp = Utilities.formatDate(new Date(), "Africa/Cairo", "yyyy-MM-dd HH:mm:ss");
    statusCell.setNote(`Saved by: ${recordData.userEmail}\nAt: ${timestamp}`);

    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function saveReturnNoteImage(dismissNumber, base64Data, userEmail) {
  // FIX: Pass userEmail to the permission checker
  if (getReturnModulePermission(userEmail) === 'viewer') {
      return { success: false, message: "Denied: Viewer permission only" };
  }

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', `Return_${dismissNumber}.png`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    // Loop to find the row by Dismiss Number
    for (let i = 12; i < data.length; i++) {
      if (data[i][0].toString() === dismissNumber.toString()) {
        // Save URL to Column L (Index 12)
        sheet.getRange(i + 1, 12).setValue(file.getUrl());
        return { success: true, url: file.getUrl() };
      }
    }
    return { success: false, message: "Row not found" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

//_____________________________________________________________________________________________
//=================================== [ ALL ORDERS SUBTAB ] =======================================
//---------------------------------------------------------------------------------------------
/**
 * BACKEND: Get All Orders Report
 * UPDATED: Includes 'confirmationState' for new UI icons
 */
function getAllOrdersReport(ignoredCriteria) {
  // Note: We ignore criteria here because we fetch EVERYTHING now.
  const dbId = typeof SALES_SPREADSHEET_ID !== 'undefined' ? SALES_SPREADSHEET_ID : '19gyEZlBxFiTBKpAt3wjwC1lNkglG9Oaovcz2d4ZnGcg'; 
  const ss = SpreadsheetApp.openById(dbId);
  const allSheets = ss.getSheets();
  const getSheetData = (name) => {
    const s = allSheets.find(sheet => sheet.getName() === name);
    return s ? s.getDataRange().getValues() : [];
  };

  const ordersData = getSheetData('Orders');
  const dismissData = getSheetData('صرف و ارتجاع');
  const invoicesData = getSheetData('Invoices');
  const clientsData = getSheetData('Clients');

  if (ordersData.length === 0) return []; // Return empty array if no data

  // 1. CLIENT MAP
  const clientMap = new Map();
  for (let i = 1; i < clientsData.length; i++) {
    const row = clientsData[i];
    const lookupName = String(row[3]).trim(); 
    if (lookupName) {
      clientMap.set(lookupName, {
        mainId: row[1], fullId: row[2], mainName: row[4], branch: row[5],
        category: row[6], paymentPlan: row[7], strategy: row[8],
        officialName: row[20], vatNo: row[22], phone: row[13],
        workHours: row[14], contactPerson: row[15], region: row[16],
        address: row[17], distance: row[18], locationUrl: row[36]
      });
    }
  }

  // 2. INVOICE MAP
  const invoiceMap = new Map();
  for (let i = 1; i < invoicesData.length; i++) {
    const row = invoicesData[i];
    // NEW MAPPING:
    // Order ID: Column T (Index 19)
    // Doc ID/Serial: Column D (Index 3)
    // Version: Column C (Index 2)
    // Total: Column H (Index 7)
    // Discount: Column J (Index 9) - Assuming Contract Discount
    // Total Before: Column N (Index 13)
    // File URL: Column U (Index 20)
    
    const orderId = String(row[19]).trim(); // Was 21
    const docId = String(row[3]).trim();    // Was 4
    
    if (orderId && docId) {
      const uniqueKey = orderId + "_" + docId;
      const ver = Number(row[2]) || 0; // Was 3
      
      if (!invoiceMap.has(uniqueKey) || ver > invoiceMap.get(uniqueKey).version) {
        invoiceMap.set(uniqueKey, {
          version: ver, 
          docId: docId, 
          totalAfter: Number(row[7])||0,   // Was 9
          discount: Number(row[9])||0,     // Was 11
          totalBefore: Number(row[13])||0, // Was 15
          fileUrl: row[20]                 // Was 22
        });
      }
    }
  }

  // 3. DISMISS MAP
  const dismissMap = new Map();
  for (let i = 1; i < dismissData.length; i++) {
    const row = dismissData[i];
    const key = String(row[0]).trim(); 
    if (key) dismissMap.set(key, { fileDismiss: row[7], status: row[9], fileReturn: row[11] });
  }

  // Column Constants
  const COL_STATUS = 2; 
  const COL_ID = 3; 
  const COL_CONFIRM = 4; // <--- NEW: Confirmation State (Column E)
  const COL_DATE = 5; 
  const COL_LOADING_DATE = 7; const COL_CLIENT = 8; const COL_VAT_TYPE = 9;
  const COL_PO_TAWOOS = 10; const COL_PO_PL = 11; const COL_FILE_1 = 12;
  const COL_COMM_INTERNAL = 16; const COL_COMM_CLIENT = 17; const COL_PROD_SUMMARY = 18;
  const COL_LOAD_ORDER = 19; const COL_TRUCK = 20; const COL_DRIVER = 21; const COL_DEL_REP = 22; const COL_SALES_REP = 23; 
  const COL_DISMISS_NUM = 24; const COL_RETURN_NUM = 25; const COL_WEIGHT = 26;
  const COL_DOC_DELIVERY = 28; const COL_DOC_TAWOOS = 29; const COL_DOC_PL = 30;
  const COL_EINV_TAWOOS = 34; const COL_EINV_PL = 35;
  const COL_AMT_BEFORE = 37; const COL_AMT_VAT = 38; const COL_AMT_WHT = 39; const COL_AMT_TOTAL = 40;
  const COL_TYPE = 41; const COL_PROD_START = 42;

  const results = [];
  const headerRow = ordersData[6];

  // Loop Rows
  for (let i = 6; i < ordersData.length; i++) {
    const row = ordersData[i];
    const orderId = String(row[COL_ID]);
    if (!orderId) continue; 
    
    // Dates
    const rowDate = row[COL_DATE] instanceof Date ? row[COL_DATE] : new Date(row[COL_DATE]);
    const isValidDate = !isNaN(rowDate.getTime());
    const dateIso = isValidDate ? Utilities.formatDate(rowDate, "GMT+2", "yyyy-MM-dd") : "";

    // Data Helpers
    const clientName = row[COL_CLIENT] || "";
    const po1 = row[COL_PO_TAWOOS] ? String(row[COL_PO_TAWOOS]) : "";
    const po2 = row[COL_PO_PL] ? String(row[COL_PO_PL]) : "";
    const dismissNum = row[COL_DISMISS_NUM] ? String(row[COL_DISMISS_NUM]).trim() : "";
    
    // Pre-build Search String
    const searchStr = (orderId + " " + clientName + " " + po1 + " " + po2 + " " + dismissNum).toLowerCase();

    // Standard Object Building
    const loadingDateRaw = row[COL_LOADING_DATE];
    const loadingDateFull = (loadingDateRaw instanceof Date) ? Utilities.formatDate(loadingDateRaw, "GMT+2", "dd-MMM-yyyy") : (loadingDateRaw || "N/A");
    const loadingDateIso = (loadingDateRaw instanceof Date) ? Utilities.formatDate(loadingDateRaw, "GMT+2", "yyyy-MM-dd") : "";
    const clientData = clientMap.get(clientName) || { strategy: "" };
    const pricingStrategy = clientData.strategy || ""; 
    const calc = { beforeVat: Number(row[COL_AMT_BEFORE])||0, vatVal: Number(row[COL_AMT_VAT])||0, whtVal: Number(row[COL_AMT_WHT])||0, total: Number(row[COL_AMT_TOTAL])||0 };
    
    let requiredCount = 0; let foundCount = 0;
    if (row[COL_DOC_TAWOOS]) { requiredCount++; if (row[COL_EINV_TAWOOS]) foundCount++; }
    if (row[COL_DOC_PL]) { requiredCount++; if (row[COL_EINV_PL]) foundCount++; }

    const docs = [];
    const docIds = [ { type: 'Delivery Note', id: row[COL_DOC_DELIVERY] }, { type: 'Tawoos Inv', id: row[COL_DOC_TAWOOS] }, { type: 'Private Label', id: row[COL_DOC_PL] } ];
    docIds.forEach(d => { if (d.id) docs.push({ type: d.type, id: d.id, details: invoiceMap.get(orderId + "_" + String(d.id).trim()) || null }); });

    const financeObj = { vatType: row[COL_VAT_TYPE], calc: calc, documents: docs, eInvoices: { tawoos: row[COL_EINV_TAWOOS], pl: row[COL_EINV_PL] }, eInvStatus: { required: requiredCount, found: foundCount, isComplete: foundCount >= requiredCount } };
    const comments = { internal: row[COL_COMM_INTERNAL], client: row[COL_COMM_CLIENT] };
    const prodSummary = row[COL_PROD_SUMMARY];
    
    const files = []; if (row[COL_FILE_1]) files.push(row[COL_FILE_1]); if (row[13]) files.push(row[13]); if (row[14]) files.push(row[14]);
    
    let dismissInfo = { status: "", fileDismiss: "", fileReturn: "" };
    if (dismissNum && dismissMap.has(dismissNum)) dismissInfo = dismissMap.get(dismissNum);

    const deliveryObj = {
      loadingDate: loadingDateRaw instanceof Date ? Utilities.formatDate(loadingDateRaw, "GMT+2", "dd/MM") : loadingDateRaw,
      loadingDateIso: loadingDateIso,
      loadOrder: row[COL_LOAD_ORDER], truck: row[COL_TRUCK], driver: row[COL_DRIVER], delRep: row[COL_DEL_REP], salesRep: row[COL_SALES_REP],
      dismissNum: dismissNum, returnNum: row[COL_RETURN_NUM], weight: row[COL_WEIGHT], dismissStatus: dismissInfo.status, fileDismiss: dismissInfo.fileDismiss, fileReturn: dismissInfo.fileReturn
    };

    let itemsCount = 0; const products = [];
    for (let j = COL_PROD_START; j < row.length; j += 5) {
      if (j + 1 >= row.length) break;
      const qty = Number(row[j]) || 0;
      if (qty > 0 || Number(row[j+1]) > 0) { products.push({ name: (headerRow && headerRow[j]) ? headerRow[j] : `Item ${j}`, qty: qty, bonus: Number(row[j+1])||0 }); itemsCount += qty; }
    }
    
    results.push({
      orderId: orderId,
      dateIso: dateIso, 
      searchStr: searchStr, 
      dateFormatted: isValidDate ? Utilities.formatDate(rowDate, "GMT+2", "dd-MMM-yyyy") : "N/A",
      loadingDateFull: loadingDateFull,
      clientName: clientName, pricingStrategy: pricingStrategy, clientData: clientData,
      poData: { text: [po1, po2].filter(Boolean).join(' / '), hasFiles: files.length > 0, files: files },
      deliveryData: deliveryObj, financeData: financeObj, comments: comments, prodSummary: prodSummary,
      
      // --- UPDATED FIELDS ---
      type: row[COL_TYPE] || "", 
      status: row[COL_STATUS], 
      confirmationState: row[COL_CONFIRM] || 'Not Confirmed', // <--- Added Here
      
      itemsCount: itemsCount, products: products
    });
  }
  return results.reverse();
}

/**
 * GENERATE EXCEL REPORT (v8.0 - RTL & Auto-Fit)
 * - Adds Right-to-Left direction for Arabic text fields
 * - Auto-fits row heights for wrapped text
 */
function generateExcelReport(records, columnConfig) {
  try {
    if (!records || records.length === 0) throw new Error("No data to export.");
    if (!columnConfig || columnConfig.length === 0) throw new Error("No columns selected.");

    // 1. SETUP & PALETTE
    const timestamp = Utilities.formatDate(new Date(), "GMT+2", "yyyy-MM-dd_HHmm");
    const filename = `Orders_Report_${timestamp}`;
    const ss = SpreadsheetApp.create(filename);
    const sheet = ss.getActiveSheet();
    const ssId = ss.getId();

    const THEMES = {
      "ORDER BASICS": { dark: "#20124d", light: "#e6e1ff" },
      "CLIENT INFO":  { dark: "#003366", light: "#cfe2ff" },
      "FINANCIALS":   { dark: "#0f5132", light: "#d1e7dd" },
      "LOGISTICS":    { dark: "#664d03", light: "#fff3cd" },
      "METADATA":     { dark: "#495057", light: "#f8f9fa" }
    };

    // 2. EXPAND CONFIG
    const finalConfig = [];
    columnConfig.forEach(col => {
      if (col.path === 'poData.files') {
        finalConfig.push({ ...col, label: 'PO File 1', subIndex: 0 });
        finalConfig.push({ ...col, label: 'PO File 2', subIndex: 1 });
        finalConfig.push({ ...col, label: 'PO File 3', subIndex: 2 });
      } else {
        finalConfig.push(col);
      }
    });

    // 3. BUILD HEADERS & STYLES
    const row1 = [];
    const row2 = [];
    const mergeRanges = [];
    const bgColors = [[], []];
    const fontColors = [[], []];
    const fontWeights = [[], []];
    
    let lastGroup = "";
    let mergeStart = 1;

    finalConfig.forEach((col, index) => {
      row2.push(col.label);
      
      const theme = THEMES[col.group] || { dark: "#000000", light: "#cccccc" };
      
      // Row 1
      bgColors[0].push(theme.dark);
      fontColors[0].push("#ffffff");
      fontWeights[0].push("bold");

      // Row 2
      bgColors[1].push(theme.light);
      fontColors[1].push(theme.dark);
      fontWeights[1].push("bold");

      if (col.group !== lastGroup) {
        if (index > 0 && (index + 1) - mergeStart > 1) {
             mergeRanges.push({ r: 1, c: mergeStart, num: (index + 1) - mergeStart });
        }
        row1.push(col.group);
        lastGroup = col.group;
        mergeStart = index + 1;
      } else {
        row1.push(""); 
      }
    });
    if ((finalConfig.length + 1) - mergeStart > 1) {
        mergeRanges.push({ r: 1, c: mergeStart, num: (finalConfig.length + 1) - mergeStart });
    }

    // 4. TRANSFORM DATA
    const rows = records.map(rec => {
      return finalConfig.map(col => {
        // Doc Links
        if (col.path === 'doc' || col.path === 'docLink') {
           const docs = rec.financeData?.documents || [];
           const found = docs.find(d => d.type === col.sub);
           if (!found) return '';
           if (col.path === 'doc') return found.id;
           if (col.path === 'docLink') return (found.details && found.details.fileUrl) ? found.details.fileUrl : '';
        }

        // PO Files
        if (col.path === 'poData.files') {
           const files = rec.poData?.files || [];
           return files[col.subIndex] || '';
        }

        let val = col.path.split('.').reduce((o, k) => (o || {})[k], rec);
        if (val === null || val === undefined) return '';

        // Date Objects
        if (col.path === 'dateIso' || col.path.includes('loadingDate') || col.path.includes('DateIso')) {
             if (val && String(val).length > 5) {
                const d = new Date(val);
                if (!isNaN(d.getTime())) return d; 
             }
        }
        return String(val);
      });
    });

    // 5. WRITE & FORMAT
    if (rows.length > 0) {
      const totalCols = finalConfig.length;
      
      sheet.getRange(1, 1, 1, row1.length).setValues([row1]);
      sheet.getRange(2, 1, 1, row2.length).setValues([row2]);
      sheet.getRange(3, 1, rows.length, rows[0].length).setValues(rows);

      // Merges
      mergeRanges.forEach(m => { try { sheet.getRange(m.r, m.c, 1, m.num).merge(); } catch(e){} });

      // Headers Style
      const headerRange = sheet.getRange(1, 1, 2, totalCols);
      headerRange.setBackgrounds(bgColors)
                 .setFontColors(fontColors)
                 .setFontWeights(fontWeights)
                 .setHorizontalAlignment("center");
      
      sheet.getRange(2, 1, 1, totalCols).setBorder(true, true, true, true, true, true);

      // Date Format
      finalConfig.forEach((col, i) => {
          if (col.path === 'dateIso' || col.path.includes('loadingDate')) {
              sheet.getRange(3, i + 1, rows.length, 1).setNumberFormat("dd-mmm-yyyy");
          }
      });
      
      SpreadsheetApp.flush();

      // --- SIZING, WRAPPING & RTL LOGIC ---
      sheet.autoResizeColumns(1, totalCols); // Base resize

      finalConfig.forEach((col, i) => {
          const colIndex = i + 1;
          const lowerLabel = col.label.toLowerCase();
          const dataRange = sheet.getRange(3, colIndex, rows.length, 1);
          
          // 1. LINKS (Clip)
          if (col.path.includes('Link') || lowerLabel.includes('link') || col.path === 'poData.files') {
              sheet.setColumnWidth(colIndex, 150);
              dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
          }

          // 2. SUMMARY & COMMENTS (Wrap + RTL)
          if (col.path === 'prodSummary' || col.path === 'comments.internal' || col.path === 'comments.client') {
               // Set wider columns for text
               if (col.path === 'prodSummary') sheet.setColumnWidth(colIndex, 400); 
               else sheet.setColumnWidth(colIndex, 300);

               // Apply Wrap, Alignment & RTL
               dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
               dataRange.setVerticalAlignment("top");
               dataRange.setTextDirection(SpreadsheetApp.TextDirection.RIGHT_TO_LEFT);
          }
      });

      // 3. AUTO-FIT ROWS (Crucial Step: Must run AFTER WrapStrategy is set)
      sheet.autoResizeRows(3, rows.length);
      
      sheet.setFrozenRows(2);
    }
    
    // 6. SAVE & EXPORT
    SpreadsheetApp.flush();
    Utilities.sleep(1000); 

    const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx`;
    const params = {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, params);
    
    DriveApp.getFileById(ssId).setTrashed(true);

    return { 
      success: true, 
      base64: Utilities.base64Encode(response.getBlob().getBytes()),
      filename: `${filename}.xlsx`
    };

  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getClientsList() {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID); 
    const sheet = ss.getSheetByName('Clients');
    
    if (!sheet) return { success: false, message: 'Clients sheet not found' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return { success: true, clients: [] };
    
    // Fetch Data Range: Row 4 to Last Row, Columns A to I (1 to 9)
    // We need Col D (Index 3 in array) and Col I (Index 8 in array)
    const data = sheet.getRange(4, 1, lastRow - 3, 9).getValues();
    
    const clients = data.map(row => {
      return {
        name: row[3].toString().trim(),     // Column D (Name)
        strategy: row[8].toString().trim()  // Column I (Strategy)
      };
    }).filter(c => c.name !== ''); 
    
    return { success: true, clients: clients };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getProductsList() {
  try {
    // Open Sales spreadsheet
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const productsSheet = salesSpreadsheet.getSheetByName('Products');
    
    if (!productsSheet) {
      return { success: false, message: 'Products sheet not found' };
    }
    
    const lastRow = productsSheet.getLastRow();
    if (lastRow < 3) {
      return { success: true, products: [] };
    }
    
    // Get data from A3:E (last row) - Columns A and E
    const productsRange = productsSheet.getRange(3, 1, lastRow - 2, 5);
    const productsData = productsRange.getValues();
    
    // Build array with both column A (for dropdown) and column E (for summary)
    const products = productsData
      .filter(row => row[0] && row[0].toString().trim() !== '')
      .map(row => ({
        fullName: row[0].toString().trim(), // Column A
        shortName: row[4] ? row[4].toString().trim() : row[0].toString().trim() // Column E, fallback to A
      }));
    
    return { success: true, products: products };
    
  } catch (error) {
    Logger.log('Get products error: ' + error.toString());
    return { success: false, message: 'System error: ' + error.toString() };
  }
}

function getSalesRepList() {
  try {
    // Open Sales spreadsheet
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const indexSheet = salesSpreadsheet.getSheetByName('Index');
    
    if (!indexSheet) {
      return { success: false, message: 'Index sheet not found' };
    }
    
    // Get data from J3:J30
    const salesRepRange = indexSheet.getRange(3, 10, 28, 1); // J3:J30 (10 = column J, 28 rows)
    const salesRepData = salesRepRange.getValues();
    
    // Flatten array and filter out empty values
    const salesReps = salesRepData
      .map(row => row[0])
      .filter(rep => rep && rep.toString().trim() !== '');
    
    return { success: true, salesReps: salesReps };
    
  } catch (error) {
    Logger.log('Get sales rep error: ' + error.toString());
    return { success: false, message: 'System error: ' + error.toString() };
  }
}

function getProductPrice(productName, clientName, invoiceDate, vatType) {
  try {
    Logger.log('=== PRICE DETECTION START ===');
    Logger.log('Product: ' + productName);
    Logger.log('Client: ' + clientName);
    Logger.log('Invoice Date (input): ' + invoiceDate);
    Logger.log('VAT Type: ' + vatType);
    
    // Open Sales spreadsheet
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const clientsSheet = salesSpreadsheet.getSheetByName('Clients');
    const pricesSheet = salesSpreadsheet.getSheetByName('Prices');
    const ordersSheet = salesSpreadsheet.getSheetByName('Orders');
    const productsSheet = salesSpreadsheet.getSheetByName('Products');
    
    if (!clientsSheet || !pricesSheet) {
      Logger.log('ERROR: Required sheets not found');
      return { success: false, message: 'Required sheets not found' };
    }
    
    // Convert invoice date to Date object
    const invoiceDateObj = new Date(invoiceDate);
    invoiceDateObj.setHours(0, 0, 0, 0);
    Logger.log('Invoice Date (parsed): ' + invoiceDateObj.toString());
    
    // Step 1: Get client data (strategy, client ID) from Clients sheet
    const clientData = getClientData(clientsSheet, clientName);
    if (!clientData) {
      Logger.log('ERROR: Client not found');
      return { success: false, message: 'Client not found' };
    }
    Logger.log('Client Data: ' + JSON.stringify(clientData));
    
    // Step 2: Get product column index from Prices sheet Row 3 (starting column K)
    const productColumnIndex = getProductColumnIndex(pricesSheet, productName);
    if (!productColumnIndex) {
      Logger.log('ERROR: Product not found in prices sheet');
      return { success: false, message: 'Product not found in prices sheet' };
    }
    Logger.log('Product Column Index: ' + productColumnIndex);
    
    // Step 3: Get all price rows from Prices sheet
    const priceRows = getAllPriceRows(pricesSheet, productColumnIndex);
    Logger.log('Total price rows found: ' + priceRows.length);
    
    // Step 4: Apply exact priority sequence to find price
    const detectedPrice = detectPriceByPriority(priceRows, clientData, invoiceDateObj);
    Logger.log('Detected Price: ' + detectedPrice);
    
    if (detectedPrice === null || detectedPrice === '') {
      Logger.log('ERROR: No valid price found');
      return { success: false, message: 'No valid price found' };
    }
    
    // Step 5: Get product VAT rate
    const productVatRate = getProductVatRate(ordersSheet, productsSheet, productName);
    Logger.log('Product VAT Rate: ' + productVatRate);
    
    // Step 6: Apply VAT adjustment if needed
    let finalPrice = detectedPrice;
    
    // Check if VAT type is inclusive
    const isInclusiveVat = vatType === 'سعر شامل الضريبة' || vatType === 'سعر شامل و خصم 1%';
    Logger.log('Is Inclusive VAT: ' + isInclusiveVat);
    
    // Check if product VAT rate is 14%
    const is14PercentVat = productVatRate === 0.14 || productVatRate === 14 || productVatRate === '14%' || productVatRate === '0.14';
    Logger.log('Is 14% VAT: ' + is14PercentVat);
    
    if (isInclusiveVat && is14PercentVat) {
      finalPrice = detectedPrice / 1.14;
      Logger.log('Applied VAT adjustment: ' + detectedPrice + ' / 1.14 = ' + finalPrice);
    }
    
    Logger.log('Final Price: ' + finalPrice);
    Logger.log('=== PRICE DETECTION END ===');
    
    return { success: true, price: finalPrice };
    
  } catch (error) {
    Logger.log('EXCEPTION: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    return { success: false, message: 'System error: ' + error.toString() };
  }
}

// Helper: Get client data from Clients sheet
function getClientData(clientsSheet, clientName) {
  const lastRow = clientsSheet.getLastRow();
  if (lastRow < 4) return null;
  
  // Get data from row 4 onwards
  const clientNames = clientsSheet.getRange(4, 4, lastRow - 3, 1).getValues(); // Column D
  const pricingStrategies = clientsSheet.getRange(4, 9, lastRow - 3, 1).getValues(); // Column I
  const clientIds = clientsSheet.getRange(4, 27, lastRow - 3, 1).getValues(); // Column AA
  
  // Find client row
  for (let i = 0; i < clientNames.length; i++) {
    if (clientNames[i][0] && clientNames[i][0].toString().trim().toLowerCase() === clientName.toLowerCase()) {
      return {
        name: clientNames[i][0].toString().trim(),
        pricingStrategy: pricingStrategies[i][0] ? pricingStrategies[i][0].toString().trim() : '',
        clientId: clientIds[i][0] ? clientIds[i][0].toString().trim() : ''
      };
    }
  }
  
  return null;
}

// Helper: Get product column index from Prices sheet Row 3
function getProductColumnIndex(pricesSheet, productName) {
  const row3Data = pricesSheet.getRange(3, 11, 1, 200).getValues()[0]; // Column K onwards (11 = K)
  
  for (let i = 0; i < row3Data.length; i++) {
    if (row3Data[i] && row3Data[i].toString().trim() === productName) {
      return 11 + i; // Return 1-indexed column number
    }
  }
  
  return null;
}

// Helper: Get all price rows from Prices sheet
function getAllPriceRows(pricesSheet, productColumnIndex) {
  // Find last row with data in column A (not formula rows)
  const columnAData = pricesSheet.getRange(4, 1, 1000, 1).getValues(); // Column A
  let lastDataRow = 3; // Default to row 3 (before data starts)
  
  for (let i = 0; i < columnAData.length; i++) {
    if (columnAData[i][0] && columnAData[i][0].toString().trim() !== '') {
      lastDataRow = 4 + i;
    }
  }
  
  const rowCount = lastDataRow - 3;
  Logger.log('Last row with data in Column A: ' + lastDataRow);
  Logger.log('Reading ' + rowCount + ' rows from Prices sheet');
  
  if (rowCount <= 0) return [];
  
  // Get all columns we need (A, B, D, F, G, H, J, and product price column)
  const strategies = pricesSheet.getRange(4, 1, rowCount, 1).getValues(); // Column A
  const clientNames = pricesSheet.getRange(4, 2, rowCount, 1).getValues(); // Column B
  const clientIds = pricesSheet.getRange(4, 4, rowCount, 1).getValues(); // Column D
  const allBranches = pricesSheet.getRange(4, 6, rowCount, 1).getValues(); // Column F
  const priceTypes = pricesSheet.getRange(4, 7, rowCount, 1).getValues(); // Column G
  const startDates = pricesSheet.getRange(4, 8, rowCount, 1).getValues(); // Column H
  const endDates = pricesSheet.getRange(4, 10, rowCount, 1).getValues(); // Column J
  const prices = pricesSheet.getRange(4, productColumnIndex, rowCount, 1).getValues(); // Product column
  
  const rows = [];
  
  for (let i = 0; i < strategies.length; i++) {
    // Only include rows where column A (strategy) has data
    if (!strategies[i][0] || strategies[i][0].toString().trim() === '') continue;
    
    rows.push({
      strategy: strategies[i][0].toString().trim(),
      clientName: clientNames[i][0] ? clientNames[i][0].toString().trim() : '',
      clientId: clientIds[i][0] ? clientIds[i][0].toString().trim() : '',
      allBranches: allBranches[i][0] ? allBranches[i][0].toString().trim() : '',
      priceType: priceTypes[i][0] ? priceTypes[i][0].toString().trim() : '',
      startDate: startDates[i][0] instanceof Date ? startDates[i][0] : null,
      endDate: endDates[i][0] instanceof Date ? endDates[i][0] : null,
      price: prices[i][0]
    });
  }
  
  Logger.log('Actual rows with data in column A: ' + rows.length);
  return rows;
}

// Helper: Detect price using exact priority sequence
function detectPriceByPriority(priceRows, clientData, invoiceDateObj) {
  Logger.log('--- Starting priority detection ---');
  
  // Convert invoice date to string format YYYY-MM-DD for comparison
  const invoiceDateStr = Utilities.formatDate(invoiceDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  Logger.log('Invoice Date String for comparison: ' + invoiceDateStr);
  
  // Filter rows that match invoice date period
  const validRows = priceRows.filter(row => {
    if (!row.startDate || !row.endDate) {
      return false;
    }
    
    // Convert dates to string format YYYY-MM-DD
    const startStr = Utilities.formatDate(row.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const endStr = Utilities.formatDate(row.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // String comparison: invoiceDate >= startDate AND invoiceDate <= endDate
    const isValid = invoiceDateStr >= startStr && invoiceDateStr <= endStr;
    
    if (isValid) {
      Logger.log('Valid row found: Strategy=' + row.strategy + ', Client=' + row.clientName + ', ClientID=' + row.clientId + ', AllBranches=' + row.allBranches + ', Type=' + row.priceType + ', Start=' + startStr + ', End=' + endStr + ', Price=' + row.price);
    }
    
    return isValid;
  });
  
  Logger.log('Valid rows after date filter: ' + validRows.length);
  
  // Priority 1: Special + Client Name + Branch + Promotional
  Logger.log('Checking Priority 1: Special + Client Name + Branch + Promotional');
  let match = validRows.find(row => 
    row.strategy === 'اسعار خاصة' &&
    row.clientName.toLowerCase() === clientData.name.toLowerCase() &&
    row.allBranches !== 'Yes' &&
    row.priceType === 'عرض' &&
    row.price !== null && row.price !== ''
  );
  if (match) {
    Logger.log('MATCH FOUND at Priority 1: ' + match.price);
    return match.price;
  }
  
  // Priority 2: Special + Client Name + Branch + Official
  Logger.log('Checking Priority 2: Special + Client Name + Branch + Official');
  match = validRows.find(row => 
    row.strategy === 'اسعار خاصة' &&
    row.clientName.toLowerCase() === clientData.name.toLowerCase() &&
    row.allBranches !== 'Yes' &&
    row.priceType === 'رسمي' &&
    row.price !== null && row.price !== ''
  );
  if (match) {
    Logger.log('MATCH FOUND at Priority 2: ' + match.price);
    return match.price;
  }
  
  // Priority 3: Special + Client ID + All Branches + Promotional
  Logger.log('Checking Priority 3: Special + Client ID + All Branches + Promotional');
  match = validRows.find(row => 
    row.strategy === 'اسعار خاصة' &&
    row.clientId === clientData.clientId &&
    row.allBranches === 'Yes' &&
    row.priceType === 'عرض' &&
    row.price !== null && row.price !== ''
  );
  if (match) {
    Logger.log('MATCH FOUND at Priority 3: ' + match.price);
    return match.price;
  }
  
  // Priority 4: Special + Client ID + All Branches + Official
  Logger.log('Checking Priority 4: Special + Client ID + All Branches + Official');
  match = validRows.find(row => 
    row.strategy === 'اسعار خاصة' &&
    row.clientId === clientData.clientId &&
    row.allBranches === 'Yes' &&
    row.priceType === 'رسمي' &&
    row.price !== null && row.price !== ''
  );
  if (match) {
    Logger.log('MATCH FOUND at Priority 4: ' + match.price);
    return match.price;
  }
  
  // VALIDATION: If client has "Special Price" strategy, stop here
  // Special Price clients should ONLY get prices from priorities 1-4
  if (clientData.pricingStrategy === 'اسعار خاصة') {
    Logger.log('Client has Special Price strategy - skipping priorities 5-8');
    Logger.log('NO MATCH FOUND - Special Price clients require explicit client-specific pricing');
    return null;
  }
  
  // Priority 5: Strategy + Same client name + Branch + Promotional
  Logger.log('Checking Priority 5: Strategy + Branch + Promotional (no client check)');
  match = validRows.find(row => 
    row.strategy === clientData.pricingStrategy &&
    row.allBranches !== 'Yes' &&
    row.priceType === 'عرض' &&
    row.price !== null && row.price !== ''
  );
  if (match) {
    Logger.log('MATCH FOUND at Priority 5: ' + match.price);
    return match.price;
  }
  
  // Priority 6: Strategy + Same client name + Branch + Official
  Logger.log('Checking Priority 6: Strategy + Branch + Official (no client check)');
  match = validRows.find(row => 
    row.strategy === clientData.pricingStrategy &&
    row.allBranches !== 'Yes' &&
    row.priceType === 'رسمي' &&
    row.price !== null && row.price !== ''
  );
  if (match) {
    Logger.log('MATCH FOUND at Priority 6: ' + match.price);
    return match.price;
  }
  
  // Priority 7: Strategy + Client ID + All Branches + Promotional
  Logger.log('Checking Priority 7: Strategy + All Branches + Promotional (no client check)');
  match = validRows.find(row => 
    row.strategy === clientData.pricingStrategy &&
    row.allBranches === 'Yes' &&
    row.priceType === 'عرض' &&
    row.price !== null && row.price !== ''
  );
  if (match) {
    Logger.log('MATCH FOUND at Priority 7: ' + match.price);
    return match.price;
  }
  
  // Priority 8: Strategy + Client ID + All Branches + Official
  Logger.log('Checking Priority 8: Strategy + All Branches + Official (no client check)');
  match = validRows.find(row => 
    row.strategy === clientData.pricingStrategy &&
    row.allBranches === 'Yes' &&
    row.priceType === 'رسمي' &&
    row.price !== null && row.price !== ''
  );
  if (match) {
    Logger.log('MATCH FOUND at Priority 8: ' + match.price);
    return match.price;
  }
  
  // No match found
  Logger.log('NO MATCH FOUND in any priority level');
  return null;
}

// Helper: Get product VAT rate
function getProductVatRate(ordersSheet, productsSheet, productName) {
  let vatRate = null;
  
  try {
    // Primary: Read from Orders sheet Row 2, starting Column AQ
    if (ordersSheet) {
      const row2Data = ordersSheet.getRange(2, 43, 1, 200).getValues()[0]; // Column AQ onwards (43 = AQ)
      const row3Data = ordersSheet.getRange(3, 43, 1, 200).getValues()[0]; // Product names in Row 3
      
      for (let i = 0; i < row3Data.length; i++) {
        if (row3Data[i] && row3Data[i].toString().trim() === productName) {
          vatRate = row2Data[i];
          break;
        }
      }
    }
    
    // Fallback: Read from Products sheet Column F, starting Row 3
    if ((vatRate === null || vatRate === '') && productsSheet) {
      const lastRow = productsSheet.getLastRow();
      if (lastRow >= 3) {
        const productNames = productsSheet.getRange(3, 1, lastRow - 2, 1).getValues(); // Column A
        const vatRates = productsSheet.getRange(3, 6, lastRow - 2, 1).getValues(); // Column F
        
        for (let i = 0; i < productNames.length; i++) {
          if (productNames[i][0] && productNames[i][0].toString().trim() === productName) {
            vatRate = vatRates[i][0];
            break;
          }
        }
      }
    }
  } catch (error) {
    Logger.log('Error reading VAT rate: ' + error.toString());
  }
  
  return vatRate;
}

/**
 * Saves New Order - SECURE VERSION
 * 1. Checks permissions explicitly server-side.
 * 2. Logs distinction between "True Viewer" and "Fallback Viewer".
 * 3. Uses the passed userEmail for logging (ignoring execution identity).
 */
function saveNewOrder(formData, userEmail) {
  try {
    // --- 1. VALIDATE IDENTITY ---
    if (!userEmail || typeof userEmail !== 'string' || userEmail.trim() === "") {
        Logger.log('[Security Alert] Save attempt rejected: No user email provided.');
        return { success: false, message: "System Error: User identity missing. Please re-login." };
    }

    // --- 2. PERMISSION & ROLE DIAGNOSIS ---
    // We fetch all permissions for this user to determine their specific access to 'orders'
    const permResult = getUserPermissions(userEmail);
    
    let role = 'viewer'; // Default deny/safe state
    let permissionDebugStatus = ''; 

    if (!permResult.success) {
        // Case A: User email does not exist in the 'Users' sheet at all
        role = 'viewer';
        permissionDebugStatus = 'FALLBACK_VIEWER (User Not Found in DB)';
    } else {
        // User exists, look for the specific 'orders' main tab permission
        // Note: We filter for !p.isSubtab to get the main module permission
        const orderPerm = permResult.permissions.find(p => p.tabId === 'orders' && !p.isSubtab);

        if (!orderPerm || !orderPerm.permission || orderPerm.permission.trim() === '') {
             // Case B: User exists, but the 'orders' column is empty
             role = 'viewer';
             permissionDebugStatus = 'FALLBACK_VIEWER (Permission Cell is Empty)';
        } else {
             // Case C: User has an explicit permission set
             role = orderPerm.permission.toLowerCase();
             if (role === 'viewer') {
                 permissionDebugStatus = 'TRUE_VIEWER (Explicitly assigned in DB)';
             } else {
                 permissionDebugStatus = 'AUTHORIZED_' + role.toUpperCase();
             }
        }
    }

    // --- 3. CONSOLE LOGGING (Requested Feature) ---
    Logger.log('------------------------------------------------');
    Logger.log(`[Order Save Attempt] User: ${userEmail}`);
    Logger.log(`[Permission Diagnosis] Status: ${permissionDebugStatus}`);
    Logger.log(`[Effective Role] ${role}`);
    Logger.log('------------------------------------------------');

    // --- 4. ENFORCE SECURITY ---
    if (role === 'viewer') {
       return { success: false, message: "Permission Denied: Viewers cannot create new orders." };
    }

    // =========================================================
    // --- 5. EXISTING LOGIC (Preserved but using validated email) ---
    // =========================================================
    
    Logger.log('--- START SAVE PROCESS ---');
    const salesSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const ordersSheet = salesSpreadsheet.getSheetByName('Orders');
    
    if (!ordersSheet) {
      return { success: false, message: 'Orders sheet not found' };
    }
    
    // Smart Row Finder
    const lastRow = ordersSheet.getLastRow();
    let newRow = lastRow + 1; 
    
    if (lastRow >= 7) {
       const rangeData = ordersSheet.getRange(7, 1, lastRow - 6, 9).getValues();
       for (let i = 0; i < rangeData.length; i++) {
          const rowData = rangeData[i];
          // Check if key columns are empty to find gaps
          const isAEmpty = !rowData[0] || rowData[0].toString().trim() === ''; // User Email Column
          const isDEmpty = !rowData[3] || rowData[3].toString().trim() === ''; // Order ID Column
          
          if (isAEmpty && isDEmpty) {
             newRow = 7 + i;
             Logger.log('Gap found at Row: ' + newRow);
             break;
          }
       }
    }
    
    // Generate Order ID
    let nextOrderId = 1;
    if (lastRow >= 7) {
       const idData = ordersSheet.getRange(7, 4, lastRow - 6, 1).getValues();
       for (let i = 0; i < idData.length; i++) {
          const val = parseInt(idData[i][0]);
          if (!isNaN(val) && val >= nextOrderId) {
             nextOrderId = val + 1;
          }
       }
    }
    Logger.log('Generated Order ID: ' + nextOrderId);

    // Handle File Uploads
    let fileUrls = [];
    let uploadDebugMsg = "No files to upload";
    
    if (formData.files && formData.files.length > 0) {
      Logger.log(`Found ${formData.files.length} files. Uploading...`);
      const uploadResult = uploadFilesToDrive(formData.files, nextOrderId);
      if (uploadResult.success) {
        fileUrls = uploadResult.fileUrls;
        uploadDebugMsg = `Uploaded ${fileUrls.length} files`;
      } else {
        uploadDebugMsg = `UPLOAD FAILED: ${uploadResult.message}`;
        Logger.log('Upload Failed: ' + uploadResult.message);
      }
    }
    
    // --- WRITE DATA ---
    // IMPORTANT: We write the passed 'userEmail' directly to Column A.
    // We do NOT use Session.getActiveUser() or placeholders.
    
    ordersSheet.getRange(newRow, 1).setValue(userEmail);          // A (User Email)
    ordersSheet.getRange(newRow, 2).setValue(new Date());         // B (Timestamp)
    ordersSheet.getRange(newRow, 3).setValue('');                 // C
    ordersSheet.getRange(newRow, 4).setValue(nextOrderId);        // D
    ordersSheet.getRange(newRow, 5).setValue(formData.confirmationState || 'Confirmed'); // E

    if (formData.invoiceDate) ordersSheet.getRange(newRow, 6).setValue(new Date(formData.invoiceDate)); // F
    if (formData.clientName) ordersSheet.getRange(newRow, 9).setValue(formData.clientName);             // I
    if (formData.vat) ordersSheet.getRange(newRow, 10).setValue(formData.vat);                          // J
    if (formData.poTawoos) ordersSheet.getRange(newRow, 11).setValue(formData.poTawoos);                // K
    if (formData.poPrivateLabel) ordersSheet.getRange(newRow, 12).setValue(formData.poPrivateLabel);    // L

    // Files (M, N, O)
    ordersSheet.getRange(newRow, 13, 1, 3).clearContent();
    if (fileUrls.length > 0) {
        if (fileUrls[0]) ordersSheet.getRange(newRow, 13).setValue(fileUrls[0]); // M
        if (fileUrls[1]) ordersSheet.getRange(newRow, 14).setValue(fileUrls[1]); // N
        if (fileUrls[2]) ordersSheet.getRange(newRow, 15).setValue(fileUrls[2]); // O
    }

    if (formData.internalComments) ordersSheet.getRange(newRow, 17).setValue(formData.internalComments); // Q
    if (formData.externalComments) ordersSheet.getRange(newRow, 18).setValue(formData.externalComments); // R
    if (formData.orderSummary) ordersSheet.getRange(newRow, 19).setValue(formData.orderSummary);         // S
    
    if (formData.salesRep) ordersSheet.getRange(newRow, 24).setValue(formData.salesRep);                 // X
    if (formData.totalWeight) ordersSheet.getRange(newRow, 27).setValue(formData.totalWeight);           // AA
    
    ordersSheet.getRange(newRow, 42).setValue(formData.orderType || 'توريد');                            // AP

    // Write Products
    if (formData.products && formData.products.length > 0) {
      const productColumnsMap = getProductColumnsMap(ordersSheet);
      formData.products.forEach(product => {
        const columns = productColumnsMap[product.productName];
        if (columns) {
          if (columns.qty) ordersSheet.getRange(newRow, columns.qty).setValue(product.qty);
          if (columns.bonus) ordersSheet.getRange(newRow, columns.bonus).setValue(product.bonus);
          if (columns.price) ordersSheet.getRange(newRow, columns.price).setValue(product.price);
          if (columns.discount) ordersSheet.getRange(newRow, columns.discount).setValue(product.discount);
          if (columns.invoiceDisc) ordersSheet.getRange(newRow, columns.invoiceDisc).setValue(product.invoiceDisc);
        }
      });
    }
    
    return { success: true, orderId: nextOrderId, message: "Order Saved. " + uploadDebugMsg };

  } catch (error) {
    Logger.log('CRITICAL SAVE ERROR: ' + error.toString());
    return { success: false, message: 'System error: ' + error.toString() };
  }
}

function getProductColumnsMap(ordersSheet) {
  const productColumnsMap = {};
  
  try {
    // Read row 3 (product names) and row 4 (field names) starting from column AQ (column 43)
    const startCol = 43; // Column AQ
    const maxCols = 200; // Read up to 200 columns
    
    const row3Data = ordersSheet.getRange(3, startCol, 1, maxCols).getValues()[0];
    const row4Data = ordersSheet.getRange(4, startCol, 1, maxCols).getValues()[0];
    
    // Build map of product names to their column numbers
    for (let i = 0; i < row3Data.length; i++) {
      const productName = row3Data[i] ? row3Data[i].toString().trim() : '';
      const fieldName = row4Data[i] ? row4Data[i].toString().trim().toUpperCase() : '';
      
      if (productName && fieldName) {
        // Initialize product entry if not exists
        if (!productColumnsMap[productName]) {
          productColumnsMap[productName] = {};
        }
        
        // Map field name to column number (1-indexed)
        const colNumber = startCol + i;
        
        if (fieldName === 'QTY') {
          productColumnsMap[productName].qty = colNumber;
        } else if (fieldName === 'BONUS') {
          productColumnsMap[productName].bonus = colNumber;
        } else if (fieldName === 'PRICE') {
          productColumnsMap[productName].price = colNumber;
        } else if (fieldName === 'DISCOUNT') {
          productColumnsMap[productName].discount = colNumber;
        } else if (fieldName === 'INVOICE DISC' || fieldName === 'INVOICE DISCOUNT' || fieldName === 'INV DISC') {
          productColumnsMap[productName].invoiceDisc = colNumber;
        }
      }
    }
    
    // Fallback: If row 4 reading failed, use fixed 5-column sequence
    const productsWithIncompleteMapping = Object.keys(productColumnsMap).filter(
      productName => !productColumnsMap[productName].qty || 
                     !productColumnsMap[productName].bonus || 
                     !productColumnsMap[productName].price || 
                     !productColumnsMap[productName].discount
    );
    
    if (productsWithIncompleteMapping.length > 0) {
      Logger.log('Warning: Some products have incomplete field mapping, using fallback');
      
      // Rebuild with fallback logic
      // UPDATED: Iterate by 5 to skip the new column structure correctly
      for (let i = 0; i < row3Data.length; i += 5) { 
        const productName = row3Data[i] ?
          row3Data[i].toString().trim() : '';
        
        if (productName) {
          // UPDATED: Map 5 columns based on "Qty, Bonus, Price, InvDisc, Discount" order
          productColumnsMap[productName] = {
            qty: startCol + i,
            bonus: startCol + i + 1,
            price: startCol + i + 2,
            discount: startCol + i + 3,     // Moved to 5th position
            invoiceDisc: startCol + i + 4 // New Column
          };
        }
      }
    }
    
  } catch (error) {
    Logger.log('Error building product columns map: ' + error.toString());
  }
  
  return productColumnsMap;
}

// Helper: Get product weight from Products sheet
function getProductWeight(productName) {
  try {
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const productsSheet = salesSpreadsheet.getSheetByName('Products');
    
    if (!productsSheet) {
      Logger.log('Products sheet not found');
      return 0;
    }
    
    const lastRow = productsSheet.getLastRow();
    if (lastRow < 3) return 0;
    
    // Get product names (Column A) and weights (Column G)
    const productNames = productsSheet.getRange(3, 1, lastRow - 2, 1).getValues();
    const weights = productsSheet.getRange(3, 7, lastRow - 2, 1).getValues();
    
    // Find matching product
    for (let i = 0; i < productNames.length; i++) {
      if (productNames[i][0] && productNames[i][0].toString().trim() === productName) {
        return parseFloat(weights[i][0]) || 0;
      }
    }
    
    return 0;
  } catch (error) {
    Logger.log('Error getting product weight: ' + error.toString());
    return 0;
  }
}

// Helper: Get client region from Clients sheet
function getClientRegion(clientName) {
  try {
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const clientsSheet = salesSpreadsheet.getSheetByName('Clients');
    
    if (!clientsSheet) {
      Logger.log('Clients sheet not found');
      return '';
    }
    
    const lastRow = clientsSheet.getLastRow();
    if (lastRow < 4) return '';
    
    // Get client names (Column D) and regions (Column Q)
    const clientNames = clientsSheet.getRange(4, 4, lastRow - 3, 1).getValues();
    const regions = clientsSheet.getRange(4, 17, lastRow - 3, 1).getValues();
    
    // Find matching client
    for (let i = 0; i < clientNames.length; i++) {
      if (clientNames[i][0] && clientNames[i][0].toString().trim().toLowerCase() === clientName.toLowerCase()) {
        return regions[i][0] ? regions[i][0].toString().trim() : '';
      }
    }
    
    return '';
  } catch (error) {
    Logger.log('Error getting client region: ' + error.toString());
    return '';
  }
}

function uploadFilesToDrive(filesData, orderId) {
  try {
    // 1. Validate Input
    if (!filesData || !Array.isArray(filesData) || filesData.length === 0) {
      return { success: true, fileUrls: [] };
    }

    // 2. Validate Folder Access
    const folderId = '1eRIkOEnHkhgUJytnGXda9IEOWprPqtRN'; 
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      return { success: false, message: "Invalid Folder ID or Permission Denied." };
    }

    const fileUrls = [];

    // 3. Process Files
    for (let i = 0; i < filesData.length && i < 3; i++) {
      const fileData = filesData[i];
      if (!fileData || !fileData.data) continue;

      try {
        const bytes = Utilities.base64Decode(fileData.data);
        const blob = Utilities.newBlob(bytes, fileData.mimeType, fileData.name);
        const newFileName = `Order_${orderId}_${fileData.name}`;
        
        const file = folder.createFile(blob);
        file.setName(newFileName);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // PUSH CLEAN URL
        fileUrls.push(file.getUrl());
      } catch (fileErr) {
        Logger.log(`Failed to create file ${i}: ${fileErr.toString()}`);
        // Continue to next file instead of failing completely
      }
    }

    return { success: true, fileUrls: fileUrls };

  } catch (error) {
    return { success: false, message: 'Upload Error: ' + error.toString() };
  }
}


// =============================================================================
// ALL ORDERS: EDIT & SMART SAVE (History Tracking)
// =============================================================================

/**
 * Fetches raw data for a specific order to populate the Edit Modal.
 */
function getOrderDataForEdit(orderId) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Orders');
    const lastRow = getOrdersLastRow(sheet);
    const orderIds = sheet.getRange(7, 4, lastRow - 6, 1).getValues(); // Column D
    
    let rowIndex = -1;
    for (let i = 0; i < orderIds.length; i++) {
      if (orderIds[i][0].toString() === orderId.toString()) {
        rowIndex = i + 7;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: 'Order not found' };

    // Fetch Full Row
    const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerRow = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Extract Products (Scan from Col AQ/42)
    const products = [];
    // Start scanning from index 42 (Column AQ) - Pattern: Qty, Bonus, Price, Disc, InvDisc
    for (let c = 42; c < rowValues.length; c += 5) {
        const qty = parseFloat(rowValues[c]) || 0;
        const bonus = parseFloat(rowValues[c+1]) || 0;
        
        if (qty > 0 || bonus > 0) {
            const prodName = headerRow[c] ? headerRow[c].toString().trim() : '';
            if (prodName) {
                products.push({
                    productName: prodName,
                    qty: qty,
                    bonus: bonus,
                    price: parseFloat(rowValues[c+2]) || 0,
                    discount: parseFloat(rowValues[c+3]) || 0,
                    invoiceDisc: parseFloat(rowValues[c+4]) || 0
                });
            }
        }
    }

    // Construct Payload
    const data = {
        invoiceDate: rowValues[5] ? Utilities.formatDate(new Date(rowValues[5]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '', // Col F
        clientName: rowValues[8], // Col I
        vat: rowValues[9], // Col J
        poTawoos: rowValues[10], // Col K
        poPrivateLabel: rowValues[11], // Col L
        internalComments: rowValues[16], // Col Q
        externalComments: rowValues[17], // Col R
        orderType: rowValues[41], // Col AP
        confirmationState: rowValues[4], // Col E (Version/Confirmation) - Assuming E is reused or check your map
        salesRep: rowValues[23], // Col X
        products: products
    };

    return { success: true, data: data };

  } catch (e) {
    Logger.log("Error in getOrderDataForEdit: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * SMART UPDATE: Compares new data vs old data.
 * Updates changed cells and appends a Note with history.
 */
function updateExistingOrder(orderId, formData, userEmail) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Orders');
    
    // 1. Find Row
    const lastRow = getOrdersLastRow(sheet);
    const orderIds = sheet.getRange(7, 4, lastRow - 6, 1).getValues();
    let rowIndex = -1;
    for (let i = 0; i < orderIds.length; i++) {
      if (orderIds[i][0].toString() === orderId.toString()) {
        rowIndex = i + 7;
        break;
      }
    }
    if (rowIndex === -1) return { success: false, message: 'Order not found' };

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm');
    const userLabel = userEmail.split('@')[0];

    // 2. Helper to Update Cell with Audit Note
    function smartUpdate(colIndex, newValue, oldValue) {
      // Normalize comparison (handle dates, strings, numbers)
      let valA = newValue;
      let valB = oldValue;
      
      if (valA instanceof Date) valA = valA.getTime();
      if (valB instanceof Date) valB = valB.getTime();
      if (valA === valB) return; // No change

      const cell = sheet.getRange(rowIndex, colIndex);
      
      // Update Value
      cell.setValue(newValue);

      // Append Note
      const oldNote = cell.getNote() || "";
      const changeLog = `[${timestamp} ${userLabel}] ${valB} -> ${newValue}`;
      cell.setNote(oldNote ? oldNote + "\n" + changeLog : changeLog);
    }

    // 3. Read Current Row Data
    const currentRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerRow = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues()[0];

    // 4. Update Standard Fields
    // Map: Field -> Col Index (1-based)
    // Date(F/6), Client(I/9), VAT(J/10), PO T(K/11), PO P(L/12), Int(Q/17), Ext(R/18), Type(AP/42), SalesRep(X/24)
    
    smartUpdate(6, new Date(formData.invoiceDate), currentRow[5]);
    smartUpdate(9, formData.clientName, currentRow[8]);
    smartUpdate(10, formData.vat, currentRow[9]);
    smartUpdate(11, formData.poTawoos, currentRow[10]);
    smartUpdate(12, formData.poPrivateLabel, currentRow[11]);
    smartUpdate(17, formData.internalComments, currentRow[16]);
    smartUpdate(18, formData.externalComments, currentRow[17]);
    smartUpdate(42, formData.orderType, currentRow[41]);
    smartUpdate(24, formData.salesRep, currentRow[23]);

    // 5. Update Confirmation State (Col E / 5)
    // Note: If you store "Confirmed/Revised" in Column E (Version), use 5. 
    // If it's stored in Status (Col G/7), use 7. Based on your code, let's assume Col E (4) or a new column.
    // The previous code mapped `row[4]` to `version`. If `confirmationState` is separate, please specify column.
    // For now, I will assume it maps to Column E (index 4) as per your object structure.
    smartUpdate(5, formData.confirmationState, currentRow[4]);


    // 6. Update Products (Complex)
    // Strategy: 
    // A. Create a map of Incoming Products: { "ProdName": {qty, bonus, price...} }
    // B. Loop through Sheet Headers (from 42). If header matches incoming -> update.
    // C. If header matches existing sheet value but NOT incoming (removed product) -> set to 0.

    const incomingMap = {};
    formData.products.forEach(p => {
        incomingMap[p.productName] = p;
    });

    for (let c = 42; c < headerRow.length; c += 5) {
        const prodName = headerRow[c] ? headerRow[c].toString().trim() : '';
        if (!prodName) continue;

        const incoming = incomingMap[prodName];
        
        // Indices (0-based from row array)
        const idxQty = c;
        const idxBonus = c+1;
        const idxPrice = c+2;
        const idxDisc = c+3;
        const idxInvD = c+4;

        if (incoming) {
            // Update values
            smartUpdate(c + 1, incoming.qty, currentRow[idxQty]);
            smartUpdate(c + 2, incoming.bonus, currentRow[idxBonus]);
            smartUpdate(c + 3, incoming.price, currentRow[idxPrice]);
            smartUpdate(c + 4, incoming.discount, currentRow[idxDisc]);
            smartUpdate(c + 5, incoming.invoiceDisc, currentRow[idxInvD]);
            
            // Remove from map to track processed items
            delete incomingMap[prodName];
        } else {
            // Product not in form, but might exist in sheet. 
            // If sheet has value > 0, set to 0 (Removal)
            if (currentRow[idxQty] > 0 || currentRow[idxBonus] > 0) {
                smartUpdate(c + 1, 0, currentRow[idxQty]);
                smartUpdate(c + 2, 0, currentRow[idxBonus]);
                smartUpdate(c + 3, 0, currentRow[idxPrice]); // Reset price too? Optional.
                smartUpdate(c + 4, 0, currentRow[idxDisc]);
                smartUpdate(c + 5, 0, currentRow[idxInvD]);
            }
        }
    }
    
    // 7. Update Total Weight (Col AA / 27)
    smartUpdate(27, formData.totalWeight, currentRow[26]);

    // 8. Update Order Summary Text (Col S / 19)
    smartUpdate(19, formData.orderSummary, currentRow[18]);

    return { success: true, message: 'Order updated successfully with history log.' };

  } catch (e) {
    Logger.log("Update Error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Updates just the confirmation status.
 */
function updateOrderConfirmationStatus(orderId, newStatus, userEmail) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Orders');
    const lastRow = getOrdersLastRow(sheet);
    const orderIds = sheet.getRange(7, 4, lastRow - 6, 1).getValues();
    
    let rowIndex = -1;
    for (let i = 0; i < orderIds.length; i++) {
      if (orderIds[i][0].toString() === orderId.toString()) {
        rowIndex = i + 7;
        break;
      }
    }
    if (rowIndex === -1) return { success: false, message: 'Order not found' };

    // Update Col E (5) for Confirmation Status
    // Also append note
    const cell = sheet.getRange(rowIndex, 5);
    const oldVal = cell.getValue();
    
    if (oldVal !== newStatus) {
        cell.setValue(newStatus);
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm');
        const userLabel = userEmail.split('@')[0];
        const changeLog = `[${timestamp} ${userLabel}] Status: ${oldVal} -> ${newStatus}`;
        const oldNote = cell.getNote() || "";
        cell.setNote(oldNote ? oldNote + "\n" + changeLog : changeLog);
    }

    return { success: true };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
