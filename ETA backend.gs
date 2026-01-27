/**
 * =============================================================================
 * ETA E-INVOICING MODULE (PRODUCTION READY)
 * =============================================================================
 * Backend aligned with Stacked Row Interface & Version Control
 */

// --- 1. GLOBAL CONFIGURATION ---

const ETA_MODULE_CONFIG = {
  // ⚠️ PRE-PRODUCTION CREDENTIALS (Keep your existing ones)
  clientId: '9855fa54-a54b-4de8-a5bd-132a4d6d7ee2', 
  clientSecret: '45ab0ab3-ca95-4636-b8fd-820d81ff0535',
  authUrl: 'https://id.preprod.eta.gov.eg/connect/token',
  apiUrl: 'https://api.preprod.invoicing.eta.gov.eg',
  
  // FILE SYSTEM & SHEETS
  targetFolderId: '1s7OyKDsUnNcd9iVSUzniMZSkAmcIAjWd',
  spreadsheetId: '19gyEZlBxFiTBKpAt3wjwC1lNkglG9Oaovcz2d4ZnGcg', 
  
  syncSheetName: 'ETA Invoices',   // Cache Sheet
  dataSheetName: 'Invoices'        // Internal Data Sheet
};

// INTERNAL SHEET COLUMNS (Based on New Requirements)
// Indices are 0-based (Column A = 0)
const INTERNAL_COLS = {
  TYPE: 1,        // Col B
  VERSION: 2,     // Col C
  SERIAL: 3,      // Col D
  DATE: 4,        // Col E
  FULL_NAME: 5,   // Col F (NEW: Full Client Name)
  SUMMARY: 6,     // Col G
  TOTAL: 7,       // Col H
  DISC: 9,        // Col J
  PRE_TAX: 13,    // Col N
  TAX: 14,        // Col O
  PO: 16,         // Col Q
  CLIENT: 17,     // Col R
  TAX_ID: 18,     // Col S
  ORDER: 19,      // Col T
  LINK: 20        // Col U
};

/**
 * --- DASHBOARD DATA API ---
 */
function getReconciliationDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(ETA_MODULE_CONFIG.spreadsheetId);
    
    // A. FETCH INTERNAL DATA (With Versioning & Filtering)
    const intSheet = ss.getSheetByName(ETA_MODULE_CONFIG.dataSheetName);
    if (!intSheet) return JSON.stringify({ error: "Internal sheet not found" });

    const intData = intSheet.getDataRange().getValues();
    const internalMap = {};

    // Start from Row 5 (Index 4)
    for (let i = 4; i < intData.length; i++) {
      const row = intData[i];
      const type = String(row[INTERNAL_COLS.TYPE] || '').trim();
      if (type === 'DeliveryNote') continue;

      const serial = String(row[INTERNAL_COLS.SERIAL] || '').trim();
      if (!serial) continue;

      const currentVersion = Number(row[INTERNAL_COLS.VERSION]) || 0;

      if (!internalMap[serial] || currentVersion > internalMap[serial].version) {
         internalMap[serial] = {
            serial: serial,
            version: currentVersion,
            date: row[INTERNAL_COLS.DATE],
            fullName: row[INTERNAL_COLS.FULL_NAME],
            summary: row[INTERNAL_COLS.SUMMARY],
            total: Number(row[INTERNAL_COLS.TOTAL]) || 0,
            discount: Number(row[INTERNAL_COLS.DISC]) || 0,
            preTax: Number(row[INTERNAL_COLS.PRE_TAX]) || 0,
            tax: Number(row[INTERNAL_COLS.TAX]) || 0,
            po: row[INTERNAL_COLS.PO],
            client: row[INTERNAL_COLS.CLIENT],
            taxId: row[INTERNAL_COLS.TAX_ID],
            order: row[INTERNAL_COLS.ORDER],
            link: row[INTERNAL_COLS.LINK]
         };
      }
    }

    // B. FETCH ETA CACHE DATA
    // UPDATED: Now reading 22 Columns (A to V)
    const cacheSheet = ss.getSheetByName(ETA_MODULE_CONFIG.syncSheetName);
    const results = [];
    
    // FIX: Use getRealLastRow on Column A (1) to ignore Array Formula blanks
    const realLastRow = getRealLastRow(cacheSheet, 1);

    if (cacheSheet && realLastRow > 1) {
      const cacheData = cacheSheet.getRange(2, 1, realLastRow - 1, 22).getValues();
      for (let i = 0; i < cacheData.length; i++) {
        const row = cacheData[i];
        const key = String(row[0]).trim(); // Col A: Internal ID
        
        const etaObj = {
          uuid: row[1],             // Col B
          status: row[2],           // Col C
          pdf: row[4],              // Col E
          issueDate: row[7],        // Col H
          client: row[8],           // Col I
          taxId: row[9],            // Col J
          total: Number(row[10]),   // Col K
          typeAR: row[15],          // Col P
          versionName: row[21]      // Col V (NEW: Type Version)
        };

        const internalObj = internalMap[key] || null;
        let matchStatus = 'MATCHED';

        if (!internalObj) {
          matchStatus = 'MISSING_INTERNAL';
        } else if (etaObj.status === 'Cancelled' || etaObj.status === 'Invalid') {
          matchStatus = 'INVALID_ETA';
        } else if (internalObj && Math.abs(etaObj.total - internalObj.total) > 0.5) {
          matchStatus = 'MISMATCH_PRICE';
        }

        results.push({
          key: key,
          eta: etaObj,
          internal: internalObj,
          status: matchStatus
        });

        if (internalMap[key]) delete internalMap[key];
      }
    }

    // C. REMAINING
    for (let key in internalMap) {
      results.push({
        key: key,
        eta: null,
        internal: internalMap[key],
        status: 'MISSING_ON_ETA'
      });
    }

    results.sort((a,b) => {
      const score = s => (s === 'MATCHED') ? 3 : (s === 'MISSING_ON_ETA' ? 1 : 2);
      return score(a.status) - score(b.status);
    });

    return JSON.stringify({ data: results });

  } catch (err) {
    Logger.log("Dashboard Error: " + err);
    return JSON.stringify({ error: err.toString() });
  }
}


/**
 * --- 2. AUTHENTICATION ---
 */
function getEtaModuleAccessToken() {
  const payloadString = [
    'grant_type=client_credentials',
    'client_id=' + encodeURIComponent(ETA_MODULE_CONFIG.clientId.trim()),
    'client_secret=' + encodeURIComponent(ETA_MODULE_CONFIG.clientSecret.trim()),
    'scope=InvoicingAPI'
  ].join('&');

  const options = {
    'method': 'post',
    'headers': { 'Content-Type': 'application/x-www-form-urlencoded' },
    'payload': payloadString,
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(ETA_MODULE_CONFIG.authUrl, options);
    const json = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200 || json.error) throw new Error(json.error);
    return json.access_token;
  } catch (e) {
    Logger.log('ETA Auth Error: ' + e.toString());
    return null;
  }
}

/**
 * --- 4. SYNC ACTION (SMART MERGE & HISTORY PRESERVE) ---
 * Updates:
 * 1. Merges new Portal Data with Existing History (Doesn't wipe old rows).
 * 2. Updates status for any record found in the fetch (covering the 3-day grace period).
 * 3. Preserves PDF links unless UUID/Status changes.
 */
function refreshEtaCache(userEmail) {
  if (!userEmail) return { success: false, message: "Security Error: User email missing." };
  
  const perm = getEtaModulePermission(userEmail);
  if (perm === 'none' || perm === 'viewer') return { success: false, message: "Permission Denied." };

  try {
    const ss = SpreadsheetApp.openById(ETA_MODULE_CONFIG.spreadsheetId);
    let cacheSheet = ss.getSheetByName(ETA_MODULE_CONFIG.syncSheetName);
    
    // Create Sheet if missing
    if (!cacheSheet) {
      cacheSheet = ss.insertSheet(ETA_MODULE_CONFIG.syncSheetName);
      const headers = [
        "Internal ID", "ETA UUID", "Status", "Public URL", "PDF Link", "Verification",
        "Date Rec", "Date Issued", "Receiver Name", "Receiver ID", "Total Amount",
        "Net Amount", "Total Sales", "Total Disc", "Doc Type", "Type Name AR",
        "Issuer Name", "Issuer ID", "Long ID", "Sub UUID", "Raw JSON", "Type Version"
      ];
      cacheSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    }

    const token = getEtaModuleAccessToken();
    if (!token) throw new Error("Authentication failed");

    // 1. Fetch Recent Docs (Active Window)
    const docs = fetchEtaModuleRecentDocs(token);
    
    // 2. Load Existing History into a Map (Key: InternalID)
    const lastRow = cacheSheet.getLastRow();
    const existingMap = new Map(); // Stores the full row array
    
    if (lastRow > 1) {
      const data = cacheSheet.getRange(2, 1, lastRow - 1, 22).getValues();
      data.forEach(row => {
        const iId = String(row[0]).trim();
        if (iId) existingMap.set(iId, row);
      });
    }

    if (!docs.length && existingMap.size === 0) {
        return { success: true, count: 0, message: "No data found." };
    }

    // 3. Process Portal Data (Group by InternalID to find Winners)
    const groupedDocs = {};
    docs.forEach(doc => {
      const iId = doc.internalId ? String(doc.internalId).trim() : "MISSING";
      if (iId === "MISSING") return;
      if (!groupedDocs[iId]) groupedDocs[iId] = [];
      groupedDocs[iId].push(doc);
    });

    const fmtDate = (d) => d ? d.replace('T', ' ').replace('Z', '') : '';

    // 4. Update/Insert Records
    // We loop through the Portal Data and update the Map
    for (const [id, versions] of Object.entries(groupedDocs)) {
       
       // Enrich with 'Smart Status' (Check for Pending Cancellation)
       versions.forEach(v => {
           v.smartStatus = v.status;
           if (v.status === 'Valid' && v.cancelRequestDate) v.smartStatus = 'Cancelling';
           if (v.status === 'Valid' && v.rejectRequestDate) v.smartStatus = 'Rejecting';
       });

       // Sort: Cancelled/ing > Valid > Submitted
       versions.sort((a, b) => {
          const getScore = (s) => {
              if (s === 'Cancelled' || s === 'Cancelling' || s === 'Rejecting') return 5;
              if (s === 'Valid') return 4;
              if (s === 'Submitted') return 3;
              return 1;
          };
          const scoreDiff = getScore(b.smartStatus) - getScore(a.smartStatus);
          if (scoreDiff !== 0) return scoreDiff;
          return new Date(b.dateTimeIssued) - new Date(a.dateTimeIssued);
       });
       
       const winner = versions[0];
       
       // Check against Existing
       const existingRow = existingMap.get(id);
       let pdfUrl = "";
       
       if (existingRow) {
          const oldUuid = existingRow[1]; // Col B
          const oldStatus = existingRow[2]; // Col C
          const oldPdf = existingRow[4]; // Col E
          
          // Rule: Keep PDF unless UUID changed OR Status Changed (e.g. Valid -> Cancelling)
          // This ensures we refresh the PDF to get the "Cancelled" watermark
          if (oldUuid !== winner.uuid || oldStatus !== winner.smartStatus) {
             pdfUrl = ""; 
          } else {
             pdfUrl = oldPdf;
          }
       }

       // Construct the New Row
       const newRow = [
          id,
          winner.uuid || '',
          winner.smartStatus, // Updated Status
          winner.publicUrl || '',
          pdfUrl, 
          "✅ Synced",
          fmtDate(winner.dateTimeReceived),
          fmtDate(winner.dateTimeIssued),
          winner.receiverName || '',
          winner.receiverId || '',
          Number(winner.total) || 0,
          Number(winner.netAmount) || 0,
          Number(winner.totalSalesAmount || winner.totalSales) || 0,
          Number(winner.totalDiscountAmount || winner.totalDiscount) || 0,
          winner.typeName || '',
          winner.documentTypeNameSecondaryLang || '',
          winner.issuerName || '',
          winner.issuerId || '',
          winner.longId || '',
          winner.submissionUUID || '',
          JSON.stringify(winner),
          winner.typeVersionName || ''
       ];

       // UPDATE the Map (Upsert)
       existingMap.set(id, newRow);
    }

    // 5. Convert Map back to Array & Write
    const finalRows = Array.from(existingMap.values());
    
    // Sort by Internal ID Descending (Newest first)
    finalRows.sort((a, b) => String(b[0]).localeCompare(String(a[0])));

    if (finalRows.length > 0) {
       // Clear Sheet & Rewrite All (Safe because finalRows includes History + Updates)
       if (cacheSheet.getLastRow() > 1) {
          cacheSheet.getRange(2, 1, cacheSheet.getLastRow() - 1, 22).clearContent();
       }
       cacheSheet.getRange(2, 1, finalRows.length, 22).setValues(finalRows);
    }

    return { success: true, count: finalRows.length, message: "Sync complete. History preserved & Recent statuses updated." };

  } catch (e) {
    Logger.log("Sync Error: " + e);
    return { success: false, message: e.toString() };
  }
}

/**
 * --- 5. BATCH PDF SAVER (NEW) ---
 * Downloads multiple PDFs in parallel using UrlFetchApp.fetchAll
 * Updates the 'Invoices' sheet with the new links.
 * @param {Array} requests - Array of { uuid, internalId }
 */
function batchSaveEtaPdfs(requests) {
  if (!requests || requests.length === 0) return { success: true, saved: 0 };
  
  try {
    const token = getEtaModuleAccessToken();
    if (!token) throw new Error("Auth Failed");
    
    const folder = DriveApp.getFolderById(ETA_MODULE_CONFIG.targetFolderId);
    const ss = SpreadsheetApp.openById(ETA_MODULE_CONFIG.spreadsheetId);
    const sheet = ss.getSheetByName(ETA_MODULE_CONFIG.syncSheetName);
    
    // Map InternalID -> Row Index for fast updating
    // We read columns A (ID) and E (PDF URL)
    const lastRow = sheet.getLastRow();
    const sheetData = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // Cols A to E
    const rowMap = new Map();
    sheetData.forEach((r, i) => {
       if (r[0]) rowMap.set(String(r[0]).trim(), i + 2); // Store actual row number
    });

    let savedCount = 0;
    const CHUNK_SIZE = 15; // Process in chunks to avoid memory/timeout issues

    for (let i = 0; i < requests.length; i += CHUNK_SIZE) {
       const chunk = requests.slice(i, i + CHUNK_SIZE);
       
       // Prepare Fetch Requests
       const fetchPayloads = chunk.map(req => {
          return {
            url: `${ETA_MODULE_CONFIG.apiUrl}/api/v1.0/documents/${req.uuid}/pdf`,
            method: 'get',
            headers: { 
              'Authorization': 'Bearer ' + token, 
              'Accept': 'application/pdf',
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            },
            muteHttpExceptions: true
          };
       });
       
       // Execute Parallel Fetch
       const responses = UrlFetchApp.fetchAll(fetchPayloads);
       
       // Process Responses
       responses.forEach((res, idx) => {
          const req = chunk[idx];
          if (res.getResponseCode() === 200 && res.getHeaders()['Content-Type'].includes('pdf')) {
             const safeId = String(req.internalId).replace(/[^a-zA-Z0-9]/g, '');
             const fileName = `ETA_${safeId}_${req.uuid.substring(0, 8)}.pdf`;
             
             // Create File (Inherits Parent Permissions - Faster)
             const blob = res.getBlob().setName(fileName);
             const file = folder.createFile(blob); 
             const fileUrl = file.getUrl();
             
             // Update Sheet immediately (or batch if preferred, but immediate is safer here)
             const rowNum = rowMap.get(String(req.internalId));
             if (rowNum) {
                sheet.getRange(rowNum, 5).setValue(fileUrl); // Col E is PDF Link
             }
             savedCount++;
          }
       });
    }
    
    return { success: true, saved: savedCount };

  } catch (e) {
    Logger.log("Batch Save Error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * --- 5. HELPERS ---
 */
/**
 * --- 5. HELPERS ---
 * FIX: Added User-Agent to bypass ETA Firewall
 * FIX: Content-Type check to ensure we don't save HTML errors as PDFs
 */
function saveEtaModulePdfToDrive(uuid, internalId, token, folder) {
  try {
    const url = `${ETA_MODULE_CONFIG.apiUrl}/api/v1.0/documents/${uuid}/pdf`;
    
    const options = {
      'method': 'get',
      'headers': { 
        'Authorization': 'Bearer ' + token, 
        'Accept-Language': 'ar',
        'Accept': 'application/pdf',
        // 🛡️ CRITICAL: Mimic a standard browser to bypass ETA WAF
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
      },
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const contentType = response.getHeaders()['Content-Type'] || '';

    // 1. Check for success AND correct content type
    if (responseCode === 200 && contentType.includes('application/pdf')) {
      
      const safeId = String(internalId).replace(/[^a-zA-Z0-9]/g, '');
      const fileName = `ETA_${safeId}_${uuid.substring(0, 8)}.pdf`;
      const blob = response.getBlob().setName(fileName);
      
      // Create and Share
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      return file.getUrl();
    } else {
      // Log the error if it wasn't a PDF (likely the WAF HTML page)
      Logger.log(`PDF Download Blocked for ${internalId}. Code: ${responseCode}, Type: ${contentType}`);
      return ""; // Return empty so we can try again later
    }
  } catch (e) {
    Logger.log('PDF Save Exception: ' + e.toString());
    return "";
  }
}

function fetchEtaModuleRecentDocs(token) {
  const end = new Date();
  const start = new Date();
  start.setDate(start.getDate() - 30);
  const dFrom = start.toISOString();
  const dTo = end.toISOString();
  
  const url = `${ETA_MODULE_CONFIG.apiUrl}/api/v1.0/documents/recent?submissionDateFrom=${dFrom}&submissionDateTo=${dTo}&direction=Sent&pageSize=100&pageNo=1`;
  const options = {
    'method': 'get',
    'headers': { 'Authorization': 'Bearer ' + token },
    'muteHttpExceptions': true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    return (res.getResponseCode() === 200) ? JSON.parse(res.getContentText()).result || [] : [];
  } catch(e) { return []; }
}

function getEtaModulePermission(userEmail) {
  if (!userEmail) return 'viewer'; 
  try {
    const userPerms = getUserPermissions(userEmail); 
    if (!userPerms.success || !userPerms.permissions) return 'viewer';
    const perm = userPerms.permissions.find(p => p.tabId === 'orders.eta');
    return perm ? perm.permission.toLowerCase() : 'none';
  } catch (e) { return 'viewer'; 
  }
}

/**
 * Helper to find the ACTUAL last row based on a specific column.
 * Ignores empty strings returned by Array Formulas.
 * @param {Sheet} sheet - The sheet object
 * @param {Integer} columnNumber - 1-based column index (e.g., 1 for Col A)
 */
function getRealLastRow(sheet, columnNumber) {
  const maxRows = sheet.getMaxRows();
  if (maxRows === 0) return 0;
  
  // Fetch the entire column data
  const data = sheet.getRange(1, columnNumber, maxRows, 1).getValues();
  
  // Loop backwards from the bottom
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "" && data[i][0] != null) {
      return i + 1; // Return 1-based row index
    }
  }
  return 0; // Sheet is empty
}
