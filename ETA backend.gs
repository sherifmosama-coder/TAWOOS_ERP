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
    
    if (cacheSheet && cacheSheet.getLastRow() > 1) {
      const cacheData = cacheSheet.getRange(2, 1, cacheSheet.getLastRow() - 1, 22).getValues();
      
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
 * --- 4. SYNC ACTION ---
 * Called by frontend: refreshEtaData()
 */
function refreshEtaCache(userEmail) {
  const perm = getEtaModulePermission(userEmail);
  if (perm === 'none' || perm === 'viewer') return { success: false, message: "Permission Denied" };

  try {
    const ss = SpreadsheetApp.openById(ETA_MODULE_CONFIG.spreadsheetId);
    let cacheSheet = ss.getSheetByName(ETA_MODULE_CONFIG.syncSheetName);

    // Create Cache Sheet if missing
    if (!cacheSheet) {
      cacheSheet = ss.insertSheet(ETA_MODULE_CONFIG.syncSheetName);
      const headers = [
        "Internal ID", "ETA UUID", "Status", "Public URL", "PDF Link", "Verification",
        "Date Rec", "Date Issued", "Receiver Name", "Receiver ID", "Total Amount",
        "Net Amount", "Total Sales", "Total Disc", "Doc Type", "Type Name AR",
        "Issuer Name", "Issuer ID", "Long ID", "Sub UUID", "Raw JSON", "Type Version" // Added Col V
      ];
      cacheSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    }

    const token = getEtaModuleAccessToken();
    if (!token) throw new Error("Authentication failed");

    const docs = fetchEtaModuleRecentDocs(token);
    if (!docs.length) return { success: true, count: 0, message: "No new invoices found" };

    // PDF Map
    const existingData = cacheSheet.getDataRange().getValues();
    const pdfMap = new Map();
    if (existingData.length > 1) {
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i][0] && existingData[i][4]) {
           pdfMap.set(String(existingData[i][0]).trim(), existingData[i][4]);
        }
      }
    }

    const folder = DriveApp.getFolderById(ETA_MODULE_CONFIG.targetFolderId);
    const newRows = [];
    const fmtDate = (d) => d ? d.replace('T', ' ').replace('Z', '') : '';

    docs.forEach(doc => {
      const internalID = doc.internalId ? String(doc.internalId).trim() : "MISSING";
      let pdfUrl = pdfMap.get(internalID) || "";
      if (!pdfUrl && internalID !== "MISSING") {
         pdfUrl = saveEtaModulePdfToDrive(doc.uuid, internalID, token, folder);
      }

      newRows.push([
        internalID,                             
        doc.uuid || '',                         
        doc.status || '',                 
        doc.publicUrl || '',                    
        pdfUrl,                                 
        "✅ Synced",                          
        fmtDate(doc.dateTimeReceived),          
        fmtDate(doc.dateTimeIssued),            
        doc.receiverName || '',                 
        doc.receiverId || '',                   
        Number(doc.total) || 0,                 
        Number(doc.netAmount) || 0,             
        Number(doc.totalSalesAmount || doc.totalSales) || 0, 
        Number(doc.totalDiscountAmount || doc.totalDiscount) || 0, 
        doc.typeName || '',                     
        doc.documentTypeNameSecondaryLang || '',
        doc.issuerName || '',                   
        doc.issuerId || '',                     
        doc.longId || '',                       
        doc.submissionUUID || '',               
        JSON.stringify(doc),                    
        doc.typeVersionName || '' // Col V [cite: 1]
      ]);
    });

    if (newRows.length > 0) {
      if (cacheSheet.getLastRow() > 1) {
          // Clear 22 columns
          cacheSheet.getRange(2, 1, cacheSheet.getLastRow() - 1, 22).clearContent();
      }
      newRows.sort((a, b) => String(b[0]).localeCompare(String(a[0])));
      // Set 22 columns
      cacheSheet.getRange(2, 1, newRows.length, 22).setValues(newRows);
    }

    return { success: true, count: newRows.length };
  } catch (e) {
    Logger.log("Sync Error: " + e);
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
  } catch (e) { return 'viewer'; }
}
