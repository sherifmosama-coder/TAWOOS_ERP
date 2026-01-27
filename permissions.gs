/**
 * =============================================================================
 * MASTER PERMISSION SCHEMA (Source of Truth)
 * =============================================================================
 * defines all modules, sub-features, and allowed permission levels.
 * IDs are automatically generated (e.g., "orders.eta") to match your system.
 */
const PERMISSION_SCHEMA = {
  // --- SALES MODULE ---
  SALES: {
    id: 'orders',       // Parent ID used in code
    name: 'Sales Module',
    features: {
      VIEW: { 
        id: 'view',     // Special ID: merges with parent (Result: "orders")
        type: 'boolean', 
        desc: 'Access Sales Module' 
      },
      ALL_ORDERS: { 
        id: 'all-orders', 
        type: 'level', 
        levels: ['viewer', 'editor'], 
        desc: 'All Orders Database' 
      },
      DISTRIBUTION: { 
        id: 'distribution', 
        type: 'level', 
        levels: ['viewer', 'editor'], 
        desc: 'Distribution Planning' 
      },
      ACCOUNTING: { 
        id: 'accounting', 
        type: 'level', 
        levels: ['viewer', 'editor'], 
        desc: 'Accounting & Invoices' 
      },
      WAREHOUSE: { 
        id: 'warehouse', 
        type: 'level', 
        levels: ['viewer', 'editor'], 
        desc: 'Release Stock (Warehouse)' 
      },
      RETURN: { 
        id: 'return', 
        type: 'level', 
        levels: ['viewer', 'editor'], 
        desc: 'Return Stock' 
      },
      ETA: { 
        id: 'eta', 
        type: 'level', 
        levels: ['viewer', 'editor', 'admin'], 
        desc: 'ETA Reconciliation' 
      }
    }
  },

  // --- OPERATIONS MODULE ---
  OPERATIONS: {
    id: 'operations',
    name: 'Operations Module',
    features: {
      VIEW: { 
        id: 'view', 
        type: 'boolean', 
        desc: 'Access Operations Module' 
      },
      PURCHASES: { 
        id: 'purchases', 
        type: 'level', 
        levels: ['viewer', 'editor'], 
        desc: 'Purchasing Dashboard' 
      },
      MATERIALS: { 
        id: 'materials', 
        type: 'level', 
        levels: ['viewer', 'editor'], 
        desc: 'Materials Inventory' 
      },
      PRODUCTION: { 
        id: 'production', 
        type: 'level', 
        levels: ['viewer', 'editor'], 
        desc: 'Production Planning' 
      }
    }
  },

  // --- SYSTEM ADMIN ---
  SYSTEM: {
    id: 'system',
    name: 'System Administration',
    features: {
      USERS: { 
        id: 'users', 
        type: 'level', 
        levels: ['admin'], 
        desc: 'Manage Users & Permissions' 
      }
    }
  }
};

/**
 * SYNC FUNCTION: Run this from the Script Editor to update the "Users" sheet.
 * It rewrites Columns A, B, and C with the structure defined in PERMISSION_SCHEMA.
 */
function syncPermissionsToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) throw new Error('Users sheet not found');
  
  const rows = [];

  // Recursive function to process features at any depth
  function processFeatures(features, parentId) {
    for (const [key, data] of Object.entries(features)) {
      // Logic: If id is 'view', use parentId. Otherwise, append (parent.child)
      const currentId = data.id === 'view' ? parentId : `${parentId}.${data.id}`;
      
      if (data.features) {
        // Recursive step for deeper nesting
        processFeatures(data.features, currentId);
      } else {
        // Create the row
        const validationRule = data.type === 'boolean' 
          ? 'TRUE / FALSE' 
          : data.levels.join(' / ');
          
        rows.push([currentId, data.desc, validationRule]);
      }
    }
  }

  // Iterate through Main Modules
  for (const [moduleKey, moduleData] of Object.entries(PERMISSION_SCHEMA)) {
    // Add visual header row
    rows.push([moduleData.id, `--- ${moduleData.name} ---`, '']); 
    
    if (moduleData.features) {
      processFeatures(moduleData.features, moduleData.id);
    }
  }
  
  // CLEAR & WRITE
  // We start at Row 5 (Assuming Row 1-4 are User Headers)
  const startRow = 5;
  const lastRow = Math.max(sheet.getLastRow(), startRow); 
  
  // Clear old definitions (Columns A, B, C only)
  sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).clearContent();
  
  // Write new definitions
  if (rows.length > 0) {
    sheet.getRange(startRow, 1, rows.length, 3).setValues(rows);
  }
  
  Logger.log(`Successfully synced ${rows.length} permission rows to the Users sheet.`);
}

// ==========================================
// PERMISSION SYSTEM: MATRIX BACKEND
// ==========================================

/**
 * 1. FETCH PERMISSIONS OBJECT (For System Use)
 * Parses the "Users" Matrix into a usable JSON tree.
 * Caches result to reduce Sheet reads.
 */
function getPermissions() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('PERMISSIONS_MATRIX_FINAL');
  if (cached) return JSON.parse(cached);

  const permissions = {};
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); // USES ACTIVE SHEET
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return {};

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // Need at least rows for header (1-2) and data (5+)
    if (lastRow < 5 || lastCol < 4) return {};

    // 1. Map Users (Row 2 = Emails, starting Col D)
    const userEmails = sheet.getRange(2, 4, 1, lastCol - 3).getValues()[0];
    
    // Initialize User Objects
    userEmails.forEach(email => {
      if(email) permissions[email] = {};
    });

    // 2. Read Matrix Data (Rows 5+)
    // Col A = IDs, Col D+ = Permissions
    const ids = sheet.getRange(5, 1, lastRow - 4, 1).getValues().flat();
    const matrix = sheet.getRange(5, 4, lastRow - 4, lastCol - 3).getValues();

    // 3. Parse Matrix
    for (let r = 0; r < ids.length; r++) {
      const id = ids[r];
      if (!id) continue;

      const parts = id.split('.'); // e.g. ['sales', 'orders']
      
      for (let u = 0; u < userEmails.length; u++) {
        const email = userEmails[u];
        if (!email) continue;
        
        let val = matrix[r][u]; 
        // Normalize value
        if (val === '' || val === null) continue;
        val = String(val);

        // Build the tree for this user
        let current = permissions[email];
        
        for (let i = 0; i < parts.length; i++) {
          const part = parts[i];
          
          if (i === parts.length - 1) {
            // Final Node (The Permission Value)
            // If it's a Module (Level 0), it expects Boolean
            if (val.toUpperCase() === 'TRUE') {
               current[part] = { _enabled: true }; 
            } else if (val.toUpperCase() === 'FALSE') {
               // Do nothing (disabled)
            } else {
               // It's a Role (admin, editor, viewer)
               current[part] = val.toLowerCase();
            }
          } else {
            // Intermediate Node (Folder)
            if (!current[part] || typeof current[part] !== 'object') {
              current[part] = {};
            }
            current = current[part];
          }
        }
      }
    }
    
    cache.put('PERMISSIONS_MATRIX_FINAL', JSON.stringify(permissions), 300); // 5 min cache
    return permissions;

  } catch (e) {
    Logger.log("Matrix Parse Error: " + e);
    return {};
  }
}

/**
 * 2. RESOLVER (Helper for Modules)
 * Returns: 'admin' | 'editor' | 'viewer' | 'none'
 */
function resolveModulePermission(email, section, tab) {
  if (!email) return 'none';
  
  // Refresh cache check
  const allPerms = getPermissions();
  const userPerms = allPerms[email];
  
  if (!userPerms) return 'none';

  // Legacy Global Admin Check (Optional: Check Row 3 role if exists)
  // For now, we assume explicit matrix assignment.

  // 1. Check Section (Module)
  if (userPerms[section]) {
    const secObj = userPerms[section];
    
    // If checking just the module existence
    if (!tab) return secObj ? 'viewer' : 'none'; 
    
    // 2. Check Tab/Subtab
    if (typeof secObj === 'object' && secObj[tab]) {
      return secObj[tab]; // Returns 'admin', 'editor', etc.
    }
  }
  
  return 'none';
}

// ==========================================
// MANAGER UI HANDLERS (Called by script-permissions.html)
// ==========================================

function getPermissionMatrixData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if(!sheet) throw "Users sheet missing";
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // A. Users
  const userCount = lastCol - 3;
  const users = [];
  if (userCount > 0) {
    const names = sheet.getRange(1, 4, 1, userCount).getValues()[0];
    const emails = sheet.getRange(2, 4, 1, userCount).getValues()[0];
    for(let i=0; i<userCount; i++) {
      if(emails[i]) users.push({ name: names[i], email: emails[i], colIndex: i + 4 });
    }
  }

  // B. Rows
  const rows = [];
  if (lastRow >= 5) {
    const idData = sheet.getRange(5, 1, lastRow - 4, 1).getValues();
    const descData = sheet.getRange(5, 2, lastRow - 4, 1).getValues();
    const permData = userCount > 0 ? sheet.getRange(5, 4, lastRow - 4, userCount).getValues() : [];
    
    for(let i=0; i<idData.length; i++) {
      if(idData[i][0]) {
        rows.push({
          id: idData[i][0],
          desc: descData[i][0],
          values: permData[i] || []
        });
      }
    }
  }
  
  return { users: users, rows: rows };
}

function updateUserPermissions(userEmail, changes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const emails = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIdx = emails.indexOf(userEmail) + 1;
  
  if(colIdx < 1) return "User not found";
  
  const ids = sheet.getRange(5, 1, sheet.getLastRow()-4, 1).getValues().flat();
  
  for (const [id, val] of Object.entries(changes)) {
    const rowOffset = ids.indexOf(id);
    if(rowOffset !== -1) {
      sheet.getRange(5 + rowOffset, colIdx).setValue(val);
    }
  }
  return "✅ Saved";
}

function syncScannedStructure(flatList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const lastRow = sheet.getLastRow();
  
  let existingIds = [];
  if (lastRow >= 5) existingIds = sheet.getRange(5, 1, lastRow - 4, 1).getValues().flat();
  
  let added = 0;
  flatList.forEach(item => {
    if(!existingIds.includes(item.id)) {
      sheet.appendRow([item.id, item.desc, item.type === 'module' ? 'TRUE/FALSE' : 'admin/editor/viewer']);
      added++;
    }
  });
  return `✅ Added ${added} new items.`;
}

function createNewUserColumn(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue("New User");
  sheet.getRange(2, col).setValue(email);
  return "User Created";
}
