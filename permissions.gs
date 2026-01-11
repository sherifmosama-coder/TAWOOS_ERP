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
