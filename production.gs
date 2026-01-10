// =============================================================================
// PRODUCTION MODULE BACKEND (production.gs)
// =============================================================================

const PRODUCTION_SPREADSHEET_ID = '1uRqCohOS6XbiAkv8KohivEM6Dj9WUQey_Fl3N0swAfw'; // Production Sheet ID

/**
 * Fetches initial data for the Production Planning view
 * - Products list from Index
 * - Recent Plans
 */
function getProductionInitialData() {
  try {
    const ss = SpreadsheetApp.openById(PRODUCTION_SPREADSHEET_ID);
    const plansSheet = ss.getSheetByName('خطط الانتاج');
    
    // 1. Get Dropdown Options (Products, etc.)
    const indexSheet = ss.getSheetByName('index');
    
    let products = [];
    if (indexSheet) {
      const lastRow = indexSheet.getLastRow();
      if (lastRow >= 2) {
         const data = indexSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Column A
         products = data.map(r => r[0].toString().trim()).filter(p => p !== '');
      }
    }

    // 2. Fetch Recent Plans (Last 30 days by default)
    const today = new Date();
    const thirtyDaysAgo = new Date(today.getTime() - (30 * 24 * 60 * 60 * 1000));
    const plans = getProductionPlansRaw(plansSheet, thirtyDaysAgo, null);

    return { 
      success: true, 
      products: products,
      plans: plans,
      lists: {
        constraints: ["تقريبي", "لا يقل عن", "لا يزيد عن", "بالعدد", "حسب الخامة المتاحة", "حسب الوقت المتاح"],
        priorities: ["خلال اليوم", "مهم", "مستعجل", "تجهيز فقط", "تجهيز و تشغيل", "لو متاح وقت"]
      }
    };

  } catch (e) {
    Logger.log('getProductionInitialData Error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Search for valid recipes for a product on a specific date
 */
function searchRecipes(productName, planDateStr) {
  try {
    const ss = SpreadsheetApp.openById(PRODUCTION_SPREADSHEET_ID);
    const recipeSheet = ss.getSheetByName('RECIPE');
    
    if (!recipeSheet) return { success: false, message: 'Recipes sheet not found' };

    const lastRow = recipeSheet.getLastRow();
    if (lastRow < 5) return { success: true, recipes: [] };

    const planDate = new Date(planDateStr);
    planDate.setHours(0,0,0,0);

    // Read Data (Start Row 5)
    // Col B (1): Recipe ID
    // Col D (3): Product Name
    // Col E (4): Start Date
    // Col G (6): End Date
    // Col H+ (7+): Materials Headers are in Row 4
    
    // Fetch Headers (Row 4) to get Material Names
    const lastCol = recipeSheet.getLastColumn();
    const headers = recipeSheet.getRange(4, 8, 1, lastCol - 7).getValues()[0]; // Start Col H

    const data = recipeSheet.getRange(5, 1, lastRow - 4, lastCol).getValues();
    const matches = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const pName = row[3] ? row[3].toString().trim() : '';
      
      // 1. Check Product Name
      if (pName !== productName) continue;

      // 2. Check Date Validity
      const validFrom = row[4] ? new Date(row[4]) : null;
      const validTo = row[6] ? new Date(row[6]) : null;

      if (validFrom) validFrom.setHours(0,0,0,0);
      if (validTo) validTo.setHours(0,0,0,0);

      if (validFrom && planDate < validFrom) continue;
      if (validTo && planDate > validTo) continue;

      // 3. Extract Required Materials (Values > 0)
      const materials = [];
      for (let m = 0; m < headers.length; m++) {
        const qty = parseFloat(row[m + 7]); // Offset 7 for Col H
        if (qty > 0) {
          materials.push({
            name: headers[m], // Material Name
            qtyPerUnit: qty
          });
        }
      }

      matches.push({
        id: row[1], // Recipe ID
        materials: materials
      });
    }

    return { success: true, recipes: matches };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Save a new Production Plan
 */
function saveProductionPlan(formData) {
  try {
    const ss = SpreadsheetApp.openById(PRODUCTION_SPREADSHEET_ID);
    let plansSheet = ss.getSheetByName('خطط الانتاج');
    
    // Create sheet if missing
    if (!plansSheet) {
      plansSheet = ss.insertSheet('خطط الانتاج');
      plansSheet.getRange('A2:Z2').setValues([['المعرف', 'التاريخ', 'المنتج', 'الكمية', 'محددات الكمية', 'الأولوية', 'ملاحظات', 'رقم الوصفة', 'خامة 1', 'نوع 1', 'خامة 2', 'نوع 2', 'خامة 3', 'نوع 3', 'خامة 4', 'نوع 4', '...', '...', '...', '...']]);
      plansSheet.getRange('A2:Z2').setFontWeight('bold').setBackground('#f3f3f3');
    }

    // 1. Generate ID (MMDDNN)
    const planDate = new Date(formData.date);
    const month = String(planDate.getMonth() + 1).padStart(2, '0');
    const day = String(planDate.getDate()).padStart(2, '0');
    const prefix = `${month}${day}`;
    
    const newId = generateDailyId(plansSheet, prefix);

    // 2. Prepare Row Data
    // Fixed Cols: A:H (8 columns)
    const fixedData = [
      newId,
      formData.date,
      formData.product,
      formData.quantity,
      formData.constraint,
      formData.priority,
      formData.notes || '',
      formData.recipeId || ''
    ];

    // Dynamic Materials (Col I onwards)
    // Structure: Name, Variation, Name, Variation...
    const materialData = [];
    if (formData.materials && Array.isArray(formData.materials)) {
      formData.materials.forEach(mat => {
        materialData.push(mat.name);
        materialData.push(mat.variation || ''); // Placeholder for variation logic
      });
    }

    const rowData = fixedData.concat(materialData);

    // 3. Save
    const lastRow = plansSheet.getLastRow();
    const targetRow = lastRow + 1;
    
    // Ensure we start at row 3 (if sheet was just created)
    const actualRow = targetRow < 3 ? 3 : targetRow;
    
    plansSheet.getRange(actualRow, 1, 1, rowData.length).setValues([rowData]);

    return { success: true, message: 'تم الحفظ بنجاح', newId: newId };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Helper: Generate unique daily ID (MMDDNN)
 */
function generateDailyId(sheet, prefix) {
  const lastRow = sheet.getLastRow();
  let maxCounter = 0;

  if (lastRow >= 3) {
    const ids = sheet.getRange(3, 1, lastRow - 2, 1).getValues(); // Col A
    ids.forEach(r => {
      const idStr = r[0].toString();
      if (idStr.startsWith(prefix)) {
        const counter = parseInt(idStr.substring(4)); // Get last 2 digits
        if (!isNaN(counter) && counter > maxCounter) {
          maxCounter = counter;
        }
      }
    });
  }

  const nextCounter = String(maxCounter + 1).padStart(2, '0');
  return `${prefix}${nextCounter}`;
}

/**
 * Helper: Fetch and parse plans
 */
function getProductionPlansRaw(sheet, startDate, endDate) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const plans = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[1];
    if (!dateVal) continue;

    const rowDate = new Date(dateVal);
    rowDate.setHours(0,0,0,0);

    if (startDate && rowDate < startDate) continue;
    if (endDate && rowDate > endDate) continue;

    // Parse Materials (Start Index 8 -> Col I)
    const materials = [];
    for (let m = 8; m < row.length; m += 2) {
      const matName = row[m];
      const matVar = row[m+1];
      if (matName) {
        materials.push({ name: matName, variation: matVar });
      }
    }

    plans.push({
      id: row[0],
      date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      product: row[2],
      quantity: row[3],
      constraint: row[4],
      priority: row[5],
      notes: row[6],
      recipeId: row[7],
      materials: materials
    });
  }
  
  // Sort by ID Descending (Newest first)
  return plans.reverse();
}
