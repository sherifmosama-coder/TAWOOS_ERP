// Code.gs - Google Apps Script Backend

// references to spreadsheet:
const SALES_SPREADSHEET_ID = '19gyEZlBxFiTBKpAt3wjwC1lNkglG9Oaovcz2d4ZnGcg';


function doGet(e) {
  return HtmlService.createTemplateFromFile('app')
    .evaluate()
    .setTitle('AL TAWOOS System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function validateLogin(email, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Users');
    
    if (!usersSheet) {
      return { success: false, message: 'Users sheet not found' };
    }
    
    // Get all data from row 2 (emails) and row 3 (passwords)
    const emailRow = usersSheet.getRange(2, 4, 1, usersSheet.getLastColumn() - 3).getValues()[0];
    const passwordRow = usersSheet.getRange(3, 4, 1, usersSheet.getLastColumn() - 3).getValues()[0];
    const nameRow = usersSheet.getRange(1, 4, 1, usersSheet.getLastColumn() - 3).getValues()[0];
    
    // Find matching email
    const userIndex = emailRow.findIndex(e => e.toString().toLowerCase() === email.toLowerCase());
    
    if (userIndex === -1) {
      return { success: false, message: 'User not found' };
    }
    
    // Validate password
    if (passwordRow[userIndex].toString() !== password) {
      return { success: false, message: 'Invalid password' };
    }
    
    // Successful login
    return {
      success: true,
      userName: nameRow[userIndex],
      userEmail: email
    };
    
  } catch (error) {
    Logger.log('Login error: ' + error.toString());
    return { success: false, message: 'System error' };
  }
}

function getUserPermissions(userEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Users');
    
    if (!usersSheet) {
      return { success: false, message: 'Users sheet not found' };
    }
    
    // Get email row to find user column
    const emailRow = usersSheet.getRange(2, 4, 1, usersSheet.getLastColumn() - 3).getValues()[0];
    const userIndex = emailRow.findIndex(e => e.toString().toLowerCase() === userEmail.toLowerCase());
    
    if (userIndex === -1) {
      return { success: false, message: 'User not found' };
    }
    
    // User column is at index 4 + userIndex (D=4, E=5, etc.)
    const userColumn = 4 + userIndex;
    
    // Get all permission data starting from row 5
    const lastRow = usersSheet.getLastRow();
    if (lastRow < 5) {
      return { success: true, permissions: [] };
    }
    
    // Get columns A-C (Tab ID, Spreadsheet, Tab) and user's permission column
    const tabIdData = usersSheet.getRange(5, 1, lastRow - 4, 1).getValues();
    const spreadsheetData = usersSheet.getRange(5, 2, lastRow - 4, 1).getValues();
    const tabData = usersSheet.getRange(5, 3, lastRow - 4, 1).getValues();
    const permissionData = usersSheet.getRange(5, userColumn, lastRow - 4, 1).getValues();
    
    // Build permissions array (including both main tabs and subtabs)
    const permissions = [];
    for (let i = 0; i < tabIdData.length; i++) {
      if (tabIdData[i][0] && spreadsheetData[i][0]) { // Only if Tab ID and Spreadsheet exist
        const tabId = tabIdData[i][0].toString();
        const permission = permissionData[i][0] ? permissionData[i][0].toString().toLowerCase() : '';
        
        // Determine if this is a main tab or subtab
        const isSubtab = tabId.includes('.');
        const parentTab = isSubtab ? tabId.split('.')[0] : null;
        const subtabId = isSubtab ? tabId.split('.')[1] : null;
        
        permissions.push({
          tabId: tabId,
          spreadsheet: spreadsheetData[i][0].toString(),
          tab: tabData[i][0].toString(),
          permission: permission,
          isSubtab: isSubtab,
          parentTab: parentTab,
          subtabId: subtabId
        });
      }
    }
    
    return { success: true, permissions: permissions };
    
  } catch (error) {
    Logger.log('Permission loading error: ' + error.toString());
    return { success: false, message: 'System error' };
  }
}

function getOrders(fetchAll) {
  try {
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = salesSpreadsheet.getSheetByName('Orders');
    const productsSheet = salesSpreadsheet.getSheetByName('Products');
    
    if (!ordersSheet) return { success: false, message: 'Orders sheet not found' };
    
    // 1. Build Product Type Map (Name -> isPL)
    const productMap = {};
    if (productsSheet) {
        const pLastRow = productsSheet.getLastRow();
        if (pLastRow >= 3) {
            // Read Name (A/0) and PL Flag (I/8)
            const pData = productsSheet.getRange(3, 1, pLastRow - 2, 9).getValues();
            pData.forEach(r => {
                if(r[0]) productMap[r[0].toString().trim()] = (r[8] === true);
            });
        }
    }

    const lastRow = getOrdersLastRow(ordersSheet);
    if (lastRow < 7) return { success: true, orders: [] };
    
    const lastCol = ordersSheet.getLastColumn();
    
    // 2. Get Product Headers from Row 3 (to map columns to names)
    const headerRow = ordersSheet.getRange(3, 1, 1, lastCol).getValues()[0];
    
    // 3. Get All Order Data
    const dataRange = ordersSheet.getRange(7, 1, lastRow - 6, lastCol);
    const data = dataRange.getValues();
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const orders = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[3]) continue; // Skip if no Order ID
      
      const status = row[2] ? row[2].toString().toLowerCase() : '';
      const isUnfulfilled = status !== 'delivered' && status !== 'paid';
      
      let invoiceDate = row[5];
      let loadingDate = row[7];
      if (invoiceDate && !(invoiceDate instanceof Date)) invoiceDate = new Date(invoiceDate);
      if (loadingDate && !(loadingDate instanceof Date)) loadingDate = new Date(loadingDate);
      
      const hasRecentDate = (invoiceDate && invoiceDate >= today) ||
                            (loadingDate && loadingDate >= today);
      
      // MODIFIED: If fetchAll is true, we ignore the status filter
      if (fetchAll || isUnfulfilled || hasRecentDate) {
        // --- DETERMINE PRODUCT MIX & COUNTS ---
        let tawoosCount = 0;
        let plCount = 0;
        let tawoosItems = [];
        let plItems = [];
        
        // Scan product columns starting at AQ (Index 42)
        for (let c = 42; c < row.length; c += 5) {
            const qty = parseFloat(row[c]) || 0;
            const bonus = parseFloat(row[c+1]) || 0;
            const total = qty + bonus;
            
            if (total > 0) {
                const prodName = headerRow[c] ? headerRow[c].toString().trim() : '';
                if (prodName) {
                    const itemSummary = `${total} ${prodName}`;
                    // Check map (Default to Tawoos if not found or false)
                    if (productMap[prodName] === true) {
                        plCount += total; // Sum of Qty + Bonus
                        plItems.push(itemSummary);
                    } else {
                        tawoosCount += total; // Sum of Qty + Bonus
                        tawoosItems.push(itemSummary);
                    }
                }
            }
        }
        
        const hasTawoos = tawoosCount > 0;
        const hasPL = plCount > 0;
        
        // Fallback for legacy data
        if (!hasTawoos && !hasPL && row[18]) {
             // If summary exists but no product columns, treat as 1 generic Tawoos item
             tawoosCount = 1;
             tawoosItems.push(row[18]); 
        }

        orders.push({
          rowIndex: i + 7,
          status: row[2],
          orderId: row[3],
          // Version/Confirmation Status (Col E)
          confirmationStatus: row[4] ? row[4].toString() : '', 
          invoiceDate: invoiceDate ? Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          loadingDate: loadingDate ? Utilities.formatDate(loadingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          clientName: row[8],
          vat: row[9],
          poTawoos: row[10],
          poPrivateLabel: row[11],
          attachment1: row[12],
          attachment2: row[13],
          attachment3: row[14],
          internalComments: row[16],
          externalComments: row[17],
          orderSummary: row[18],
          loadingOrder: row[19],
          truck: row[20],
          driver: row[21],
          deliveryRep: row[22],
          salesRep: row[23],
          clientContact: row[15],
          totalWeight: row[26],
          deliveryNote: row[28],
          tawoosInvoice: row[29],
          plInvoice: row[30],
          amountExcludingTax: row[37],
          taxValue: row[38],
          deduction1Percent: row[39],
          amountIncludingTax: row[40],
          orderType: row[41],
          // NEW: ETA UUIDs (Columns AI=34, AJ=35)
          tawoosEtaUuid: row[34] ? row[34].toString() : '',
          plEtaUuid: row[35] ? row[35].toString() : '',
          // NEW FLAGS & COUNTS
          hasTawoos: hasTawoos || tawoosCount > 0,
          hasPL: hasPL || plCount > 0,
          tawoosCount: tawoosCount,
          plCount: plCount,
          tawoosSummary: tawoosItems.join('<br>'),
          plSummary: plItems.join('<br>')
        });
      }
    }
    
    return { success: true, orders: orders };
  } catch (error) {
    Logger.log('Get orders error: ' + error.toString());
    return { success: false, message: 'System error: ' + error.toString() };
  }
}

// Test function to trigger Drive authorization
function testDriveAccess() {
  try {
    const folderId = '1eRIkOEnHkhgUJytnGXda9IEOWprPqtRN';
    const folder = DriveApp.getFolderById(folderId);
    Logger.log('Successfully accessed folder: ' + folder.getName());
    return 'Success! Folder name: ' + folder.getName();
  } catch (error) {
    Logger.log('Error accessing Drive: ' + error.toString());
    return 'Error: ' + error.toString();
  }
}

// Get orders for distribution planning
function getDistributionOrders(selectedLoadingDate) {
  try {
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = salesSpreadsheet.getSheetByName('Orders');
    
    if (!ordersSheet) {
      return { success: false, message: 'Orders sheet not found' };
    }
    
    const lastRow = getOrdersLastRow(ordersSheet);
    if (lastRow < 7) {
      return { success: true, orders: [] };
    }
    
    // Parse selected loading date
    const selectedDate = new Date(selectedLoadingDate);
    selectedDate.setHours(0, 0, 0, 0);
    
    // Get all data from row 7 onwards (A to AP - 42 columns)
    const dataRange = ordersSheet.getRange(7, 1, lastRow - 6, 42);
    const data = dataRange.getValues();
    
    const orders = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows (check Order ID column D)
      if (!row[3]) continue;
      
      const status = row[2] ? row[2].toString().toLowerCase() : '';
      const isBlankStatus = status === '' || status === 'not yet';
      
      // Get dates
      let invoiceDate = row[5];
      let loadingDate = row[7];
      
      if (invoiceDate && !(invoiceDate instanceof Date)) {
        invoiceDate = new Date(invoiceDate);
      }
      if (loadingDate && !(loadingDate instanceof Date)) {
        loadingDate = new Date(loadingDate);
      }
      
      if (invoiceDate) invoiceDate.setHours(0, 0, 0, 0);
      if (loadingDate) loadingDate.setHours(0, 0, 0, 0);
      
      // Filter logic: (Status ≠ Delivered AND Status ≠ Paid) OR (Loading Date = planDate OR Invoice Date = planDate)
      const notDeliveredOrPaid = status !== 'delivered' && status !== 'paid';
      const dateMatchesExactly = (loadingDate && loadingDate.getTime() === selectedDate.getTime()) || 
                                (invoiceDate && invoiceDate.getTime() === selectedDate.getTime());

      if (notDeliveredOrPaid || dateMatchesExactly) {
        // Get client region
        const clientRegion = getClientRegion(row[8]); // Column I - Client Name
        
        // Get distance from base
        const distanceFromBase = clientRegion ? getDistanceFromBase(clientRegion) : null;
        
        // Check if order loading date matches selected date
        const isDateMismatch = loadingDate && loadingDate.getTime() !== selectedDate.getTime();
        
        // Check if has delivery data
        const hasDeliveryData = row[19] || row[20] || row[21] || row[22]; // Loading Order, Truck, Driver, Delivery Rep
        
        // Check if disabled (delivered/paid)
        const isDisabled = status === 'delivered' || status === 'paid';
        
        orders.push({
            rowIndex: i + 7,
            status: row[2],
            orderId: row[3],
            invoiceDate: invoiceDate ? Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
            loadingDate: loadingDate ? Utilities.formatDate(loadingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
            clientName: row[8],
            clientRegion: clientRegion,
            distanceFromBase: distanceFromBase,
            loadingOrder: row[19],
            truck: row[20],
            driver: row[21],
            deliveryRep: row[22],
            dismissNumber: row[24], // Column Y
            totalWeight: row[26],
            clientContact: row[15], // Column P
            orderSummary: row[18], // Column S
            internalComments: row[16], // Column Q
            isDateMismatch: isDateMismatch,
            hasDeliveryData: hasDeliveryData,
            isDisabled: isDisabled
      });
      }
    }
    
    return { success: true, orders: orders };
    
  } catch (error) {
    Logger.log('Get distribution orders error: ' + error.toString());
    return { success: false, message: 'System error: ' + error.toString() };
  }
}

// Get distance from base location (ابو رواش) to a region
function getDistanceFromBase(regionName) {
  try {
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const regionsSheet = salesSpreadsheet.getSheetByName('المناطق');
    
    if (!regionsSheet) return null;
    
    const baseLocation = 'ابو رواش';
    
    // Get all region names from Column A (starting row 3)
    const lastRow = regionsSheet.getLastRow();
    if (lastRow < 3) return null;
    
    const regionNamesCol = regionsSheet.getRange(3, 1, lastRow - 2, 1).getValues();
    const headerRow = regionsSheet.getRange(1, 4, 1, lastRow - 2).getValues()[0];
    
    // Find base location row
    let baseRowIndex = -1;
    for (let i = 0; i < regionNamesCol.length; i++) {
      if (regionNamesCol[i][0] && regionNamesCol[i][0].toString().trim() === baseLocation) {
        baseRowIndex = i + 3; // Actual row number
        break;
      }
    }
    
    if (baseRowIndex === -1) return null;
    
    // Find target region column
    let targetColIndex = -1;
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i] && headerRow[i].toString().trim() === regionName) {
        targetColIndex = i + 4; // Actual column number (D=4)
        break;
      }
    }
    
    if (targetColIndex === -1) return null;
    
    // Get distance value
    const distance = regionsSheet.getRange(baseRowIndex, targetColIndex).getValue();
    
    return distance ? parseFloat(distance) : null;
    
  } catch (error) {
    Logger.log('Get distance from base error: ' + error.toString());
    return null;
  }
}

// Get distance between two regions
function getDistanceBetweenRegions(region1, region2) {
  try {
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const regionsSheet = salesSpreadsheet.getSheetByName('المناطق');
    
    if (!regionsSheet) {
      return { success: false, message: 'Regions sheet not found' };
    }
    
    const lastRow = regionsSheet.getLastRow();
    if (lastRow < 3) {
      return { success: false, message: 'No region data found' };
    }
    
    const regionNamesCol = regionsSheet.getRange(3, 1, lastRow - 2, 1).getValues();
    const headerRow = regionsSheet.getRange(1, 4, 1, lastRow - 2).getValues()[0];
    
    // Find region1 row
    let region1RowIndex = -1;
    for (let i = 0; i < regionNamesCol.length; i++) {
      if (regionNamesCol[i][0] && regionNamesCol[i][0].toString().trim() === region1) {
        region1RowIndex = i + 3;
        break;
      }
    }
    
    if (region1RowIndex === -1) {
      return { success: false, message: 'Region 1 not found: ' + region1 };
    }
    
    // Find region2 column
    let region2ColIndex = -1;
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i] && headerRow[i].toString().trim() === region2) {
        region2ColIndex = i + 4;
        break;
      }
    }
    
    if (region2ColIndex === -1) {
      return { success: false, message: 'Region 2 not found: ' + region2 };
    }
    
    // Get distance
    const distance = regionsSheet.getRange(region1RowIndex, region2ColIndex).getValue();
    
    if (!distance && distance !== 0) {
      return { success: false, message: 'Distance not found in matrix' };
    }
    
    return { 
      success: true, 
      distance: parseFloat(distance),
      region1: region1,
      region2: region2
    };
    
  } catch (error) {
    Logger.log('Get distance between regions error: ' + error.toString());
    return { success: false, message: 'System error: ' + error.toString() };
  }
}

// Get last used loading order for a date
function getLastLoadingOrderForDate(selectedDate) {
  try {
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = salesSpreadsheet.getSheetByName('Orders');
    
    if (!ordersSheet) return { success: true, lastLoadingOrder: null };
    
    const lastRow = getOrdersLastRow(ordersSheet);
    if (lastRow < 7) return { success: true, lastLoadingOrder: null };
    
    const dateObj = new Date(selectedDate);
    dateObj.setHours(0, 0, 0, 0);
    
    // Get loading dates and loading orders
    const loadingDates = ordersSheet.getRange(7, 8, lastRow - 6, 1).getValues(); // Column H
    const loadingOrders = ordersSheet.getRange(7, 20, lastRow - 6, 1).getValues(); // Column T
    
    let maxLoadingOrder = 0;
    
    for (let i = 0; i < loadingDates.length; i++) {
      let orderDate = loadingDates[i][0];
      
      if (orderDate && orderDate instanceof Date) {
        orderDate.setHours(0, 0, 0, 0);
        
        if (orderDate.getTime() === dateObj.getTime()) {
          const loadingOrderNum = loadingOrders[i][0];
          if (loadingOrderNum && !isNaN(loadingOrderNum)) {
            const num = parseInt(loadingOrderNum);
            if (num > maxLoadingOrder) {
              maxLoadingOrder = num;
            }
          }
        }
      }
    }
    
    return { 
      success: true, 
      lastLoadingOrder: maxLoadingOrder > 0 ? maxLoadingOrder : null 
    };
    
  } catch (error) {
    Logger.log('Get last loading order error: ' + error.toString());
    return { success: false, message: 'System error: ' + error.toString() };
  }
}

// =============================================================================
// SAVE LOADING ORDERS
// =============================================================================

function saveLoadingOrders(loadingDate, loadingOrdersData) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = ss.getSheetByName('Orders');
    
    if (!ordersSheet) {
      return { success: false, message: 'Orders sheet not found' };
    }
    
    // Validate loading date
    if (!loadingDate) {
      return { success: false, message: 'Loading date is required' };
    }
    
    const dateObj = new Date(loadingDate);
    dateObj.setHours(0, 0, 0, 0);
    
    // CRITICAL VALIDATION: Check for duplicate (date + LO#) with different truck/driver
    const validation = validateLoadingOrderUniqueness(ordersSheet, loadingOrdersData);
    if (!validation.success) {
      return validation;
    }
    
    // Process each loading order
    for (let lo of loadingOrdersData) {
      const loadingOrderNumber = lo.loadingOrderNumber;
      const truck = lo.truck;
      const driver = lo.driver;
      const deliveryRep = lo.deliveryRep || '';
      const status = lo.status || 'Prepare';
      const orderIds = lo.orderIds;
      const loLoadingDate = lo.loadingDate || loadingDate;
      
      // Update each order in this loading order
      for (let orderId of orderIds) {
        const result = updateOrderLoadingInfo(
          ordersSheet,
          orderId,
          loadingOrderNumber,
          truck,
          driver,
          deliveryRep,
          status,
          loLoadingDate
        );
        
        if (!result.success) {
          return { 
            success: false, 
            message: `Failed to update Order ${orderId}: ${result.message}` 
          };
        }
      }
    }
    
    return { success: true, message: 'Loading orders saved successfully' };
    
  } catch (error) {
    Logger.log('Save loading orders error: ' + error.toString());
    return { success: false, message: 'System error: ' + error.toString() };
  }
}

function updateOrderLoadingInfo(ordersSheet, orderId, loadingOrderNumber, truck, driver, deliveryRep, newStatus, loadingDate) {
  try {
    const lastRow = getOrdersLastRow(ordersSheet);
    
    // Find order by Order ID (Column A)
    const orderIds = ordersSheet.getRange(7, 1, lastRow - 6, 1).getValues();
    
    let rowIndex = -1;
    for (let i = 0; i < orderIds.length; i++) {
      if (orderIds[i][0] && orderIds[i][0].toString() === orderId.toString()) {
        rowIndex = i + 7; // Actual row number
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: `Order ${orderId} not found` };
    }
    
// Update columns:
    // G (7): Status -> "Prepare"
    // H (8): Loading Date
    // T (20): Loading Order Number
    // U (21): Truck
    // V (22): Driver
    // W (23): Delivery Rep
    
    ordersSheet.getRange(rowIndex, 7).setValue(newStatus); // Status
    ordersSheet.getRange(rowIndex, 8).setValue(new Date(loadingDate)); // Loading Date
    ordersSheet.getRange(rowIndex, 20).setValue(loadingOrderNumber); // Loading Order
    ordersSheet.getRange(rowIndex, 21).setValue(truck); // Truck
    ordersSheet.getRange(rowIndex, 22).setValue(driver); // Driver
    ordersSheet.getRange(rowIndex, 23).setValue(deliveryRep); // Delivery Rep
    
    return { success: true };
    
  } catch (error) {
    Logger.log('Update order loading info error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function validateLoadingOrderUniqueness(ordersSheet, loadingOrdersData) {
  try {
    const lastRow = getOrdersLastRow(ordersSheet);
    if (lastRow < 7) return { success: true };
    
    // Get all existing orders with loading data
    const dataRange = ordersSheet.getRange(7, 1, lastRow - 6, 23);
    const data = dataRange.getValues();
    
    // Build map of (date + LO#) -> {truck, driver}
    const existingMap = {};
    
    for (let i = 0; i < data.length; i++) {
      const orderId = data[i][0]; // Column A
      if (!orderId) continue;
      
      const loadingDateValue = data[i][7]; // Column H
      const loNumber = data[i][19]; // Column T
      const truck = data[i][20]; // Column U
      const driver = data[i][21]; // Column V
      
      if (!loNumber || !truck || !driver) continue;
      
      let dateStr = '';
      if (loadingDateValue instanceof Date) {
        const d = new Date(loadingDateValue);
        d.setHours(0, 0, 0, 0);
        dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      
      if (!dateStr) continue;
      
      const key = dateStr + '_' + loNumber;
      
      if (!existingMap[key]) {
        existingMap[key] = { truck: truck, driver: driver, orderIds: [] };
      }
      
      // Check if same key has different truck/driver
      if (existingMap[key].truck !== truck || existingMap[key].driver !== driver) {
        return {
          success: false,
          message: 'DATABASE CONFLICT: Loading Order #' + loNumber + ' on ' + dateStr + 
                   ' already exists with different truck/driver. Existing: ' + 
                   existingMap[key].truck + '/' + existingMap[key].driver + 
                   ', Attempted: ' + truck + '/' + driver
        };
      }
      
      existingMap[key].orderIds.push(orderId.toString());
    }
    
    // Validate incoming data against existing and itself
    for (let lo of loadingOrdersData) {
      const dateStr = lo.loadingDate;
      const loNumber = lo.loadingOrderNumber;
      const truck = lo.truck;
      const driver = lo.driver;
      const orderIds = lo.orderIds || [];
      
      const key = dateStr + '_' + loNumber;
      
      // Check if this LO# already exists in database with different truck/driver
      if (existingMap[key]) {
        // Check if any orders in incoming data are NOT in existing
        const newOrders = orderIds.filter(id => !existingMap[key].orderIds.includes(id.toString()));
        
        if (newOrders.length > 0) {
          // New orders being added to existing LO - must match truck/driver
          if (existingMap[key].truck !== truck || existingMap[key].driver !== driver) {
            return {
              success: false,
              message: 'CONFLICT: Loading Order #' + loNumber + ' on ' + dateStr + 
                       ' already exists in database with Truck: ' + existingMap[key].truck + 
                       ', Driver: ' + existingMap[key].driver + 
                       '. Cannot change to Truck: ' + truck + ', Driver: ' + driver
            };
          }
        }
      }
    }
    
    // Validate incoming data against itself (no duplicates in current save)
    const incomingMap = {};
    for (let lo of loadingOrdersData) {
      const key = lo.loadingDate + '_' + lo.loadingOrderNumber;
      
      if (!incomingMap[key]) {
        incomingMap[key] = { truck: lo.truck, driver: lo.driver };
      } else {
        // Same key in incoming data - check if truck/driver match
        if (incomingMap[key].truck !== lo.truck || incomingMap[key].driver !== lo.driver) {
          return {
            success: false,
            message: 'CONFLICT IN CURRENT SAVE: Loading Order #' + lo.loadingOrderNumber + 
                     ' on ' + lo.loadingDate + ' has conflicting assignments. ' +
                     'First: ' + incomingMap[key].truck + '/' + incomingMap[key].driver + 
                     ', Second: ' + lo.truck + '/' + lo.driver
          };
        }
      }
    }
    
    return { success: true };
    
  } catch (error) {
    Logger.log('Validation error: ' + error.toString());
    return { success: false, message: 'Validation error: ' + error.toString() };
  }
}

// =============================================================================
// GET TRUCKS DATA
// =============================================================================

function getTrucksData() {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const trucksSheet = ss.getSheetByName('Trucks');
    
    if (!trucksSheet) {
      return { success: true, trucks: [] };
    }
    
    const lastRow = trucksSheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, trucks: [] };
    }
    
    const truckNames = trucksSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const drivers = trucksSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const capacities = trucksSheet.getRange(2, 7, lastRow - 1, 1).getValues();
    
    const trucks = [];
    
    for (let i = 0; i < truckNames.length; i++) {
      const truckName = truckNames[i][0];
      if (!truckName || truckName.toString().trim() === '') continue;
      
      const driver = drivers[i][0] ? drivers[i][0].toString().trim() : '';
      const capacity = capacities[i][0] && !isNaN(capacities[i][0]) ? parseFloat(capacities[i][0]) : 0;
      
      if (capacity > 0) {
        trucks.push({
          name: truckName.toString().trim(),
          driver: driver,
          capacity: capacity
        });
      }
    }
    
    Logger.log('Loaded trucks: ' + trucks.length);
    return { success: true, trucks: trucks };
    
  } catch (error) {
    Logger.log('Get trucks error: ' + error.toString());
    return { success: true, trucks: [] };
  }
}

function getDriversList() {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const indexSheet = ss.getSheetByName('Index');
    
    if (!indexSheet) {
      return { success: true, drivers: [] };
    }
    
    const driversData = indexSheet.getRange('X3:X12').getValues();
    const drivers = [];
    
    driversData.forEach(row => {
      if (row[0] && row[0].toString().trim() !== '') {
        drivers.push(row[0].toString().trim());
      }
    });
    
    Logger.log('Loaded drivers from Index X3:X12: ' + drivers.length);
    return { success: true, drivers: drivers };
    
  } catch (error) {
    Logger.log('Get drivers error: ' + error.toString());
    return { success: true, drivers: [] };
  }
}

// =============================================================================
// GET EXISTING LOADING ORDERS (NOT DELIVERED/PAID)
// =============================================================================

function getExistingLoadingOrders(loadingDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Orders');
    
    if (!ordersSheet) {
      return { success: false, message: 'Orders sheet not found' };
    }
    
    // Validate loading date
    if (!loadingDate) {
      return { success: false, message: 'Loading date is required' };
    }
    
    const dateObj = new Date(loadingDate);
    dateObj.setHours(0, 0, 0, 0);
    
    const lastRow = getOrdersLastRow(ordersSheet);
    if (lastRow < 7) {
      return { success: true, loadingOrders: [] };
    }
    
    // Get necessary columns
    const orderIds = ordersSheet.getRange(7, 1, lastRow - 6, 1).getValues(); // A: Order ID
    const statuses = ordersSheet.getRange(7, 7, lastRow - 6, 1).getValues(); // G: Status
    const loadingDates = ordersSheet.getRange(7, 8, lastRow - 6, 1).getValues(); // H: Loading Date
    const clientNames = ordersSheet.getRange(7, 4, lastRow - 6, 1).getValues(); // D: Client Name
    const loadingOrderNums = ordersSheet.getRange(7, 20, lastRow - 6, 1).getValues(); // T: Loading Order
    const trucks = ordersSheet.getRange(7, 21, lastRow - 6, 1).getValues(); // U: Truck
    const drivers = ordersSheet.getRange(7, 22, lastRow - 6, 1).getValues(); // V: Driver
    const deliveryReps = ordersSheet.getRange(7, 23, lastRow - 6, 1).getValues(); // W: Delivery Rep
    
    const loadingOrders = [];
    
    for (let i = 0; i < orderIds.length; i++) {
      const orderId = orderIds[i][0];
      const status = statuses[i][0] ? statuses[i][0].toString().trim() : '';
      const loadingDateValue = loadingDates[i][0];
      const loadingOrderNum = loadingOrderNums[i][0];
      
      // Skip if no order ID or loading order number
      if (!orderId || !loadingOrderNum) continue;
      
      // Skip if status is "Delivered" or "Paid"
      if (status.toLowerCase() === 'delivered' || status.toLowerCase() === 'paid') {
        continue;
      }
      
      // Check loading date match
      if (loadingDateValue && loadingDateValue instanceof Date) {
        const orderLoadingDate = new Date(loadingDateValue);
        orderLoadingDate.setHours(0, 0, 0, 0);
        
        if (orderLoadingDate.getTime() === dateObj.getTime()) {
          loadingOrders.push({
            orderId: orderId.toString(),
            clientName: clientNames[i][0] ? clientNames[i][0].toString() : '',
            loadingOrderNumber: parseInt(loadingOrderNum),
            truck: trucks[i][0] ? trucks[i][0].toString() : '',
            driver: drivers[i][0] ? drivers[i][0].toString() : '',
            deliveryRep: deliveryReps[i][0] ? deliveryReps[i][0].toString() : '',
            status: status,
            loadingDate: Utilities.formatDate(orderLoadingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          });
        }
      }
    }
    
    Logger.log('Found existing loading orders: ' + loadingOrders.length);
    return { success: true, loadingOrders: loadingOrders };
    
  } catch (error) {
    Logger.log('Get existing loading orders error: ' + error.toString());
    return { success: false, message: 'System error: ' + error.toString() };
  }
}
// =============================================================================
// DISTRIBUTION PLANNING - BACKEND FUNCTIONS (CORRECT COLUMN MAPPING)
// Based on your actual Orders sheet structure
// =============================================================================

function getAllUndeliveredOrders() {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = ss.getSheetByName('Orders');
    const clientsSheet = ss.getSheetByName('Clients');
    const regionsSheet = ss.getSheetByName('المناطق');
    
    if (!ordersSheet) {
      return { success: false, message: 'Orders sheet not found' };
    }
    
    const lastRow = getOrdersLastRow(ordersSheet);
    Logger.log('Orders sheet last row: ' + lastRow);
    
    if (lastRow < 7) {
      return { success: true, orders: [] };
    }
    
    const dataRange = ordersSheet.getRange(7, 1, lastRow - 6, 42);
    const data = dataRange.getValues();
    Logger.log('Retrieved ' + data.length + ' rows');
    
    // Get client regions (Column D = name, Column Q = region)
    const clientRegionMap = {};
    const clientDistanceMap = {};
    
    if (clientsSheet) {
      try {
        const clientLastRow = clientsSheet.getLastRow();
        Logger.log('Clients sheet last row: ' + clientLastRow);
        
        if (clientLastRow >= 4) {
          const clientNames = clientsSheet.getRange(4, 4, clientLastRow - 3, 1).getValues();
          const regions = clientsSheet.getRange(4, 17, clientLastRow - 3, 1).getValues();
          const distances = clientsSheet.getRange(4, 19, clientLastRow - 3, 1).getValues();
          
          for (let i = 0; i < clientNames.length; i++) {
            const name = clientNames[i][0];
            if (name && name.toString().trim() !== '') {
              const trimmedName = name.toString().trim();
              const region = regions[i] && regions[i][0] ? regions[i][0].toString().trim() : '';
              const dist = distances[i] && distances[i][0] && !isNaN(distances[i][0]) ? parseFloat(distances[i][0]) : null;
              
              clientRegionMap[trimmedName] = region;
              if (dist !== null && dist > 0) {
                clientDistanceMap[trimmedName] = dist;
              }
            }
          }
          
          Logger.log('Client region map: ' + Object.keys(clientRegionMap).length);
          Logger.log('Client distance map (Column S): ' + Object.keys(clientDistanceMap).length);
        }
      } catch (e) {
        Logger.log('Error reading clients: ' + e.toString());
      }
    }
    
    // Get region distances from المناطق
    const distanceMap = {};
    const baseLocation = 'ابو رواش';
    
    if (regionsSheet) {
      try {
        const lastRow = regionsSheet.getLastRow();
        const lastCol = regionsSheet.getLastColumn();
        
        if (lastRow >= 2 && lastCol >= 2) {
          const regionNamesCol = regionsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
          const regionNamesRow = regionsSheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
          const distanceData = regionsSheet.getRange(2, 2, lastRow - 1, lastCol - 1).getValues();
          
          let baseRowIndex = -1;
          let baseColIndex = -1;
          
          for (let i = 0; i < regionNamesCol.length; i++) {
            if (regionNamesCol[i][0] && regionNamesCol[i][0].toString().trim() === baseLocation) {
              baseRowIndex = i;
              break;
            }
          }
          
          for (let j = 0; j < regionNamesRow.length; j++) {
            if (regionNamesRow[j] && regionNamesRow[j].toString().trim() === baseLocation) {
              baseColIndex = j;
              break;
            }
          }
          
          if (baseRowIndex !== -1) {
            for (let j = 0; j < regionNamesRow.length; j++) {
              const region = regionNamesRow[j] ? regionNamesRow[j].toString().trim() : '';
              if (region && distanceData[baseRowIndex] && distanceData[baseRowIndex][j]) {
                const dist = parseFloat(distanceData[baseRowIndex][j]);
                if (!isNaN(dist)) distanceMap[region] = dist;
              }
            }
          } else if (baseColIndex !== -1) {
            for (let i = 0; i < regionNamesCol.length; i++) {
              const region = regionNamesCol[i][0] ? regionNamesCol[i][0].toString().trim() : '';
              if (region && distanceData[i] && distanceData[i][baseColIndex]) {
                const dist = parseFloat(distanceData[i][baseColIndex]);
                if (!isNaN(dist)) distanceMap[region] = dist;
              }
            }
          }
          
          Logger.log('Region distance map: ' + Object.keys(distanceMap).length);
        }
      } catch (e) {
        Logger.log('Error reading المناطق: ' + e.toString());
      }
    }
    
    // Build orders
    const orders = [];
    
    for (let i = 0; i < data.length; i++) {
      try {
        const row = data[i];
        
        const orderId = row[3];
        if (!orderId) continue;
        
        const clientName = row[8] ? row[8].toString().trim() : '';
        if (!clientName) continue;
        
        const status = row[2] ? row[2].toString().trim() : 'Not Yet';
        const clientRegion = clientRegionMap[clientName] || '';
        
        // Priority: Column S > Region distance
        let distanceFromBase = clientDistanceMap[clientName] || 0;
        if (distanceFromBase === 0 && clientRegion) {
          distanceFromBase = distanceMap[clientRegion] || 0;
        }
        
        let invoiceDate = '';
        if (row[5] && row[5] instanceof Date) {
          invoiceDate = Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        
        let loadingDate = '';
        if (row[7] && row[7] instanceof Date) {
          loadingDate = Utilities.formatDate(row[7], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        
        const totalWeight = row[26] && !isNaN(row[26]) ? parseFloat(row[26]) : 0;
        
        let loNum = null;
        if (row[19]) {
          const parsed = parseInt(row[19]);
          if (!isNaN(parsed)) loNum = parsed;
        }
        
    orders.push({
          orderId: orderId.toString(),
          invoiceDate: invoiceDate,
          clientName: clientName,
          clientRegion: clientRegion,
          distanceFromBase: distanceFromBase,
          status: status,
          loadingDate: loadingDate,
          totalWeight: totalWeight,
          loadingOrderNumber: loNum,
          truck: row[20] ? row[20].toString() : '',
          driver: row[21] ? row[21].toString() : '',
          deliveryRep: row[22] ? row[22].toString() : '',
          dismissNumber: row[24] ? row[24].toString().trim() : '', // Column Y (index 24)
          isDelivered: status.toLowerCase() === 'delivered',
          isPaid: status.toLowerCase() === 'paid'
        });
      } catch (e) {
        Logger.log('Error row ' + (i + 7) + ': ' + e.toString());
      }
    }
    
    Logger.log('Total orders: ' + orders.length);
    if (orders.length > 0) {
      Logger.log('Sample: ' + orders[0].orderId + ' - ' + orders[0].clientName + ' (' + orders[0].clientRegion + ') - ' + orders[0].distanceFromBase + 'km');
    }
    
    return { success: true, orders: orders };
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// =============================================================================
// GET DELIVERY REPS FROM INDEX SHEET
// =============================================================================

function getDeliveryReps() {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const indexSheet = ss.getSheetByName('Index');
    
    if (!indexSheet) {
      Logger.log('Index sheet not found');
      return { success: true, reps: [] };
    }
    
    try {
      const repsData = indexSheet.getRange('Y3:Y12').getValues();
      const reps = [];
      
      repsData.forEach(row => {
        if (row[0] && row[0].toString().trim() !== '') {
          reps.push(row[0].toString().trim());
        }
      });
      
      Logger.log('Loaded delivery reps: ' + reps.length);
      return { success: true, reps: reps };
      
    } catch (e) {
      Logger.log('Error reading Y3:Y12: ' + e.toString());
      return { success: true, reps: [] };
    }
    
  } catch (error) {
    Logger.log('Get delivery reps error: ' + error.toString());
    return { success: true, reps: [] };
  }
}

// =============================================================================
// GET DISTANCE MATRIX FROM REGIONS SHEET
// =============================================================================

function getDistanceMatrix() {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const regionsSheet = ss.getSheetByName('المناطق');
    
    if (!regionsSheet) {
      Logger.log('المناطق sheet not found');
      return { success: true, matrix: {} };
    }
    
    const lastRow = regionsSheet.getLastRow();
    const lastCol = regionsSheet.getLastColumn();
    
    if (lastRow < 2 || lastCol < 2) {
      return { success: true, matrix: {} };
    }
    
    try {
      const regionNamesCol = regionsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      const regionNamesRow = regionsSheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
      const distanceData = regionsSheet.getRange(2, 2, lastRow - 1, lastCol - 1).getValues();
      
      const matrix = {};
      
      for (let i = 0; i < regionNamesCol.length; i++) {
        const region1 = regionNamesCol[i][0] ? regionNamesCol[i][0].toString().trim() : '';
        if (!region1) continue;
        
        for (let j = 0; j < regionNamesRow.length; j++) {
          const region2 = regionNamesRow[j] ? regionNamesRow[j].toString().trim() : '';
          if (!region2) continue;
          
          const distance = distanceData[i] && distanceData[i][j];
          if (distance && !isNaN(distance) && distance > 0) {
            matrix[`${region1}-${region2}`] = parseFloat(distance);
          }
        }
      }
      
      Logger.log('Distance matrix entries: ' + Object.keys(matrix).length);
      return { success: true, matrix: matrix };
      
    } catch (e) {
      Logger.log('Error reading distance matrix: ' + e.toString());
      return { success: true, matrix: {} };
    }
    
  } catch (error) {
    Logger.log('Get distance matrix error: ' + error.toString());
    return { success: true, matrix: {} };
  }
}

// =============================================================================
// GET ORDER TYPE DISMISS/RETURN RULES FROM INDEX SHEET
// =============================================================================

function getOrderTypeRules() {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const indexSheet = ss.getSheetByName('Index');
    
    if (!indexSheet) {
      Logger.log('Index sheet not found');
      return { success: true, rules: {} };
    }
    
    try {
      // Read O3:O12 (Order Types), Q3:Q12 (Dismiss ref), R3:R12 (Return ref)
      const orderTypes = indexSheet.getRange('O3:O12').getValues();
      const dismissRefs = indexSheet.getRange('Q3:Q12').getValues();
      const returnRefs = indexSheet.getRange('R3:R12').getValues();
      
      const rules = {};
      
      for (let i = 0; i < orderTypes.length; i++) {
        const orderType = orderTypes[i][0] ? orderTypes[i][0].toString().trim() : '';
        if (orderType) {
          const dismissRef = dismissRefs[i][0];
          const returnRef = returnRefs[i][0];
          
          rules[orderType] = {
            needsDismiss: dismissRef === 1 || dismissRef === '1',
            needsReturn: returnRef === 1 || returnRef === '1'
          };
        }
      }
      
      Logger.log('Loaded order type rules: ' + Object.keys(rules).length);
      return { success: true, rules: rules };
      
    } catch (e) {
      Logger.log('Error reading order type rules: ' + e.toString());
      return { success: true, rules: {} };
    }
    
  } catch (error) {
    Logger.log('Get order type rules error: ' + error.toString());
    return { success: true, rules: {} };
  }
}
// =============================================================================
// SAVE DISTRIBUTION PLAN
// =============================================================================

// =============================================================================
// PART 3: BACKEND - SAVE WITH STATUS UPDATE AND RESET
// =============================================================================

function saveDistributionPlan(loadingOrdersData, ordersToReset, loadingDateChanges, invoiceDateChanges) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = ss.getSheetByName('Orders');
    
    if (!ordersSheet) {
      return { success: false, message: 'Orders sheet not found' };
    }
    
    Logger.log('Saving ' + loadingOrdersData.length + ' loading orders');
    Logger.log('Resetting ' + (ordersToReset ? ordersToReset.length : 0) + ' orders');
    Logger.log('Updating ' + (loadingDateChanges ? loadingDateChanges.length : 0) + ' loading dates');
    Logger.log('Updating ' + (invoiceDateChanges ? invoiceDateChanges.length : 0) + ' invoice dates');
    
    let successCount = 0;
    let resetCount = 0;
    let dateUpdateCount = 0;
    
    // Update loading orders with new status
    for (let lo of loadingOrdersData) {
      Logger.log('Processing LO #' + lo.loadingOrderNumber + ' - Status: ' + lo.status);
      
      for (let orderId of lo.orderIds) {
        // Get the loading date from the first order in this loading order
        // All orders in same LO should have same loading date
        const loadingDate = lo.loadingDate || new Date();
        const dateObj = new Date(loadingDate);
        dateObj.setHours(0, 0, 0, 0);
        
        const result = updateOrderDistributionInfo(
          ordersSheet,
          orderId,
          dateObj,
          lo.loadingOrderNumber,
          lo.truck,
          lo.driver,
          lo.deliveryRep || '',
          lo.status
        );
        
        if (result.success) {
          successCount++;
        }
      }
    }
    
    // Reset removed orders
    if (ordersToReset && ordersToReset.length > 0) {
      for (let orderId of ordersToReset) {
        const result = resetOrderDistributionInfo(ordersSheet, orderId);
        if (result.success) {
          resetCount++;
        }
      }
    }
    
    // Update loading dates for available orders (from left panel edits)
    if (loadingDateChanges && loadingDateChanges.length > 0) {
      for (let change of loadingDateChanges) {
        try {
          const row = change.rowIndex;
          const newDate = change.loadingDate ? new Date(change.loadingDate) : '';
          if (newDate) {
            newDate.setHours(0, 0, 0, 0);
          }
          ordersSheet.getRange(row, 8).setValue(newDate); // Column H = Loading Date
          dateUpdateCount++;
        } catch (e) {
          Logger.log('Error updating loading date for order ' + change.orderId + ': ' + e.toString());
        }
      }
    }

// Update invoice dates
    let invoiceUpdateCount = 0;
    if (invoiceDateChanges && invoiceDateChanges.length > 0) {
      Logger.log('Processing ' + invoiceDateChanges.length + ' invoice date changes');
      
      const lastRow = getOrdersLastRow(ordersSheet);
      if (lastRow >= 7) {
        const orderIds = ordersSheet.getRange(7, 4, lastRow - 6, 1).getValues();
        
        for (let change of invoiceDateChanges) {
          try {
            Logger.log('Looking for order ID: ' + change.orderId + ' (type: ' + typeof change.orderId + ')');
            
          // Find order row
            let rowIndex = -1;
            for (let i = 0; i < orderIds.length; i++) {
              const sheetOrderId = orderIds[i][0];
              if (sheetOrderId) {
                // Try multiple comparison methods
                const sheetIdStr = sheetOrderId.toString().trim();
                const changeIdStr = change.orderId.toString().trim();
                const sheetIdNum = Number(sheetOrderId);
                const changeIdNum = Number(change.orderId);
                
                Logger.log('Comparing sheet: "' + sheetIdStr + '" (num:' + sheetIdNum + ') vs change: "' + changeIdStr + '" (num:' + changeIdNum + ')');
                
                // Compare as strings OR as numbers
                if (sheetIdStr === changeIdStr || sheetIdNum === changeIdNum) {
                  rowIndex = i + 7;
                  Logger.log('MATCH FOUND at row ' + rowIndex);
                  break;
                }
              }
            }
            
            if (rowIndex !== -1) {
              // Update Column F (6) - Invoice Date
              if (change.invoiceDate) {
                const invoiceDate = new Date(change.invoiceDate);
                invoiceDate.setHours(0, 0, 0, 0);
                ordersSheet.getRange(rowIndex, 6).setValue(invoiceDate);
                Logger.log('Updated invoice date for order ' + change.orderId + ' at row ' + rowIndex);
              } else {
                ordersSheet.getRange(rowIndex, 6).clearContent();
                Logger.log('Cleared invoice date for order ' + change.orderId + ' at row ' + rowIndex);
              }
              invoiceUpdateCount++;
            } else {
              Logger.log('WARNING: Order ID ' + change.orderId + ' NOT FOUND in sheet');
            }
          } catch (e) {
            Logger.log('Error updating invoice date for order ' + change.orderId + ': ' + e.toString());
          }
        }
      }
    }

    Logger.log('Save complete: ' + successCount + ' orders updated, ' + resetCount + ' orders reset, ' + dateUpdateCount + ' loading dates updated, ' + invoiceUpdateCount + ' invoice dates updated');
    
    // Save dismiss numbers to "صرف و ارتجاع" sheet
    const dismissRecords = [];
    
    for (let lo of loadingOrdersData) {
      // Calculate dismiss number for this LO
      const loadingDate = lo.loadingDate || new Date();
      const dateObj = new Date(loadingDate);
      dateObj.setHours(0, 0, 0, 0);
      
      const month = String(dateObj.getMonth() + 1).padStart(2, '0');
      const day = String(dateObj.getDate()).padStart(2, '0');
      const loNum = String(lo.loadingOrderNumber).padStart(2, '0');
      
      const dismissNum = month + day + loNum;
      
      if (dismissNum) {
        dismissRecords.push({
          dismissNumber: dismissNum,
          loadingOrderNumber: lo.loadingOrderNumber,
          loadingDate: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        });
      }
    }
    
    if (dismissRecords.length > 0) {
      Logger.log('Saving ' + dismissRecords.length + ' dismiss records');
      const dismissResult = saveDismissReturnRecords(dismissRecords);
      if (!dismissResult.success) {
        Logger.log('Dismiss save warning: ' + dismissResult.message);
      } else {
        Logger.log('Dismiss records saved successfully');
      }
    }
    
    return { 
      success: true, 
      message: successCount + ' orders updated, ' + resetCount + ' reset, ' + dateUpdateCount + ' loading dates updated, ' + invoiceUpdateCount + ' invoice dates updated' 
    };
    
  } catch (error) {
    Logger.log('Save error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// =============================================================================
// SAVE DISMISS NUMBERS TO "صرف و ارتجاع" SHEET (FULLY OPTIMIZED)
// =============================================================================

function saveDismissReturnRecords(dismissRecords) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    let dismissSheet = ss.getSheetByName('صرف و ارتجاع');
    
    // Create sheet if it doesn't exist
    if (!dismissSheet) {
      dismissSheet = ss.insertSheet('صرف و ارتجاع');
      
      // Set headers in Row 11
      dismissSheet.getRange(11, 1, 1, 5).setValues([[
        'DISMISS NUMBER',
        'Truck, Driver, Delivery Rep',
        'Order IDs',
        'Loading Date',
        'Return Number'
      ]]);
      
      dismissSheet.getRange(11, 1, 1, 5).setFontWeight('bold');
      dismissSheet.setFrozenRows(11);
      
      // Row 12 is left empty as formula template
      Logger.log('Sheet created with headers in row 11, row 12 reserved for formulas');
    }
    
    if (!dismissRecords || dismissRecords.length === 0) {
      return { success: true, message: 'No dismiss records to save' };
    }
    
    // Sort records by dismiss number ascending
    const sortedRecords = dismissRecords.slice().sort((a, b) => {
      const numA = parseInt(a.dismissNumber) || 0;
      const numB = parseInt(b.dismissNumber) || 0;
      return numA - numB;
    });
    
    Logger.log('Saving ' + sortedRecords.length + ' dismiss records');
    
    const sheetLastRow = dismissSheet.getLastRow();
    let lastDataRow = 12; // Start at 12 (template row)
    
    // OPTIMIZED: Read Column A in ONE batch (starting from row 13)
    if (sheetLastRow >= 13) {
      const columnAData = dismissSheet.getRange(13, 1, sheetLastRow - 12, 1).getValues();
      
      // Scan backward to find last non-empty cell (faster)
      for (let i = columnAData.length - 1; i >= 0; i--) {
        const cellValue = columnAData[i][0];
        if (cellValue && cellValue.toString().trim() !== '') {
          lastDataRow = 13 + i;
          break;
        }
      }
    }
    
    Logger.log('Last data row in Column A: ' + lastDataRow);
    
    // OPTIMIZED: Build existing dismiss numbers from batch data
    const existingDismissNumbers = [];
    
    if (lastDataRow >= 13) {
      const existingData = dismissSheet.getRange(13, 1, lastDataRow - 12, 1).getValues();
      for (let i = 0; i < existingData.length; i++) {
        const dismissNum = existingData[i][0];
        if (dismissNum && dismissNum.toString().trim() !== '') {
          existingDismissNumbers.push(dismissNum.toString().trim());
        }
      }
    }
    
    Logger.log('Existing dismiss numbers: ' + existingDismissNumbers.length);
    
    let insertedCount = 0;
    
    // Insert each record in sequence
    for (let record of sortedRecords) {
      const dismissNum = record.dismissNumber;
      
      // Skip if already exists
      if (existingDismissNumbers.includes(dismissNum)) {
        Logger.log('Skipping duplicate: ' + dismissNum);
        continue;
      }
      
      const newNum = parseInt(dismissNum) || 0;
      let targetRow = 13; // Default to first data row
      let needsInsert = false;
      
      if (existingDismissNumbers.length === 0) {
        // No existing data - use row 13
        targetRow = 13;
        needsInsert = false;
      } else {
        // Find correct position
        let insertedInMiddle = false;
        
        for (let i = 0; i < existingDismissNumbers.length; i++) {
          const existingNum = existingDismissNumbers[i];
          const currentNum = parseInt(existingNum) || 0;
          
          if (newNum < currentNum) {
            // Insert BEFORE this number (in middle of sequence)
            targetRow = 13 + i;
            needsInsert = true;
            insertedInMiddle = true;
            Logger.log(dismissNum + ' goes before ' + existingNum + ' - needs insert at row ' + targetRow);
            break;
          }
        }
        
        if (!insertedInMiddle) {
          // Goes at end - check if empty row exists
          targetRow = lastDataRow + 1;
          
          if (targetRow <= sheetLastRow) {
            // Empty row exists - use it (no insert needed)
            needsInsert = false;
            Logger.log(dismissNum + ' goes at end - using existing empty row ' + targetRow);
          } else {
            // No empty row - must insert
            needsInsert = true;
            Logger.log(dismissNum + ' goes at end - no empty row, needs insert at row ' + targetRow);
          }
        }
      }
      
      // Insert row if needed
      if (needsInsert) {
        Logger.log('Inserting new row at ' + targetRow);
        dismissSheet.insertRowAfter(targetRow - 1);
        
        // Copy formulas from Row 12 (template row)
        const lastCol = dismissSheet.getLastColumn();
        if (lastCol > 0) {
          dismissSheet.getRange(12, 1, 1, lastCol)
            .copyTo(dismissSheet.getRange(targetRow, 1, 1, lastCol), 
                    SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        }
        Logger.log('Copied formulas from row 12 to row ' + targetRow);
      } else {
        Logger.log('Using existing empty row ' + targetRow + ' (no insert needed)');
      }
      
      // Write dismiss number to Column A
      dismissSheet.getRange(targetRow, 1).setValue(dismissNum);
      
      // Update tracking
      existingDismissNumbers.splice(targetRow - 13, 0, dismissNum);
      
      // Update lastDataRow if we wrote beyond it
      if (targetRow > lastDataRow) {
        lastDataRow = targetRow;
      }
      
      insertedCount++;
      Logger.log('Saved successfully: ' + dismissNum + ' at row ' + targetRow);
    }
    
    return { success: true, message: 'Saved ' + insertedCount + ' records successfully' };
    
  } catch (error) {
    Logger.log('Save dismiss error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}


function updateOrderDistributionInfo(ordersSheet, orderId, loadingDate, loadingOrderNumber, truck, driver, deliveryRep, status) {
  try {
    const lastRow = getOrdersLastRow(ordersSheet);
    const orderIds = ordersSheet.getRange(7, 4, lastRow - 6, 1).getValues(); // Column D
    
    let rowIndex = -1;
    for (let i = 0; i < orderIds.length; i++) {
      if (orderIds[i][0] && orderIds[i][0].toString() === orderId.toString()) {
        rowIndex = i + 7;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Order not found' };
    }
    
    // Update all fields including status
    ordersSheet.getRange(rowIndex, 3).setValue(status); // Column C - Status
    ordersSheet.getRange(rowIndex, 8).setValue(loadingDate); // Column H - Loading Date
    ordersSheet.getRange(rowIndex, 20).setValue(loadingOrderNumber); // Column T - Loading Order
    ordersSheet.getRange(rowIndex, 21).setValue(truck); // Column U - Truck
    ordersSheet.getRange(rowIndex, 22).setValue(driver); // Column V - Driver
    ordersSheet.getRange(rowIndex, 23).setValue(deliveryRep); // Column W - Delivery Rep
    
    return { success: true };
    
  } catch (error) {
    Logger.log('Update order error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function resetOrderDistributionInfo(ordersSheet, orderId) {
  try {
    const lastRow = getOrdersLastRow(ordersSheet);
    const orderIds = ordersSheet.getRange(7, 4, lastRow - 6, 1).getValues(); // Column D
    
    let rowIndex = -1;
    for (let i = 0; i < orderIds.length; i++) {
      if (orderIds[i][0] && orderIds[i][0].toString() === orderId.toString()) {
        rowIndex = i + 7;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Order not found' };
    }
    
    // Reset status and clear delivery info
    ordersSheet.getRange(rowIndex, 3).setValue(''); // Column C - Status (blank)
    ordersSheet.getRange(rowIndex, 8).setValue(''); // Column H - Loading Date
    ordersSheet.getRange(rowIndex, 20).setValue(''); // Column T - Loading Order
    ordersSheet.getRange(rowIndex, 21).setValue(''); // Column U - Truck
    ordersSheet.getRange(rowIndex, 22).setValue(''); // Column V - Driver
    ordersSheet.getRange(rowIndex, 23).setValue(''); // Column W - Delivery Rep
    
    Logger.log('Reset order: ' + orderId);
    return { success: true };
    
  } catch (error) {
    Logger.log('Reset order error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// =============================================================================
// END OF DISTRIBUTION PLANNING BACKEND FUNCTIONS
// =============================================================================

// =============================================================================
// ACCOUNTING MODULE - INVOICE HELPERS
// =============================================================================

/**
 * Looks up product details from the Products sheet.
 * @param {string} productName The full name of the product.
 * @returns {Object|null} An object containing vatRate and isPrivateLabel, or null.
 */
function getProductDetails(productName) {
  try {
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const productsSheet = salesSpreadsheet.getSheetByName('Products');
    
    if (!productsSheet || !productName) return null;
    
    const lastRow = productsSheet.getLastRow();
    if (lastRow < 3) return null;
    
    // Get data from A3:I (Column A: Name, F: VAT, I: Private Label)
    // Note: Reading up to Column I (9 columns)
    const dataRange = productsSheet.getRange(3, 1, lastRow - 2, 9);
    const data = dataRange.getValues();
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const name = row[0] ? row[0].toString().trim() : '';
      
      if (name === productName) {
        // Column F (index 5): VAT Rate
        let vatRate = parseFloat(row[5]) || 0;
        
        // Column I (index 8): Private Label (checkbox)
        // User specified Column I is the flag
        const isPrivateLabel = !!row[8]; 
        
        return {
          vatRate: vatRate, 
          isPrivateLabel: isPrivateLabel
        };
      }
    }
    return null;
  } catch (error) {
    Logger.log('Error in getProductDetails: ' + error.toString());
    return null;
  }
}

/**
 * Fetches valid contract discounts for a client on a specific date.
 * @param {string} clientName The name of the client.
 * @param {string} invoiceDate The invoice date (YYYY-MM-DD).
 * @returns {Object} A map of productName -> discountPercentage, or an empty object.
 */
function getInvoiceDiscounts(clientName, invoiceDate) {
  const discountMap = {};
  try {
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const discountsSheet = salesSpreadsheet.getSheetByName('Invoice Discounts');
    const clientsSheet = salesSpreadsheet.getSheetByName('Clients');
    
    if (!discountsSheet) {
      Logger.log('Invoice Discounts sheet not found.');
      return discountMap;
    }
    
    const dateObj = new Date(invoiceDate);
    dateObj.setHours(0, 0, 0, 0);
    const invoiceDateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    const lastRow = discountsSheet.getLastRow();
    if (lastRow < 3) return discountMap;

    // Get client name, dates, and product headers (Row 2, starting from J)
    const clientIds = discountsSheet.getRange(3, 2, lastRow - 2, 1).getValues(); // Column B (Client ID)
    const startDates = discountsSheet.getRange(3, 6, lastRow - 2, 1).getValues(); // Column F
    const endDates = discountsSheet.getRange(3, 8, lastRow - 2, 1).getValues();   // Column H
    
    const productHeaders = discountsSheet.getRange(2, 10, 1, discountsSheet.getLastColumn() - 9).getValues()[0]; // Row 2, starting from J
    const discountData = discountsSheet.getRange(3, 10, lastRow - 2, discountsSheet.getLastColumn() - 9).getValues(); // Data starting from J3
    
    // First get client ID from Clients sheet
    let clientId = '';
    if (clientsSheet) {
      const clientData = getClientData(clientsSheet, clientName);
      if (clientData) clientId = clientData.clientId;
    }
    
    Logger.log('Contract discount lookup - Client: ' + clientName + ', ID: ' + clientId);
    
    const clientRowIndex = clientIds.findIndex(row => 
      row[0] && row[0].toString().trim() === clientId
    );
    
    Logger.log('Contract discount row index: ' + clientRowIndex);
    
    if (clientRowIndex === -1) return discountMap;
    
    // Check date validity
    const startDate = startDates[clientRowIndex][0];
    const endDate = endDates[clientRowIndex][0];
    
    const startStr = startDate instanceof Date ? Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
    const endStr = endDate instanceof Date ? Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
    
    const isDateValid = startStr && endStr && invoiceDateStr >= startStr && invoiceDateStr <= endStr;

    if (isDateValid) {
      const rowDiscounts = discountData[clientRowIndex];
      rowDiscounts.forEach((discount, colIndex) => {
        const productName = productHeaders[colIndex] ? productHeaders[colIndex].toString().trim() : '';
        if (productName && discount) {
          // Store discount as a fraction (e.g., 5% becomes 0.05)
          const percentage = typeof discount === 'string' && discount.includes('%') 
                           ? parseFloat(discount.replace('%', '')) / 100
                           : parseFloat(discount) || 0;
          if (percentage > 0) {
            discountMap[productName] = percentage;
          }
        }
      });
    }
    
    return discountMap;
    
  } catch (error) {
    Logger.log('Error in getInvoiceDiscounts: ' + error.toString());
    return discountMap;
  }
}

/**
 * Finds the next sequential serial number for invoices or delivery notes.
 * @param {boolean} isTaxInvoice True for Tax Invoices (shared sequence AD/AE), false for Delivery Notes (separate sequence AC).
 * @returns {number} The next sequential serial number.
 */
function getNextInvoiceSerialNumber(isTaxInvoice) {
  try {
    const salesSpreadsheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = salesSpreadsheet.getSheetByName('Orders');

    if (!ordersSheet) return 1;

    const MAX_ROWS = 10000;
    const START_ROW = 7;
    const range = ordersSheet.getRange(START_ROW, 1, MAX_ROWS, ordersSheet.getLastColumn());
    const data = range.getValues();
    
    let maxSerial = 0;

    if (isTaxInvoice) {
      // Check Tax Invoice columns: AD (30) and AE (31) -- 1-indexed
      const TAWOOS_COL_INDEX = 30 - 1; // 29
      const PL_COL_INDEX = 31 - 1;     // 30
      
      data.forEach(row => {
        const tawoosSerial = parseFloat(row[TAWOOS_COL_INDEX]);
        const plSerial = parseFloat(row[PL_COL_INDEX]);
        
        if (!isNaN(tawoosSerial) && tawoosSerial > maxSerial) {
          maxSerial = tawoosSerial;
        }
        if (!isNaN(plSerial) && plSerial > maxSerial) {
          maxSerial = plSerial;
        }
      });
    } else {
      // Check Delivery Note column: AC (29) -- 1-indexed
      const DN_COL_INDEX = 29 - 1; // 28
      
      data.forEach(row => {
        const dnSerial = parseFloat(row[DN_COL_INDEX]);
        if (!isNaN(dnSerial) && dnSerial > maxSerial) {
          maxSerial = dnSerial;
        }
      });
    }
    
    return maxSerial + 1;

  } catch (error) {
    Logger.log('GET_SERIAL_ERROR: ' + error.toString());
    return 1;
  }
}

function previewInvoiceData(orderId, invoiceType, manualSerialNumber) {
  try {
    const calc = calculateAndPrepareInvoiceData(orderId, invoiceType, manualSerialNumber);
    
    if (!calc) return { success: false, message: "Backend returned null." };
    if (!calc.success) return { success: false, message: calc.message };
    
    // Add version prediction
    const nextVer = getNextInvoiceVersion(orderId, invoiceType);
    
    // Safely assign if data exists
    if (calc.data) {
        calc.data.VERSION = nextVer;
        calc.data.SERIAL_NUMBER = manualSerialNumber;
        calc.data.INVOICE_TYPE = invoiceType;
    }
    
    return { success: true, data: calc.data };
    
  } catch (e) {
    return { success: false, message: "Preview Error: " + e.toString() };
  }
}

/**
 * Detects if product has promotional discount for client
 * Uses the same 8-tier priority logic as price detection
 * @returns {Object} { hasPromotion, officialPrice, promotionalPrice, discountPct }
 */
function detectPromotionalDiscount(clientName, productName, invoiceDate, branchName, clientId, strategy) {
  Logger.log('*** PROMOTIONAL DISCOUNT FUNCTION CALLED - VERSION 2 ***');
  Logger.log('Inputs: Product=' + productName + ', Client=' + clientName + ', ClientID=' + clientId + ', Strategy=' + strategy);
  
  const result = { hasPromotion: false, officialPrice: 0, promotionalPrice: 0, discountPct: 0 };
  
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const pricesSheet = ss.getSheetByName('Prices');
    
    if (!pricesSheet) {
      Logger.log('ERROR: Prices sheet not found!');
      return result;
    }
    Logger.log('Prices sheet found');
    
    if (!clientId && !strategy) {
      Logger.log('ERROR: No client ID or strategy provided');
      return result;
    }
    if (!clientId && !strategy) return result;
    
    // Find product column in row 3 (headers)
    const headerRow = pricesSheet.getRange(3, 11, 1, pricesSheet.getLastColumn() - 10).getValues()[0];
    let productColIndex = -1;
    
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i] && headerRow[i].toString().trim() === productName) {
        productColIndex = 11 + i;
        Logger.log('Product found in row 3 at column: ' + productColIndex);
        break;
      }
    }
    
    if (productColIndex === -1) {
      Logger.log('ERROR: Product not found in row 3. Product name: "' + productName + '"');
      Logger.log('Row 3 headers sample: ' + JSON.stringify(headerRow.slice(0, 5)));
      return result;
    }
    
    // Get all price data starting from row 4
    const lastRow = pricesSheet.getLastRow();
    if (lastRow < 4) return result;
    
    const data = pricesSheet.getRange(4, 1, lastRow - 3, productColIndex).getValues();
    Logger.log('Reading price data from row 4 to row ' + lastRow);
    
    // Build price rows
    const priceRows = [];
    const invoiceDateObj = new Date(invoiceDate);
    const invoiceDateStr = Utilities.formatDate(invoiceDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // Skip if strategy column is empty
      
      const startDate = row[7];
      const endDate = row[9];
      
      if (!(startDate instanceof Date) || !(endDate instanceof Date)) continue;
      
      const startStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const endStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      if (invoiceDateStr < startStr || invoiceDateStr > endStr) continue;
      
      const price = parseFloat(row[productColIndex - 1]) || 0;
      if (price === 0) continue;
      
      priceRows.push({
        strategy: row[0] ? row[0].toString().trim() : '',      // Column A (index 0)
        clientName: row[1] ? row[1].toString().trim() : '',    // Column B (index 1)
        clientId: row[3] ? row[3].toString().trim() : '',      // Column D (index 3)
        allBranches: row[5] ? row[5].toString().trim() : '',   // Column F (index 5)
        priceType: row[6] ? row[6].toString().trim() : '',     // Column G (index 6)
        price: price
      });
    }
    
    Logger.log('Total valid price rows found: ' + priceRows.length);
    if (priceRows.length > 0) {
      Logger.log('All price rows for this product:');
      priceRows.forEach((r, i) => {
        Logger.log(`  Row ${i}: Strategy="${r.strategy}", ClientName="${r.clientName}", ClientID="${r.clientId}", AllBranches="${r.allBranches}", Type="${r.priceType}", Price=${r.price}`);
      });
    }
    
    Logger.log('Looking for client: "' + clientName + '", ID: "' + clientId + '", Strategy: "' + strategy + '"');
    if (priceRows.length > 0) {
      Logger.log('All price rows for this product:');
      priceRows.forEach((r, i) => {
        Logger.log(`  Row ${i}: Strategy="${r.strategy}", ClientName="${r.clientName}", ClientID="${r.clientId}", AllBranches="${r.allBranches}", Type="${r.priceType}", Price=${r.price}`);
      });
    }
    
    Logger.log('Looking for client: "' + clientName + '", ID: "' + clientId + '", Strategy: "' + strategy + '"');
    
    // Check each priority tier for BOTH promotional and official prices
    const priorities = [
      // Priority 1: Special + Client Name + Not All Branches
      { check: r => r.strategy === 'اسعار خاصة' && r.clientName.toLowerCase() === clientName.toLowerCase() && r.allBranches !== 'Yes' },
      // Priority 2: Special + Client ID + Not All Branches
      { check: r => r.strategy === 'اسعار خاصة' && r.clientId === clientId && r.allBranches !== 'Yes' },
      // Priority 3: Special + Client ID + All Branches
      { check: r => r.strategy === 'اسعار خاصة' && r.clientId === clientId && r.allBranches === 'Yes' },
      // Priority 4: (Skip - same as 3 for special prices)
      // Priority 5: Strategy + Not All Branches
      { check: r => r.strategy === strategy && r.allBranches !== 'Yes' },
      // Priority 6: (Same as 5)
      // Priority 7: Strategy + All Branches
      { check: r => r.strategy === strategy && r.allBranches === 'Yes' }
    ];
    
    // Stop at special prices if client has special pricing strategy
    const maxPriority = strategy === 'اسعار خاصة' ? 3 : priorities.length;
    
    for (let p = 0; p < maxPriority; p++) {
      const priority = priorities[p];
      if (!priority) continue;
      
      Logger.log(`Checking priority tier ${p + 1}`);
      
      // Find promotional price at this tier
      const promoRow = priceRows.find(r => priority.check(r) && r.priceType === 'عرض');
      // Find official price at this tier
      const officialRow = priceRows.find(r => priority.check(r) && r.priceType === 'رسمي');
      
      Logger.log(`  Promo found: ${!!promoRow}, Official found: ${!!officialRow}`);
      if (promoRow) Logger.log(`  Promo price: ${promoRow.price}`);
      if (officialRow) Logger.log(`  Official price: ${officialRow.price}`);
      
      if (promoRow && officialRow && promoRow.price < officialRow.price) {
        result.hasPromotion = true;
        result.officialPrice = officialRow.price;
        result.promotionalPrice = promoRow.price;
        result.discountPct = ((officialRow.price - promoRow.price) / officialRow.price) * 100;
        
        Logger.log('Promotional discount found at priority ' + (p + 1) + ': ' + result.discountPct.toFixed(2) + '%');
        return result;
      }
    }
    
  } catch(e) {
    Logger.log('detectPromotionalDiscount error: ' + e.toString());
  }
  
  return result;
}

function calculateAndPrepareInvoiceData(ORDER_ID, INVOICE_TYPE, SERIAL_NUMBER) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = ss.getSheetByName('Orders');
    const clientsSheet = ss.getSheetByName('Clients');
    const productsSheet = ss.getSheetByName('Products');
    // 1. Find Order
    // Fetch all columns including products (starting at column 43/AQ)
    const lastCol = ordersSheet.getLastColumn();
    const ordersData = ordersSheet.getRange(7, 1, ordersSheet.getLastRow()-6, lastCol).getValues();
    const orderRow = ordersData.find(r => r[3] && r[3].toString() === ORDER_ID.toString());
    if (!orderRow) return { success: false, message: 'Order not found' };
    
    // 2. Build Product Details Map
    const productDetailsMap = {};
    if (productsSheet && productsSheet.getLastRow() >= 3) {
        const pData = productsSheet.getRange(3, 1, productsSheet.getLastRow()-2, 17).getValues();
        pData.forEach(row => {
            if(row[0]) {
                productDetailsMap[row[0].toString().trim()] = {
                    barcode: row[7] || '',
                    nameAr: row[4] || row[0],
                    nameEn: row[16] || row[0],
                    unitAr: row[11] || '',
                    unitEn: row[15] || '',
                    vatRate: parseFloat(row[5]) || 0,
                    isPrivateLabel: !!row[8]
                };
            }
        });
    }

    // 3. Client Data
    const clientNameOrder = orderRow[8] ? orderRow[8].toString() : '';
    let clientInfo = { name: clientNameOrder, branch: '', address: '', vat: '' };
    // --- SMART CLIENT LOOKUP (Parent/Branch Logic) ---
    let clientIdForDiscounts = '';
    let pricingStrategy = '';
    let resolvedClientName = clientNameOrder; // Default to order name

    if (clientsSheet && clientsSheet.getLastRow() >= 4) {
        // Read up to Col AA (Client ID)
        const cData = clientsSheet.getRange(4, 1, clientsSheet.getLastRow()-3, 27).getValues();
        
        // 1. Try EXACT Match first
        let clientRow = cData.find(r => 
            (r[3] && r[3].toString() === clientNameOrder) || 
            (r[25] && r[25].toString() === clientNameOrder)
        );

        // 2. If no exact match, try PARENT/BRANCH Match
        // Look for a registered client name that is contained within the order's client name
        // e.g. Order: "Hyper One - Zayed", Database: "Hyper One" -> MATCH
        if (!clientRow) {
            // Sort by length (descending) to match "Hyper One Market" before "Hyper One" if both exist
            const potentialParents = cData.filter(r => r[3] && clientNameOrder.includes(r[3].toString()));
            potentialParents.sort((a, b) => b[3].toString().length - a[3].toString().length);
            
            if (potentialParents.length > 0) {
                clientRow = potentialParents[0];
                Logger.log(`Smart Lookup: Mapped branch "${clientNameOrder}" to parent "${clientRow[3]}"`);
            }
        }

        if (clientRow) {
            // Basic Info
            clientInfo.name = clientRow[20] || clientNameOrder; // Official Name
            clientInfo.branch = clientRow[5] || '';
            clientInfo.address = (clientRow[16] || '') + ' ' + (clientRow[17] || '');
            clientInfo.vat = clientRow[22] || '';
            
            // Critical: Capture the ID and Strategy of the PARENT
            clientIdForDiscounts = clientRow[26] ? clientRow[26].toString() : '';
            pricingStrategy = clientRow[8] ? clientRow[8].toString() : '';
            resolvedClientName = clientRow[3].toString(); // Use Parent Name for discount lookup
        }
    }
    
    // 4. Date & Discounts
    const rawDate = orderRow[5];
    let invoiceDateStr = rawDate instanceof Date ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : new Date().toISOString().slice(0,10);
    let displayDateStr = rawDate instanceof Date ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'dd MMM yyyy') : String(rawDate);
    let contractDiscounts = {};
    try { 
        if (typeof getInvoiceDiscounts === 'function') {
            // Pass the RESOLVED name (Parent) to find the contract
            contractDiscounts = getInvoiceDiscounts(resolvedClientName, invoiceDateStr);
        }
    } catch(e) {
        Logger.log('Contract discount error: ' + e.toString());
    }
    const branchName = clientInfo.branch || '';

    // 5. Process Products
    const productsMap = getProductColumnsMap(ordersSheet);
    const isTawoosInv = (INVOICE_TYPE === 'Tawoos');
    let productsDetail = [];
    let total14 = 0;
    let total0 = 0;
    
    for (const [prodName, cols] of Object.entries(productsMap)) {
       if (!cols || !cols.qty || !cols.bonus || !cols.price || !cols.discount) continue;
       
       const qty = parseFloat(orderRow[cols.qty - 1]) || 0;
       const bonus = parseFloat(orderRow[cols.bonus - 1]) || 0;
       const detectedPrice = parseFloat(orderRow[cols.price - 1]) || 0;
       const manualDiscountAmount = parseFloat(orderRow[cols.discount - 1]) || 0;

       if (qty + bonus <= 0) continue;

       const details = productDetailsMap[prodName];
       const vatRate = details ? details.vatRate : 0;
       const isPL = details ? details.isPrivateLabel : false;
       const barcode = details ? details.barcode : '';
       const nameAr = details ? details.nameAr : prodName;
       const nameEn = details ? details.nameEn : prodName;
       const unitAr = details ? details.unitAr : 'قطعة';
       const unitEn = details ? details.unitEn : 'Pcs';

       // Filter by invoice type
       if (isTawoosInv && isPL) continue;
       if (INVOICE_TYPE === 'PL' && !isPL) continue;

       // === DISCOUNT DETECTION ===
       let promotionalDiscountPct = 0;
       let officialPrice = 0;
       let promotionalPrice = 0;
       let hasPromotionalDiscount = false;
       
       try {
         // Pass resolvedClientName to ensure Name-based strategies work for the parent
         const priceData = detectPromotionalDiscount(resolvedClientName, prodName, invoiceDateStr, branchName, clientIdForDiscounts, pricingStrategy);
         if (priceData.hasPromotion) {
           hasPromotionalDiscount = true;
           officialPrice = priceData.officialPrice;
           promotionalPrice = priceData.promotionalPrice;
           promotionalDiscountPct = priceData.discountPct;
         }
       } catch(e) {
         Logger.log('Promotional detection error: ' + e.toString());
       }
       
       const contractDiscountPct = contractDiscounts[prodName] || 0;
       
       // === FINANCIAL CALCULATION [UPDATED] ===
       // 1. Base price uses official price if promo exists, else detected
       const basePrice = hasPromotionalDiscount ? officialPrice : detectedPrice;
       
       // 2. IMPORTANT: Calculate line subtotal on QTY ONLY (Bonus is free)
       const totalQty = qty + bonus;
       let lineSubtotal = qty * basePrice; // UPDATED: Changed from totalQty to qty

       // Apply promotional discount first (if exists)
       let promotionalDiscountAmount = 0;
       if (hasPromotionalDiscount) {
         promotionalDiscountAmount = lineSubtotal * (promotionalDiscountPct / 100);
         lineSubtotal = lineSubtotal - promotionalDiscountAmount;
       }
       
       // 3. Manual Discount
       const manualDiscountPerItem = manualDiscountAmount;
       // Total manual discount based on QTY ONLY (since bonus is free)
       const totalManualDiscount = manualDiscountPerItem * qty; 
       const manualDiscountPct = basePrice > 0 ? (manualDiscountPerItem / basePrice) * 100 : 0;
       
       // Apply contract discount
       const contractDiscountAmount = lineSubtotal * contractDiscountPct;
       lineSubtotal = lineSubtotal - contractDiscountAmount;
       
       // Apply manual discount
       lineSubtotal = lineSubtotal - totalManualDiscount;
       const lineFinal = lineSubtotal;

       // Accumulate by VAT rate
       if (vatRate > 0) total14 += lineFinal;
       else total0 += lineFinal;

       productsDetail.push({
         name: prodName,
         nameAr: nameAr,
         nameEn: nameEn,
         barcode: barcode,
         barcodeSVG: generateBarcodeSVG(barcode),
         unitAr: unitAr,
         unitEn: unitEn,
         qty: qty,
         bonus: bonus,
         totalQty: totalQty,
         unitPrice: hasPromotionalDiscount ? officialPrice : detectedPrice,
         
         hasPromotionalDiscount: hasPromotionalDiscount,
         promotionalDiscountAmount: promotionalDiscountAmount,
         officialPrice: officialPrice,
         promotionalPrice: promotionalPrice,
         promotionalDiscountPct: promotionalDiscountPct,
         
         hasContractDiscount: contractDiscountPct > 0,
         contractDiscountPct: contractDiscountPct * 100,
         contractDiscountAmount: contractDiscountAmount,
         
         hasManualDiscount: manualDiscountPerItem > 0,
         manualDiscountPct: manualDiscountPct,
         manualDiscountPerItem: manualDiscountPerItem,
         manualDiscountAmount: totalManualDiscount,
         
         subtotalBeforeDiscount: qty * detectedPrice, // Updated to match new logic
         finalAmount: lineFinal,
         vatRate: vatRate
       });
    }

    productsDetail.sort((a, b) => b.vatRate - a.vatRate);

    // Totals
    const totalOriginal = total14 + total0;
    
    // === VAT & DEDUCTION LOGIC [UPDATED] ===
    let vatValue = total14 * 0.14;
    
    // Only deduct 1% if VAT type indicates it
    const vatType = orderRow[9] ? orderRow[9].toString() : '';
    const shouldDeduct1Percent = vatType.includes('خصم') || vatType.includes('-1%') || vatType.includes('1%');
    let deduction = shouldDeduct1Percent ? totalOriginal * 0.01 : 0;
    
    // FORCE ZERO FOR DELIVERY NOTES
    if (INVOICE_TYPE === 'DeliveryNote') {
        vatValue = 0;
        deduction = 0;
    }
    
    const grandTotal = totalOriginal + vatValue - deduction;

    return {
        success: true,
        data: {
            INVOICE_DATE: displayDateStr,
            RAW_DATE: invoiceDateStr,
            SERIAL_NUMBER: SERIAL_NUMBER,
            PRODUCTS_DETAIL: productsDetail, 
            TOTAL_INVOICE_FINAL: grandTotal,
            TOTAL_ORIGINAL_AMT: totalOriginal,
            TOTAL_VAT_AMT: vatValue,
            TOTAL_DEDUCTION: deduction,
            ORIGINAL_14_AMT: total14,
            ORIGINAL_0_AMT: total0,
            PO_NUMBER: isTawoosInv ? orderRow[10] : (INVOICE_TYPE === 'PL' ? orderRow[11] : ''),
            CLIENT_DETAILS: clientInfo,
            CLIENT_NAME: clientNameOrder,
            CLIENT_NAME_INV: clientInfo.name,
            TAX_CARD: clientInfo.vat,
            ORDER_SUMMARY_TEXT: productsDetail.map(p => `${p.qty+p.bonus} ${p.name}`).join('\n')
        }
    };
  } catch (e) {
    Logger.log('CALC_ERROR: ' + e.toString());
    return { success: false, message: "Calc Error: " + e.toString() };
  }
}

/**
 * Updates the invoice serial numbers for a specific order.
 * @param {string} orderId The Order ID.
 * @param {Object} invoiceData Object containing {tawoosInvoice, plInvoice, deliveryNote}.
 * @returns {Object} Success status.
 */
function updateOrderInvoiceNumbers(orderId, invoiceData) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = ss.getSheetByName('Orders');
    
    if (!ordersSheet) return { success: false, message: 'Orders sheet not found' };
    
    const lastRow = getOrdersLastRow(ordersSheet);
    const orderIds = ordersSheet.getRange(7, 4, lastRow - 6, 1).getValues(); // Column D
    
    let rowIndex = -1;
    for (let i = 0; i < orderIds.length; i++) {
      if (orderIds[i][0] && orderIds[i][0].toString() === orderId.toString()) {
        rowIndex = i + 7;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, message: 'Order not found' };
    
    // Update specific columns
    // Delivery Note -> Column AC (29)
    if (invoiceData.deliveryNote !== undefined) {
        ordersSheet.getRange(rowIndex, 29).setValue(invoiceData.deliveryNote);
    }
    
    // Tawoos Invoice -> Column AD (30)
    if (invoiceData.tawoosInvoice !== undefined) {
        ordersSheet.getRange(rowIndex, 30).setValue(invoiceData.tawoosInvoice);
    }
    
    // PL Invoice -> Column AE (31)
    if (invoiceData.plInvoice !== undefined) {
        ordersSheet.getRange(rowIndex, 31).setValue(invoiceData.plInvoice);
    }
    
    return { success: true };
    
  } catch (error) {
    Logger.log('Error updating invoice numbers: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// =============================================================================
// ACCOUNTING MODULE - SERIAL TRACKING
// =============================================================================

/**
 * Gets the last used serial numbers for Tax Invoices and Delivery Notes.
 * Used to initialize the accounting dashboard trackers.
 * @returns {Object} { success: boolean, taxSerial: number, noteSerial: number }
 */
/**
 * Gets the MIN and MAX serial numbers for Tax Invoices and Delivery Notes.
 */
function getLastInvoiceSerialNumbers() {
  try {
    const ordersSheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID).getSheetByName('Orders');
    
    if (!ordersSheet) return { success: false, message: 'Sheet not found' };

    const MAX_ROWS = 5000;
    const lastRow = getOrdersLastRow(ordersSheet);
    
    // Default: Max is 0, Min is Infinity (to be lowered)
    let result = { 
        success: true, 
        taxSerial: 0, 
        noteSerial: 0,
        taxMin: Infinity, 
        noteMin: Infinity 
    };

    if (lastRow < 7) {
        // No data, reset Min to 0 so logic knows it's fresh
        result.taxMin = 0;
        result.noteMin = 0;
        return result;
    }

    // Get Columns AC (29), AD (30), AE (31)
    const data = ordersSheet.getRange(7, 29, lastRow - 6, 3).getValues();
    
    data.forEach(row => {
      // Delivery Note (AC -> index 0)
      const noteVal = parseInt(row[0]);
      if (!isNaN(noteVal) && noteVal > 0) {
          if (noteVal > result.noteSerial) result.noteSerial = noteVal;
          if (noteVal < result.noteMin) result.noteMin = noteVal;
      }

      // Tax Invoices (AD -> index 1, AE -> index 2)
      [row[1], row[2]].forEach(val => {
          const taxVal = parseInt(val);
          if (!isNaN(taxVal) && taxVal > 0) {
              if (taxVal > result.taxSerial) result.taxSerial = taxVal;
              if (taxVal < result.taxMin) result.taxMin = taxVal;
          }
      });
    });

    // If no data was found, reset Min to 0
    if (result.taxMin === Infinity) result.taxMin = 0;
    if (result.noteMin === Infinity) result.noteMin = 0;

    return result;

  } catch (error) {
    Logger.log('Error fetching serials: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// =============================================================================
// ACCOUNTING DASHBOARD DATA BATCH FETCHER
// =============================================================================



function getAccountingDashboardData() {
  Logger.log('=== getAccountingDashboardData START ===');
  let orders = [];
  let serials = { taxSerial: 0, noteSerial: 0, taxMin: 0, noteMin: 0 };
  let invoices = [];
  
  try {
    // 1. Fetch orders
    Logger.log('Step 1: Fetching orders...');
    try {
      // MODIFIED: Pass 'true' to load ALL history for Global Search
      const ordersResult = getOrders(true);
      Logger.log('Orders result type: ' + typeof ordersResult);
      Logger.log('Orders result: ' + JSON.stringify(ordersResult).substring(0, 200));
      
      if (ordersResult && ordersResult.orders) {
        orders = ordersResult.orders;
        Logger.log('Got ' + orders.length + ' orders');
      } else {
        Logger.log('WARNING: ordersResult.orders is missing');
      }
    } catch (orderError) {
      Logger.log('ERROR fetching orders: ' + orderError.toString());
      // Continue with empty orders
    }
    
    // 2. Fetch serials
    Logger.log('Step 2: Fetching serials...');
    try {
      const serialsResult = getLastInvoiceSerialNumbers();
      Logger.log('Serials result: ' + JSON.stringify(serialsResult));
      
      if (serialsResult) {
        serials = {
          taxSerial: serialsResult.taxSerial || 0,
          noteSerial: serialsResult.noteSerial || 0,
          taxMin: serialsResult.taxMin || 0,
          noteMin: serialsResult.noteMin || 0
        };
        Logger.log('Got serials: tax=' + serials.taxSerial + ', note=' + serials.noteSerial);
      }
    } catch (serialError) {
      Logger.log('ERROR fetching serials: ' + serialError.toString());
      // Continue with default serials
    }
    
    // 3. Fetch invoices
    Logger.log('Step 3: Fetching invoices...');
    try {
      const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
      const invoicesSheet = ss.getSheetByName('Invoices');
      if (invoicesSheet) {
        const lastRow = getInvoicesLastRow(invoicesSheet);
        Logger.log('Invoices sheet found, last data row: ' + lastRow);
        
        if (lastRow >= 4) {
          // Read up to AE (31 columns) to be safe
          const data = invoicesSheet.getRange(4, 1, lastRow - 3, 31).getValues();
          Logger.log('Read ' + data.length + ' invoice rows');
          
          for (let i = 0; i < data.length; i++) {
            const row = data[i];
            
            // NEW MAPPING:
            // Order ID: T (19)
            // Canceled: X (23)
            // Type: B (1)
            // Serial: D (3)
            // File URL: U (20)
            // Version: C (2)
            
            if (!row[19] || row[19] === '') continue; // Order ID
            if (row[23] === true) continue;           // Canceled
            
            invoices.push({
              orderId: row[19].toString(),      // T - Order ID
              invoiceType: row[1].toString(),   // B - Type
              serialNumber: row[3],             // D - Serial
              fileUrl: row[20].toString(),      // U - File URL
              fileName: '',                     // Name column removed/empty
              version: parseInt(row[2]) || 1    // C - Version
            });
          }
          
          Logger.log('Loaded ' + invoices.length + ' invoices');
        }
      } else {
        Logger.log('Invoices sheet not found');
      }
    } catch (invoiceError) {
      Logger.log('ERROR fetching invoices: ' + invoiceError.toString());
    }
    
    // 4. Build result
    const result = {
      orders: orders,
      serials: serials,
      invoices: invoices
    };
    
    Logger.log('=== RETURNING RESULT ===');
    Logger.log('Orders: ' + result.orders.length);
    Logger.log('Serials: ' + JSON.stringify(result.serials));
    Logger.log('Invoices: ' + result.invoices.length);
    
    return result;
    
  } catch (error) {
    Logger.log('=== CRITICAL ERROR ===');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    
    // NEVER return null - always return safe object
    return {
      orders: [],
      serials: { taxSerial: 0, noteSerial: 0, taxMin: 0, noteMin: 0 },
      invoices: []
    };
  }
}
/**
 * Generates a short date string (e.g., "24 Nov") from a Date object.
 */
function getShortDateString(dateObj) {
  if (!dateObj) return '';
  const d = new Date(dateObj);
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${d.getDate()} ${months[d.getMonth()]}`;
}

/**
 * Determines the next version number for a specific invoice by scanning the Invoices log.
 */
// DEPRECATED: Use getNextVersionForSerial() instead
// This function is kept for backward compatibility but redirects to new logic
function getNextInvoiceVersion(orderId, invoiceType) {
  // This function is no longer used - version is now per serial, not per order
  // Keeping for backward compatibility
  return 1;
}

/**
 * Saves the Invoice Log and converts the Frontend HTML directly to PDF.
 * @param {string} orderId
 * @param {Object} invoiceData (Optional/Legacy)
 * @param {Object} settings Settings object
 * @param {string} rawHtmlContent THE EXACT HTML FROM FRONTEND PREVIEW
 */
function commitInvoiceGeneration(orderId, invoiceData, settings, payload, userEmail) {
  try {
    const targetType = settings.targetType;
    const serial = settings.serialNumber;
    const fileType = settings.fileType || 'PNG';
    
    const calc = calculateAndPrepareInvoiceData(orderId, targetType, serial);
    if (!calc.success) return calc;
    const data = calc.data;
    
    // Get version based on serial number (not order ID)
    data.VERSION = getNextVersionForSerial(serial, targetType);
    data.SERIAL_NUMBER = serial;

    // --- File Generation ---
    const FOLDER_ID = '1lyfSYYZAIQTW6i8q59DXqoLjQ5e563Oc';
    const folder = DriveApp.getFolderById(FOLDER_ID);
    
    const safeClient = data.CLIENT_NAME_INV ? data.CLIENT_NAME_INV.replace(/[^a-zA-Z0-9\u0600-\u06FF ]/g, "").trim() : 'Unknown';
    const invoiceDateObj = data.RAW_DATE ? new Date(data.RAW_DATE) : new Date(data.INVOICE_DATE);
    const shortDate = getShortDateString(invoiceDateObj);
    const baseName = `${safeClient}_${serial}_v${data.VERSION}_${shortDate}`;
    
    let fileBlob, fileName;
    if (fileType === 'PNG') {
        const bytes = Utilities.base64Decode(payload);
        fileBlob = Utilities.newBlob(bytes, MimeType.PNG);
        fileName = `${baseName}.png`;
        fileBlob.setName(fileName);
    } else {
        fileBlob = Utilities.newBlob(payload, MimeType.HTML).getAs(MimeType.PDF);
        fileName = `${baseName}.pdf`;
    }
    
    const savedFile = folder.createFile(fileBlob);
    savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = savedFile.getUrl();

    // --- NEW STRUCTURE: Write to Invoices Sheet ---
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const invoicesSheet = ss.getSheetByName('Invoices');
    
    // Find next available row by scanning Column D (Serial)
    const targetRow = getNextInvoiceRow(invoicesSheet);

    const orderSummary = data.ORDER_SUMMARY_TEXT || '';
    const poNumber = targetType === 'PL' ? (data.PO_PL || '') : (data.PO_TAWOOS || '');
    
    let totalPromotionDisc = 0;
    let totalContractDisc = 0;
    let totalManualDisc = 0;
    
    if (data.PRODUCTS_DETAIL && data.PRODUCTS_DETAIL.length > 0) {
      data.PRODUCTS_DETAIL.forEach(p => {
        totalPromotionDisc += (p.promotionalDiscountAmount || 0);
        totalContractDisc += (p.contractDiscountAmount || 0);
        totalManualDisc += (p.manualDiscountAmount || 0);
      });
    }
    
    const vatAmount = targetType === 'DeliveryNote' ? 0 : (data.TOTAL_VAT_AMT || 0);
    
    let clientNameForInvoice = data.CLIENT_NAME_INV || '';
    if (settings.isCashSales) {
        clientNameForInvoice = settings.language === 'ar' ? 'بيع نقدي' : 'Cash Sales';
    }
    
    // NEW COLUMN MAPPING ARRAY (31 Columns: A to AE)
    const rowData = [
      '',                                // A (0) - Empty
      targetType,                        // B (1) - Invoice Type (was C)
      data.VERSION,                      // C (2) - Version (was D)
      serial,                            // D (3) - Serial (was E)
      data.INVOICE_DATE,                 // E (4) - Date (was F)
      data.CLIENT_NAME || '',            // F (5) - Client Name (was H)
      orderSummary,                      // G (6) - Summary (was I)
      data.TOTAL_INVOICE_FINAL || 0,     // H (7) - Total (was J)
      -Math.abs(totalPromotionDisc),     // I (8) - Promo (was K)
      -Math.abs(totalContractDisc),      // J (9) - Contract (was L)
      -Math.abs(totalManualDisc),        // K (10)- Manual (was M)
      data.ORIGINAL_14_AMT || 0,         // L (11)- Sub 14% (was N)
      data.ORIGINAL_0_AMT || 0,          // M (12)- Sub 0% (was O)
      data.TOTAL_ORIGINAL_AMT || 0,      // N (13)- Before Tax (was P)
      vatAmount,                         // O (14)- VAT (was Q)
      -(data.TOTAL_DEDUCTION || 0),      // P (15)- Deduct (was R)
      poNumber,                          // Q (16)- PO (was S)
      clientNameForInvoice,              // R (17)- Official Client (was T)
      data.TAX_CARD || '',               // S (18)- Tax Card (was U)
      orderId,                           // T (19)- Order ID (was V)
      fileUrl,                           // U (20)- File URL (was W)
      userEmail,                         // V (21)- User Email (was X)
      new Date(),                        // W (22)- Timestamp (was Y)
      false,                             // X (23)- Canceled (Moved from AD)
      '',                                // Y (24)- Reserved
      '',                                // Z (25)- Reserved
      '',                                // AA(26)- Reserved
      '',                                // AB(27)- Reserved
      '',                                // AC(28)- Reserved
      '',                                // AD(29)- New Column 1
      ''                                 // AE(30)- New Column 2
    ];
    
    invoicesSheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    
    // --- Update Orders Sheet ---
    const invoiceUpdate = {};
    if (targetType === 'Tawoos') invoiceUpdate.tawoosInvoice = serial;
    if (targetType === 'PL') invoiceUpdate.plInvoice = serial;
    if (targetType === 'DeliveryNote') invoiceUpdate.deliveryNote = serial;
    
    updateOrderInvoiceNumbers(orderId, invoiceUpdate);

    return { 
      success: true, 
      message: 'Saved & Generated Successfully',
      fileUrl: fileUrl,
      fileName: fileName
    };
  } catch (e) {
    Logger.log('COMMIT_ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Gets next available row in Invoices sheet by scanning Column E
 */
function getNextInvoiceRow(invoicesSheet) {
  // Scan Column D (Index 3) for Serials
  const serialColumn = invoicesSheet.getRange('D4:D').getValues();
  for (let i = 0; i < serialColumn.length; i++) {
    if (serialColumn[i][0] === '' || serialColumn[i][0] === null) {
      return i + 4;
    }
  }
  return invoicesSheet.getLastRow() + 1;
}

/**
 * Gets next version for a specific serial number (across all orders)
 * Version is per serial number, not per order
 */
function getNextVersionForSerial(serialNumber, invoiceType) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Invoices');
    
    if (!sheet || sheet.getLastRow() < 4) return 1;

    const lastRow = getInvoicesLastRow(sheet);
    if (lastRow < 4) return 1;
    
    // Read up to X (Canceled) at least. 24 columns is enough.
    const data = sheet.getRange(4, 1, lastRow - 3, 24).getValues(); 
    
    let maxVersion = 0;
    data.forEach(row => {
      // NEW MAPPING:
      // Canceled: X (23)
      // Serial: D (3)
      // Type: B (1)
      // Version: C (2)
      
      const canceled = row[23]; 
      const rowSerial = row[3]; 
      const rowType = row[1];   
      
      if (canceled === true) return;
      
      if (rowSerial.toString() === serialNumber.toString() && 
          rowType === invoiceType) {
        const version = parseInt(row[2]) || 0; 
        if (version > maxVersion) maxVersion = version;
      }
    });
    return maxVersion + 1;
    
  } catch (e) {
    Logger.log('VERSION_ERROR: ' + e.toString());
    return 1;
  }
}

/**
 * Gets actual last row with data in Invoices sheet by scanning Column E
 */
function getInvoicesLastRow(sheet) {
  // Scan Column D
  const serialColumn = sheet.getRange('D4:D').getValues();
  let lastRow = 3; 
  for (let i = 0; i < serialColumn.length; i++) {
    if (serialColumn[i][0] === '' || serialColumn[i][0] === null) {
      break;
    }
    lastRow = i + 4;
  }
  return lastRow;
}

/**
 * Gets actual last row with data in Invoices sheet by scanning Column E
 */
function getOrdersLastRow(sheet) {
  const serialColumn = sheet.getRange('D7:D').getValues();
  
  let lastRow = 6; // Start from row 3 (headers)
  for (let i = 0; i < serialColumn.length; i++) {
    if (serialColumn[i][0] === '' || serialColumn[i][0] === null) {
      break;
    }
    lastRow = i + 7;
  }
  return lastRow;
}
function getAccountingData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // ... existing orders code ...
    
    // Load invoices
    const invoicesSheet = ss.getSheetByName('Invoices');
    const invoicesData = invoicesSheet.getDataRange().getValues();
    const invoicesHeaders = invoicesData[0];
    
    const invoices = [];
    for (let i = 1; i < invoicesData.length; i++) {
      const row = invoicesData[i];
      
      // Skip soft-deleted (column AB = true)
      if (row[27] === true) continue;
      
      invoices.push({
        orderId: row[17],           // Column R
        invoiceType: row[5],        // Column F
        serialNumber: row[headers.indexOf('SERIAL_NUMBER')], // Find serial column
        fileUrl: row[19],           // Column T
        fileName: row[18],          // Column S
        invoiceDate: row[4],        // Column E
        total: row[7]              // Column H
      });
    }
    
    return {
      success: true,
      orders: ordersArray,
      serials: serialsData,
      invoices: invoices  // ADD THIS
    };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Generates SVG barcode using Code128B algorithm
 * @param {string} text - The barcode text/number
 * @returns {string} - SVG string
 */
function generateBarcodeSVG(text) {
  if (!text) return '';
  
  text = text.toString();
  
  // Code128B patterns (values 0-105)
  const patterns = [
    '11011001100', '11001101100', '11001100110', '10010011000', '10010001100',
    '10001001100', '10011001000', '10011000100', '10001100100', '11001001000',
    '11001000100', '11000100100', '10110011100', '10011011100', '10011001110',
    '10111001100', '10011101100', '10011100110', '11001110010', '11001011100',
    '11001001110', '11011100100', '11001110100', '11101101110', '11101001100',
    '11100101100', '11100100110', '11101100100', '11100110100', '11100110010',
    '11011011000', '11011000110', '11000110110', '10100011000', '10001011000',
    '10001000110', '10110001000', '10001101000', '10001100010', '11010001000',
    '11000101000', '11000100010', '10110111000', '10110001110', '10001101110',
    '10111011000', '10111000110', '10001110110', '11101110110', '11010001110',
    '11000101110', '11011101000', '11011100010', '11011101110', '11101011000',
    '11101000110', '11100010110', '11101101000', '11101100010', '11100011010',
    '11101111010', '11001000010', '11110001010', '10100110000', '10100001100',
    '10010110000', '10010000110', '10000101100', '10000100110', '10110010000',
    '10110000100', '10011010000', '10011000010', '10000110100', '10000110010',
    '11000010010', '11001010000', '11110111010', '11000010100', '10001111010',
    '10100111100', '10010111100', '10010011110', '10111100100', '10011110100',
    '10011110010', '11110100100', '11110010100', '11110010010', '11011011110',
    '11011110110', '11110110110', '10101111000', '10100011110', '10001011110',
    '10111101000', '10111100010', '11110101000', '11110100010', '10111011110',
    '10111101110', '11101011110', '11110101110', '11010000100', '11010010000',
    '11010011100', '1100011101011'
  ];
  
  // Start with Code B (104)
  let encoded = patterns[104];
  let checksum = 104;
  
  // Encode each character
  for (let i = 0; i < text.length; i++) {
    const charCode = text.charCodeAt(i);
    let value;
    
    if (charCode >= 32 && charCode <= 127) {
      value = charCode - 32;
    } else {
      value = 0; // Space for invalid chars
    }
    
    encoded += patterns[value];
    checksum += value * (i + 1);
  }
  
  // Add checksum
  checksum = checksum % 103;
  encoded += patterns[checksum];
  
  // Add stop pattern
  encoded += patterns[106];
  
  // Convert to SVG bars
  let bars = '';
  let x = 0;
  const barWidth = 0.7; // Reduced from 2 to 1
  const height = 23; // As you set
  
  for (let i = 0; i < encoded.length; i++) {
    if (encoded.charAt(i) === '1') {
      bars += `<rect x="${x}" y="0" width="${barWidth}" height="${height}" fill="black"/>`;
    }
    x += barWidth;
  }
  
  const totalWidth = x;
  const textPadding = 4; // Space between barcode and text
  
  return `<svg width="${totalWidth}" height="${height + 20}" viewBox="0 0 ${totalWidth} ${height + 20}" xmlns="http://www.w3.org/2000/svg">
    ${bars}
    <text x="${totalWidth/2}" y="${height + textPadding + 11}" text-anchor="middle" font-family="Arial, sans-serif" font-size="9" font-weight="500" fill="black" letter-spacing="4">${text}</text>
  </svg>`;
}

// =============================================================================
// BACKEND FUNCTION - Add to code.gs
// =============================================================================

/**
 * Get URL of saved invoice file
 * @param {string} orderId - Order ID
 * @param {string} invoiceType - 'Tawoos', 'PL', or 'DeliveryNote'
 * @returns {object} - {success: boolean, url: string}
 */
function getSavedInvoiceUrl(orderId, invoiceType) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ordersSheet = ss.getSheetByName('Orders');
    
    // Find order row
    const data = ordersSheet.getDataRange().getValues();
    const headers = data[0];
    const orderIdCol = headers.indexOf('orderId');
    
    let orderRow = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][orderIdCol].toString() === orderId.toString()) {
        orderRow = data[i];
        break;
      }
    }
    
    if (!orderRow) {
      return { success: false, message: 'Order not found' };
    }
    
    // Get file URL based on invoice type
    let fileUrlCol;
    if (invoiceType === 'Tawoos') {
      fileUrlCol = headers.indexOf('tawoosInvoiceFile');
    } else if (invoiceType === 'PL') {
      fileUrlCol = headers.indexOf('plInvoiceFile');
    } else if (invoiceType === 'DeliveryNote') {
      fileUrlCol = headers.indexOf('deliveryNoteFile');
    }
    
    if (fileUrlCol === -1) {
      return { success: false, message: 'File column not found' };
    }
    
    const fileUrl = orderRow[fileUrlCol];
    
    if (!fileUrl) {
      return { success: false, message: 'File not saved yet' };
    }
    
    return { success: true, url: fileUrl };
    
  } catch (error) {
    Logger.log('Error getting saved invoice URL: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Gets the last saved invoice URL from Invoices sheet
 * Matches by Order ID, Invoice Type, and Serial Number
 * Returns the most recent entry (highest row number) from Column V
 * @param {string} orderId - Order ID
 * @param {string} invoiceType - 'Tawoos', 'PL', or 'DeliveryNote'
 * @param {number} serialNumber - Invoice serial number
 * @returns {Object} {success: boolean, url: string, message: string}
 */
function getLastSavedInvoiceUrl(orderId, invoiceType, serialNumber) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const invoicesSheet = ss.getSheetByName('Invoices');
    
    if (!invoicesSheet || invoicesSheet.getLastRow() < 4) {
      return { success: false, message: 'No invoices found' };
    }
    
    // NEW STRUCTURE: Headers at row 3, data starts at row 4
    // Columns: V=Order ID (22), C=Type (3), E=Serial (5), W=File URL (23), D=Version (4), AD=Canceled (30)
    const lastRow = getInvoicesLastRow(invoicesSheet);
    if (lastRow < 4) return { success: false, message: 'No invoices found' };

    // Read up to X (24 columns) covers everything we need
    const data = invoicesSheet.getRange(4, 1, lastRow - 3, 24).getValues();
    
    let matchedUrl = null;
    let maxVersion = 0;

    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      // NEW MAPPING:
      // Order ID: T (19)
      // Type: B (1)
      // Serial: D (3)
      // URL: U (20)
      // Version: C (2)
      // Canceled: X (23)
      
      const rowOrderId = row[19];
      const rowType = row[1];
      const rowSerial = row[3];
      const rowUrl = row[20];
      const rowVersion = parseInt(row[2]) || 0;
      const canceled = row[23];
      
      if (canceled === true) continue;

      if (rowOrderId.toString() === orderId.toString() && 
          rowType === invoiceType && 
          parseFloat(rowSerial) === parseFloat(serialNumber) &&
          rowVersion > maxVersion) {
        matchedUrl = rowUrl;
        maxVersion = rowVersion;
      }
    }
    
    if (matchedUrl) {
      return { success: true, url: matchedUrl };
    } else {
      return { success: false, message: 'No matching invoice found' };
    }
    
  } catch (error) {
    Logger.log('Error in getLastSavedInvoiceUrl: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}


/**
 * Generates a PDF blob from HTML content and returns it as a Base64 string.
 * Used for the "Quick actions" feature in the frontend.
 */
function generatePdfBase64(htmlContent) {
  try {
    var blob = Utilities.newBlob(htmlContent, MimeType.HTML).getAs(MimeType.PDF);
    return Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    throw new Error("PDF Generation Failed: " + e.message);
  }
}
// =============================================================================
// WAREHOUSE MODULE - RELEASE NOTE MANAGEMENT
// =============================================================================
/**
 * Optimized Warehouse Data Fetcher
 * - Default: Returns Unsaved (Drafts) + Today's Records only.
 * - History Mode: If startDate/endDate provided, returns Unsaved + Records in range.
 * - Performance: Batch resolves client names and compresses product data.
 * @param {string} startDate - Optional 'YYYY-MM-DD' start of range
 * @param {string} endDate - Optional 'YYYY-MM-DD' end of range
 */
function getWarehouseReleaseRecords(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const dismissSheet = ss.getSheetByName('صرف و ارتجاع');
    
    if (!dismissSheet) return { success: false, message: 'Sheet "صرف و ارتجاع" not found' };

    // --- 1. Load Master Data & Helpers ---
    
    const lastCol = dismissSheet.getLastColumn();
    // Row 1: Product Names (Dashboard) - Start at Column S (Index 19)
    const productNamesRow1 = dismissSheet.getRange(1, 19, 1, lastCol - 18).getValues()[0];
    // Row 6: Stock Balance - Start at Column S (Index 19)
    const stockBalanceRow6 = dismissSheet.getRange(6, 19, 1, lastCol - 18).getValues()[0];
    // Row 10: Record Headers - Start at Column S (Index 19)
    const productNamesRow10 = dismissSheet.getRange(10, 19, 1, lastCol - 18).getValues()[0];

    // Build Master Product List & Stock Map
    const products = [];
    const stockMap = {};
    
    // Map stock from dashboard rows
    for (let i = 0; i < productNamesRow1.length; i++) {
      const name = productNamesRow1[i];
      if (name && String(name).trim()) {
        stockMap[String(name).trim()] = parseFloat(stockBalanceRow6[i]) || 0;
      }
    }

    // Build product structure from header row
    for (let i = 0; i < productNamesRow10.length; i += 7) {
      const name = productNamesRow10[i];
      if (name && String(name).trim()) {
        const cleanName = String(name).trim();
        products.push({
          name: cleanName,
          stockBalance: stockMap[cleanName] || 0,
          columnOffset: i
        });
      }
    }

    // --- 2. Pre-load Client Data for Batch Resolution ---
    const clientMap = getClientMapForWarehouse();

    // --- 3. Fetch & Process Records ---
    const lastRow = dismissSheet.getLastRow();
    if (lastRow < 13) return { success: true, records: [], products: products };

    // Bulk read all transaction data (Row 13 to End)
    const fullData = dismissSheet.getRange(13, 1, lastRow - 12, lastCol).getValues();
    
    const records = [];
    
    // Define "Today" string for default filtering
    const sessionTimeZone = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), sessionTimeZone, 'yyyy-MM-dd');

    for (let i = 0; i < fullData.length; i++) {
      const row = fullData[i];
      const dismissNumber = row[0]; // Col A
      
      // Skip empty rows
      if (!dismissNumber || String(dismissNumber).trim() === '') continue;

      const pdfUrl = row[7] ? String(row[7]) : ''; // Col H (PDF URL)
      const loadingDateVal = row[4]; // Col E (Loading Date)
      
      let loadingDateStr = '';
      if (loadingDateVal instanceof Date) {
        loadingDateStr = Utilities.formatDate(loadingDateVal, sessionTimeZone, 'yyyy-MM-dd');
      }

      // --- FILTERING LOGIC ---
      // 1. Always include Unsaved (Drafts) - Check if PDF URL is empty
      const isUnsaved = !pdfUrl || pdfUrl === '';
      
      // 2. Date Logic
      let isDateMatch = false;
      
      if (startDate) {
        // History Mode: Range Check
        // If only startDate provided, check exact match. If both, check range.
        if (endDate) {
           if (loadingDateStr >= startDate && loadingDateStr <= endDate) isDateMatch = true;
        } else {
           if (loadingDateStr === startDate) isDateMatch = true;
        }
      } else {
        // Default Mode: Today Only
        if (loadingDateStr === todayStr) isDateMatch = true;
      }

      // Final Decision: Keep if Unsaved OR Date Matches
      if (isUnsaved || isDateMatch) {
        
        // --- 4. Sparse Product Parsing (Optimization) ---
        const rowProducts = [];
        
        // Product data starts at array index 18 (Column S)
        for (let p = 0; p < products.length; p++) {
          const offset = products[p].columnOffset;
          const baseIdx = 18 + offset; 
          
          // Read values directly from bulk array
          const orderedQty = parseFloat(row[baseIdx]) || 0;
          const returnQty = parseFloat(row[baseIdx + 1]) || 0;
          const loadedQty = parseFloat(row[baseIdx + 2]) || 0;
          const pallets = row[baseIdx + 3] ? String(row[baseIdx + 3]) : '';
          const extraQty = parseFloat(row[baseIdx + 4]) || 0;
          const expectedReturn = parseFloat(row[baseIdx + 5]) || 0;
          
          // Only push product if it has activity
          if (orderedQty > 0 || loadedQty > 0 || returnQty > 0 || extraQty > 0 || pallets !== '') {
             rowProducts.push({
              name: products[p].name,
              stockBalance: products[p].stockBalance, // Needed for frontend validation
              orderedQty: orderedQty,
              returnQty: returnQty,
              loadedQty: loadedQty,
              pallets: pallets,
              extraQty: extraQty,
              expectedReturn: expectedReturn,
              columnOffset: offset
            });
          }
        }

        // --- 5. Batch Client Name Resolution ---
        let clientNames = '';
        const orderIds = row[3] ? String(row[3]) : ''; // Col D
        if (orderIds) {
           const ids = orderIds.split('-').map(s => s.trim());
           const names = new Set();
           ids.forEach(id => {
             const numericId = parseInt(id);
             if (!isNaN(numericId) && clientMap[numericId]) {
               names.add(clientMap[numericId]);
             }
           });
           clientNames = Array.from(names).join(', ');
        }

        // Build Record
        records.push({
          rowIndex: i + 13,
          dismissNumber: String(dismissNumber),
          status: row[1] ? String(row[1]) : '',
          truckDriverRep: row[2] ? String(row[2]) : '',
          orderIds: orderIds,
          loadingDate: loadingDateStr,
          invoiceDates: row[5] ? String(row[5]) : '',
          releaseSummary: row[6] ? String(row[6]) : '',
          pdfUrl: pdfUrl,
          returnNumber: row[8] ? String(row[8]) : '',
          recordStatus: row[9], // Boolean or null
          expectedReturns: row[10] ? String(row[10]) : '',
          returnPdfUrl: row[11] ? String(row[11]) : '',
          returnDate: row[12] ? String(row[12]) : '',
          loadingLabor: row[13] ? String(row[13]).split(',').map(l => l.trim()) : [],
          unloadingLabor: row[14] ? String(row[14]).split(',').map(l => l.trim()) : [],
          userEmail: row[15] ? String(row[15]) : '',
          timestamp: row[16] ? String(row[16]) : '',
          comments: row[17] ? String(row[17]) : '',
          products: rowProducts,
          clientNames: clientNames // Pre-resolved
        });
      }
    }

    Logger.log('Optimized fetch: Returned ' + records.length + ' records');
    return { success: true, records: records, products: products };

  } catch (error) {
    Logger.log('Error in getWarehouseReleaseRecords: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Helper to cache Order ID -> Client Name map
 * Scans Orders sheet quickly to build a lookup table.
 */
/**
 * Helper to cache Order ID -> Client Name map
 * Scans Orders sheet quickly to build a lookup table.
 */
function getClientMapForWarehouse() {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = ss.getSheetByName('Orders');
    if (!ordersSheet) return {};
    
    const lastRow = getOrdersLastRow(ordersSheet);
    if (lastRow < 7) return {};
    
    // Read Order ID (Col D/4) and Client Name (Col I/9)
    // We read columns D to I (6 columns total)
    const data = ordersSheet.getRange(7, 4, lastRow - 6, 6).getValues();
    const map = {};
    
    for (let i = 0; i < data.length; i++) {
      const oid = parseInt(data[i][0]); // Col D is index 0 in this range
      const client = data[i][5];        // Col I is index 5 in this range
      
      if (!isNaN(oid) && client) {
        map[oid] = String(client).trim();
      }
    }
    return map;
  } catch (e) {
    Logger.log('Client Map Error: ' + e.toString());
    return {};
  }
}

function saveWarehouseReleaseNote(recordData) {
  try {
    // --- SAFETY CHECK START ---
    if (!recordData || !recordData.products || !Array.isArray(recordData.products)) {
        return { success: false, message: 'بيانات غير صالحة (Invalid Payload)' };
    }

    // Check if we are trying to save "Zero" loaded quantities for everything
    const hasLoadedQty = recordData.products.some(p => p.loadedQty > 0);
    // It's possible to save 0 if ordered is 0, but usually we expect something loaded.
    // If you want strict enforcement:
    // if (!hasLoadedQty) return { success: false, message: 'لا توجد كميات محملة للحفظ' };
    // --- SAFETY CHECK END ---

    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    // ... rest of the function ...
    const dismissSheet = ss.getSheetByName('صرف و ارتجاع');
    
    batchData.forEach(data => {
       try {
         // Reuse existing logic per item (refactored or called directly)
         // Note: Calling saveWarehouseReleaseNote in a loop is okay for <10 items
         // For larger batches, you'd rewrite to use setValues() once.
         // For Phase 1, we simply wrap the logic:
         
         const res = saveWarehouseReleaseNote(data); // Reusing existing function
         results.push({ 
             dismissNumber: data.dismissNumber, 
             success: res.success, 
             pdfUrl: res.pdfUrl 
         });
         
       } catch (e) {
         results.push({ dismissNumber: data.dismissNumber, success: false, message: e.toString() });
       }
    });
    
    return results;
  } catch (e) {
    throw new Error("Batch failed: " + e.toString());
  }
}

/**
 * Get labor list from Index sheet
 */
function getLaborList() {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const indexSheet = ss.getSheetByName('Index');
    
    if (!indexSheet) {
      return { success: false, message: 'Index sheet not found' };
    }
    
    // Get range O14:O100
    const laborRange = indexSheet.getRange('O14:O100').getValues();
    const laborList = [];
    
    for (let i = 0; i < laborRange.length; i++) {
      const labor = laborRange[i][0];
      if (labor && labor.toString().trim() !== '') {
        laborList.push(labor.toString().trim());
      }
    }
    
    return { success: true, laborList: laborList };
    
  } catch (error) {
    Logger.log('Error in getLaborList: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Resolve client names from order IDs
 */
function resolveClientNames(orderIds, dismissNumber) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const ordersSheet = ss.getSheetByName('Orders');
    
    if (!ordersSheet) {
      return { success: false, message: 'Orders sheet not found' };
    }
    
    const clientNames = [];
    
    // Method 1: Lookup by Order IDs
    if (orderIds && orderIds.trim() !== '') {
      const orderIdArray = orderIds.split('-').map(id => id.trim());
      
      const lastRow = getOrdersLastRow(ordersSheet);
      const orderIdColumn = ordersSheet.getRange(7, 4, lastRow - 6, 1).getValues(); // Column D
      const clientColumn = ordersSheet.getRange(7, 9, lastRow - 6, 1).getValues(); // Column I
      
      for (let oid of orderIdArray) {
        const orderIdNum = parseInt(oid);
        if (isNaN(orderIdNum)) continue;
        
        for (let i = 0; i < orderIdColumn.length; i++) {
          if (orderIdColumn[i][0] && parseInt(orderIdColumn[i][0]) === orderIdNum) {
            const clientName = clientColumn[i][0] ? clientColumn[i][0].toString() : '';
            if (clientName && !clientNames.includes(clientName)) {
              clientNames.push(clientName);
            }
            break;
          }
        }
      }
    }
    
    // Method 2: Fallback - Lookup by Dismiss Number
    if (clientNames.length === 0 && dismissNumber) {
      const lastRow = getOrdersLastRow(ordersSheet);
      const dismissColumn = ordersSheet.getRange(7, 25, lastRow - 6, 1).getValues(); // Column Y
      const clientColumn = ordersSheet.getRange(7, 9, lastRow - 6, 1).getValues(); // Column I
      
      for (let i = 0; i < dismissColumn.length; i++) {
        if (dismissColumn[i][0] && dismissColumn[i][0].toString() === dismissNumber.toString()) {
          const clientName = clientColumn[i][0] ? clientColumn[i][0].toString() : '';
          if (clientName && !clientNames.includes(clientName)) {
            clientNames.push(clientName);
          }
        }
      }
    }
    
    return { success: true, clientNames: clientNames };
    
  } catch (error) {
    Logger.log('Error in resolveClientNames: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Save warehouse release note
 */
function saveWarehouseReleaseNote(recordData) {
  try {
    // --- SAFETY CHECK ---
    if (!recordData || !recordData.products || !Array.isArray(recordData.products)) {
        return { success: false, message: 'بيانات غير صالحة (Invalid Payload)' };
    }

    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const dismissSheet = ss.getSheetByName('صرف و ارتجاع');
    
    if (!dismissSheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const rowIndex = recordData.rowIndex;
    Logger.log('Saving warehouse data text at row ' + rowIndex);
    
    // 1. Write product data (loaded qty and pallets)
    for (let product of recordData.products) {
      if (product.loadedQty && product.loadedQty > 0) {
        const colIndex = 19 + product.columnOffset; // Column S = 19
        
        // Write Loaded Qty to offset +2 (Column U, AB, AI...)
        dismissSheet.getRange(rowIndex, colIndex + 2).setValue(product.loadedQty);
        
        // Write Pallets to offset +3 (Column V, AC, AJ...)
        if (product.pallets) {
          dismissSheet.getRange(rowIndex, colIndex + 3).setValue(product.pallets);
        }
      }
    }
    
    // 2. Generate and write release summary (Column G)
    const summary = generateReleaseSummary(recordData.products);
    dismissSheet.getRange(rowIndex, 7).setValue(summary);
    
    // 3. Write labor (Columns N, O)
    if (recordData.loadingLabor && recordData.loadingLabor.length > 0) {
      dismissSheet.getRange(rowIndex, 14).setValue(recordData.loadingLabor.join(', '));
    }
    if (recordData.unloadingLabor && recordData.unloadingLabor.length > 0) {
      dismissSheet.getRange(rowIndex, 15).setValue(recordData.unloadingLabor.join(', '));
    }
    
    // 4. Write comments (Column R)
    if (recordData.comments) {
      dismissSheet.getRange(rowIndex, 18).setValue(recordData.comments);
    }
    
    // 5. Write metadata (P, Q)
    dismissSheet.getRange(rowIndex, 16).setValue(recordData.userEmail);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    dismissSheet.getRange(rowIndex, 17).setValue(timestamp);
    
    // NOTE: Column H (URL) is NOT updated here. 
    // It will be updated by saveReleaseNoteImage after frontend generation.
    
    return { success: true, message: 'Data saved successfully' };
    
  } catch (error) {
    Logger.log('Error in saveWarehouseReleaseNote: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Generate release summary text for Column G
 */
function generateReleaseSummary(products) {
  const lines = [];
  
  for (let product of products) {
    if (product.loadedQty && product.loadedQty > 0) {
      let line = product.loadedQty + ' ' + product.name;
      
      if (product.extraQty && product.extraQty > 0) {
        line += ' [ ' + product.extraQty + ' زيادة ]';
      } else {
        line += ' [بدون زيادات]';
      }
      
      lines.push(line);
    }
  }
  
  return lines.join('\n');
}

/**
 * Generate release note PDF and upload to Drive
 */
function generateReleaseNotePDF(recordData) {
  try {
    // Get client names
    const clientResult = resolveClientNames(recordData.orderIds, recordData.dismissNumber);
    const clientNames = clientResult.success ? clientResult.clientNames.join(', ') : '';
    
    // Parse truck/driver/rep
    const parts = recordData.truckDriverRep.split('-').map(p => p.trim());
    const truck = parts[0] || '';
    const driver = parts[1] || '';
    const deliveryRep = parts[2] || '';
    
    // Build HTML for PDF
    let html = '<html dir="rtl"><head><meta charset="UTF-8">';
    html += '<style>';
    html += 'body { font-family: Tajawal, Arial; font-size: 12pt; direction: rtl; }';
    html += 'h1 { text-align: center; color: #1e3a5f; }';
    html += 'table { width: 100%; border-collapse: collapse; margin: 20px 0; }';
    html += 'th, td { border: 1px solid #ddd; padding: 8px; text-align: right; }';
    html += 'th { background-color: #1e3a5f; color: white; }';
    html += '.header { margin-bottom: 20px; }';
    html += '.info-row { margin: 10px 0; }';
    html += '</style></head><body>';
    
    html += '<h1>إذن صرف منتج تام</h1>';
    
    html += '<div class="header">';
    html += '<div class="info-row"><strong>رقم الصرف:</strong> ' + recordData.dismissNumber + '</div>';
    html += '<div class="info-row"><strong>تاريخ التحميل:</strong> ' + recordData.loadingDate + '</div>';
    html += '<div class="info-row"><strong>السيارة:</strong> ' + truck + '</div>';
    html += '<div class="info-row"><strong>السائق:</strong> ' + driver + '</div>';
    html += '<div class="info-row"><strong>مندوب التسليم:</strong> ' + deliveryRep + '</div>';
    html += '<div class="info-row"><strong>العملاء:</strong> ' + clientNames + '</div>';
    html += '</div>';
    
    html += '<h3>المنتجات المحملة</h3>';
    html += '<table>';
    html += '<tr><th>#</th><th>المنتج</th><th>الكمية المحملة</th><th>البالتات</th><th>الزيادة</th></tr>';
    
    let rowNum = 1;
    for (let product of recordData.products) {
      if (product.loadedQty && product.loadedQty > 0) {
        html += '<tr>';
        html += '<td>' + rowNum + '</td>';
        html += '<td>' + product.name + '</td>';
        html += '<td>' + product.loadedQty + '</td>';
        html += '<td>' + (product.pallets || '-') + '</td>';
        html += '<td>' + (product.extraQty || 0) + '</td>';
        html += '</tr>';
        rowNum++;
      }
    }
    
    html += '</table>';
    
    if (recordData.loadingLabor && recordData.loadingLabor.length > 0) {
      html += '<div class="info-row"><strong>عمال التحميل:</strong> ' + recordData.loadingLabor.join(', ') + '</div>';
    }
    if (recordData.unloadingLabor && recordData.unloadingLabor.length > 0) {
      html += '<div class="info-row"><strong>عمال التفريغ:</strong> ' + recordData.unloadingLabor.join(', ') + '</div>';
    }
    if (recordData.comments) {
      html += '<div class="info-row"><strong>ملاحظات:</strong> ' + recordData.comments + '</div>';
    }
    
    html += '<div class="info-row" style="margin-top: 30px;"><strong>تاريخ الطباعة:</strong> ' + new Date().toLocaleString('ar-EG') + '</div>';
    html += '<div class="info-row"><strong>المستخدم:</strong> ' + recordData.userEmail + '</div>';
    
    html += '</body></html>';
    
    // Convert HTML to PDF blob
    const blob = Utilities.newBlob(html, 'text/html', 'release_note.html');
    const pdfBlob = blob.getAs('application/pdf');
    
    // Upload to Google Drive
    const fileName = 'Release_' + recordData.dismissNumber + '_' + new Date().getTime() + '.pdf';
    const folder = DriveApp.getRootFolder(); // TODO: Use specific folder
    const file = folder.createFile(pdfBlob);
    file.setName(fileName);
    
    // Get shareable link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getUrl();
    
    Logger.log('PDF created: ' + url);
    
    return { success: true, url: url };
    
  } catch (error) {
    Logger.log('Error in generateReleaseNotePDF: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Update existing warehouse release note
 */
function updateWarehouseReleaseNote(recordData) {
  try {
    // Same as save, but overwrites existing data
    const result = saveWarehouseReleaseNote(recordData);
    
    if (result.success) {
      result.message = 'تم التحديث بنجاح';
    }
    
    return result;
    
  } catch (error) {
    Logger.log('Error in updateWarehouseReleaseNote: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Batch Save Warehouse Release Notes
 * REQUIRED for "Save All" button to work.
 */
/**
 * Batch Save Warehouse Release Notes
 * REQUIRED for "Save All" button to work.
 */
function batchSaveWarehouseReleaseNotes(batchData) {
  const results = [];
  try {
    batchData.forEach(data => {
       try {
         // Reuse the single save logic
         const res = saveWarehouseReleaseNote(data); 
         
         results.push({ 
             dismissNumber: data.dismissNumber, 
             success: res.success, 
             pdfUrl: res.pdfUrl,
             message: res.message
         });
         
       } catch (e) {
         results.push({ 
             dismissNumber: data.dismissNumber, 
             success: false, 
             message: e.toString() 
         });
       }
    });
    
    return results;
  } catch (e) {
    throw new Error("Batch operation failed: " + e.toString());
  }
}

/**
 * Optimized Batch Save - Single Spreadsheet Open + Batch Write
 */
function batchSaveWarehouseReleaseNotes(batchData) {
  try {
    if (!batchData || batchData.length === 0) {
      return { success: false, message: 'No data to save' };
    }

    // OPTIMIZATION: Open spreadsheet ONCE
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const dismissSheet = ss.getSheetByName('صرف و ارتجاع');
    
    if (!dismissSheet) {
      return { success: false, message: 'Sheet not found' };
    }

    const results = [];
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const userEmail = batchData[0].userEmail || ''; // Assume same user for batch

    // BATCH WRITE: Prepare all updates
    const updates = [];

    batchData.forEach(data => {
      try {
        // Validation
        if (!data || !data.products || !Array.isArray(data.products)) {
          results.push({ 
            dismissNumber: data.dismissNumber, 
            success: false, 
            message: 'بيانات غير صالحة'
          });
          return;
        }

        const rowIndex = data.rowIndex;

        // 1. Product Data
        data.products.forEach(product => {
          if (product.loadedQty && product.loadedQty > 0) {
            const colIndex = 19 + product.columnOffset;
            
            // Loaded Qty (Column U+offset)
            updates.push({
              range: dismissSheet.getRange(rowIndex, colIndex + 2),
              value: product.loadedQty
            });
            
            // Pallets (Column V+offset)
            if (product.pallets) {
              updates.push({
                range: dismissSheet.getRange(rowIndex, colIndex + 3),
                value: product.pallets
              });
            }
          }
        });

        // 2. Release Summary (Column G)
        const summary = generateReleaseSummary(data.products);
        updates.push({
          range: dismissSheet.getRange(rowIndex, 7),
          value: summary
        });

        // 3. Labor (Columns N, O)
        if (data.loadingLabor && data.loadingLabor.length > 0) {
          updates.push({
            range: dismissSheet.getRange(rowIndex, 14),
            value: data.loadingLabor.join(', ')
          });
        }
        if (data.unloadingLabor && data.unloadingLabor.length > 0) {
          updates.push({
            range: dismissSheet.getRange(rowIndex, 15),
            value: data.unloadingLabor.join(', ')
          });
        }

        // 4. Comments (Column R)
        if (data.comments) {
          updates.push({
            range: dismissSheet.getRange(rowIndex, 18),
            value: data.comments
          });
        }

        // 5. Metadata (Columns P, Q)
        updates.push({
          range: dismissSheet.getRange(rowIndex, 16),
          value: userEmail
        });
        updates.push({
          range: dismissSheet.getRange(rowIndex, 17),
          value: timestamp
        });

        results.push({ 
          dismissNumber: data.dismissNumber, 
          success: true, 
          message: 'تم الحفظ بنجاح',
          needsPrint: true // Flag for frontend to generate PNG
        });

      } catch (e) {
        results.push({ 
          dismissNumber: data.dismissNumber, 
          success: false, 
          message: e.toString() 
        });
      }
    });

    // BATCH WRITE: Execute all updates in one go
    if (updates.length > 0) {
      updates.forEach(update => {
        update.range.setValue(update.value);
      });
    }

    Logger.log('Batch save completed: ' + results.length + ' records');
    return results;

  } catch (e) {
    Logger.log('Batch save error: ' + e.toString());
    throw new Error('Batch operation failed: ' + e.toString());
  }
}
/**
 * Fetches deep data for the Release Note Printout
 * Joins Warehouse Record + Order Details + Client Details
 */
function getReleaseNoteDeepData(dismissNumber) {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const dismissSheet = ss.getSheetByName('صرف و ارتجاع');
    const ordersSheet = ss.getSheetByName('Orders');
    const clientsSheet = ss.getSheetByName('Clients');
    
    // 1. Get Warehouse Record
    const dismissData = dismissSheet.getDataRange().getValues();
    const dismissRow = dismissData.find((r, i) => i >= 12 && r[0].toString() === dismissNumber.toString());
    
    if (!dismissRow) return { success: false, message: 'Release note not found' };
    
    const releaseData = {
    dismissNumber: dismissRow[0],
    status: dismissRow[1],
    truckDriver: dismissRow[2],
    orderIdsStr: dismissRow[3],
    loadingDate: dismissRow[4] instanceof Date ? Utilities.formatDate(dismissRow[4], Session.getScriptTimeZone(), "yyyy-MM-dd") : dismissRow[4],
    loadingLabor: dismissRow[13],
    unloadingLabor: dismissRow[14],
    comments: dismissRow[17] || '',
    totalLoadedMap: {}
    };

    // Extract Loaded Quantities
    const productHeaders = dismissData[9]; 
    for (let c = 18; c < dismissRow.length; c += 7) {
        const prodName = productHeaders[c];
        if (prodName) {
            const loaded = parseFloat(dismissRow[c+2]) || 0;
            if (loaded > 0) {
                releaseData.totalLoadedMap[prodName.trim()] = loaded;
            }
        }
    }

    // 2. Fetch Orders & Products
    // CRITICAL FIX: Convert concatenated IDs to Array of Numbers
    // "0005-0006" -> [5, 6]
    const rawOrderString = releaseData.orderIdsStr ? releaseData.orderIdsStr.toString() : '';
    // Split by non-digit characters (handles "-", ",", " ", etc)
    const targetOrderIds = rawOrderString.split(/[^0-9]+/).filter(s => s.trim() !== '').map(s => parseInt(s, 10));
    
    Logger.log('Looking for Numeric Order IDs: ' + JSON.stringify(targetOrderIds));

    const ordersData = ordersSheet.getDataRange().getValues();
    const productColsMap = getProductColumnsMap(ordersSheet);
    
    const clientsMap = {}; 

    // Data starts Row 7. Col D=Order ID (index 3), I=Client (index 8), Q=Internal Comments (index 16)
    const clientsList = [];

    for (let i = 6; i < ordersData.length; i++) {
        const row = ordersData[i];
        const sheetIdVal = row[3];
        
        if (!sheetIdVal) continue;

        const sheetIdNum = parseInt(sheetIdVal, 10);
        
        if (targetOrderIds.includes(sheetIdNum)) {
            const clientName = row[8] ? row[8].toString().trim() : 'Unknown';
            const internalComments = row[16] ? row[16].toString().trim() : '';
            
            const clientOrder = {
                name: clientName,
                orderId: sheetIdNum,
                internalComments: internalComments,
                products: {}
            };
            
            // Extract Products
            for (const [prodName, cols] of Object.entries(productColsMap)) {
                if (cols.qty && row[cols.qty-1]) {
                    const qty = parseFloat(row[cols.qty-1]) || 0;
                    const bonus = parseFloat(row[cols.bonus-1]) || 0;
                    if (qty + bonus > 0) {
                        clientOrder.products[prodName] = (qty + bonus);
                    }
                }
            }
            
            clientsList.push(clientOrder);
        }
    }

// 3. Fetch Client Details with Distance (OPTIMIZED - Load regions ONCE)
const clientsData = clientsSheet.getDataRange().getValues();
const regionsSheet = ss.getSheetByName('المناطق');
const regionsData = regionsSheet ? regionsSheet.getDataRange().getValues() : null;

// Pre-calculate base region column index (ONCE, not per client)
let baseColIndex = -1;
if (regionsData) {
    const baseRegion = 'ابو رواش';
    const headerRow = regionsData[0];
    for (let c = 3; c < headerRow.length; c++) {
        if (headerRow[c] && headerRow[c].toString().trim() === baseRegion) {
            baseColIndex = c;
            break;
        }
    }
}

clientsList.forEach(client => {
    const clientRow = clientsData.find((r, i) => i >= 3 && r[3] && r[3].toString().trim() === client.name);
    
    if (clientRow) {
        client.paymentPlan = clientRow[7] || '';
        client.phone = clientRow[13] || '';
        client.workHours = clientRow[14] || '';
        client.personInCharge = clientRow[15] || '';
        client.region = clientRow[16] || '';
        client.address = clientRow[17] || '';
        client.specificDistance = parseFloat(clientRow[18]) || null; // Col S
        client.locationUrl = clientRow[36] || '';
        
      // Distance Logic: Try specific first, fallback to region (OPTIMIZED)
        if (client.specificDistance && client.specificDistance > 0) {
            client.distance = client.specificDistance;
            client.distanceType = 'specific';
        } else if (client.region && regionsData && baseColIndex > -1) {
            // Lookup region distance (base column already found above)
            const regionName = client.region.toString().trim();
            
            for (let r = 2; r < regionsData.length; r++) {
                if (regionsData[r][0] && regionsData[r][0].toString().trim() === regionName) {
                    const distance = parseFloat(regionsData[r][baseColIndex]) || 0;
                    if (distance > 0) {
                        client.distance = distance;
                        client.distanceType = 'region';
                    }
                    break;
                }
            }
        }
    }
});

    return { 
        success: true, 
        data: {
            release: releaseData,
            clients: clientsList
        }
    };

  } catch (e) {
    Logger.log("Deep Data Error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Saves the PNG image to Drive and updates the sheet
 */
function saveReleaseNoteImage(dismissNumber, base64Data) {
  try {
    const folderId = '1lyfSYYZAIQTW6i8q59DXqoLjQ5e563Oc'; // Invoice Folder (or change if needed)
    const folder = DriveApp.getFolderById(folderId);
    
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.PNG, `Release_${dismissNumber}.png`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Update Sheet
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('صرف و ارتجاع');
    const data = sheet.getDataRange().getValues();
    
    // Find row
    let rowIndex = -1;
    for(let i=12; i<data.length; i++) {
        if(data[i][0].toString() === dismissNumber.toString()) {
            rowIndex = i + 1;
            break;
        }
    }
    
    if(rowIndex > 0) {
        sheet.getRange(rowIndex, 8).setValue(file.getUrl()); // Col H
        return { success: true, url: file.getUrl() };
    }
    
    return { success: false, message: 'Record not found to update URL' };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Saves SVG document to Google Drive
 */
/**
 * Saves HTML document to Google Drive
 */
function saveHtmlDocument(dismissNumber, htmlContent) {
  try {
    const folderId = '1lyfSYYZAIQTW6i8q59DXqoLjQ5e563Oc';
    const folder = DriveApp.getFolderById(folderId);
    
    // Create HTML file
    const blob = Utilities.newBlob(htmlContent, 'text/html', `Release_${dismissNumber}.html`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const url = file.getUrl();
    
    // Update Sheet Column H
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('صرف و ارتجاع');
    const data = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    for(let i = 12; i < data.length; i++) {
        if(data[i][0].toString() === dismissNumber.toString()) {
            rowIndex = i + 1;
            break;
        }
    }
    
    if(rowIndex > 0) {
        sheet.getRange(rowIndex, 8).setValue(url); // Col H
        return { success: true, url: url };
    }
    
    return { success: false, message: 'Record not found to update URL' };

  } catch (e) {
    Logger.log('HTML save error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}
// =============================================================================
// ADMIN RETURN EDIT FEATURES
// =============================================================================

function getReturnOrderDetails(orderIdsStr) {
  try {
    // 1. Robust ID Extraction
    // Convert input to string and extract only numeric sequences (digits)
    // This handles "1001 - 1002", "1001, 1002", or "Order 1001" automatically.
    const rawStr = String(orderIdsStr || '');
    const targetIds = rawStr.match(/\d+/g); 

    if (!targetIds || targetIds.length === 0) {
        return JSON.stringify({ success: true, orders: [] });
    }

    if (targetIds.length === 0) return JSON.stringify({ success: true, orders: [] });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Orders');
    
    // 2. Map Columns (Start at AQ = Index 42)
    // We assume Headers are at Row 3
    const headerRow = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues()[0];
    const productMap = {};
    for (let c = 42; c < headerRow.length; c += 5) {
        const prodName = headerRow[c];
        if (prodName) {
            productMap[String(prodName).trim()] = {
                qtyCol: c + 1,   // 1-based index
                bonusCol: c + 2
            };
        }
    }

    // 3. Fetch Data
    const lastRow = getOrdersLastRow(sheet);
    if (lastRow < 7) {
        Logger.log("Error: Orders sheet has less than 7 rows");
        return JSON.stringify({ success: true, orders: [] });
    }

    // Getting ALL data to ensure we don't miss anything (Rows 7 to Last)
    const data = sheet.getRange(7, 1, lastRow - 6, sheet.getLastColumn()).getValues();
    Logger.log(`Searching through ${data.length} rows in 'Orders' sheet`);

    const resultOrders = [];

    targetIds.forEach(targetId => {
        // NUMERIC MATCHING: Handles "0001" vs 1 vs "1"
        const row = data.find(r => {
             // r[3] is Column D
             // Convert both to Integers (Base 10) so 0001 == 1
             const sheetId = parseInt(r[3], 10);
             const searchId = parseInt(targetId, 10);
             return !isNaN(sheetId) && !isNaN(searchId) && sheetId === searchId;
        });

        if (row) {
            Logger.log(`FOUND Match for ID: ${targetId}`);
            const products = [];
            Object.keys(productMap).forEach(prodName => {
                const cols = productMap[prodName];
                const qtyVal = row[cols.qtyCol - 1];
                const bonusVal = row[cols.bonusCol - 1];
                
                // Include if it has values (number or string)
                if ((qtyVal !== "" && qtyVal !== null) || (bonusVal !== "" && bonusVal !== null)) {
                     // Check if greater than 0
                    if (Number(qtyVal) > 0 || Number(bonusVal) > 0) {
                        products.push({
                            name: prodName,
                            qty: Number(qtyVal) || 0,
                            bonus: Number(bonusVal) || 0
                        });
                    }
                }
            });
            resultOrders.push({
                orderId: targetId,
                clientName: row[8], // Col I
                products: products
            });
        } else {
            Logger.log(`WARNING: Could NOT find ID '${targetId}' in Column D`);
        }
    });

    Logger.log(`Final Result: Found ${resultOrders.length} orders`);
    return JSON.stringify({ success: true, orders: resultOrders });

  } catch (e) {
    Logger.log("FATAL ERROR: " + e.toString());
    return JSON.stringify({ success: false, message: e.toString() });
  }
}

/**
 * Updates Qty and Bonus (Accepts JSON String, Returns JSON String)
 * AUTO-UPDATES: Order Summary (Col S) and Total Weight (Col AA)
 * ADVANCED VALIDATION: Checks specific product types (Tawoos vs PL) against their respective invoices.
 */
function updateOriginalOrderQuantities(updatesJson, userEmail) {
  try {
    const updates = JSON.parse(updatesJson); 
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Orders');
    const productsSheet = ss.getSheetByName('Products'); 

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm');
    const userLabel = userEmail ? userEmail.split('@')[0] : 'Admin';
    
    // Default Success Message
    let returnMessage = "تم تحديث الكميات والوزن وملخص الطلب بنجاح"; 
    let hasWarnings = false;

    // 1. Build Product Meta Map (Weight & PL Status)
    const productMeta = {};
    if (productsSheet) {
        const pLastRow = productsSheet.getLastRow();
        if (pLastRow >= 3) {
            // Col A=Name, Col G=Weight(6), Col I=IsPrivateLabel(8)
            // Reading up to Column I (9 columns)
            const pData = productsSheet.getRange(3, 1, pLastRow - 2, 9).getValues();
            pData.forEach(r => {
                if(r[0]) {
                    productMeta[String(r[0]).trim()] = {
                        weight: parseFloat(r[6]) || 0,
                        isPL: (r[8] === true) // Checkbox in Col I
                    };
                }
            });
        }
    }

    // 2. Map Product Columns in Orders Sheet
    const headerRow = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues()[0];
    const productMap = {};
    for (let c = 42; c < headerRow.length; c += 5) {
        const prodName = headerRow[c];
        if (prodName) {
            productMap[String(prodName).trim()] = {
                qtyCol: c + 1,
                bonusCol: c + 2
            };
        }
    }

    const lastRow = getOrdersLastRow(sheet);
    if (lastRow < 7) return JSON.stringify({ success: false, message: 'No orders found' });

    const orderIdsData = sheet.getRange(7, 4, lastRow - 6, 1).getValues(); 
    
    updates.forEach(update => {
        const rowIndex = orderIdsData.findIndex(r => {
            const sheetId = parseInt(r[0], 10);
            const searchId = parseInt(update.orderId, 10);
            return !isNaN(sheetId) && !isNaN(searchId) && sheetId === searchId;
        });

        if (rowIndex !== -1) {
            const sheetRow = rowIndex + 7;
            const fullRowRange = sheet.getRange(sheetRow, 1, 1, sheet.getLastColumn());
            const rowValues = fullRowRange.getValues()[0]; 

            // TRACKING FLAGS
            let tawoosChanged = false;
            let plChanged = false;

            // A. Apply Updates
            update.products.forEach(p => {
                const prodName = String(p.name).trim();
                const cols = productMap[prodName];
                
                // Determine Type
                const isPL = productMeta[prodName] ? productMeta[prodName].isPL : false;

                if (cols) {
                    let changed = false;
                    // Update Qty
                    const currentQty = rowValues[cols.qtyCol - 1]; 
                    if (currentQty != p.qty) {
                        sheet.getRange(sheetRow, cols.qtyCol).setValue(p.qty);
                        const n = sheet.getRange(sheetRow, cols.qtyCol).getNote();
                        sheet.getRange(sheetRow, cols.qtyCol).setNote((n?n+"\n":"") + `[${timestamp} ${userLabel}] Qty -> ${p.qty}`);
                        rowValues[cols.qtyCol - 1] = p.qty;
                        changed = true;
                    }
                    // Update Bonus
                    const currentBonus = rowValues[cols.bonusCol - 1];
                    if (currentBonus != p.bonus) {
                        sheet.getRange(sheetRow, cols.bonusCol).setValue(p.bonus);
                        const n = sheet.getRange(sheetRow, cols.bonusCol).getNote();
                        sheet.getRange(sheetRow, cols.bonusCol).setNote((n?n+"\n":"") + `[${timestamp} ${userLabel}] Bonus -> ${p.bonus}`);
                        rowValues[cols.bonusCol - 1] = p.bonus;
                        changed = true;
                    }

                    // Mark flag if changed
                    if (changed) {
                        if (isPL) plChanged = true;
                        else tawoosChanged = true;
                    }
                }
            });

            // B. Recalculate Summary & Weight
            let newSummary = [];
            let newWeight = 0;

            Object.keys(productMap).forEach(pName => {
                const cols = productMap[pName];
                const q = parseFloat(rowValues[cols.qtyCol - 1]) || 0;
                const b = parseFloat(rowValues[cols.bonusCol - 1]) || 0;
                const total = q + b;
                
                if (total > 0) {
                    newSummary.push(`${total} ${pName}`);
                    const w = productMeta[pName] ? productMeta[pName].weight : 0;
                    newWeight += (total * w);
                }
            });

            const summaryStr = newSummary.join('\n');
            if (rowValues[18] !== summaryStr) sheet.getRange(sheetRow, 19).setValue(summaryStr); 
            if (parseFloat(rowValues[26]) !== newWeight) sheet.getRange(sheetRow, 27).setValue(newWeight);

            // ==========================================================================================
            // C. PRECISE INVOICE VALIDATION
            // ==========================================================================================
            const alerts = [];

            // 1. Validate TAWOOS Changes
            if (tawoosChanged) {
                const hasTawoosInv = rowValues[29] && String(rowValues[29]).trim() !== ""; // Col AD
                const hasTawoosETA = rowValues[34] && String(rowValues[34]).trim() !== ""; // Col AI (ETA UUID)

                if (hasTawoosETA) {
                    alerts.push("⚠️ منتجات (الطاووس): الفاتورة الإلكترونية معتمدة. يجب عمل إشعار خصم/إضافة (Credit Note).");
                } else if (hasTawoosInv) {
                    alerts.push("⚠️ منتجات (الطاووس): توجد فاتورة داخلية. يجب إعادة حفظ الفاتورة بنفس الرقم.");
                }
            }

            // 2. Validate PL Changes
            if (plChanged) {
                const hasPLInv = rowValues[30] && String(rowValues[30]).trim() !== ""; // Col AE
                const hasPLETA = rowValues[35] && String(rowValues[35]).trim() !== ""; // Col AJ (ETA UUID)

                if (hasPLETA) {
                    alerts.push("⚠️ منتجات (Private Label): الفاتورة الإلكترونية معتمدة. يجب عمل إشعار خصم/إضافة (Credit Note).");
                } else if (hasPLInv) {
                    alerts.push("⚠️ منتجات (Private Label): توجد فاتورة داخلية. يجب إعادة حفظ الفاتورة بنفس الرقم.");
                }
            }

            // 3. General Fallback (Delivery Note Check)
            if (alerts.length === 0 && (tawoosChanged || plChanged)) {
                 const hasDN = rowValues[28] && String(rowValues[28]).trim() !== ""; // Col AC
                 if (hasDN) {
                     alerts.push("⚠️ إذن التسليم: يجب إعادة الحفظ لتحديث القيم.");
                 }
            }

            if (alerts.length > 0) {
                hasWarnings = true;
                returnMessage = "تم التحديث. يرجى الانتباه للإجراءات التالية:\n\n" + alerts.join("\n");
            }
        }
    });

    return JSON.stringify({ success: true, message: returnMessage });

  } catch (e) {
    return JSON.stringify({ success: false, message: e.toString() });
  }
}

// =============================================================================
// END OF WAREHOUSE MODULE
// =============================================================================
