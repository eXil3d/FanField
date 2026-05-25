// ============================================================
// FAN FIELD - Google Apps Script Backend API v3.0
// All features: Delivery, Vouchers, Affiliate Commission,
// Sale Reversal, Stock Log, Delivery Management
// ============================================================

const ss = SpreadsheetApp.getActiveSpreadsheet();

// ============================================================
// NEW SHEETS NEEDED (create these in your spreadsheet):
// - StockLog: ID, Date, ProductID, ProductName, Action, Quantity, OldStock, NewStock, Reference, Notes
// - Vouchers: ID, Code, Name, DiscountType, DiscountValue, MinOrder, MaxUses, TimesUsed, TotalDiscountGiven, ValidFrom, ValidTo, Active
// - AffiliateCommissions: ID, Date, AffiliateID, AffiliateCode, AffiliateName, InvoiceNo, Edition, CommissionAmount, Status, PaidDate, ExpenseID, Notes
// - Deliveries: ID, InvoiceNo, Date, CustomerName, CustomerPhone, Address, DeliveryType, DeliveryCharge, CourierService, TrackingNo, DeliveryStatus, Notes
// Sales sheet: Add columns 22=DeliveryType, 23=DeliveryCharge, 24=VoucherCode, 25=VoucherDiscount, 26=DeliveryStatus, 27=CourierService, 28=TrackingNo
// Affiliates sheet: Add columns 9=CommissionFanEdition, 10=CommissionPlayerEdition, 11=CommissionSpecialEdition, 12=CommissionBDPremium, 13=TotalCommission, 14=PaidCommission, 15=PendingCommission
// ============================================================

function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    switch(action) {
      case 'getDashboard': result = getDashboard(); break;
      case 'getInventory': result = getInventory(); break;
      case 'getBatches': result = getBatches(e.parameter.productId); break;
      case 'getAllBatches': result = getAllBatches(); break;
      case 'getCustomers': result = getCustomers(); break;
      case 'getSales': result = getSales(); break;
      case 'getSalesItems': result = getSalesItems(e.parameter.invoiceNo); break;
      case 'getPayments': result = getPayments(); break;
      case 'getSuppliers': result = getSuppliers(); break;
      case 'getSupplierLedger': result = getSupplierLedger(); break;
      case 'getSupplierLedgerById': result = getSupplierLedgerById(e.parameter.supplierId); break;
      case 'getExpenses': result = getExpenses(); break;
      case 'getAffiliates': result = getAffiliates(); break;
      case 'getSettings': result = getSettings(); break;
      case 'getNextInvoice': result = getNextInvoiceNumber(); break;
      case 'getCustomerLedger': result = getCustomerLedger(e.parameter.customerId); break;
      case 'getStockLog': result = getStockLog(); break;
      case 'getVouchers': result = getVouchers(); break;
      case 'getAffiliateCommissions': result = getAffiliateCommissions(e.parameter.affiliateId); break;
      case 'getDeliveries': result = getDeliveries(); break;
      case 'getPreorders': result = getPreorders(); break;
      case 'getDeadStockHistory': result = getDeadStockHistory(); break;
      case 'getChronicDeadStock': result = getChronicDeadStock(); break;
      default: result = {error:'Unknown action: '+action};
    }
  } catch(err) {
    result = {error: err.toString(), stack: err.stack};
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  let result;
  try {
    switch(action) {
      case 'addProduct': result = addProduct(data); break;
      case 'updateProduct': result = updateProduct(data); break;
      case 'deleteProduct': result = deleteProduct(data); break;
      case 'updateStock': result = updateStock(data); break;
      case 'addBatch': result = addBatch(data); break;
      case 'addCustomer': result = addCustomer(data); break;
      case 'updateCustomer': result = updateCustomer(data); break;
      case 'deleteCustomer': result = deleteCustomer(data); break;
      case 'createSale': result = createSale(data); break;
      case 'updateSaleProfit': result = updateSaleProfit(data); break;
      case 'adjustSaleAmount': result = adjustSaleAmount(data); break;
      case 'deleteSale': result = deleteSale(data); break;
      case 'reverseSale': result = reverseSale(data); break;
      case 'partialReverseSale': result = partialReverseSale(data); break;
      case 'exchangeAndReplace': result = exchangeAndReplace(data); break;
      case 'addManualSale': result = addManualSale(data); break;
      case 'addPayment': result = addPayment(data); break;
      case 'addSupplier': result = addSupplier(data); break;
      case 'updateSupplier': result = updateSupplier(data); break;
      case 'deleteSupplier': result = deleteSupplier(data); break;
      case 'addSupplierLedgerEntry': result = addSupplierLedgerEntry(data); break;
      case 'addExpense': result = addExpense(data); break;
      case 'deleteExpense': result = deleteExpense(data); break;
      case 'addAffiliate': result = addAffiliate(data); break;
      case 'updateAffiliate': result = updateAffiliate(data); break;
      case 'deleteAffiliate': result = deleteAffiliate(data); break;
      case 'updateSettings': result = updateSettings(data); break;
      case 'addVoucher': result = addVoucher(data); break;
      case 'updateVoucher': result = updateVoucher(data); break;
      case 'deleteVoucher': result = deleteVoucher(data); break;
      case 'payAffiliateCommission': result = payAffiliateCommission(data); break;
      case 'reverseAffiliatePayment': result = reverseAffiliatePayment(data); break;
      case 'updateDeliveryStatus': result = updateDeliveryStatus(data); break;
      case 'updateSaleDelivery': result = updateSaleDelivery(data); break;
      case 'migratePhoneFormats': result = migratePhoneFormats(); break;
      case 'addPreorder': result = addPreorder(data); break;
      case 'updatePreorder': result = updatePreorder(data); break;
      case 'deletePreorder': result = deletePreorder(data); break;
      case 'saveDeadStockSnapshot': result = saveDeadStockSnapshot(data); break;
      case 'deleteDeadStockSnapshot': result = deleteDeadStockSnapshot(data); break;
      default: result = {error:'Unknown action: '+action};
    }
  } catch(err) {
    result = {error: err.toString(), stack: err.stack};
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// PHONE NUMBER FORMATTER & MIGRATION
// ============================================================

// Format phone to 017-1234-5678
function formatPhoneBackend(phone) {
  if (!phone) return '';
  
  // Strip all non-numeric characters
  let digits = String(phone).replace(/\D/g, '');
  
  // Handle country code variations
  if (digits.indexOf('880') === 0) digits = digits.substring(3);
  if (digits.indexOf('00880') === 0) digits = digits.substring(5);
  
  // Add leading 0 if missing
  if (digits.length === 10 && digits.charAt(0) !== '0') {
    digits = '0' + digits;
  }
  
  // Limit to 11 digits
  if (digits.length > 11) digits = digits.substring(0, 11);
  
  // Don't format if too short (incomplete numbers preserved as-is)
  if (digits.length < 11) return digits;
  
  // Format: 017-1234-5678
  return digits.substring(0, 3) + '-' + digits.substring(3, 7) + '-' + digits.substring(7);
}

function migratePhoneFormats() {
  let customersUpdated = 0;
  let suppliersUpdated = 0;
  let salesUpdated = 0;
  let paymentsUpdated = 0;
  let deliveriesUpdated = 0;
  
  // ============================================================
  // 1. Update Customers (Phone is in column 3)
  // ============================================================
  const custSheet = ss.getSheetByName('Customers');
  if (custSheet) {
    const lastRow = custSheet.getLastRow();
    if (lastRow > 1) {
      const phoneRange = custSheet.getRange(2, 3, lastRow - 1, 1);
      const phones = phoneRange.getValues();
      const newPhones = [];
      let changed = false;
      
      phones.forEach(row => {
        const original = String(row[0] || '');
        const formatted = formatPhoneBackend(original);
        if (formatted !== original && formatted !== '') {
          changed = true;
          customersUpdated++;
        }
        newPhones.push([formatted || original]);
      });
      
      if (changed) phoneRange.setValues(newPhones);
    }
  }
  
  // ============================================================
  // 2. Update Suppliers (Phone is in column 3)
  // ============================================================
  const supSheet = ss.getSheetByName('Suppliers');
  if (supSheet) {
    const lastRow = supSheet.getLastRow();
    if (lastRow > 1) {
      const phoneRange = supSheet.getRange(2, 3, lastRow - 1, 1);
      const phones = phoneRange.getValues();
      const newPhones = [];
      let changed = false;
      
      phones.forEach(row => {
        const original = String(row[0] || '');
        const formatted = formatPhoneBackend(original);
        if (formatted !== original && formatted !== '') {
          changed = true;
          suppliersUpdated++;
        }
        newPhones.push([formatted || original]);
      });
      
      if (changed) phoneRange.setValues(newPhones);
    }
  }
  
  // ============================================================
  // 3. Update Sales (CustomerPhone is in column 6)
  // ============================================================
  const salesSheet = ss.getSheetByName('Sales');
  if (salesSheet) {
    const lastRow = salesSheet.getLastRow();
    if (lastRow > 1) {
      const phoneRange = salesSheet.getRange(2, 6, lastRow - 1, 1);
      const phones = phoneRange.getValues();
      const newPhones = [];
      let changed = false;
      
      phones.forEach(row => {
        const original = String(row[0] || '');
        const formatted = formatPhoneBackend(original);
        if (formatted !== original && formatted !== '') {
          changed = true;
          salesUpdated++;
        }
        newPhones.push([formatted || original]);
      });
      
      if (changed) phoneRange.setValues(newPhones);
    }
  }
  
  // ============================================================
  // 4. Update Deliveries (CustomerPhone is in column 5)
  // ============================================================
  const dlvSheet = ss.getSheetByName('Deliveries');
  if (dlvSheet) {
    const lastRow = dlvSheet.getLastRow();
    if (lastRow > 1) {
      const phoneRange = dlvSheet.getRange(2, 5, lastRow - 1, 1);
      const phones = phoneRange.getValues();
      const newPhones = [];
      let changed = false;
      
      phones.forEach(row => {
        const original = String(row[0] || '');
        const formatted = formatPhoneBackend(original);
        if (formatted !== original && formatted !== '') {
          changed = true;
          deliveriesUpdated++;
        }
        newPhones.push([formatted || original]);
      });
      
      if (changed) phoneRange.setValues(newPhones);
    }
  }
  
  return {
    success: true,
    customersUpdated,
    suppliersUpdated,
    salesUpdated,
    deliveriesUpdated,
    message: `Phone numbers formatted: ${customersUpdated} customers, ${suppliersUpdated} suppliers, ${salesUpdated} sales, ${deliveriesUpdated} deliveries.`
  };
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================

function getSheetData(sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      }
      obj[headers[j]] = val;
    }
    obj._row = i + 1;
    rows.push(obj);
  }
  return rows;
}

function generateId(prefix) {
  if (!prefix) {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
  }
  const sheetMap = {
    'PRD': 'Inventory', 'SUP': 'Suppliers', 'CUS': 'Customers',
    'BAT': 'InventoryBatches', 'PAY': 'Payments', 'EXP': 'Expenses',
    'AFF': 'Affiliates', 'SLE': 'SupplierLedger', 'VCH': 'Vouchers',
    'AFCM': 'AffiliateCommissions', 'DLV': 'Deliveries', 'SLOG': 'StockLog'
  };
  const sheetName = sheetMap[prefix];
  if (!sheetName) {
    return prefix + '-' + Date.now().toString(36).substr(-4).toUpperCase();
  }
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return prefix + '-' + Date.now().toString(36).substr(-4).toUpperCase();
  const lastRow = sheet.getLastRow();
  let maxNum = 0;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(row => {
      const id = String(row[0]);
      const match = id.match(new RegExp('^' + prefix + '-(\\d+)'));
      if (match) {
        const num = parseInt(match[1]);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    });
  }
  return prefix + '-' + String(maxNum + 1).padStart(4, '0');
}

function findRowById(sheetName, id) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function findRowByCol(sheetName, colIndex, value) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(value)) return i + 1;
  }
  return -1;
}

function fmtDate(date) {
  if (!date) date = new Date();
  if (typeof date === 'string') return date;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function buildProductName(club, type, edition, size) {
  return [club, type, edition, size].filter(Boolean).join(' ');
}

// ============================================================
// INVENTORY SORTING
// ============================================================

function getSizeOrder(size) {
  const order = {'S':1,'M':2,'L':3,'XL':4,'2XL':5,'3XL':6,'4XL':7,'5XL':8};
  return order[String(size).toUpperCase()] || 99;
}
function getTypeOrder(type) {
  const order = {'Home':1,'Away':2,'Third':3,'Retro':4,'GK':5,'Training':6,'Terrance':7};
  return order[type] || 99;
}
function getEditionOrder(edition) {
  const order = {'Player Edition':1,'Fan Edition':2,'Special Edition':3,'BD Premium':4};
  return order[edition] || 99;
}

function sortInventorySheet() {
  const sheet = ss.getSheetByName('Inventory');
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const numCols = sheet.getLastColumn();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, numCols);
  const data = dataRange.getValues();
  data.sort(function(a, b) {
    const clubA = String(a[2] || '').toLowerCase();
    const clubB = String(b[2] || '').toLowerCase();
    if (clubA !== clubB) return clubA < clubB ? -1 : 1;
    const typeA = getTypeOrder(a[4]);
    const typeB = getTypeOrder(b[4]);
    if (typeA !== typeB) return typeA - typeB;
    const edA = getEditionOrder(a[3]);
    const edB = getEditionOrder(b[3]);
    if (edA !== edB) return edA - edB;
    const sizeA = getSizeOrder(a[5]);
    const sizeB = getSizeOrder(b[5]);
    return sizeA - sizeB;
  });
  dataRange.setValues(data);
}

// ============================================================
// STOCK LOG
// ============================================================

function logStockChange(productId, productName, action, qty, oldStock, newStock, reference, notes) {
  const sheet = ss.getSheetByName('StockLog');
  if (!sheet) return;
  sheet.appendRow([
    generateId('SLOG'), fmtDate(new Date()),
    productId, productName || '',
    action, qty, oldStock, newStock,
    reference || '', notes || ''
  ]);
}

function getStockLog() {
  return getSheetData('StockLog');
}

// ============================================================
// DASHBOARD
// ============================================================

function getDashboard() {
  const inventory = getSheetData('Inventory');
  const sales = getSheetData('Sales');
  const customers = getSheetData('Customers');
  const payments = getSheetData('Payments');
  const expenses = getSheetData('Expenses');
  const suppliers = getSheetData('Suppliers');

  const today = new Date();
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const thisMonth = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM");

  let totalStockValue = 0, totalRetailValue = 0, totalItems = 0;
  let lowStockItems = [], outOfStockItems = [];

  inventory.forEach(item => {
    const stock = Number(item.Stock) || 0;
    const cost = Number(item.CostPrice) || 0;
    const sell = Number(item.SellPrice) || 0;
    const minStock = Number(item.MinStock) || 5;
    totalItems += stock;
    totalStockValue += stock * cost;
    totalRetailValue += stock * sell;
    if (stock === 0) outOfStockItems.push(item);
    else if (stock <= minStock) lowStockItems.push(item);
  });

  let totalRevenue = 0, todaySales = 0, monthlySales = 0, totalPending = 0;
  let todaySaleCount = 0, monthlySaleCount = 0;
  let totalProfit = 0, monthlyProfit = 0, todayProfit = 0;
  let totalCost = 0;

  sales.forEach(sale => {
    if (sale.Status === 'Cancelled') return;
    const saleDate = String(sale.Date).substring(0, 10);
    const saleMonth = String(sale.Date).substring(0, 7);
    const finalAmt = Number(sale.FinalAmount) || 0;
    const dueAmt = Number(sale.DueAmount) || 0;
    const profit = Number(sale.ProfitOverride) || Number(sale.Profit) || 0;
    const costAmt = Number(sale.CostTotal) || 0;
    totalRevenue += finalAmt;
    totalPending += dueAmt;
    totalProfit += profit;
    totalCost += costAmt;
    if (saleDate === todayStr) { todaySales += finalAmt; todaySaleCount++; todayProfit += profit; }
    if (saleMonth === thisMonth) { monthlySales += finalAmt; monthlySaleCount++; monthlyProfit += profit; }
  });

  let totalExpenses = 0, monthlyExpenses = 0;
  expenses.forEach(exp => {
    const amt = Number(exp.Amount) || 0;
    totalExpenses += amt;
    const expMonth = String(exp.Date).substring(0, 7);
    if (expMonth === thisMonth) monthlyExpenses += amt;
  });

  let supplierPending = 0;
  suppliers.forEach(sup => { supplierPending += Number(sup.PendingAmount) || 0; });

  const recentSales = sales.filter(s => s.Status !== 'Cancelled').slice(-15).reverse();

  const salesItems = getSheetData('SalesItems');
  const productSales = {};
  salesItems.forEach(item => {
    const name = item.ProductName;
    if (!productSales[name]) productSales[name] = {name, qty:0, revenue:0};
    productSales[name].qty += Number(item.Quantity) || 0;
    productSales[name].revenue += Number(item.LineTotal) || 0;
  });
  const topProducts = Object.values(productSales).sort((a,b) => b.revenue - a.revenue).slice(0, 10);
  const totalUnitsSold = salesItems.reduce((sum, item) => sum + (Number(item.Quantity) || 0), 0);

  const monthlySalesData = {};
  for (let i = 5; i >= 0; i--) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM");
    monthlySalesData[key] = {month: Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM yyyy"), sales:0, profit:0, count:0};
  }
  sales.forEach(sale => {
    if (sale.Status === 'Cancelled') return;
    const saleMonth = String(sale.Date).substring(0, 7);
    if (monthlySalesData[saleMonth]) {
      monthlySalesData[saleMonth].sales += Number(sale.FinalAmount) || 0;
      monthlySalesData[saleMonth].profit += Number(sale.ProfitOverride) || Number(sale.Profit) || 0;
      monthlySalesData[saleMonth].count++;
    }
  });

  return {
    totalProducts: inventory.length, totalItems,
    totalStockValue: Math.round(totalStockValue), totalRetailValue: Math.round(totalRetailValue),
    totalRevenue: Math.round(totalRevenue), todaySales: Math.round(todaySales), todaySaleCount,
    monthlySales: Math.round(monthlySales), monthlySaleCount,
    totalPending: Math.round(totalPending), totalExpenses: Math.round(totalExpenses),
    monthlyExpenses: Math.round(monthlyExpenses),
    totalProfit: Math.round(totalProfit), monthlyProfit: Math.round(monthlyProfit),
    todayProfit: Math.round(todayProfit), totalCost: Math.round(totalCost),
    supplierPending: Math.round(supplierPending),
    totalCustomers: customers.length, totalUnitsSold, totalSuppliers: suppliers.length,
    lowStockItems, outOfStockItems,
    lowStockCount: lowStockItems.length, outOfStockCount: outOfStockItems.length,
    recentSales, topProducts,
    monthlySalesChart: Object.values(monthlySalesData)
  };
}

// ============================================================
// INVENTORY
// ============================================================

function getInventory() { return getSheetData('Inventory'); }

function addProduct(data) {
  const sheet = ss.getSheetByName('Inventory');
  const id = generateId('PRD');
  const now = new Date();
  const name = buildProductName(data.Club_Country, data.Type, data.Edition, data.Size);
  sheet.appendRow([
    id, name, data.Club_Country || '', data.Edition || '', data.Type || '', data.Size || '',
    data.Color || '', Number(data.CostPrice) || 0, Number(data.SellPrice) || 0,
    Number(data.Stock) || 0, Number(data.MinStock) || 5,
    data.SupplierID || '', data.SupplierName || '', fmtDate(now), fmtDate(now)
  ]);
  if (Number(data.Stock) > 0) {
    logStockChange(id, name, 'Initial Stock', Number(data.Stock), 0, Number(data.Stock), '', 'Product created');
  }
  sortInventorySheet();
  return {success:true, id, ProductName:name};
}

function updateProduct(data) {
  const sheet = ss.getSheetByName('Inventory');
  const row = findRowById('Inventory', data.ID);
  if (row === -1) return {error:'Product not found'};
  const name = buildProductName(data.Club_Country, data.Type, data.Edition, data.Size);
  const existing = sheet.getRange(row, 14).getValue();
  const vals = [
    data.ID, name, data.Club_Country || '', data.Edition || '', data.Type || '', data.Size || '',
    data.Color || '', Number(data.CostPrice) || 0, Number(data.SellPrice) || 0,
    Number(data.Stock) || 0, Number(data.MinStock) || 5,
    data.SupplierID || '', data.SupplierName || '',
    existing || fmtDate(new Date()), fmtDate(new Date())
  ];
  sheet.getRange(row, 1, 1, vals.length).setValues([vals]);
  sortInventorySheet();
  return {success:true, ProductName:name};
}

function deleteProduct(data) {
  const sheet = ss.getSheetByName('Inventory');
  const row = findRowById('Inventory', data.ID);
  if (row === -1) return {error:'Product not found'};
  sheet.deleteRow(row);
  return {success:true};
}

function updateStock(data) {
  const sheet = ss.getSheetByName('Inventory');
  const row = findRowById('Inventory', data.ID);
  if (row === -1) return {error:'Product not found'};
  const currentStock = Number(sheet.getRange(row, 10).getValue()) || 0;
  const productName = sheet.getRange(row, 2).getValue();
  let newStock = currentStock;
  const qty = Number(data.Stock) || 0;
  if (data.mode === 'set') newStock = qty;
  else if (data.mode === 'add') newStock = currentStock + qty;
  else if (data.mode === 'subtract') newStock = Math.max(0, currentStock - qty);
  sheet.getRange(row, 10).setValue(newStock);
  sheet.getRange(row, 15).setValue(fmtDate(new Date()));
  logStockChange(data.ID, productName, 'Manual ' + data.mode, qty, currentStock, newStock, '', data.Notes || '');
  return {success:true, newStock};
}

// ============================================================
// INVENTORY BATCHES
// ============================================================

function getAllBatches() { return getSheetData('InventoryBatches'); }
function getBatches(productId) {
  return getSheetData('InventoryBatches').filter(b => String(b.ProductID) === String(productId));
}

// REPLACE existing addBatch - now creates sub-product even from same supplier:
function addBatch(data) {
  const invSheet = ss.getSheetByName('Inventory');
  const batchSheet = ss.getSheetByName('InventoryBatches');
  const now = new Date();
  
  const origRow = findRowById('Inventory', data.ProductID);
  if (origRow === -1) return {error: 'Original product not found'};
  
  const origData = invSheet.getRange(origRow, 1, 1, 15).getValues()[0];
  const baseId = String(origData[0]);
  
  // Strip any existing suffix to get the TRUE base ID
  const trueBase = baseId.replace(/[A-Z]+$/, '');
  
  const suffix = getNextBatchSuffix(trueBase);
  const batchProductId = trueBase + suffix;
  
  // Build a clean variant name with supplier + cost identifier
  const supplierLabel = data.SupplierName || 'New Batch';
  const costLabel = data.CostPrice ? ` @${data.CostPrice}` : '';
  const variantName = origData[1].replace(/\s*\[.*?\]$/, '') + ' [' + supplierLabel + costLabel + ']';
  
  invSheet.appendRow([
    batchProductId, variantName,
    origData[2], origData[3], origData[4], origData[5],
    origData[6],
    Number(data.CostPrice) || 0,
    Number(data.SellPrice || origData[8]) || 0,
    Number(data.Quantity) || 0,
    Number(origData[10]) || 5,
    data.SupplierID || origData[11] || '',
    data.SupplierName || origData[12] || '',
    fmtDate(now), fmtDate(now)
  ]);
  
  const batchId = generateId('BAT');
  batchSheet.appendRow([
    batchId, batchProductId, variantName,
    data.SupplierID || '', data.SupplierName || '',
    fmtDate(now),
    Number(data.Quantity) || 0,
    Number(data.CostPrice) || 0,
    Number(data.Quantity) || 0,
    data.Notes || ('New price batch from ' + supplierLabel)
  ]);
  
  logStockChange(batchProductId, variantName, 'Batch Added', 0, Number(data.Quantity) || 0, 'New batch from ' + supplierLabel, batchId);
  try { sortInventorySheet(); } catch(e) {}
  return {success:true, batchId, batchProductId};
}

// REPLACE existing getNextBatchSuffix to handle the true base properly:
function getNextBatchSuffix(baseId) {
  const sheet = ss.getSheetByName('Inventory');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'A';
  
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
  const used = [];
  
  ids.forEach(id => {
    // Match exact base or base + single letter suffix
    if (id === baseId) return; // base exists, that's fine
    if (id.startsWith(baseId)) {
      const suffix = id.substring(baseId.length);
      if (/^[A-Z]+$/.test(suffix)) used.push(suffix);
    }
  });
  
  if (!used.length) return 'A';
  
  // Sort and find next available letter
  used.sort();
  const lastChar = used[used.length - 1];
  
  if (lastChar.length === 1) {
    const code = lastChar.charCodeAt(0);
    if (code < 90) return String.fromCharCode(code + 1); // B, C, D...
    return 'AA'; // After Z, go to AA
  }
  // For double letters (AA, AB...), increment last char
  const lastLetter = lastChar.charAt(lastChar.length - 1);
  if (lastLetter.charCodeAt(0) < 90) {
    return lastChar.substring(0, lastChar.length - 1) + String.fromCharCode(lastLetter.charCodeAt(0) + 1);
  }
  return lastChar + 'A';
}

// CUSTOMERS
// ============================================================

function getCustomers() { return getSheetData('Customers'); }

function addCustomer(data) {
  const sheet = ss.getSheetByName('Customers');
  const id = generateId('CUS');
  const formattedPhone = formatPhoneBackend(data.Phone || '');
  sheet.appendRow([id, data.Name, formattedPhone, data.Email || '', data.Address || '', 0, 0, 0, fmtDate(new Date()), data.Notes || '']);
  return {success:true, id, Name:data.Name};
}

function updateCustomer(data) {
  const sheet = ss.getSheetByName('Customers');
  const row = findRowById('Customers', data.ID);
  if (row === -1) return {error:'Customer not found'};
  const formattedPhone = formatPhoneBackend(data.Phone || '');
  sheet.getRange(row, 2).setValue(data.Name);
  sheet.getRange(row, 3).setValue(formattedPhone);
  sheet.getRange(row, 4).setValue(data.Email || '');
  sheet.getRange(row, 5).setValue(data.Address || '');
  sheet.getRange(row, 10).setValue(data.Notes || '');
  return {success:true};
}

function deleteCustomer(data) {
  const sheet = ss.getSheetByName('Customers');
  const row = findRowById('Customers', data.ID);
  if (row === -1) return {error:'Customer not found'};
  sheet.deleteRow(row);
  return {success:true};
}

function getCustomerLedger(customerId) {
  const sales = getSheetData('Sales').filter(s => String(s.CustomerID) === String(customerId));
  const payments = getSheetData('Payments').filter(p => String(p.CustomerID) === String(customerId));
  return {sales, payments};
}

// ============================================================
// VOUCHERS
// ============================================================

function getVouchers() { return getSheetData('Vouchers'); }

function addVoucher(data) {
  const sheet = ss.getSheetByName('Vouchers');
  if (!sheet) return {error:'Vouchers sheet not found'};
  const id = generateId('VCH');
  sheet.appendRow([
    id, data.Code, data.Name || '',
    data.DiscountType || 'Fixed', Number(data.DiscountValue) || 0,
    Number(data.MinOrder) || 0, Number(data.MaxUses) || 0,
    0, 0,
    data.ValidFrom || '', data.ValidTo || '',
    data.Active !== false ? 'Yes' : 'No'
  ]);
  return {success:true, id};
}

function updateVoucher(data) {
  const sheet = ss.getSheetByName('Vouchers');
  if (!sheet) return {error:'Vouchers sheet not found'};
  const row = findRowById('Vouchers', data.ID);
  if (row === -1) return {error:'Voucher not found'};
  sheet.getRange(row, 2).setValue(data.Code);
  sheet.getRange(row, 3).setValue(data.Name || '');
  sheet.getRange(row, 4).setValue(data.DiscountType || 'Fixed');
  sheet.getRange(row, 5).setValue(Number(data.DiscountValue) || 0);
  sheet.getRange(row, 6).setValue(Number(data.MinOrder) || 0);
  sheet.getRange(row, 7).setValue(Number(data.MaxUses) || 0);
  sheet.getRange(row, 11).setValue(data.ValidFrom || '');
  sheet.getRange(row, 12).setValue(data.ValidTo || '');
  sheet.getRange(row, 12).setValue(data.Active ? 'Yes' : 'No');
  return {success:true};
}

function deleteVoucher(data) {
  const sheet = ss.getSheetByName('Vouchers');
  if (!sheet) return {error:'Vouchers sheet not found'};
  const row = findRowById('Vouchers', data.ID);
  if (row === -1) return {error:'Voucher not found'};
  sheet.deleteRow(row);
  return {success:true};
}

// ============================================================
// AFFILIATE COMMISSIONS
// ============================================================

function getAffiliateCommissions(affiliateId) {
  const all = getSheetData('AffiliateCommissions');
  if (affiliateId) return all.filter(c => String(c.AffiliateID) === String(affiliateId));
  return all;
}

function addAffiliateCommission(affiliateId, affiliateCode, affiliateName, invoiceNo, edition, commissionAmount) {
  const sheet = ss.getSheetByName('AffiliateCommissions');
  if (!sheet) return;
  const id = generateId('AFCM');
  sheet.appendRow([
    id, fmtDate(new Date()), affiliateId, affiliateCode, affiliateName,
    invoiceNo, edition, commissionAmount, 'Unpaid', '', '', ''
  ]);
  
  // Update affiliate totals
  const affSheet = ss.getSheetByName('Affiliates');
  const affData = affSheet.getDataRange().getValues();
  for (let i = 1; i < affData.length; i++) {
    if (String(affData[i][0]) === String(affiliateId)) {
      const totalComm = (Number(affData[i][12]) || 0) + commissionAmount;
      const paidComm = Number(affData[i][13]) || 0;
      affSheet.getRange(i + 1, 13).setValue(totalComm);
      affSheet.getRange(i + 1, 15).setValue(totalComm - paidComm);
      break;
    }
  }
}

function payAffiliateCommission(data) {
  const commSheet = ss.getSheetByName('AffiliateCommissions');
  const affSheet = ss.getSheetByName('Affiliates');
  if (!commSheet || !affSheet) return {error:'Required sheets not found'};
  
  const amount = Number(data.Amount) || 0;
  if (amount <= 0) return {error:'Invalid amount'};
  
  // Mark unpaid commissions as paid
  const commData = commSheet.getDataRange().getValues();
  let remaining = amount;
  const paidIds = [];
  
  for (let i = 1; i < commData.length; i++) {
    if (String(commData[i][2]) === String(data.AffiliateID) && commData[i][8] === 'Unpaid' && remaining > 0) {
      const commAmt = Number(commData[i][7]) || 0;
      if (commAmt <= remaining) {
        commSheet.getRange(i + 1, 9).setValue('Paid');
        commSheet.getRange(i + 1, 10).setValue(fmtDate(new Date()));
        remaining -= commAmt;
        paidIds.push(commData[i][0]);
      }
    }
  }
  
  // Update affiliate paid/pending
  const affRow = findRowById('Affiliates', data.AffiliateID);
  if (affRow !== -1) {
    const r = affSheet.getRange(affRow, 1, 1, 15).getValues()[0];
    const newPaid = (Number(r[13]) || 0) + amount;
    const totalComm = Number(r[12]) || 0;
    affSheet.getRange(affRow, 14).setValue(newPaid);
    affSheet.getRange(affRow, 15).setValue(Math.max(0, totalComm - newPaid));
  }
  
  // Add expense record
  const expSheet = ss.getSheetByName('Expenses');
  const expId = generateId('EXP');
  expSheet.appendRow([
    expId, fmtDate(new Date()), 'Affiliate Payout',
    'Commission payout to ' + (data.AffiliateName || data.AffiliateCode || ''),
    amount, data.AffiliateName || '', data.PaymentMethod || 'Cash',
    'Affiliate commission payout'
  ]);
  
  // Store expense ID in commission records
  paidIds.forEach(pid => {
    const prow = findRowById('AffiliateCommissions', pid);
    if (prow !== -1) commSheet.getRange(prow, 11).setValue(expId);
  });
  
  return {success:true, expenseId: expId, paidCount: paidIds.length};
}

function reverseAffiliatePayment(data) {
  const commSheet = ss.getSheetByName('AffiliateCommissions');
  const affSheet = ss.getSheetByName('Affiliates');
  if (!commSheet || !affSheet) return {error:'Required sheets not found'};
  
  const row = findRowById('AffiliateCommissions', data.CommissionID);
  if (row === -1) return {error:'Commission record not found'};
  
  const commData = commSheet.getRange(row, 1, 1, 12).getValues()[0];
  if (commData[8] !== 'Paid') return {error:'Commission is not in Paid status'};
  
  const commAmt = Number(commData[7]) || 0;
  
  commSheet.getRange(row, 9).setValue('Unpaid');
  commSheet.getRange(row, 10).setValue('');
  
  // Reverse affiliate totals
  const affRow = findRowById('Affiliates', commData[2]);
  if (affRow !== -1) {
    const r = affSheet.getRange(affRow, 1, 1, 15).getValues()[0];
    const newPaid = Math.max(0, (Number(r[13]) || 0) - commAmt);
    const totalComm = Number(r[12]) || 0;
    affSheet.getRange(affRow, 14).setValue(newPaid);
    affSheet.getRange(affRow, 15).setValue(totalComm - newPaid);
  }
  
  // Delete the associated expense if exists
  const expId = commData[10];
  if (expId) {
    const expRow = findRowById('Expenses', expId);
    if (expRow !== -1) {
      ss.getSheetByName('Expenses').deleteRow(expRow);
    }
  }
  
  return {success:true};
}

// ============================================================
// DELIVERIES
// ============================================================

function getDeliveries() { return getSheetData('Deliveries'); }

function updateDeliveryStatus(data) {
  const dlvSheet = ss.getSheetByName('Deliveries');
  const salesSheet = ss.getSheetByName('Sales');
  if (!dlvSheet) return {error:'Deliveries sheet not found'};
  
  const row = findRowById('Deliveries', data.ID);
  if (row === -1) return {error:'Delivery not found'};
  
  if (data.DeliveryStatus) dlvSheet.getRange(row, 11).setValue(data.DeliveryStatus);
  if (data.CourierService) dlvSheet.getRange(row, 9).setValue(data.CourierService);
  if (data.TrackingNo) dlvSheet.getRange(row, 10).setValue(data.TrackingNo);
  if (data.Notes) dlvSheet.getRange(row, 12).setValue(data.Notes);
  
  // Also update in Sales sheet
  if (data.InvoiceNo && salesSheet) {
    const salesData = salesSheet.getDataRange().getValues();
    for (let i = 1; i < salesData.length; i++) {
      if (String(salesData[i][1]) === String(data.InvoiceNo)) {
        if (data.DeliveryStatus) salesSheet.getRange(i + 1, 26).setValue(data.DeliveryStatus);
        if (data.CourierService) salesSheet.getRange(i + 1, 27).setValue(data.CourierService);
        if (data.TrackingNo) salesSheet.getRange(i + 1, 28).setValue(data.TrackingNo);
        break;
      }
    }
  }
  
  return {success:true};
}

function updateSaleDelivery(data) {
  const salesSheet = ss.getSheetByName('Sales');
  if (!salesSheet) return {error:'Sales sheet not found'};
  const salesData = salesSheet.getDataRange().getValues();
  for (let i = 1; i < salesData.length; i++) {
    if (String(salesData[i][1]) === String(data.InvoiceNo)) {
      if (data.DeliveryStatus) salesSheet.getRange(i + 1, 26).setValue(data.DeliveryStatus);
      if (data.CourierService) salesSheet.getRange(i + 1, 27).setValue(data.CourierService);
      if (data.TrackingNo) salesSheet.getRange(i + 1, 28).setValue(data.TrackingNo);
      break;
    }
  }
  return {success:true};
}

// ============================================================
// SALES & INVOICING (Enhanced with delivery, vouchers, commissions)
// ============================================================

function getSales() { return getSheetData('Sales'); }

function getSalesItems(invoiceNo) {
  const all = getSheetData('SalesItems');
  if (invoiceNo) return all.filter(i => String(i.InvoiceNo) === String(invoiceNo));
  return all;
}

function getNextInvoiceNumber() {
  const settings = getSettings();
  const prefix = settings.invoicePrefix || 'FF-';
  const nextNum = Number(settings.nextInvoiceNum) || 1001;
  return {invoiceNo: prefix + nextNum};
}

function createSale(data) {
  const salesSheet = ss.getSheetByName('Sales');
  const salesItemsSheet = ss.getSheetByName('SalesItems');
  const inventorySheet = ss.getSheetByName('Inventory');
  const settingsSheet = ss.getSheetByName('Settings');

  const settingsData = settingsSheet.getDataRange().getValues();
  let prefix = 'FF-', nextNum = 1001, numRow = -1;
  for (let i = 1; i < settingsData.length; i++) {
    if (settingsData[i][0] === 'invoicePrefix') prefix = settingsData[i][1];
    if (settingsData[i][0] === 'nextInvoiceNum') { nextNum = Number(settingsData[i][1]); numRow = i + 1; }
  }

  const invoiceNo = data.InvoiceNo || (prefix + nextNum);
  const saleId = generateId('SALEID');
  const now = new Date();

  const totalAmount = Number(data.TotalAmount) || 0;
  const discount = Number(data.Discount) || 0;
  const affDiscount = Number(data.AffiliateDiscount) || 0;
  const voucherDiscount = Number(data.VoucherDiscount) || 0;
  const deliveryCharge = Number(data.DeliveryCharge) || 0;
  const finalAmount = totalAmount - discount - affDiscount - voucherDiscount + deliveryCharge;
  const paidAmount = Number(data.PaidAmount) || 0;
  const dueAmount = Math.max(0, finalAmount - paidAmount);
  const status = dueAmount <= 0 ? 'Paid' : (paidAmount > 0 ? 'Partial' : 'Unpaid');

  let totalCost = 0;
  const itemsSummary = data.Items.map(i => {
    const pName = buildProductName(i.Club_Country, i.Type, i.Edition, i.Size);
    return pName + ' x' + i.Quantity;
  }).join(', ');

  const modSummary = data.Items.filter(i => i.Patch || i.NamePrint).map(i => {
    const parts = [];
    if (i.Patch) parts.push('Patch');
    if (i.NamePrint) parts.push('Name: ' + i.NamePrint);
    return buildProductName(i.Club_Country, i.Type, i.Edition, i.Size) + ' (' + parts.join(', ') + ')';
  }).join('; ');

  const inventoryData = inventorySheet.getDataRange().getValues();
  const settingsObj = getSettings();
  const patchCostFromSettings = Number(settingsObj.patchDefaultCost) || 0;
  const namePrintCostFromSettings = Number(settingsObj.namePrintDefaultCost) || 0;

  data.Items.forEach(item => {
    const itemId = generateId('SLI');
    const pName = buildProductName(item.Club_Country, item.Type, item.Edition, item.Size);
    const qty = Number(item.Quantity) || 1;
    const unitPrice = Number(item.UnitPrice) || 0;
    const costPrice = Number(item.CostPrice) || 0;
    const patchPrice = item.Patch ? (Number(item.PatchPrice) || 0) : 0;
    const namePrintPrice = item.NamePrint ? (Number(item.NamePrintPrice) || 0) : 0;
    const modTotal = (patchPrice + namePrintPrice) * qty;
    const lineTotal = (unitPrice * qty) + modTotal;

    const patchCost = item.Patch ? patchCostFromSettings : 0;
    const namePrintCost = item.NamePrint ? namePrintCostFromSettings : 0;
    const modCostTotal = (patchCost + namePrintCost) * qty;
    totalCost += (costPrice * qty) + modCostTotal;

    salesItemsSheet.appendRow([
      itemId, invoiceNo, item.ProductID || '', pName,
      item.Club_Country || '', item.Edition || '', item.Type || '', item.Size || '',
      qty, unitPrice, costPrice, unitPrice * qty,
      item.Patch ? 'Yes' : 'No', item.NamePrint || '',
      patchPrice, namePrintPrice, modTotal, lineTotal
    ]);

    if (item.ProductID) {
      for (let i = 1; i < inventoryData.length; i++) {
        if (String(inventoryData[i][0]) === String(item.ProductID)) {
          const curStock = Number(inventoryData[i][9]) || 0;
          const newStock = Math.max(0, curStock - qty);
          inventorySheet.getRange(i + 1, 10).setValue(newStock);
          inventorySheet.getRange(i + 1, 15).setValue(fmtDate(now));
          logStockChange(item.ProductID, pName, 'Sale', qty, curStock, newStock, invoiceNo, '');
          inventoryData[i][9] = newStock;
          break;
        }
      }
    }
  });

  // Affiliate commission calculation
  let totalCommission = 0;
  if (data.AffiliateCode) {
    const affSheet = ss.getSheetByName('Affiliates');
    const affData = affSheet.getDataRange().getValues();
    let affiliateId = '', affiliateName = '';
    const commRates = {};
    
    for (let i = 1; i < affData.length; i++) {
      if (String(affData[i][1]).toLowerCase() === String(data.AffiliateCode).toLowerCase()) {
        affiliateId = affData[i][0];
        affiliateName = affData[i][2];
        commRates['Fan Edition'] = Number(affData[i][8]) || 0;
        commRates['Player Edition'] = Number(affData[i][9]) || 0;
        commRates['Special Edition'] = Number(affData[i][10]) || 0;
        commRates['BD Premium'] = Number(affData[i][11]) || 0;
        
        // Update usage count and discount given
        affSheet.getRange(i + 1, 6).setValue((Number(affData[i][5]) || 0) + 1);
        affSheet.getRange(i + 1, 7).setValue((Number(affData[i][6]) || 0) + affDiscount);
        break;
      }
    }
    
    // Calculate commission per item based on edition
    if (affiliateId) {
      data.Items.forEach(item => {
        const edition = item.Edition || '';
        const qty = Number(item.Quantity) || 1;
        const commRate = commRates[edition] || 0;
        if (commRate > 0) {
          const itemComm = commRate * qty;
          totalCommission += itemComm;
          addAffiliateCommission(affiliateId, data.AffiliateCode, affiliateName, invoiceNo, edition, itemComm);
        }
      });
    }
  }
  
  const profit = finalAmount - totalCost - totalCommission;

  // Sale row with extended columns
  salesSheet.appendRow([
    saleId, invoiceNo, fmtDate(now),
    data.CustomerID || '', data.CustomerName || 'Walk-in Customer', formatPhoneBackend(data.CustomerPhone || ''),
    itemsSummary, modSummary,
    totalAmount, discount,
    data.AffiliateCode || '', affDiscount,
    finalAmount, paidAmount, dueAmount,
    data.PaymentMethod || 'Cash',
    status, totalCost, profit, '', data.Notes || '',
    // Extended columns (22-29)
    data.DeliveryType || '', deliveryCharge,
    data.VoucherCode || '', voucherDiscount,
    '', '', '', // DeliveryStatus, CourierService, TrackingNo (cols 26-28)
    data.DeliveryAddress || '' // DeliveryAddress (col 29)
  ]);

  // Customer update
  if (data.CustomerID) {
    const custSheet = ss.getSheetByName('Customers');
    const custRow = findRowById('Customers', data.CustomerID);
    if (custRow !== -1) {
      const r = custSheet.getRange(custRow, 1, 1, 10).getValues()[0];
      custSheet.getRange(custRow, 6).setValue((Number(r[5]) || 0) + 1);
      custSheet.getRange(custRow, 7).setValue((Number(r[6]) || 0) + finalAmount);
      custSheet.getRange(custRow, 8).setValue((Number(r[7]) || 0) + dueAmount);
    }
  }

  // Payment record
  if (paidAmount > 0) {
    const paymentsSheet = ss.getSheetByName('Payments');
    paymentsSheet.appendRow([
      generateId('PAY'), fmtDate(now), invoiceNo,
      data.CustomerID || '', data.CustomerName || 'Walk-in Customer',
      paidAmount, data.PaymentMethod || 'Cash', '', 'Payment at sale'
    ]);
  }

  // Voucher tracking
  if (data.VoucherCode && voucherDiscount > 0) {
    const vchSheet = ss.getSheetByName('Vouchers');
    if (vchSheet) {
      const vchData = vchSheet.getDataRange().getValues();
      for (let i = 1; i < vchData.length; i++) {
        if (String(vchData[i][1]).toLowerCase() === String(data.VoucherCode).toLowerCase()) {
          vchSheet.getRange(i + 1, 8).setValue((Number(vchData[i][7]) || 0) + 1);
          vchSheet.getRange(i + 1, 9).setValue((Number(vchData[i][8]) || 0) + voucherDiscount);
          break;
        }
      }
    }
  }

  // Delivery record
  if (data.DeliveryType && data.DeliveryType !== 'None') {
    const dlvSheet = ss.getSheetByName('Deliveries');
    if (dlvSheet) {
      dlvSheet.appendRow([
        generateId('DLV'), invoiceNo, fmtDate(now),
        data.CustomerName || 'Walk-in', formatPhoneBackend(data.CustomerPhone || ''),
        data.DeliveryAddress || '', data.DeliveryType, deliveryCharge,
        '', '', 'Processing', data.Notes || ''
      ]);
    }
  }

  if (numRow !== -1) {
    settingsSheet.getRange(numRow, 2).setValue(nextNum + 1);
  }

  return {success:true, invoiceNo, saleId, profit: Math.round(profit), commission: Math.round(totalCommission)};
}

// ============================================================
// REVERSE/CANCEL SALE
// ============================================================

function reverseSale(data) {
  const salesSheet = ss.getSheetByName('Sales');
  const salesItemsSheet = ss.getSheetByName('SalesItems');
  const inventorySheet = ss.getSheetByName('Inventory');
  const paymentsSheet = ss.getSheetByName('Payments');
  const custSheet = ss.getSheetByName('Customers');
  
  // Find sale
  const saleRow = findRowById('Sales', data.ID);
  if (saleRow === -1) return {error:'Sale not found'};
  
  const saleData = salesSheet.getRange(saleRow, 1, 1, 28).getValues()[0];
  const invoiceNo = saleData[1];
  const customerId = saleData[3];
  const finalAmount = Number(saleData[12]) || 0;
  const paidAmount = Number(saleData[13]) || 0;
  const dueAmount = Number(saleData[14]) || 0;
  const affiliateCode = saleData[10];
  const affDiscount = Number(saleData[11]) || 0;
  
  // 1. Restore stock for each item
  const allItems = salesItemsSheet.getDataRange().getValues();
  const headers = allItems[0];
  for (let i = 1; i < allItems.length; i++) {
    if (String(allItems[i][1]) === String(invoiceNo)) {
      const productId = allItems[i][2];
      const qty = Number(allItems[i][8]) || 0;
      if (productId && qty > 0) {
        const invRow = findRowById('Inventory', productId);
        if (invRow !== -1) {
          const curStock = Number(inventorySheet.getRange(invRow, 10).getValue()) || 0;
          const newStock = curStock + qty;
          inventorySheet.getRange(invRow, 10).setValue(newStock);
          inventorySheet.getRange(invRow, 15).setValue(fmtDate(new Date()));
          logStockChange(productId, allItems[i][3], 'Sale Reversed', qty, curStock, newStock, invoiceNo, 'Sale cancelled');
        }
      }
    }
  }
  
  // 2. Update customer totals
  if (customerId) {
    const custRow = findRowById('Customers', customerId);
    if (custRow !== -1) {
      const r = custSheet.getRange(custRow, 1, 1, 10).getValues()[0];
      custSheet.getRange(custRow, 6).setValue(Math.max(0, (Number(r[5]) || 0) - 1));
      custSheet.getRange(custRow, 7).setValue(Math.max(0, (Number(r[6]) || 0) - finalAmount));
      custSheet.getRange(custRow, 8).setValue(Math.max(0, (Number(r[7]) || 0) - dueAmount));
    }
  }
  
  // 3. Reverse affiliate commission
  if (affiliateCode) {
    const commSheet = ss.getSheetByName('AffiliateCommissions');
    if (commSheet) {
      const commData = commSheet.getDataRange().getValues();
      for (let i = commData.length - 1; i >= 1; i--) {
        if (String(commData[i][5]) === String(invoiceNo) && commData[i][8] === 'Unpaid') {
          commSheet.deleteRow(i + 1);
        }
      }
    }
    
    // Reverse affiliate usage count
    const affSheet = ss.getSheetByName('Affiliates');
    if (affSheet) {
      const affData = affSheet.getDataRange().getValues();
      for (let i = 1; i < affData.length; i++) {
        if (String(affData[i][1]).toLowerCase() === String(affiliateCode).toLowerCase()) {
          affSheet.getRange(i + 1, 6).setValue(Math.max(0, (Number(affData[i][5]) || 0) - 1));
          affSheet.getRange(i + 1, 7).setValue(Math.max(0, (Number(affData[i][6]) || 0) - affDiscount));
          // Recalc commission totals
          const remainingComms = getSheetData('AffiliateCommissions').filter(c => String(c.AffiliateID) === String(affData[i][0]));
          const totalComm = remainingComms.reduce((s, c) => s + (Number(c.CommissionAmount) || 0), 0);
          const paidComm = remainingComms.filter(c => c.Status === 'Paid').reduce((s, c) => s + (Number(c.CommissionAmount) || 0), 0);
          affSheet.getRange(i + 1, 13).setValue(totalComm);
          affSheet.getRange(i + 1, 14).setValue(paidComm);
          affSheet.getRange(i + 1, 15).setValue(totalComm - paidComm);
          break;
        }
      }
    }
  }
  
  // 4. Mark sale as Cancelled
  salesSheet.getRange(saleRow, 17).setValue('Cancelled');
  salesSheet.getRange(saleRow, 14).setValue(0);
  salesSheet.getRange(saleRow, 15).setValue(0);
  salesSheet.getRange(saleRow, 19).setValue(0);
  salesSheet.getRange(saleRow, 20).setValue(0);
  
  // 5. Delete delivery record
  const dlvSheet = ss.getSheetByName('Deliveries');
  if (dlvSheet) {
    const dlvData = dlvSheet.getDataRange().getValues();
    for (let i = dlvData.length - 1; i >= 1; i--) {
      if (String(dlvData[i][1]) === String(invoiceNo)) {
        dlvSheet.deleteRow(i + 1);
      }
    }
  }
  
  return {success:true, invoiceNo, message: 'Sale reversed. Stock restored. Customer updated.'};
}

// ============================================================
// PARTIAL REVERSE / RETURN (Return specific items from a sale)
// ============================================================
function partialReverseSale(data) {
  const salesSheet = ss.getSheetByName('Sales');
  const salesItemsSheet = ss.getSheetByName('SalesItems');
  const inventorySheet = ss.getSheetByName('Inventory');
  const custSheet = ss.getSheetByName('Customers');
  
  const saleRow = findRowById('Sales', data.SaleID);
  if (saleRow === -1) return { error: 'Sale not found' };
  
  const saleData = salesSheet.getRange(saleRow, 1, 1, 28).getValues()[0];
  const invoiceNo = saleData[1];
  const customerId = saleData[3];
  const customerName = saleData[4];
  const currentFinal = Number(saleData[12]) || 0;
  const currentPaid = Number(saleData[13]) || 0;
  const currentDue = Number(saleData[14]) || 0;
  const currentCost = Number(saleData[17]) || 0;
  const currentProfit = Number(saleData[18]) || 0;
  const affiliateCode = saleData[10];
  const isExchange = data.IsExchange === true;
  
  let totalRefund = 0;
  let totalCostReturned = 0;
  let totalQtyReturned = 0;
  const returnLog = [];
  
  // Process each returned item
  data.Items.forEach(retItem => {
    const retQty = Number(retItem.ReturnQty) || 0;
    if (retQty <= 0) return;
    
    const unitPrice = Number(retItem.UnitPrice) || 0;
    const costPrice = Number(retItem.CostPrice) || 0;
    const modPerUnit = Number(retItem.ModPerUnit) || 0;
    
    const itemRefund = (unitPrice + modPerUnit) * retQty;
    const itemCost = costPrice * retQty;
    
    totalRefund += itemRefund;
    totalCostReturned += itemCost;
    totalQtyReturned += retQty;
    
    // 1. Restore stock
    if (retItem.ProductID) {
      const invRow = findRowById('Inventory', retItem.ProductID);
      if (invRow !== -1) {
        const curStock = Number(inventorySheet.getRange(invRow, 10).getValue()) || 0;
        const newStock = curStock + retQty;
        inventorySheet.getRange(invRow, 10).setValue(newStock);
        inventorySheet.getRange(invRow, 15).setValue(fmtDate(new Date()));
        logStockChange(
          retItem.ProductID,
          retItem.ProductName,
          isExchange ? 'Partial Exchange' : 'Partial Return',
          retQty,
          curStock,
          newStock,
          invoiceNo,
          (data.Notes || '') + (isExchange ? ' [EXCHANGE]' : ' [RETURN]')
        );
      }
    }
    
    // 2. Update SalesItems sheet — reduce qty or remove row
    if (retItem.SalesItemID) {
      const siRow = findRowById('SalesItems', retItem.SalesItemID);
      if (siRow !== -1) {
        const siData = salesItemsSheet.getRange(siRow, 1, 1, 18).getValues()[0];
        const origQty = Number(siData[8]) || 0;
        const newQty = origQty - retQty;
        
        if (newQty <= 0) {
          // Fully returned — append a marker note in the row but keep for audit
          salesItemsSheet.getRange(siRow, 9).setValue(0);
          salesItemsSheet.getRange(siRow, 12).setValue(0);
          salesItemsSheet.getRange(siRow, 17).setValue(0);
          salesItemsSheet.getRange(siRow, 18).setValue(0);
        } else {
          // Partial return — reduce quantity and recalculate line totals
          const origUnitPrice = Number(siData[9]) || 0;
          const origPatchPrice = Number(siData[14]) || 0;
          const origNamePrintPrice = Number(siData[15]) || 0;
          const newSubtotal = origUnitPrice * newQty;
          const newModTotal = (origPatchPrice + origNamePrintPrice) * newQty;
          const newLineTotal = newSubtotal + newModTotal;
          
          salesItemsSheet.getRange(siRow, 9).setValue(newQty);
          salesItemsSheet.getRange(siRow, 12).setValue(newSubtotal);
          salesItemsSheet.getRange(siRow, 17).setValue(newModTotal);
          salesItemsSheet.getRange(siRow, 18).setValue(newLineTotal);
        }
      }
    }
    
    returnLog.push(`${retItem.ProductName} x${retQty}`);
  });
  
  // 3. Update Sales record
  const newFinal = Math.max(0, currentFinal - totalRefund);
  const newCost = Math.max(0, currentCost - totalCostReturned);
  
  // Recalculate paid/due based on refund (only if NOT exchange)
  let newPaid = currentPaid;
  let newDue = currentDue;
  
  if (!isExchange) {
    // Refund reduces paid amount first, then due
    if (currentPaid >= totalRefund) {
      newPaid = currentPaid - totalRefund;
      newDue = Math.max(0, newFinal - newPaid);
    } else {
      // Partial refund covered by paid, rest reduces due
      const remainingRefund = totalRefund - currentPaid;
      newPaid = 0;
      newDue = Math.max(0, currentDue - remainingRefund);
    }
  } else {
    // Exchange: keep paid as is, just reduce due if applicable
    newDue = Math.max(0, newFinal - newPaid);
  }
  
  const newProfit = newFinal - newCost;
  let newStatus = 'Paid';
  if (newFinal === 0) newStatus = 'Cancelled';
  else if (newDue > 0 && newPaid > 0) newStatus = 'Partial';
  else if (newDue > 0) newStatus = 'Unpaid';
  
  salesSheet.getRange(saleRow, 13).setValue(newFinal);
  salesSheet.getRange(saleRow, 14).setValue(newPaid);
  salesSheet.getRange(saleRow, 15).setValue(newDue);
  salesSheet.getRange(saleRow, 17).setValue(newStatus);
  salesSheet.getRange(saleRow, 18).setValue(newCost);
  salesSheet.getRange(saleRow, 19).setValue(newProfit);
  
  // Append return note
  const existingNotes = saleData[20] || '';
  const returnNote = `[${isExchange ? 'EXCHANGE' : 'RETURN'} ${fmtDate(new Date())}] Returned: ${returnLog.join(', ')}. Refund: ${totalRefund}. ${data.Notes || ''}`;
  salesSheet.getRange(saleRow, 21).setValue(existingNotes ? existingNotes + ' | ' + returnNote : returnNote);
  
  // Rebuild items summary in Sales row
  const remainingItems = getSheetData('SalesItems').filter(i => 
    String(i.InvoiceNo) === String(invoiceNo) && Number(i.Quantity) > 0
  );
  const newItemsSummary = remainingItems.map(i => `${i.ProductName} x${i.Quantity}`).join(', ');
  salesSheet.getRange(saleRow, 7).setValue(newItemsSummary || 'All items returned');
  
  // 4. Update customer totals
  if (customerId) {
    const custRow = findRowById('Customers', customerId);
    if (custRow !== -1) {
      const r = custSheet.getRange(custRow, 1, 1, 10).getValues()[0];
      const newTotalSpent = Math.max(0, (Number(r[6]) || 0) - totalRefund);
      const pendingDelta = currentDue - newDue;
      const newPending = Math.max(0, (Number(r[7]) || 0) - pendingDelta);
      custSheet.getRange(custRow, 7).setValue(newTotalSpent);
      custSheet.getRange(custRow, 8).setValue(newPending);
    }
  }
  
  // 5. Reverse proportional affiliate commissions
  if (affiliateCode) {
    const commSheet = ss.getSheetByName('AffiliateCommissions');
    const affSheet = ss.getSheetByName('Affiliates');
    if (commSheet && affSheet) {
      const commData = commSheet.getDataRange().getValues();
      const affData = affSheet.getDataRange().getValues();
      
      // Find affiliate
      let affRow = -1, affId = '';
      for (let i = 1; i < affData.length; i++) {
        if (String(affData[i][1]).toLowerCase() === String(affiliateCode).toLowerCase()) {
          affRow = i + 1;
          affId = affData[i][0];
          break;
        }
      }
      
      // For each returned item, reverse commission for that edition proportionally
      data.Items.forEach(retItem => {
        const retQty = Number(retItem.ReturnQty) || 0;
        const edition = retItem.Edition || '';
        if (retQty <= 0 || !edition) return;
        
        // Find unpaid commissions matching this invoice + edition, deduct proportionally
        let qtyToReverse = retQty;
        for (let i = commData.length - 1; i >= 1 && qtyToReverse > 0; i--) {
          if (String(commData[i][5]) === String(invoiceNo) &&
              String(commData[i][6]) === String(edition) &&
              commData[i][8] === 'Unpaid') {
            const commAmt = Number(commData[i][7]) || 0;
            // Get commission rate per unit from affiliate
            let commRatePerUnit = 0;
            if (affRow !== -1) {
              const editionColMap = {
                'Fan Edition': 8,
                'Player Edition': 9,
                'Special Edition': 10,
                'BD Premium': 11
              };
              const colIdx = editionColMap[edition];
              if (colIdx !== undefined) commRatePerUnit = Number(affData[affRow - 1][colIdx]) || 0;
            }
            
            const reverseAmt = commRatePerUnit * qtyToReverse;
            const newCommAmt = Math.max(0, commAmt - reverseAmt);
            
            if (newCommAmt <= 0) {
              commSheet.deleteRow(i + 1);
              commData.splice(i, 1);
            } else {
              commSheet.getRange(i + 1, 8).setValue(newCommAmt);
            }
            qtyToReverse = 0;
            break;
          }
        }
      });
      
      // Recalculate affiliate totals
      if (affRow !== -1 && affId) {
        const remainingComms = getSheetData('AffiliateCommissions').filter(c => String(c.AffiliateID) === String(affId));
        const totalComm = remainingComms.reduce((s, c) => s + (Number(c.CommissionAmount) || 0), 0);
        const paidComm = remainingComms.filter(c => c.Status === 'Paid').reduce((s, c) => s + (Number(c.CommissionAmount) || 0), 0);
        affSheet.getRange(affRow, 13).setValue(totalComm);
        affSheet.getRange(affRow, 14).setValue(paidComm);
        affSheet.getRange(affRow, 15).setValue(Math.max(0, totalComm - paidComm));
      }
    }
  }
  
  // 6. Record refund as negative payment (only if not exchange and refund > 0)
  if (!isExchange && totalRefund > 0) {
    const paymentsSheet = ss.getSheetByName('Payments');
    paymentsSheet.appendRow([
      generateId('PAY'), fmtDate(new Date()), invoiceNo,
      customerId || '', customerName || '',
      -totalRefund,
      data.RefundMethod || 'Cash',
      '',
      `REFUND — Partial return: ${returnLog.join(', ')}. ${data.Notes || ''}`
    ]);
  }
  
  // 7. Update delivery record if all items returned
  if (newFinal === 0) {
    const dlvSheet = ss.getSheetByName('Deliveries');
    if (dlvSheet) {
      const dlvData = dlvSheet.getDataRange().getValues();
      for (let i = 1; i < dlvData.length; i++) {
        if (String(dlvData[i][1]) === String(invoiceNo)) {
          dlvSheet.getRange(i + 1, 11).setValue('Returned');
          break;
        }
      }
    }
  }
  
  return {
    success: true,
    invoiceNo,
    itemsReturned: totalQtyReturned,
    refundAmount: Math.round(totalRefund),
    isExchange,
    newFinalAmount: Math.round(newFinal),
    message: isExchange
      ? `Exchange processed. ${totalQtyReturned} item(s) returned to stock. Create new sale for replacement.`
      : `Return processed. ${totalQtyReturned} item(s) refunded. Refund: ${totalRefund}`
  };
}

// ============================================================
// EXCHANGE & REPLACE — Atomic operation: return + new sale
// ============================================================
function exchangeAndReplace(data) {
  const salesSheet = ss.getSheetByName('Sales');
  const salesItemsSheet = ss.getSheetByName('SalesItems');
  const inventorySheet = ss.getSheetByName('Inventory');
  const custSheet = ss.getSheetByName('Customers');
  const settingsSheet = ss.getSheetByName('Settings');
  
  // Validate the original sale
  const origSaleRow = findRowById('Sales', data.SaleID);
  if (origSaleRow === -1) return { error: 'Original sale not found' };
  
  const origSaleData = salesSheet.getRange(origSaleRow, 1, 1, 28).getValues()[0];
  const origInvoiceNo = origSaleData[1];
  const customerId = data.CustomerID || origSaleData[3];
  const customerName = data.CustomerName || origSaleData[4];
  const customerPhone = data.CustomerPhone || origSaleData[5];
  const affiliateCode = origSaleData[10];
  
  const now = new Date();
  
  // ========================================
  // PART 1: Process Returns
  // ========================================
  let totalReturnCredit = 0;
  let totalReturnCost = 0;
  let totalReturnQty = 0;
  const returnLog = [];
  
  data.ReturnItems.forEach(retItem => {
    const retQty = Number(retItem.ReturnQty) || 0;
    if (retQty <= 0) return;
    
    const unitPrice = Number(retItem.UnitPrice) || 0;
    const costPrice = Number(retItem.CostPrice) || 0;
    const modPerUnit = Number(retItem.ModPerUnit) || 0;
    
    const itemCredit = (unitPrice + modPerUnit) * retQty;
    const itemCost = costPrice * retQty;
    
    totalReturnCredit += itemCredit;
    totalReturnCost += itemCost;
    totalReturnQty += retQty;
    
    // Restore stock
    if (retItem.ProductID) {
      const invRow = findRowById('Inventory', retItem.ProductID);
      if (invRow !== -1) {
        const curStock = Number(inventorySheet.getRange(invRow, 10).getValue()) || 0;
        const newStock = curStock + retQty;
        inventorySheet.getRange(invRow, 10).setValue(newStock);
        inventorySheet.getRange(invRow, 15).setValue(fmtDate(now));
        logStockChange(
          retItem.ProductID,
          retItem.ProductName,
          'Exchange Return',
          retQty,
          curStock,
          newStock,
          origInvoiceNo,
          'Exchange & Replace: ' + (data.Notes || '')
        );
      }
    }
    
    // Update SalesItems for original sale
    if (retItem.SalesItemID) {
      const siRow = findRowById('SalesItems', retItem.SalesItemID);
      if (siRow !== -1) {
        const siData = salesItemsSheet.getRange(siRow, 1, 1, 18).getValues()[0];
        const origQty = Number(siData[8]) || 0;
        const newQty = origQty - retQty;
        
        if (newQty <= 0) {
          salesItemsSheet.getRange(siRow, 9).setValue(0);
          salesItemsSheet.getRange(siRow, 12).setValue(0);
          salesItemsSheet.getRange(siRow, 17).setValue(0);
          salesItemsSheet.getRange(siRow, 18).setValue(0);
        } else {
          const origUnitPrice = Number(siData[9]) || 0;
          const origPatchPrice = Number(siData[14]) || 0;
          const origNamePrintPrice = Number(siData[15]) || 0;
          const newSubtotal = origUnitPrice * newQty;
          const newModTotal = (origPatchPrice + origNamePrintPrice) * newQty;
          const newLineTotal = newSubtotal + newModTotal;
          
          salesItemsSheet.getRange(siRow, 9).setValue(newQty);
          salesItemsSheet.getRange(siRow, 12).setValue(newSubtotal);
          salesItemsSheet.getRange(siRow, 17).setValue(newModTotal);
          salesItemsSheet.getRange(siRow, 18).setValue(newLineTotal);
        }
      }
    }
    
    returnLog.push(`${retItem.ProductName} x${retQty}`);
  });
  
  // Update original sale's adjusted totals
  const origFinal = Number(origSaleData[12]) || 0;
  const origCost = Number(origSaleData[17]) || 0;
  const newOrigFinal = Math.max(0, origFinal - totalReturnCredit);
  const newOrigCost = Math.max(0, origCost - totalReturnCost);
  const newOrigProfit = newOrigFinal - newOrigCost;
  
  salesSheet.getRange(origSaleRow, 13).setValue(newOrigFinal);
  salesSheet.getRange(origSaleRow, 18).setValue(newOrigCost);
  salesSheet.getRange(origSaleRow, 19).setValue(newOrigProfit);
  
  // Update items summary
  const remainingItems = getSheetData('SalesItems').filter(i =>
    String(i.InvoiceNo) === String(origInvoiceNo) && Number(i.Quantity) > 0
  );
  const newItemsSummary = remainingItems.map(i => `${i.ProductName} x${i.Quantity}`).join(', ');
  salesSheet.getRange(origSaleRow, 7).setValue(newItemsSummary || 'All items returned');
  
  // ========================================
  // PART 2: Reverse proportional commissions on returned items
  // ========================================
  if (affiliateCode) {
    const commSheet = ss.getSheetByName('AffiliateCommissions');
    const affSheet = ss.getSheetByName('Affiliates');
    if (commSheet && affSheet) {
      const affData = affSheet.getDataRange().getValues();
      let affRow = -1, affId = '';
      for (let i = 1; i < affData.length; i++) {
        if (String(affData[i][1]).toLowerCase() === String(affiliateCode).toLowerCase()) {
          affRow = i + 1;
          affId = affData[i][0];
          break;
        }
      }
      
      data.ReturnItems.forEach(retItem => {
        const retQty = Number(retItem.ReturnQty) || 0;
        const edition = retItem.Edition || '';
        if (retQty <= 0 || !edition || affRow === -1) return;
        
        const editionColMap = {
          'Fan Edition': 8,
          'Player Edition': 9,
          'Special Edition': 10,
          'BD Premium': 11
        };
        const colIdx = editionColMap[edition];
        if (colIdx === undefined) return;
        const commRatePerUnit = Number(affData[affRow - 1][colIdx]) || 0;
        if (commRatePerUnit <= 0) return;
        
        const reverseAmt = commRatePerUnit * retQty;
        
        // Find unpaid commission row matching invoice + edition
        const commData = commSheet.getDataRange().getValues();
        for (let i = commData.length - 1; i >= 1; i--) {
          if (String(commData[i][5]) === String(origInvoiceNo) &&
              String(commData[i][6]) === String(edition) &&
              commData[i][8] === 'Unpaid') {
            const commAmt = Number(commData[i][7]) || 0;
            const newCommAmt = Math.max(0, commAmt - reverseAmt);
            if (newCommAmt <= 0) {
              commSheet.deleteRow(i + 1);
            } else {
              commSheet.getRange(i + 1, 8).setValue(newCommAmt);
            }
            break;
          }
        }
      });
      
      // Recalculate affiliate totals
      if (affRow !== -1 && affId) {
        const remainingComms = getSheetData('AffiliateCommissions').filter(c => String(c.AffiliateID) === String(affId));
        const totalComm = remainingComms.reduce((s, c) => s + (Number(c.CommissionAmount) || 0), 0);
        const paidComm = remainingComms.filter(c => c.Status === 'Paid').reduce((s, c) => s + (Number(c.CommissionAmount) || 0), 0);
        affSheet.getRange(affRow, 13).setValue(totalComm);
        affSheet.getRange(affRow, 14).setValue(paidComm);
        affSheet.getRange(affRow, 15).setValue(Math.max(0, totalComm - paidComm));
      }
    }
  }
  
  // ========================================
  // PART 3: Create Replacement Sale
  // ========================================
  const settingsData = settingsSheet.getDataRange().getValues();
  let prefix = 'FF-', nextNum = 1001, numRow = -1;
  for (let i = 1; i < settingsData.length; i++) {
    if (settingsData[i][0] === 'invoicePrefix') prefix = settingsData[i][1];
    if (settingsData[i][0] === 'nextInvoiceNum') { nextNum = Number(settingsData[i][1]); numRow = i + 1; }
  }
  
  const newInvoiceNo = prefix + nextNum + '-EX';
  const newSaleId = generateId('SALEID');
  
  let replaceSubtotal = 0;
  let replaceCost = 0;
  const replaceItemsSummary = [];
  const replaceModSummary = [];
  
  const settingsObj = getSettings();
  const patchCostFromSettings = Number(settingsObj.patchDefaultCost) || 0;
  const namePrintCostFromSettings = Number(settingsObj.namePrintDefaultCost) || 0;
  
  data.ReplaceItems.forEach(item => {
    const itemId = generateId('SLI');
    const pName = buildProductName(item.Club_Country, item.Type, item.Edition, item.Size);
    const qty = Number(item.Quantity) || 1;
    const unitPrice = Number(item.UnitPrice) || 0;
    const costPrice = Number(item.CostPrice) || 0;
    const patchPrice = item.Patch ? (Number(item.PatchPrice) || 0) : 0;
    const namePrintPrice = item.NamePrint ? (Number(item.NamePrintPrice) || 0) : 0;
    const modTotal = (patchPrice + namePrintPrice) * qty;
    const lineTotal = (unitPrice * qty) + modTotal;
    
    const patchCost = item.Patch ? patchCostFromSettings : 0;
    const namePrintCost = item.NamePrint ? namePrintCostFromSettings : 0;
    const modCostTotal = (patchCost + namePrintCost) * qty;
    
    replaceSubtotal += lineTotal;
    replaceCost += (costPrice * qty) + modCostTotal;
    replaceItemsSummary.push(pName + ' x' + qty);
    
    if (item.Patch || item.NamePrint) {
      const parts = [];
      if (item.Patch) parts.push('Patch');
      if (item.NamePrint) parts.push('Name: ' + item.NamePrint);
      replaceModSummary.push(pName + ' (' + parts.join(', ') + ')');
    }
    
    salesItemsSheet.appendRow([
      itemId, newInvoiceNo, item.ProductID || '', pName,
      item.Club_Country || '', item.Edition || '', item.Type || '', item.Size || '',
      qty, unitPrice, costPrice, unitPrice * qty,
      item.Patch ? 'Yes' : 'No', item.NamePrint || '',
      patchPrice, namePrintPrice, modTotal, lineTotal
    ]);
    
    // Deduct stock
    if (item.ProductID) {
      const invRow = findRowById('Inventory', item.ProductID);
      if (invRow !== -1) {
        const curStock = Number(inventorySheet.getRange(invRow, 10).getValue()) || 0;
        const newStock = Math.max(0, curStock - qty);
        inventorySheet.getRange(invRow, 10).setValue(newStock);
        inventorySheet.getRange(invRow, 15).setValue(fmtDate(now));
        logStockChange(item.ProductID, pName, 'Exchange Sale', qty, curStock, newStock, newInvoiceNo, 'Replacement for ' + origInvoiceNo);
      }
    }
  });
  
  const extraDiscount = Number(data.ExtraDiscount) || 0;
  const adjustedReplaceTotal = Math.max(0, replaceSubtotal - extraDiscount);
  const settlement = adjustedReplaceTotal - totalReturnCredit;
  // settlement > 0 = customer pays, settlement < 0 = refund, settlement === 0 = even
  
  let paidAmt = 0;
  let dueAmt = 0;
  let newStatus = 'Paid';
  
  if (settlement > 0) {
    // Customer pays the difference — assume paid in full at exchange time
    paidAmt = settlement;
    dueAmt = 0;
    newStatus = 'Paid';
  } else if (settlement < 0) {
    // Refund — record as paid to keep accounting clean (refund handled via negative payment)
    paidAmt = adjustedReplaceTotal;
    dueAmt = 0;
    newStatus = 'Paid';
  } else {
    // Even exchange
    paidAmt = adjustedReplaceTotal;
    dueAmt = 0;
    newStatus = 'Paid';
  }
  
  const replaceProfit = adjustedReplaceTotal - replaceCost;
  
  // Append the replacement sale
  salesSheet.appendRow([
    newSaleId, newInvoiceNo, fmtDate(now),
    customerId || '', customerName, customerPhone || '',
    replaceItemsSummary.join(', '), replaceModSummary.join('; '),
    replaceSubtotal, extraDiscount,
    '', 0, // no affiliate on exchange replacement
    adjustedReplaceTotal, paidAmt, dueAmt,
    data.SettleMethod || 'Cash',
    newStatus, replaceCost, replaceProfit, '',
    `EXCHANGE for ${origInvoiceNo}. Returned: ${returnLog.join(', ')}. ${data.Notes || ''}`,
    '', 0, '', 0, '', '', '', '' // delivery columns + address
  ]);
  
  // Update invoice counter
  if (numRow !== -1) {
    settingsSheet.getRange(numRow, 2).setValue(nextNum + 1);
  }
  
  // ========================================
  // PART 4: Adjust customer totals
  // ========================================
  if (customerId) {
    const custRow = findRowById('Customers', customerId);
    if (custRow !== -1) {
      const r = custSheet.getRange(custRow, 1, 1, 10).getValues()[0];
      // Net change: -returnCredit (from original) + adjustedReplaceTotal (from new sale)
      const totalSpentDelta = adjustedReplaceTotal - totalReturnCredit;
      const newTotalSpent = Math.max(0, (Number(r[6]) || 0) + totalSpentDelta);
      // Purchase count: replacement adds 1
      const newPurchases = (Number(r[5]) || 0) + 1;
      custSheet.getRange(custRow, 6).setValue(newPurchases);
      custSheet.getRange(custRow, 7).setValue(newTotalSpent);
    }
  }
  
  // ========================================
  // PART 5: Record settlement payment
  // ========================================
  const paymentsSheet = ss.getSheetByName('Payments');
  
  if (settlement > 0) {
    // Customer paid extra
    paymentsSheet.appendRow([
      generateId('PAY'), fmtDate(now), newInvoiceNo,
      customerId || '', customerName,
      settlement, data.SettleMethod || 'Cash', '',
      `Exchange settlement (extra paid by customer)`
    ]);
  } else if (settlement < 0) {
    // Refund issued
    paymentsSheet.appendRow([
      generateId('PAY'), fmtDate(now), origInvoiceNo,
      customerId || '', customerName,
      settlement, data.SettleMethod || 'Cash', '',
      `Exchange refund — replaced with ${newInvoiceNo}`
    ]);
  }
  // For even exchange: log full replacement as payment paired with original credit (no net cash)
  if (settlement === 0 && adjustedReplaceTotal > 0) {
    paymentsSheet.appendRow([
      generateId('PAY'), fmtDate(now), newInvoiceNo,
      customerId || '', customerName,
      adjustedReplaceTotal, 'Exchange Credit', '',
      `Even exchange — credit from ${origInvoiceNo}`
    ]);
    paymentsSheet.appendRow([
      generateId('PAY'), fmtDate(now), origInvoiceNo,
      customerId || '', customerName,
      -totalReturnCredit, 'Exchange Credit', '',
      `Credit applied to ${newInvoiceNo}`
    ]);
  }
  
  // ========================================
  // PART 6: Append exchange note to original sale
  // ========================================
  const existingNotes = origSaleData[20] || '';
  const exchangeNote = `[EXCHANGE ${fmtDate(now)}] Replaced with ${newInvoiceNo}. Returned: ${returnLog.join(', ')}. ${data.Notes || ''}`;
  salesSheet.getRange(origSaleRow, 21).setValue(existingNotes ? existingNotes + ' | ' + exchangeNote : exchangeNote);
  
  return {
    success: true,
    originalInvoiceNo: origInvoiceNo,
    newInvoiceNo: newInvoiceNo,
    newSaleId: newSaleId,
    itemsReturned: totalReturnQty,
    returnCredit: Math.round(totalReturnCredit),
    replaceTotal: Math.round(adjustedReplaceTotal),
    settlement: Math.round(settlement), // positive=customer pays, negative=refund, 0=even
    message: `Exchange complete. New invoice: ${newInvoiceNo}`
  };
}

function addManualSale(data) {
  const salesSheet = ss.getSheetByName('Sales');
  const saleId = generateId('SALEID');
  const now = new Date();
  const finalAmt = Number(data.FinalAmount) || 0;
  const paidAmt = Number(data.PaidAmount) || 0;
  const dueAmt = Math.max(0, finalAmt - paidAmt);
  const status = dueAmt <= 0 ? 'Paid' : (paidAmt > 0 ? 'Partial' : 'Unpaid');
  const costTotal = Number(data.CostTotal) || 0;
  const profit = finalAmt - costTotal;

  salesSheet.appendRow([
    saleId, data.InvoiceNo || 'MANUAL-' + saleId.substr(0,6), fmtDate(now),
    data.CustomerID || '', data.CustomerName || 'Walk-in', formatPhoneBackend(data.CustomerPhone || ''),
    data.Items || 'Manual Entry', data.Modifications || '',
    finalAmt, Number(data.Discount) || 0, '', 0,
    finalAmt, paidAmt, dueAmt, data.PaymentMethod || 'Cash',
    status, costTotal, profit, '', data.Notes || 'Manually added',
    '', 0, '', 0, '', '', '', '' // delivery columns + address
  ]);

  if (data.CustomerID) {
    const custSheet = ss.getSheetByName('Customers');
    const custRow = findRowById('Customers', data.CustomerID);
    if (custRow !== -1) {
      const r = custSheet.getRange(custRow, 1, 1, 10).getValues()[0];
      custSheet.getRange(custRow, 6).setValue((Number(r[5]) || 0) + 1);
      custSheet.getRange(custRow, 7).setValue((Number(r[6]) || 0) + finalAmt);
      custSheet.getRange(custRow, 8).setValue((Number(r[7]) || 0) + dueAmt);
    }
  }

  if (paidAmt > 0) {
    ss.getSheetByName('Payments').appendRow([
      generateId('PAY'), fmtDate(now), data.InvoiceNo || 'MANUAL-' + saleId.substr(0,6),
      data.CustomerID || '', data.CustomerName || 'Walk-in',
      paidAmt, data.PaymentMethod || 'Cash', '', 'Manual sale payment'
    ]);
  }
  return {success:true, saleId};
}

function updateSaleProfit(data) {
  const sheet = ss.getSheetByName('Sales');
  const row = findRowById('Sales', data.ID);
  if (row === -1) return {error:'Sale not found'};
  sheet.getRange(row, 19).setValue(Number(data.Profit) || 0);
  sheet.getRange(row, 20).setValue(Number(data.ProfitOverride) || Number(data.Profit) || 0);
  return {success:true};
}

// ============================================================
// ADJUST SALE AMOUNT — Edit final amount of an existing sale
// (e.g., customer paid less than invoiced, on-the-spot negotiation, etc.)
// ============================================================
function adjustSaleAmount(data) {
  const salesSheet = ss.getSheetByName('Sales');
  const custSheet = ss.getSheetByName('Customers');
  const paymentsSheet = ss.getSheetByName('Payments');
  
  const saleRow = findRowById('Sales', data.ID);
  if (saleRow === -1) return { error: 'Sale not found' };
  
  // Read current sale data
  const saleData = salesSheet.getRange(saleRow, 1, 1, 28).getValues()[0];
  const invoiceNo = saleData[1];
  const customerId = saleData[3];
  const customerName = saleData[4];
  const currentTotal = Number(saleData[8]) || 0;       // col 9: TotalAmount (subtotal)
  const currentDiscount = Number(saleData[9]) || 0;    // col 10: Discount
  const affDiscount = Number(saleData[11]) || 0;       // col 12: AffiliateDiscount
  const voucherDiscount = Number(saleData[24]) || 0;   // col 25: VoucherDiscount
  const deliveryCharge = Number(saleData[22]) || 0;    // col 23: DeliveryCharge
  const currentFinal = Number(saleData[12]) || 0;      // col 13: FinalAmount
  const currentPaid = Number(saleData[13]) || 0;       // col 14: PaidAmount
  const currentDue = Number(saleData[14]) || 0;        // col 15: DueAmount
  const currentCost = Number(saleData[17]) || 0;       // col 18: CostTotal
  const status = saleData[16];                          // col 17: Status
  const affiliateCode = saleData[10];                   // col 11: AffiliateCode
  const existingNotes = saleData[20] || '';            // col 21: Notes
  
  // Validation
  if (status === 'Cancelled') return { error: 'Cannot adjust a cancelled sale' };
  
  const newFinal = Number(data.NewFinalAmount);
  if (isNaN(newFinal) || newFinal < 0) return { error: 'Invalid new amount' };
  
  const adjustmentAmount = currentFinal - newFinal; // positive = discount given
  const adjustmentReason = data.Reason || 'Manual adjustment';
  const adjustmentMethod = data.AdjustmentMethod || 'Price Reduction'; 
  // Methods: 'Price Reduction' (treat as extra discount), 'Refund' (record negative payment)
  
  if (Math.abs(adjustmentAmount) < 0.01) {
    return { error: 'New amount is the same as current amount — no change needed' };
  }
  
  // ============================================================
  // 1. Update Sales sheet
  // ============================================================
  
  // Determine new paid/due based on adjustment method
  let newPaid = currentPaid;
  let newDue;
  
  if (adjustmentMethod === 'Refund' && adjustmentAmount > 0) {
    // Customer was over-paid → refund the difference
    // The over-payment is recorded as a negative payment
    newPaid = currentPaid - adjustmentAmount;
    if (newPaid < 0) newPaid = 0;
    newDue = Math.max(0, newFinal - newPaid);
  } else {
    // Price Reduction: just lower the final amount, paid stays same
    // Due is recalculated
    newDue = Math.max(0, newFinal - newPaid);
    // If customer overpaid (paid > new final), cap paid at new final
    if (newPaid > newFinal) {
      newPaid = newFinal;
      newDue = 0;
    }
  }
  
  // Calculate the "effective extra discount" — for the discount column
  // Original subtotal stays the same, but we record adjustment as an extra discount
  const newAdjustedDiscount = currentDiscount + adjustmentAmount;
  // Note: we keep TotalAmount (subtotal) unchanged, but increase Discount column
  
  // Recalculate cost & profit
  // Cost doesn't change (items sold are same), only revenue/profit changes
  const newProfit = newFinal - currentCost;
  
  // New status
  let newStatus = 'Paid';
  if (newDue > 0 && newPaid > 0) newStatus = 'Partial';
  else if (newDue > 0) newStatus = 'Unpaid';
  else if (newFinal === 0) newStatus = 'Cancelled';
  
  // Write updates
  salesSheet.getRange(saleRow, 10).setValue(newAdjustedDiscount); // Discount column
  salesSheet.getRange(saleRow, 13).setValue(newFinal);            // FinalAmount
  salesSheet.getRange(saleRow, 14).setValue(newPaid);             // PaidAmount
  salesSheet.getRange(saleRow, 15).setValue(newDue);              // DueAmount
  salesSheet.getRange(saleRow, 17).setValue(newStatus);           // Status
  salesSheet.getRange(saleRow, 19).setValue(newProfit);           // Profit (col 19)
  salesSheet.getRange(saleRow, 20).setValue(newProfit);           // ProfitOverride (col 20) - sync both
  
  // Append adjustment note
  const sign = adjustmentAmount > 0 ? '-' : '+';
  const adjLabel = adjustmentAmount > 0 ? 'Reduced' : 'Increased';
  const adjNote = `[ADJUSTMENT ${fmtDate(new Date())}] ${adjLabel} ${invoiceNo} by ${sign}${Math.abs(adjustmentAmount)}. Old: ${currentFinal} → New: ${newFinal}. Reason: ${adjustmentReason}`;
  const combinedNotes = existingNotes ? existingNotes + ' | ' + adjNote : adjNote;
  salesSheet.getRange(saleRow, 21).setValue(combinedNotes);
  
  // ============================================================
  // 2. Update customer totals
  // ============================================================
  if (customerId) {
    const custRow = findRowById('Customers', customerId);
    if (custRow !== -1) {
      const r = custSheet.getRange(custRow, 1, 1, 10).getValues()[0];
      const oldCustSpent = Number(r[6]) || 0;
      const oldCustPending = Number(r[7]) || 0;
      
      // Total spent changes by -adjustmentAmount (or rather, by newFinal - currentFinal)
      const spentDelta = newFinal - currentFinal;
      const newCustSpent = Math.max(0, oldCustSpent + spentDelta);
      
      // Pending changes by newDue - currentDue
      const dueDelta = newDue - currentDue;
      const newCustPending = Math.max(0, oldCustPending + dueDelta);
      
      custSheet.getRange(custRow, 7).setValue(newCustSpent);   // TotalSpent
      custSheet.getRange(custRow, 8).setValue(newCustPending); // PendingAmount
    }
  }
  
  // ============================================================
  // 3. Record adjustment as a negative/positive payment for audit
  // ============================================================
  if (adjustmentMethod === 'Refund' && adjustmentAmount > 0) {
    // Record a refund (negative payment)
    paymentsSheet.appendRow([
      generateId('PAY'),
      fmtDate(new Date()),
      invoiceNo,
      customerId || '',
      customerName || '',
      -adjustmentAmount,
      data.RefundMethod || 'Cash',
      '',
      `REFUND — Sale adjustment: ${adjustmentReason}`
    ]);
  } else if (data.LogAsPayment !== false) {
    // Record adjustment as a memo (zero-amount payment for the trail)
    paymentsSheet.appendRow([
      generateId('PAY'),
      fmtDate(new Date()),
      invoiceNo,
      customerId || '',
      customerName || '',
      0,
      'Adjustment',
      '',
      `MEMO — ${adjLabel} sale by ${Math.abs(adjustmentAmount)}. Reason: ${adjustmentReason}`
    ]);
  }
  
  // ============================================================
  // 4. Adjust affiliate commission proportionally (optional)
  // ============================================================
  // Only reverse commission if explicitly requested and price was reduced
  if (data.AdjustCommission && affiliateCode && adjustmentAmount > 0 && currentFinal > 0) {
    const commSheet = ss.getSheetByName('AffiliateCommissions');
    const affSheet = ss.getSheetByName('Affiliates');
    if (commSheet && affSheet) {
      const ratio = newFinal / currentFinal; // shrink commissions by this ratio
      const commData = commSheet.getDataRange().getValues();
      let affId = '';
      
      for (let i = 1; i < commData.length; i++) {
        if (String(commData[i][5]) === String(invoiceNo) && commData[i][8] === 'Unpaid') {
          affId = commData[i][2];
          const oldComm = Number(commData[i][7]) || 0;
          const newComm = Math.round(oldComm * ratio);
          if (newComm <= 0) {
            commSheet.deleteRow(i + 1);
          } else {
            commSheet.getRange(i + 1, 8).setValue(newComm);
          }
        }
      }
      
      // Recalc affiliate totals
      if (affId) {
        const affRow = findRowById('Affiliates', affId);
        if (affRow !== -1) {
          const remaining = getSheetData('AffiliateCommissions').filter(c => String(c.AffiliateID) === String(affId));
          const totalComm = remaining.reduce((s, c) => s + (Number(c.CommissionAmount) || 0), 0);
          const paidComm = remaining.filter(c => c.Status === 'Paid').reduce((s, c) => s + (Number(c.CommissionAmount) || 0), 0);
          affSheet.getRange(affRow, 13).setValue(totalComm);
          affSheet.getRange(affRow, 14).setValue(paidComm);
          affSheet.getRange(affRow, 15).setValue(Math.max(0, totalComm - paidComm));
        }
      }
    }
  }
  
  return {
    success: true,
    invoiceNo,
    oldAmount: currentFinal,
    newAmount: newFinal,
    adjustmentAmount: Math.round(adjustmentAmount),
    newPaid: Math.round(newPaid),
    newDue: Math.round(newDue),
    newStatus,
    newProfit: Math.round(newProfit),
    message: adjustmentAmount > 0 
      ? `Sale reduced by ${Math.abs(adjustmentAmount)}. New total: ${newFinal}`
      : `Sale increased by ${Math.abs(adjustmentAmount)}. New total: ${newFinal}`
  };
}

function deleteSale(data) {
  const sheet = ss.getSheetByName('Sales');
  const row = findRowById('Sales', data.ID);
  if (row === -1) return {error:'Sale not found'};
  sheet.deleteRow(row);
  return {success:true};
}

// ============================================================
// PAYMENTS
// ============================================================

function getPayments() { return getSheetData('Payments'); }

function addPayment(data) {
  const paymentsSheet = ss.getSheetByName('Payments');
  const salesSheet = ss.getSheetByName('Sales');
  const id = generateId('PAY');
  paymentsSheet.appendRow([
    id, fmtDate(new Date()), data.InvoiceNo,
    data.CustomerID || '', data.CustomerName || '',
    Number(data.Amount), data.PaymentMethod || 'Cash', '', data.Notes || ''
  ]);
  const salesData = salesSheet.getDataRange().getValues();
  for (let i = 1; i < salesData.length; i++) {
    if (String(salesData[i][1]) === String(data.InvoiceNo)) {
      const currentPaid = Number(salesData[i][13]) || 0;
      const finalAmount = Number(salesData[i][12]) || 0;
      const newPaid = currentPaid + Number(data.Amount);
      const newDue = Math.max(0, finalAmount - newPaid);
      salesSheet.getRange(i + 1, 14).setValue(newPaid);
      salesSheet.getRange(i + 1, 15).setValue(newDue);
      salesSheet.getRange(i + 1, 17).setValue(newDue <= 0 ? 'Paid' : 'Partial');
      break;
    }
  }
  if (data.CustomerID) {
    const custSheet = ss.getSheetByName('Customers');
    const custRow = findRowById('Customers', data.CustomerID);
    if (custRow !== -1) {
      const curPending = Number(custSheet.getRange(custRow, 8).getValue()) || 0;
      custSheet.getRange(custRow, 8).setValue(Math.max(0, curPending - Number(data.Amount)));
    }
  }
  return {success:true, id};
}

// ============================================================
// SUPPLIERS
// ============================================================

function getSuppliers() { return getSheetData('Suppliers'); }

function addSupplier(data) {
  const sheet = ss.getSheetByName('Suppliers');
  const id = generateId('SUP');
  const formattedPhone = formatPhoneBackend(data.Phone || '');
  sheet.appendRow([id, data.Name, formattedPhone, data.Email || '', data.Address || '', 0, 0, 0, data.Notes || '']);
  return {success:true, id};
}

function updateSupplier(data) {
  const sheet = ss.getSheetByName('Suppliers');
  const row = findRowById('Suppliers', data.ID);
  if (row === -1) return {error:'Supplier not found'};
  const formattedPhone = formatPhoneBackend(data.Phone || '');
  sheet.getRange(row, 2).setValue(data.Name);
  sheet.getRange(row, 3).setValue(formattedPhone);
  sheet.getRange(row, 4).setValue(data.Email || '');
  sheet.getRange(row, 5).setValue(data.Address || '');
  sheet.getRange(row, 9).setValue(data.Notes || '');
  return {success:true};
}

function deleteSupplier(data) {
  const sheet = ss.getSheetByName('Suppliers');
  const row = findRowById('Suppliers', data.ID);
  if (row === -1) return {error:'Supplier not found'};
  sheet.deleteRow(row);
  return {success:true};
}

// ============================================================
// SUPPLIER LEDGER
// ============================================================

function getSupplierLedger() { return getSheetData('SupplierLedger'); }

function getSupplierLedgerById(supplierId) {
  return getSheetData('SupplierLedger').filter(e => String(e.SupplierID) === String(supplierId));
}

function addSupplierLedgerEntry(data) {
  const sheet = ss.getSheetByName('SupplierLedger');
  const supplierSheet = ss.getSheetByName('Suppliers');
  const id = generateId('SLE');
  const debit = Number(data.Debit) || 0;
  const credit = Number(data.Credit) || 0;
  const allEntries = getSheetData('SupplierLedger').filter(e => String(e.SupplierID) === String(data.SupplierID));
  let lastBalance = allEntries.length > 0 ? (Number(allEntries[allEntries.length - 1].Balance) || 0) : 0;
  const newBalance = lastBalance + debit - credit;
  sheet.appendRow([id, fmtDate(new Date()), data.SupplierID, data.SupplierName || '', data.Description || '', debit, credit, newBalance, data.Notes || '']);
  const supRow = findRowById('Suppliers', data.SupplierID);
  if (supRow !== -1) {
    const r = supplierSheet.getRange(supRow, 1, 1, 9).getValues()[0];
    if (debit > 0) {
      supplierSheet.getRange(supRow, 6).setValue((Number(r[5]) || 0) + 1);
      supplierSheet.getRange(supRow, 7).setValue((Number(r[6]) || 0) + debit);
    }
    supplierSheet.getRange(supRow, 8).setValue(newBalance);
  }
  return {success:true, id};
}

// ============================================================
// EXPENSES
// ============================================================

function getExpenses() { return getSheetData('Expenses'); }

function addExpense(data) {
  const sheet = ss.getSheetByName('Expenses');
  const id = generateId('EXP');
  sheet.appendRow([id, fmtDate(new Date()), data.Category || '', data.Description || '', Number(data.Amount) || 0, data.PaidTo || '', data.PaymentMethod || 'Cash', data.Notes || '']);
  return {success:true, id};
}

function deleteExpense(data) {
  const sheet = ss.getSheetByName('Expenses');
  const row = findRowById('Expenses', data.ID);
  if (row === -1) return {error:'Expense not found'};
  sheet.deleteRow(row);
  return {success:true};
}

// ============================================================
// AFFILIATES (Enhanced with commission rates)
// ============================================================

function getAffiliates() { return getSheetData('Affiliates'); }

function addAffiliate(data) {
  const sheet = ss.getSheetByName('Affiliates');
  const id = generateId('AFF');
  sheet.appendRow([
    id, data.Code, data.Name || '',
    data.DiscountType || 'Percentage', Number(data.DiscountValue) || 0,
    0, 0, data.Active !== false ? 'Yes' : 'No',
    Number(data.CommFan) || 0, Number(data.CommPlayer) || 0,
    Number(data.CommSpecial) || 0, Number(data.CommBD) || 0,
    0, 0, 0
  ]);
  return {success:true, id};
}

function updateAffiliate(data) {
  const sheet = ss.getSheetByName('Affiliates');
  const row = findRowById('Affiliates', data.ID);
  if (row === -1) return {error:'Affiliate not found'};
  sheet.getRange(row, 2).setValue(data.Code);
  sheet.getRange(row, 3).setValue(data.Name || '');
  sheet.getRange(row, 4).setValue(data.DiscountType || 'Percentage');
  sheet.getRange(row, 5).setValue(Number(data.DiscountValue) || 0);
  sheet.getRange(row, 8).setValue(data.Active ? 'Yes' : 'No');
  sheet.getRange(row, 9).setValue(Number(data.CommFan) || 0);
  sheet.getRange(row, 10).setValue(Number(data.CommPlayer) || 0);
  sheet.getRange(row, 11).setValue(Number(data.CommSpecial) || 0);
  sheet.getRange(row, 12).setValue(Number(data.CommBD) || 0);
  return {success:true};
}

function deleteAffiliate(data) {
  const sheet = ss.getSheetByName('Affiliates');
  const row = findRowById('Affiliates', data.ID);
  if (row === -1) return {error:'Affiliate not found'};
  sheet.deleteRow(row);
  return {success:true};
}

// ============================================================
// SETTINGS
// ============================================================

function getSettings() {
  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) { settings[data[i][0]] = data[i][1]; }
  return settings;
}

function updateSettings(data) {
  const sheet = ss.getSheetByName('Settings');
  const sheetData = sheet.getDataRange().getValues();
  for (const key in data.settings) {
    let found = false;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0] === key) { sheet.getRange(i + 1, 2).setValue(data.settings[key]); found = true; break; }
    }
    if (!found) sheet.appendRow([key, data.settings[key]]);
  }
  return {success:true};
}

// ============================================================
// DEAD STOCK HISTORY — Snapshot tracking
// ============================================================
function getDeadStockHistory() {
  const data = getSheetData('DeadStockHistory');
  return data;
}

function saveDeadStockSnapshot(data) {
  const sheet = ss.getSheetByName('DeadStockHistory');
  if (!sheet) return { error: 'DeadStockHistory sheet not found. Please create it.' };
  
  const items = data.items || [];
  if (!items.length) return { error: 'No items to save' };
  
  // Generate a unique snapshot ID for this batch
  const snapshotId = 'SNAP-' + new Date().getTime();
  const snapshotDate = fmtDate(new Date());
  const monthKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  
  // Check if a snapshot for this month already exists — if yes, delete it first
  const existing = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = existing.length - 1; i >= 1; i--) {
    if (String(existing[i][2]) === monthKey) {  // Column C = Month
      rowsToDelete.push(i + 1);
    }
  }
  rowsToDelete.forEach(row => sheet.deleteRow(row));
  
  // Add all items in batch
  const rows = items.map(item => [
    snapshotId,
    snapshotDate,
    monthKey,
    item.ProductID || '',
    item.ProductName || '',
    item.Club_Country || '',
    item.Edition || '',
    item.Type || '',
    item.Size || '',
    Number(item.Stock) || 0,
    Number(item.CostPrice) || 0,
    Number(item.SellPrice) || 0,
    Number(item.CapitalTied) || 0,
    Number(item.DaysInactive) || 0,
    item.Urgency || 'LOW'
  ]);
  
  if (rows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  }
  
  return { 
    success: true, 
    snapshotId, 
    monthKey, 
    itemCount: rows.length,
    replaced: rowsToDelete.length > 0
  };
}

function deleteDeadStockSnapshot(data) {
  const sheet = ss.getSheetByName('DeadStockHistory');
  if (!sheet) return { error: 'Sheet not found' };
  
  const monthKey = data.monthKey;
  if (!monthKey) return { error: 'monthKey required' };
  
  const allData = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = allData.length - 1; i >= 1; i--) {
    if (String(allData[i][2]) === monthKey) {
      rowsToDelete.push(i + 1);
    }
  }
  rowsToDelete.forEach(row => sheet.deleteRow(row));
  
  return { success: true, deleted: rowsToDelete.length };
}

// ============================================================
// CHRONIC DEAD STOCK — Items appearing in 2+ snapshots
// ============================================================
function getChronicDeadStock() {
  const history = getSheetData('DeadStockHistory');
  if (!history.length) return { items: [], categoryTrends: {} };
  
  // Group by ProductID to count appearances
  const productMap = {};
  const monthsSeen = new Set();
  
  history.forEach(row => {
    const pid = row.ProductID;
    const month = row.Month;
    monthsSeen.add(month);
    
    if (!productMap[pid]) {
      productMap[pid] = {
        ProductID: pid,
        ProductName: row.ProductName,
        Club_Country: row.Club_Country,
        Edition: row.Edition,
        Type: row.Type,
        Size: row.Size,
        appearances: 0,
        months: [],
        latestDays: 0,
        latestStock: 0,
        latestCapital: 0,
        totalCapitalAcrossMonths: 0,
        avgDaysInactive: 0,
        firstSeen: month,
        lastSeen: month
      };
    }
    
    productMap[pid].appearances++;
    productMap[pid].months.push(month);
    productMap[pid].latestDays = Number(row.DaysInactive) || 0;
    productMap[pid].latestStock = Number(row.Stock) || 0;
    productMap[pid].latestCapital = Number(row.CapitalTied) || 0;
    productMap[pid].totalCapitalAcrossMonths += Number(row.CapitalTied) || 0;
    
    if (month < productMap[pid].firstSeen) productMap[pid].firstSeen = month;
    if (month > productMap[pid].lastSeen) productMap[pid].lastSeen = month;
  });
  
  // Calculate avg days inactive
  Object.values(productMap).forEach(p => {
    p.avgDaysInactive = Math.round(p.totalCapitalAcrossMonths / Math.max(p.appearances, 1));
  });
  
  // Filter to "chronic" = appeared in 2+ snapshots
  const chronicItems = Object.values(productMap)
    .filter(p => p.appearances >= 2)
    .sort((a, b) => b.appearances - a.appearances || b.totalCapitalAcrossMonths - a.totalCapitalAcrossMonths);
  
  // Category trends — group by Club, Edition, Type
  const clubTrends = {};
  const editionTrends = {};
  const typeTrends = {};
  
  history.forEach(row => {
    const club = row.Club_Country || 'Unknown';
    const edition = row.Edition || 'Unknown';
    const type = row.Type || 'Unknown';
    const capital = Number(row.CapitalTied) || 0;
    
    if (!clubTrends[club]) clubTrends[club] = { count: 0, totalCapital: 0, products: new Set() };
    clubTrends[club].count++;
    clubTrends[club].totalCapital += capital;
    clubTrends[club].products.add(row.ProductID);
    
    if (!editionTrends[edition]) editionTrends[edition] = { count: 0, totalCapital: 0, products: new Set() };
    editionTrends[edition].count++;
    editionTrends[edition].totalCapital += capital;
    editionTrends[edition].products.add(row.ProductID);
    
    if (!typeTrends[type]) typeTrends[type] = { count: 0, totalCapital: 0, products: new Set() };
    typeTrends[type].count++;
    typeTrends[type].totalCapital += capital;
    typeTrends[type].products.add(row.ProductID);
  });
  
  // Convert sets to counts
  const convertTrends = (obj) => Object.entries(obj).map(([key, val]) => ({
    name: key,
    appearances: val.count,
    uniqueProducts: val.products.size,
    totalCapital: Math.round(val.totalCapital)
  })).sort((a, b) => b.totalCapital - a.totalCapital);
  
  return {
    items: chronicItems,
    totalSnapshots: monthsSeen.size,
    monthsList: [...monthsSeen].sort(),
    categoryTrends: {
      clubs: convertTrends(clubTrends).slice(0, 15),
      editions: convertTrends(editionTrends).slice(0, 10),
      types: convertTrends(typeTrends).slice(0, 10)
    }
  };
}

// ============================================================
// PRE-ORDERS
// ============================================================
function getPreorders() {
  return getSheetData('Preorders');
}

function addPreorder(data) {
  const sheet = ss.getSheetByName('Preorders');
  if (!sheet) return { error: 'Preorders sheet not found. Please create it.' };
  const id = generateId('PO');
  sheet.appendRow([
    id, fmtDate(new Date()),
    data.CustomerID || '', data.CustomerName || '',
    formatPhoneBackend(data.CustomerPhone || ''),
    data.ProductName || '', data.Size || '',
    Number(data.Quantity) || 1,
    data.Status || 'Pending',
    data.ExpectedDate || '', '',
    data.Notes || ''
  ]);
  return { success: true, id };
}

function updatePreorder(data) {
  const sheet = ss.getSheetByName('Preorders');
  if (!sheet) return { error: 'Preorders sheet not found' };
  const row = findRowById('Preorders', data.ID);
  if (row === -1) return { error: 'Pre-order not found' };
  
  if (data.CustomerID !== undefined) sheet.getRange(row, 3).setValue(data.CustomerID || '');
  if (data.CustomerName !== undefined) sheet.getRange(row, 4).setValue(data.CustomerName || '');
  if (data.CustomerPhone !== undefined) sheet.getRange(row, 5).setValue(formatPhoneBackend(data.CustomerPhone || ''));
  if (data.ProductName !== undefined) sheet.getRange(row, 6).setValue(data.ProductName || '');
  if (data.Size !== undefined) sheet.getRange(row, 7).setValue(data.Size || '');
  if (data.Quantity !== undefined) sheet.getRange(row, 8).setValue(Number(data.Quantity) || 1);
  if (data.Status !== undefined) sheet.getRange(row, 9).setValue(data.Status);
  if (data.ExpectedDate !== undefined) sheet.getRange(row, 10).setValue(data.ExpectedDate || '');
  if (data.FulfilledDate !== undefined) sheet.getRange(row, 11).setValue(data.FulfilledDate || fmtDate(new Date()));
  if (data.Notes !== undefined) sheet.getRange(row, 12).setValue(data.Notes || '');
  
  return { success: true };
}

function deletePreorder(data) {
  const sheet = ss.getSheetByName('Preorders');
  if (!sheet) return { error: 'Preorders sheet not found' };
  const row = findRowById('Preorders', data.ID);
  if (row === -1) return { error: 'Pre-order not found' };
  sheet.deleteRow(row);
  return { success: true };
}
