// src/index.js — Fan Field API on Cloudflare Workers + D1
// FULLY AUDITED & FIXED to match frontend exactly
// Schema matches FanField_DB.xlsx

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type,Authorization,X-API-Key',
};

const json = (data, status = 200) => new Response(JSON.stringify(data), {
  status, headers: { 'Content-Type': 'application/json', ...CORS }
});

const uid = (prefix = '') => prefix + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const now = () => new Date().toISOString();
const today = () => now().slice(0, 10);
const num = (v) => { const n = Number(v); return isFinite(n) ? n : 0; };

// Sequential ID generator: SLOG-0107, SLOG-0108, etc.
// Reads max existing ID for the prefix and increments.
async function seqId(db, table, idCol, prefix, pad = 4) {
  const row = await one(db, 
    `SELECT ${idCol} FROM ${table} WHERE ${idCol} LIKE ? ORDER BY ${idCol} DESC LIMIT 1`,
    prefix + '%'
  );
  let next = 1;
  if (row) {
    const m = String(row[idCol]).match(new RegExp('^' + prefix + '(\\d+)'));
    if (m) next = parseInt(m[1], 10) + 1;
  }
  return prefix + String(next).padStart(pad, '0');
}

// Phone formatter: 01712345678 -> 017-1234-5678
const fmtPhone = (p) => {
  if (!p) return '';
  const d = String(p).replace(/\D/g, '');
  if (d.length === 11 && d.startsWith('01')) return `${d.slice(0,3)}-${d.slice(3,7)}-${d.slice(7)}`;
  if (d.length === 13 && d.startsWith('8801')) return `0${d.slice(2,5)}-${d.slice(5,9)}-${d.slice(9)}`;
  return p;
};

// Next batch suffix: A → B → ... → Z → AA → AB → ...
function nextBatchSuffix(existingSuffixes) {
  if (!existingSuffixes.length) return 'A';
  // Sort by length first, then alphabetically (so AA > Z)
  const sorted = [...existingSuffixes].sort((a, b) => 
    a.length !== b.length ? a.length - b.length : (a < b ? -1 : 1)
  );
  const last = sorted[sorted.length - 1];
  
  if (last.length === 1) {
    const code = last.charCodeAt(0);
    if (code < 90) return String.fromCharCode(code + 1); // B, C, ... Z
    return 'AA'; // After Z
  }
  // Double-letter: increment last char
  const lastChar = last.charAt(last.length - 1);
  if (lastChar.charCodeAt(0) < 90) {
    return last.slice(0, -1) + String.fromCharCode(lastChar.charCodeAt(0) + 1);
  }
  // After AZ, BZ, etc → go up another letter
  return last + 'A';
}

// ============================================================
// MAIN HANDLER
// ============================================================
export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') return new Response(null, { headers: CORS });

    const url = new URL(request.url);
    let params = {};
    let action = url.searchParams.get('action');

    if (request.method === 'GET') {
      params = Object.fromEntries(url.searchParams.entries());
    } else {
      const body = await safeJson(request);
      params = body || {};
      action = action || params.action;
    }

    if (!action) return json({ success: false, error: 'No action specified' }, 400);

    try {
      const handler = ACTIONS[action];
      if (!handler) return json({ success: false, error: `Unknown action: ${action}` }, 404);
      const result = await handler(params, env.DB, env);
      return json(result);
    } catch (e) {
      console.error(action, e);
      return json({ success: false, error: e.message, stack: e.stack }, 500);
    }
  },
  // 🆕 ADD THIS — runs on cron schedule
  async scheduled(event, env, ctx) {
    ctx.waitUntil(syncToGoogleSheets(env));
  }
};

async function safeJson(req) {
  try { return await req.json(); } catch { return null; }
}

// ============================================================
// DB HELPERS
// ============================================================
const all = async (db, sql, ...binds) => (await db.prepare(sql).bind(...binds).all()).results || [];
const one = async (db, sql, ...binds) => (await db.prepare(sql).bind(...binds).first()) || null;
const run = async (db, sql, ...binds) => await db.prepare(sql).bind(...binds).run();

async function getSetting(db, key, def = '') {
  const r = await one(db, 'SELECT Value FROM Settings WHERE Key=?', key);
  return r ? r.Value : def;
}
async function setSetting(db, key, val) {
  await run(db, 'INSERT OR REPLACE INTO Settings (Key,Value) VALUES (?,?)', key, String(val ?? ''));
}

async function nextInvoice(db) {
  const prefix = await getSetting(db, 'invoicePrefix', 'FF-');
  const n = parseInt(await getSetting(db, 'nextInvoiceNum', '1001'), 10);
  await setSetting(db, 'nextInvoiceNum', String(n + 1));
  return `${prefix}${n}`;
}

// Next booking number: B-00125, B-00126, ...
async function nextBookingNumber(db) {
  const prefix = await getSetting(db, 'bookingPrefix', 'B-');
  const n = parseInt(await getSetting(db, 'nextBookingNum', '101'), 10);
  await setSetting(db, 'nextBookingNum', String(n + 1));
  return `${prefix}${String(n).padStart(5, '0')}`;
}

// Available stock = Physical (Stock) - Reserved (ReservedQty)
function availableStock(prod) {
  return num(prod.Stock) - num(prod.ReservedQty);
}

// Write a stock-audit record for reservation tracking
async function auditStock(db, opts) {
  await insertRow(db, 'StockAudit', {
    AuditID: uid('AUD-'),
    ProductID: opts.ProductID || '',
    ProductName: opts.ProductName || '',
    Action: opts.Action || '',
    ReferenceType: opts.ReferenceType || 'Booking',
    ReferenceID: opts.ReferenceID || '',
    QtyChange: num(opts.QtyChange),
    BeforeReserved: num(opts.BeforeReserved),
    AfterReserved: num(opts.AfterReserved),
    BeforePhysical: num(opts.BeforePhysical),
    AfterPhysical: num(opts.AfterPhysical),
    UserID: opts.UserID || '',
    Timestamp: now()
  });
}


// FIFO cost consumption: deduct `qty` units from the product's oldest batches.
// Returns the ACTUAL total cost of those units (sum of batch cost × units taken).
// Falls back to the product's CostPrice for any units not covered by batches.
async function consumeBatchesFIFO(db, productId, qty, fallbackCost) {
  qty = num(qty);
  if (qty <= 0 || !productId) return { actualCost: 0, breakdown: [] };

  // Oldest batches first
  const batches = await all(db,
    'SELECT * FROM InventoryBatches WHERE ProductID=? AND RemainingQty > 0 ORDER BY BatchDate ASC, ID ASC',
    productId);

  let remaining = qty;
  let actualCost = 0;
  const breakdown = [];

  for (const b of batches) {
    if (remaining <= 0) break;
    const avail = num(b.RemainingQty);
    const take = Math.min(avail, remaining);
    if (take <= 0) continue;

    actualCost += take * num(b.CostPrice);
    remaining -= take;
    breakdown.push({ batchId: b.ID, units: take, unitCost: num(b.CostPrice) });

    // Reduce the batch's remaining quantity
    await run(db, 'UPDATE InventoryBatches SET RemainingQty = RemainingQty - ? WHERE ID = ?',
      take, b.ID);
  }

  // If there weren't enough batch units (e.g. legacy stock with no batches),
  // cost the leftover at the product's current CostPrice.
  if (remaining > 0) {
    actualCost += remaining * num(fallbackCost);
    breakdown.push({ batchId: null, units: remaining, unitCost: num(fallbackCost) });
  }

  return { actualCost, breakdown };
}

// Reverse of FIFO: when a sale is cancelled/returned, put units BACK into batches.
// We restore to the most-recently-consumed batches (LIFO restore = mirror of FIFO).
// Simpler reliable approach: restore to newest batches first.
async function restoreBatchesFIFO(db, productId, qty) {
  qty = num(qty);
  if (qty <= 0 || !productId) return;

  // Newest first so we "undo" the most recent depletions
  const batches = await all(db,
    'SELECT * FROM InventoryBatches WHERE ProductID=? ORDER BY BatchDate DESC, ID DESC',
    productId);

  let remaining = qty;
  for (const b of batches) {
    if (remaining <= 0) break;
    const room = num(b.Quantity) - num(b.RemainingQty); // how much was consumed
    if (room <= 0) continue;
    const give = Math.min(room, remaining);
    await run(db, 'UPDATE InventoryBatches SET RemainingQty = RemainingQty + ? WHERE ID = ?',
      give, b.ID);
    remaining -= give;
  }
  // If still remaining (no batches / over-restore), we just leave it —
  // physical Stock is the source of truth and gets corrected separately.
}


// Add a timeline entry
async function addTimeline(db, bookingId, action, description, by = '') {
  await insertRow(db, 'BookingTimeline', {
    TimelineID: uid('BTL-'),
    BookingID: bookingId,
    Action: action,
    Description: description,
    CreatedBy: by,
    CreatedAt: now()
  });
}

// Reserve stock for ONE product, concurrency-safe via guarded UPDATE.
// Returns { reserved, backordered }. Never lets ReservedQty exceed Stock.
async function reserveStock(db, productId, wantQty, refId, productName) {
  wantQty = num(wantQty);
  if (wantQty <= 0 || !productId) return { reserved: 0, backordered: wantQty };

  const prod = await one(db, 'SELECT * FROM Inventory WHERE ID=?', productId);
  if (!prod) return { reserved: 0, backordered: wantQty };

  const avail = availableStock(prod);
  const toReserve = Math.max(0, Math.min(wantQty, avail));
  const backorder = wantQty - toReserve;

  if (toReserve > 0) {
    const beforeRes = num(prod.ReservedQty);
    // Guarded update: only succeeds if reserved won't exceed physical stock.
    const res = await db.prepare(
      `UPDATE Inventory SET ReservedQty = ReservedQty + ?, LastUpdated = ?
       WHERE ID = ? AND (ReservedQty + ?) <= Stock`
    ).bind(toReserve, now(), productId, toReserve).run();

    if (res.meta && res.meta.changes === 0) {
      // Lost a race — re-read and retry once with fresh availability
      const fresh = await one(db, 'SELECT * FROM Inventory WHERE ID=?', productId);
      const freshAvail = availableStock(fresh);
      const retryReserve = Math.max(0, Math.min(wantQty, freshAvail));
      if (retryReserve > 0) {
        await run(db,
          'UPDATE Inventory SET ReservedQty = ReservedQty + ?, LastUpdated = ? WHERE ID = ?',
          retryReserve, now(), productId);
        await auditStock(db, {
          ProductID: productId, ProductName: prod.ProductName,
          Action: 'BOOK_RESERVED', ReferenceID: refId, QtyChange: retryReserve,
          BeforeReserved: num(fresh.ReservedQty), AfterReserved: num(fresh.ReservedQty) + retryReserve,
          BeforePhysical: num(fresh.Stock), AfterPhysical: num(fresh.Stock)
        });
      }
      return { reserved: retryReserve, backordered: wantQty - retryReserve };
    }

    await auditStock(db, {
      ProductID: productId, ProductName: prod.ProductName,
      Action: 'BOOK_RESERVED', ReferenceID: refId, QtyChange: toReserve,
      BeforeReserved: beforeRes, AfterReserved: beforeRes + toReserve,
      BeforePhysical: num(prod.Stock), AfterPhysical: num(prod.Stock)
    });
  }
  return { reserved: toReserve, backordered: backorder };
}

// Release reserved stock for ONE product (e.g. cancel, reduce qty, remove item)
async function releaseStock(db, productId, qty, refId, action = 'BOOK_RELEASED') {
  qty = num(qty);
  if (qty <= 0 || !productId) return;
  const prod = await one(db, 'SELECT * FROM Inventory WHERE ID=?', productId);
  if (!prod) return;
  const beforeRes = num(prod.ReservedQty);
  const release = Math.min(qty, beforeRes); // never go negative
  if (release <= 0) return;
  await run(db,
    'UPDATE Inventory SET ReservedQty = ReservedQty - ?, LastUpdated = ? WHERE ID = ?',
    release, now(), productId);
  await auditStock(db, {
    ProductID: productId, ProductName: prod.ProductName,
    Action: action, ReferenceID: refId, QtyChange: -release,
    BeforeReserved: beforeRes, AfterReserved: beforeRes - release,
    BeforePhysical: num(prod.Stock), AfterPhysical: num(prod.Stock)
  });
}

// Recompute a booking's overall status from its items
function computeBookingStatus(items) {
  if (!items.length) return 'Draft';
  let anyBackorder = false, anyReserved = false, allReserved = true;
  for (const it of items) {
    if (num(it.BackorderedQty) > 0) anyBackorder = true;
    if (num(it.ReservedQty) > 0) anyReserved = true;
    if (num(it.ReservedQty) < num(it.RequestedQty)) allReserved = false;
  }
  if (anyBackorder && anyReserved) return 'Partial';
  if (anyBackorder && !anyReserved) return 'Need Restocking';
  if (allReserved) return 'Active';
  return 'Partial';
}

// Set per-item status
function computeItemStatus(it) {
  const req = num(it.RequestedQty), res = num(it.ReservedQty);
  if (res === 0) return 'Backordered';
  if (res < req) return 'Partial';
  return 'Reserved';
}

// Recompute booking header status + refresh in DB
async function refreshBookingStatus(db, bookingId) {
  const items = await all(db, 'SELECT * FROM BookingItems WHERE BookingID=?', bookingId);
  const status = computeBookingStatus(items);
  await run(db, 'UPDATE Bookings SET Status=?, UpdatedAt=? WHERE BookingID=?',
    status, now(), bookingId);
  return status;
}

async function insertRow(db, table, obj) {
  const keys = Object.keys(obj);
  if (!keys.length) return;
  const placeholders = keys.map(() => '?').join(',');
  const sql = `INSERT INTO ${table} (${keys.map(k=>`"${k}"`).join(',')}) VALUES (${placeholders})`;
  await db.prepare(sql).bind(...keys.map(k => obj[k] ?? null)).run();
}

async function updateRow(db, table, id, obj, idCol = 'ID') {
  const keys = Object.keys(obj).filter(k => k !== idCol);
  if (!keys.length) return;
  const sets = keys.map(k => `"${k}"=?`).join(',');
  await db.prepare(`UPDATE ${table} SET ${sets} WHERE "${idCol}"=?`).bind(...keys.map(k => obj[k] ?? null), id).run();
}

// Commission rate based on edition
function commissionRate(aff, edition) {
  const e = String(edition || '').toLowerCase();
  if (/bd|premium/.test(e)) return num(aff.CommissionBDPremium);
  if (/special/.test(e)) return num(aff.CommissionSpecialEdition);
  if (/player/.test(e)) return num(aff.CommissionPlayerEdition);
  if (/fan/.test(e)) return num(aff.CommissionFanEdition);
  return num(aff.CommissionFanEdition);
}

// Recompute customer aggregates from sales
async function recomputeCustomer(db, customerId) {
  if (!customerId) return;
  const r = await one(db,
    `SELECT 
      COUNT(*) as cnt, 
      COALESCE(SUM(FinalAmount),0) as spent, 
      COALESCE(SUM(DueAmount),0) as pending 
    FROM Sales WHERE CustomerID=? AND Status!='Cancelled'`, customerId);
  await run(db,
    'UPDATE Customers SET TotalPurchases=?, TotalSpent=?, PendingAmount=? WHERE ID=?',
    r.cnt || 0, r.spent || 0, r.pending || 0, customerId);
}


// ============================================================
// ACTIONS
// ============================================================
const ACTIONS = {

  // ============== INVENTORY ==============
  getInventory: async (p, db) => await all(db, 'SELECT * FROM Inventory ORDER BY Club_Country, Type, Edition, Size'),

  addProduct: async (p, db) => {
    const id = p.ID || uid('PRD-');
    const productName = p.ProductName || [p.Club_Country, p.Type, p.Edition, p.Size].filter(Boolean).join(' ');
    const row = {
      ID: id,
      ProductName: productName,
      Club_Country: p.Club_Country || '',
      Edition: p.Edition || '',
      Type: p.Type || '',
      Size: p.Size || '',
      Color: p.Color || '',
      CostPrice: num(p.CostPrice),
      SellPrice: num(p.SellPrice),
      Stock: num(p.Stock),
      MinStock: num(p.MinStock),
      SupplierID: p.SupplierID || '',
      SupplierName: p.SupplierName || '',
      DateAdded: now(),
      LastUpdated: now()
    };
    await insertRow(db, 'Inventory', row);
    if (row.Stock > 0) {
      await insertRow(db, 'StockLog', {
        ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'), Date: now(), ProductID: id, ProductName: row.ProductName,
        Action: 'Initial', Quantity: row.Stock, OldStock: 0, NewStock: row.Stock,
        Reference: 'Product creation', Notes: ''
      });
    }
    return { success: true, id, product: row };
  },

  getSupplierLedgerById: async (p, db) => {
    const rows = await all(db, 'SELECT * FROM SupplierLedger WHERE SupplierID=? ORDER BY Date', p.supplierId);
    return rows;
  },

  updateProduct: async (p, db) => {
    const cur = await one(db, 'SELECT * FROM Inventory WHERE ID=?', p.ID);
    if (!cur) return { success: false, error: 'Product not found' };
    const upd = {
      ProductName: p.ProductName || [p.Club_Country ?? cur.Club_Country, p.Type ?? cur.Type, p.Edition ?? cur.Edition, p.Size ?? cur.Size].filter(Boolean).join(' '),
      Club_Country: p.Club_Country ?? cur.Club_Country,
      Edition: p.Edition ?? cur.Edition,
      Type: p.Type ?? cur.Type,
      Size: p.Size ?? cur.Size,
      Color: p.Color ?? cur.Color,
      CostPrice: num(p.CostPrice ?? cur.CostPrice),
      SellPrice: num(p.SellPrice ?? cur.SellPrice),
      Stock: p.Stock !== undefined ? num(p.Stock) : cur.Stock,
      MinStock: num(p.MinStock ?? cur.MinStock),
      SupplierID: p.SupplierID ?? cur.SupplierID,
      SupplierName: p.SupplierName ?? cur.SupplierName,
      LastUpdated: now()
    };
    await updateRow(db, 'Inventory', p.ID, upd);
    // Log stock change if Stock provided & differs
    if (p.Stock !== undefined && num(p.Stock) !== num(cur.Stock)) {
      await insertRow(db, 'StockLog', {
        ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'), Date: now(), ProductID: p.ID, ProductName: upd.ProductName,
        Action: 'Edit', Quantity: num(p.Stock) - num(cur.Stock),
        OldStock: num(cur.Stock), NewStock: num(p.Stock),
        Reference: 'Product edit', Notes: ''
      });
    }
    return { success: true };
  },

  // FIXED: frontend sends {ID, mode, Stock} where mode = 'set'|'add'|'subtract'
  updateStock: async (p, db) => {
    const prod = await one(db, 'SELECT * FROM Inventory WHERE ID=?', p.ID);
    if (!prod) return { success: false, error: 'Product not found' };
    const oldStock = num(prod.Stock);
    const mode = p.mode || 'set';
    const inputVal = num(p.Stock);
    let newStock;
    if (mode === 'add') newStock = oldStock + inputVal;
    else if (mode === 'subtract') newStock = oldStock - inputVal;
    else newStock = inputVal; // 'set'
    const change = newStock - oldStock;
    await run(db, 'UPDATE Inventory SET Stock=?, LastUpdated=? WHERE ID=?', newStock, now(), p.ID);
    await insertRow(db, 'StockLog', {
      ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'), Date: now(), ProductID: p.ID, ProductName: prod.ProductName,
      Action: 'Manual ' + mode, Quantity: change, OldStock: oldStock, NewStock: newStock,
      Reference: p.Reference || 'Manual update', Notes: p.Notes || ''
    });
    // Auto-allocate freshly added stock to waiting bookings
    if (change > 0) {
      try { await ACTIONS.allocateRestock({ productId: p.ID }, db); } catch (e) { console.error('allocate', e); }
    }
    return { success: true, newStock };
  },

  deleteProduct: async (p, db) => {
    const prod = await one(db, 'SELECT * FROM Inventory WHERE ID=?', p.ID);
    if (!prod) return { success: false, error: 'Product not found' };

    // Keep only the booking guard — deleting a reserved product would
    // break active bookings. Past SALES are safe (SalesItems stores a full snapshot).
    if (num(prod.ReservedQty) > 0) {
      return { success: false, error: `Cannot delete — ${num(prod.ReservedQty)} unit(s) reserved by active bookings. Cancel those bookings first.` };
    }

    const oldStock = num(prod.Stock);

    // Log the deletion (sales history remains fully intact)
    await insertRow(db, 'StockLog', {
      ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'),
      Date: now(),
      ProductID: prod.ID,
      ProductName: prod.ProductName,
      Action: 'Product Deleted',
      Quantity: -oldStock,
      OldStock: oldStock,
      NewStock: 0,
      Reference: 'Product deletion',
      Notes: 'Product removed from inventory (past sales history preserved)'
    });

    await run(db, 'DELETE FROM Inventory WHERE ID=?', p.ID);
    return { success: true };
  },

  // ============== INVENTORY BATCHES ==============
  getBatches: async (p, db) => p.productId
    ? await all(db, 'SELECT * FROM InventoryBatches WHERE ProductID=? ORDER BY BatchDate DESC', p.productId)
    : await all(db, 'SELECT * FROM InventoryBatches ORDER BY BatchDate DESC'),

  // FIXED: frontend's addBatch creates a NEW Inventory variant (suffix like PRD-001A)
  // not a stock-only batch. Looking at frontend code carefully:
  //   "Creates a variant with suffix ID (e.g., PRD-0006 → PRD-0006A)"
  // Adds a batch to an EXISTING product: bumps its Stock, recalculates
  // weighted-average CostPrice, and records batch history. No new variant.
  // NEW MODEL: A batch adds stock to the EXISTING product and records a
  // cost layer in InventoryBatches. No new Inventory variant is created.
  addBatch: async (p, db) => {
    const prod = await one(db, 'SELECT * FROM Inventory WHERE ID=?', p.ProductID);
    if (!prod) return { success: false, error: 'Product not found' };

    const qty = num(p.Quantity);
    if (qty <= 0) return { success: false, error: 'Quantity must be greater than 0' };

    const batchCost = num(p.CostPrice);
    // Optional: a batch can update the product's sell price; otherwise keep current
    const newSellPrice = (p.SellPrice !== undefined && p.SellPrice !== '')
      ? num(p.SellPrice) : num(prod.SellPrice);

    const oldStock = num(prod.Stock);
    const newStock = oldStock + qty;

    // ── Weighted-average cost across all stock ──
    // (oldStock × oldCost + qty × batchCost) / newStock
    const oldCost = num(prod.CostPrice);
    const weightedCost = newStock > 0
      ? ((oldStock * oldCost) + (qty * batchCost)) / newStock
      : batchCost;

    // Update the base product: more stock, new weighted cost, optional sell price
    await run(db,
      'UPDATE Inventory SET Stock=?, CostPrice=?, SellPrice=?, SupplierID=?, SupplierName=?, LastUpdated=? WHERE ID=?',
      newStock,
      Math.round(weightedCost * 100) / 100,   // round to 2 decimals
      newSellPrice,
      p.SupplierID || prod.SupplierID || '',
      p.SupplierName || prod.SupplierName || '',
      now(),
      p.ProductID
    );

    // Record the cost layer in InventoryBatches (history)
    const batchId = uid('BT-');
    await insertRow(db, 'InventoryBatches', {
      ID: batchId,
      ProductID: p.ProductID,            // ← points to the SAME product now
      ProductName: prod.ProductName,
      SupplierID: p.SupplierID || '',
      SupplierName: p.SupplierName || '',
      BatchDate: now(),
      Quantity: qty,
      CostPrice: batchCost,
      RemainingQty: qty,
      Notes: p.Notes || ''
    });

    // Log the stock increase
    await insertRow(db, 'StockLog', {
      ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'),
      Date: now(),
      ProductID: p.ProductID,
      ProductName: prod.ProductName,
      Action: 'Batch Added',
      Quantity: qty,
      OldStock: oldStock,
      NewStock: newStock,
      Reference: 'Batch ' + batchId,
      Notes: `Added ${qty} @ ${batchCost} (avg cost now ${Math.round(weightedCost*100)/100})${p.Notes ? ' — ' + p.Notes : ''}`
    });

    // Auto-allocate this new stock to any waiting bookings
    try { await ACTIONS.allocateRestock({ productId: p.ProductID }, db); } catch (e) { console.error('allocate', e); }

    return { success: true, productId: p.ProductID, newStock, weightedCost: Math.round(weightedCost*100)/100, batchId };
  },

  deleteBatch: async (p, db) => {
    const batch = await one(db, 'SELECT * FROM InventoryBatches WHERE ID=?', p.ID);
    if (!batch) return { success: false, error: 'Batch not found' };

    // Optionally subtract this batch's remaining stock from the product
    if (p.SubtractStock && batch.ProductID) {
      const prod = await one(db, 'SELECT * FROM Inventory WHERE ID=?', batch.ProductID);
      if (prod) {
        const remove = Math.min(num(batch.RemainingQty), num(prod.Stock));
        const available = num(prod.Stock) - num(prod.ReservedQty);
        if (remove > available) {
          return { success: false, error: `Cannot remove ${remove} units — only ${available} are unreserved.` };
        }
        const oldStock = num(prod.Stock);
        const newStock = oldStock - remove;
        await run(db, 'UPDATE Inventory SET Stock=?, LastUpdated=? WHERE ID=?', newStock, now(), batch.ProductID);
        await insertRow(db, 'StockLog', {
          ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'),
          Date: now(), ProductID: batch.ProductID, ProductName: batch.ProductName,
          Action: 'Batch Deleted', Quantity: -remove, OldStock: oldStock, NewStock: newStock,
          Reference: batch.ID, Notes: 'Batch record removed with stock'
        });
      }
    }

    await run(db, 'DELETE FROM InventoryBatches WHERE ID=?', p.ID);
    return { success: true };
  },

  // ============== CUSTOMERS ==============
  getCustomers: async (p, db) => await all(db, 'SELECT * FROM Customers ORDER BY Name'),

  addCustomer: async (p, db) => {
    const id = uid('CUST-');
    const row = {
      ID: id, Name: p.Name, Phone: fmtPhone(p.Phone || ''), Email: p.Email || '',
      Address: p.Address || '', TotalPurchases: 0, TotalSpent: 0, PendingAmount: 0,
      DateAdded: now(), Notes: p.Notes || ''
    };
    await insertRow(db, 'Customers', row);
    return { success: true, id, customer: row };
  },

  updateCustomer: async (p, db) => {
    const cur = await one(db, 'SELECT * FROM Customers WHERE ID=?', p.ID);
    if (!cur) return { success: false, error: 'Customer not found' };
    await updateRow(db, 'Customers', p.ID, {
      Name: p.Name ?? cur.Name,
      Phone: fmtPhone(p.Phone ?? cur.Phone),
      Email: p.Email ?? cur.Email,
      Address: p.Address ?? cur.Address,
      Notes: p.Notes ?? cur.Notes
    });
    return { success: true };
  },

  deleteCustomer: async (p, db) => {
    await run(db, 'DELETE FROM Customers WHERE ID=?', p.ID);
    return { success: true };
  },

  // NEW: getCustomerLedger — frontend calls this with {customerId}
  getCustomerLedger: async (p, db) => {
    const sales = await all(db, 'SELECT * FROM Sales WHERE CustomerID=? ORDER BY Date DESC', p.customerId);
    const payments = await all(db, 'SELECT * FROM Payments WHERE CustomerID=? ORDER BY Date DESC', p.customerId);
    return { sales, payments };
  },

  // ============== SALES ==============
  getSales: async (p, db) => await all(db, 'SELECT * FROM Sales ORDER BY Date DESC'),

  // FIXED: frontend calls 'getSalesItems' (with s), supports optional invoiceNo
  getSalesItems: async (p, db) => p.invoiceNo
    ? await all(db, 'SELECT * FROM SalesItems WHERE InvoiceNo=?', p.invoiceNo)
    : await all(db, 'SELECT * FROM SalesItems'),
  // Alias for safety
  getSaleItems: async (p, db) => p.invoiceNo
    ? await all(db, 'SELECT * FROM SalesItems WHERE InvoiceNo=?', p.invoiceNo)
    : await all(db, 'SELECT * FROM SalesItems'),

  // FIXED: frontend sends Items array under "Items" key
  createSale: async (p, db) => {
    // 🆕 BOOKING CONVERSION: if this sale fulfills a booking, release its
    // reservations FIRST so the physical-stock deduction below is correct.
    let _convBooking = null;
    if (p.FromBookingID) {
      _convBooking = await one(db, 'SELECT * FROM Bookings WHERE BookingID=?', p.FromBookingID);
      if (_convBooking && _convBooking.Status !== 'Completed' && _convBooking.Status !== 'Cancelled') {
        const _bItems = await all(db, 'SELECT * FROM BookingItems WHERE BookingID=?', p.FromBookingID);
        for (const _bi of _bItems) {
          if (num(_bi.ReservedQty) > 0) {
            await releaseStock(db, _bi.ProductID, num(_bi.ReservedQty),
              _convBooking.BookingNumber, 'BOOK_CONVERTED');
          }
        }
      }
    }
    // Frontend sends "Items" (capital), but accept all variants
    let items = p.Items || p.items || [];
    if (typeof items === 'string') {
      try { items = JSON.parse(items); } catch { items = []; }
    }
    if (!Array.isArray(items) || !items.length) {
      return { success: false, error: 'No items in sale' };
    }

    const invoiceNo = await nextInvoice(db);
    const saleId = uid('SALE-');
    const date = p.Date || now();

    // Normalize items
    const normItems = items.map(it => ({
      ProductID: it.ProductID || '',
      ProductName: it.ProductName || '',
      Club_Country: it.Club_Country || '',
      Edition: it.Edition || '',
      Type: it.Type || '',
      Size: it.Size || '',
      Quantity: num(it.Quantity),
      UnitPrice: num(it.UnitPrice),
      CostPrice: num(it.CostPrice),
      Patch: it.Patch === true || it.Patch === 'Yes' || it.Patch === 'true' ? 'Yes' : '',
      PatchPrice: num(it.PatchPrice),
      NamePrint: it.NamePrint || '',
      NamePrintPrice: num(it.NamePrintPrice),
      IsCustomizationService: !!it.IsCustomizationService,
      ServiceName: it.ServiceName || ''
    }));

    // 🆕 Build ProductName from Inventory parts (Club_Country → Type → Edition → Size) using ProductID
    for (const it of normItems) {
      // Customization service items: use the service name, skip inventory lookup
      const isService = it.IsCustomizationService === true
        || it.ProductID === '__CUSTOMIZATION__'
        || !it.ProductID;
      if (isService) {
        if (!it.ProductName) it.ProductName = it.ServiceName || 'Customization service';
        continue;
      }
      // Look up parts from Inventory using ProductID
      const prod = await one(db, 'SELECT * FROM Inventory WHERE ID=?', it.ProductID);
      if (prod) {
        it.Club_Country = prod.Club_Country || '';
        it.Edition      = prod.Edition || '';
        it.Type         = prod.Type || '';
        it.Size         = prod.Size || '';
      }
      // Build the name in your order: Club → Type → Edition → Size
      it.ProductName = [it.Club_Country, it.Type, it.Edition, it.Size]
        .filter(Boolean)
        .join(' ')
        || it.ProductName;  // fallback to whatever frontend sent, if any
    }

    // Read customization costs from settings
    const patchCost = num(await getSetting(db, 'patchDefaultCost', 130));
    const namePrintCost = num(await getSetting(db, 'namePrintDefaultCost', 130));

    let totalAmount = 0;

    // Revenue only — cost comes later after FIFO
    for (const it of normItems) {
      const patchP = it.Patch === 'Yes' ? it.PatchPrice : 0;
      const nameP = it.NamePrint ? it.NamePrintPrice : 0;

      totalAmount +=
        it.UnitPrice * it.Quantity +
        (patchP + nameP) * it.Quantity;
    }

    // Temporary placeholders until FIFO finishes
    let costTotal = 0;

    const discount = num(p.Discount);
    const affDisc = num(p.AffiliateDiscount);
    const vchDisc = num(p.VoucherDiscount);
    const delCharge = num(p.DeliveryCharge);
    const finalAmount = totalAmount - discount - affDisc - vchDisc + delCharge;
    const paid = num(p.PaidAmount);
    const due = Math.max(0, finalAmount - paid);
    const profit = 0; // recalculated after FIFO

    let status = 'Unpaid';
    if (paid >= finalAmount) status = 'Paid';
    else if (paid > 0) status = 'Partial';

    const sale = {
      ID: saleId, InvoiceNo: invoiceNo, Date: date,
      CustomerID: p.CustomerID || '', CustomerName: p.CustomerName || 'Walk-in Customer',
      CustomerPhone: fmtPhone(p.CustomerPhone || ''),
      Items: JSON.stringify(normItems.map(it => ({
        n: it.ProductName, q: it.Quantity, p: it.UnitPrice
      }))),
      Modifications: JSON.stringify(normItems.filter(it => it.Patch || it.NamePrint).map(it => ({
        n: it.ProductName, patch: it.Patch, np: it.NamePrint
      }))),
      TotalAmount: totalAmount,
      Discount: discount,
      AffiliateCode: p.AffiliateCode || '',
      AffiliateDiscount: affDisc,
      FinalAmount: finalAmount,
      PaidAmount: paid,
      DueAmount: due,
      PaymentMethod: p.PaymentMethod || 'Cash',
      Status: status,
      CostTotal: costTotal,
      Profit: profit,
      ProfitOverride: null,
      Notes: p.Notes || '',
      DeliveryType: p.DeliveryType || 'None',
      DeliveryCharge: delCharge,
      VoucherCode: p.VoucherCode || '',
      VoucherDiscount: vchDisc,
      DeliveryStatus: (p.DeliveryType && p.DeliveryType !== 'None') ? 'Processing' : '',
      CourierService: '',
      TrackingNo: '',
      DeliveryAddress: p.DeliveryAddress || ''
    };
    await insertRow(db, 'Sales', sale);

    // SalesItems + stock decrement
    for (const it of normItems) {
      const patchP = it.Patch === 'Yes' ? it.PatchPrice : 0;
      const nameP = it.NamePrint ? it.NamePrintPrice : 0;
      const modTotal = (patchP + nameP) * it.Quantity;
      const lineTotal = it.UnitPrice * it.Quantity + modTotal;
      await insertRow(db, 'SalesItems', {
        ID: uid('SI-'), InvoiceNo: invoiceNo,
        ProductID: (it.ProductID === '__CUSTOMIZATION__') ? '' : (it.ProductID || ''),
        ProductName: it.ProductName || it.ServiceName || '',
        Club_Country: it.Club_Country, Edition: it.Edition,
        Type: it.Type, Size: it.Size,
        Quantity: it.Quantity, UnitPrice: it.UnitPrice,
        CostPrice: it.CostPrice,
        TotalPrice: it.UnitPrice * it.Quantity,
        Patch: it.Patch, NamePrint: it.NamePrint,
        PatchPrice: it.PatchPrice, NamePrintPrice: it.NamePrintPrice,
        ModTotal: modTotal, LineTotal: lineTotal
      });
      // Decrement stock only for real products (not services)
      // NEW CODE — add explicit check for the magic string:
      const isService = it.IsCustomizationService === true 
        || it.ProductID === '__CUSTOMIZATION__' 
        || !it.ProductID;
        
      if (!isService) {
        const prod = await one(db, 'SELECT Stock,ProductName,CostPrice FROM Inventory WHERE ID=?', it.ProductID);
        if (prod) {
          const oldS = num(prod.Stock);
          const newS = oldS - it.Quantity;
          await run(db, 'UPDATE Inventory SET Stock=?, LastUpdated=? WHERE ID=?', newS, now(), it.ProductID);

          // ── FIFO: consume oldest batches & get the ACTUAL product cost ──
          const { actualCost } = await consumeBatchesFIFO(db, it.ProductID, it.Quantity, prod.CostPrice);
          // Store the real per-unit cost on the item (used by profit calc below)
          it._fifoUnitCost = it.Quantity > 0 ? (actualCost / it.Quantity) : num(prod.CostPrice);

          await insertRow(db, 'StockLog', {
            ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'), Date: now(), ProductID: it.ProductID, ProductName: prod.ProductName,
            Action: 'Sale', Quantity: -it.Quantity, OldStock: oldS, NewStock: newS,
            Reference: invoiceNo, Notes: ''
          });
        }
      }
    }

    // ─────────────────────────────────────────────
    // Recompute cost using ACTUAL FIFO costs
    // ─────────────────────────────────────────────

    let costTotalFIFO = 0;

    for (const it of normItems) {

      const isService =
        it.IsCustomizationService === true ||
        it.ProductID === '__CUSTOMIZATION__' ||
        !it.ProductID;

      const baseUnitCost =
        isService
          ? 0
          : (
              it._fifoUnitCost !== undefined
                ? it._fifoUnitCost
                : it.CostPrice
            );

      const patchC =
        it.Patch === 'Yes'
          ? patchCost
          : 0;

      const nameC =
        it.NamePrint
          ? namePrintCost
          : 0;

      costTotalFIFO +=
        (baseUnitCost + patchC + nameC) *
        it.Quantity;
    }

    const profitFIFO =
      finalAmount -
      costTotalFIFO -
      delCharge;

    // Patch Sales row with real FIFO values
    await run(
      db,
      `
      UPDATE Sales
      SET CostTotal=?,
          Profit=?
      WHERE ID=?
      `,
      Math.round(costTotalFIFO * 100) / 100,
      Math.round(profitFIFO * 100) / 100,
      saleId
    );

    // Payment
    if (paid > 0) {
      await insertRow(db, 'Payments', {
        ID: uid('PAY-'), Date: date, InvoiceNo: invoiceNo,
        CustomerID: p.CustomerID || '', CustomerName: p.CustomerName || 'Walk-in',
        Amount: paid, PaymentMethod: p.PaymentMethod || 'Cash',
        ReceivedBy: '', Notes: 'Initial payment at sale'
      });
    }

    // Customer aggregates
    if (p.CustomerID) await recomputeCustomer(db, p.CustomerID);

    // Affiliate commissions
    let totalComm = 0;
    if (p.AffiliateCode) {
      // NEW: case-insensitive match
      const aff = await one(db, 'SELECT * FROM Affiliates WHERE LOWER(Code)=LOWER(?)', p.AffiliateCode);
      if (aff) {
        const editionGroups = {};
        for (const it of normItems) {
          const edition = it.Edition || '';
          if (!editionGroups[edition]) editionGroups[edition] = 0;
          editionGroups[edition] += it.Quantity;
        }
        for (const [edition, qty] of Object.entries(editionGroups)) {
          const rate = commissionRate(aff, edition);
          const commAmt = rate * qty;
          if (commAmt > 0) {
            await insertRow(db, 'AffiliateCommissions', {
              ID: uid('AC-'), Date: date,
              AffiliateID: aff.ID, AffiliateCode: aff.Code, AffiliateName: aff.Name,
              InvoiceNo: invoiceNo, Edition: edition,
              CommissionAmount: commAmt, Status: 'Unpaid',
              PaidDate: '', ExpenseID: '', Notes: ''
            });
            totalComm += commAmt;
          }
        }
        await run(db,
          'UPDATE Affiliates SET TotalUsed=TotalUsed+1, TotalDiscountGiven=TotalDiscountGiven+?, PendingCommission=PendingCommission+? WHERE ID=?',
          affDisc, totalComm, aff.ID);
      }
    }

    // Voucher usage
    if (p.VoucherCode) {
      // NEW: case-insensitive
      await run(db,
        'UPDATE Vouchers SET TimesUsed=TimesUsed+1, TotalDiscountGiven=TotalDiscountGiven+? WHERE LOWER(Code)=LOWER(?)',
        vchDisc, p.VoucherCode);
    }

    // Delivery record
    if (p.DeliveryType && p.DeliveryType !== 'None') {
      await insertRow(db, 'Deliveries', {
        ID: uid('DEL-'), InvoiceNo: invoiceNo, Date: date,
        CustomerName: p.CustomerName || '', CustomerPhone: fmtPhone(p.CustomerPhone || ''),
        Address: p.DeliveryAddress || '', DeliveryType: p.DeliveryType,
        DeliveryCharge: delCharge, CourierService: '',
        TrackingNo: '', DeliveryStatus: 'Processing', Notes: ''
      });
    }

    // 🆕 BOOKING CONVERSION: finalize the linked booking now that the sale exists.
    if (_convBooking) {
      // Recompute booking item statuses: reserved items are now "Sold",
      // backordered items remain so the booking can stay open if needed.
      const bItems = await all(db, 'SELECT * FROM BookingItems WHERE BookingID=?', _convBooking.BookingID);
      let anyBackorderLeft = false;
      for (const bi of bItems) {
        if (num(bi.ReservedQty) > 0) {
          // These were the units just sold
          await updateRow(db, 'BookingItems', bi.BookingItemID, {
            Status: 'Sold', ReservedQty: 0, UpdatedAt: now()
          }, 'BookingItemID');
        }
        if (num(bi.BackorderedQty) > 0) anyBackorderLeft = true;
      }

      const newBookingStatus = anyBackorderLeft ? 'Partial' : 'Completed';
      await run(db,
        'UPDATE Bookings SET Status=?, LinkedSaleID=?, LinkedInvoiceNo=?, UpdatedAt=? WHERE BookingID=?',
        newBookingStatus, saleId, invoiceNo, now(), _convBooking.BookingID);

      await addTimeline(db, _convBooking.BookingID, 'BOOK_CONVERTED',
        `Converted to sale ${invoiceNo} via POS`);
      await addTimeline(db, _convBooking.BookingID, 'SALE_COMPLETED',
        `Sale completed. Invoice ${invoiceNo}${anyBackorderLeft ? ' (backorders remain — booking kept Partial)' : ''}`);

      return { success: true, invoiceNo, saleId, profit: Math.round(profitFIFO * 100) / 100, commission: totalComm,
               bookingConverted: true, bookingStatus: newBookingStatus };
    }

    return { success: true, invoiceNo, saleId, profit: Math.round(profitFIFO * 100) / 100, commission: totalComm };
  },

  // Manual sale (free-text items, just totals). Writes a single
  // summary SalesItems row (no ProductID → no stock decrement) so the
  // sale still shows up in invoice/item lookups consistently.
  addManualSale: async (p, db) => {
    const invoiceNo = await nextInvoice(db);
    const saleId = uid('SALE-');
    const date = p.Date || now();
    const finalAmount = num(p.FinalAmount);
    const costTotal = num(p.CostTotal);
    const paid = num(p.PaidAmount);
    const due = Math.max(0, finalAmount - paid);
    let status = 'Unpaid';
    if (paid >= finalAmount && finalAmount > 0) status = 'Paid';
    else if (paid > 0) status = 'Partial';

    const itemsText = p.Items || 'Manual sale';

    await insertRow(db, 'Sales', {
      ID: saleId, InvoiceNo: invoiceNo, Date: date,
      CustomerID: p.CustomerID || '', CustomerName: p.CustomerName || 'Walk-in',
      CustomerPhone: fmtPhone(p.CustomerPhone || ''),
      Items: itemsText, Modifications: '',
      TotalAmount: finalAmount, Discount: 0,
      AffiliateCode: '', AffiliateDiscount: 0,
      FinalAmount: finalAmount, PaidAmount: paid, DueAmount: due,
      PaymentMethod: p.PaymentMethod || 'Cash', Status: status,
      CostTotal: costTotal, Profit: finalAmount - costTotal,
      ProfitOverride: null, Notes: p.Notes || '',
      DeliveryType: 'None', DeliveryCharge: 0,
      VoucherCode: '', VoucherDiscount: 0, DeliveryStatus: '',
      CourierService: '', TrackingNo: '', DeliveryAddress: ''
    });

    // Single summary line item (no ProductID → stock untouched on reverse/delete)
    await insertRow(db, 'SalesItems', {
      ID: uid('SI-'), InvoiceNo: invoiceNo,
      ProductID: '', ProductName: itemsText,
      Club_Country: '', Edition: '', Type: '', Size: '',
      Quantity: 1, UnitPrice: finalAmount, CostPrice: costTotal,
      TotalPrice: finalAmount, Patch: '', NamePrint: '',
      PatchPrice: 0, NamePrintPrice: 0,
      ModTotal: 0, LineTotal: finalAmount
    });

    if (paid > 0) {
      await insertRow(db, 'Payments', {
        ID: uid('PAY-'), Date: date, InvoiceNo: invoiceNo,
        CustomerID: p.CustomerID || '', CustomerName: p.CustomerName || 'Walk-in',
        Amount: paid, PaymentMethod: p.PaymentMethod || 'Cash',
        ReceivedBy: '', Notes: 'Manual sale entry'
      });
    }
    if (p.CustomerID) await recomputeCustomer(db, p.CustomerID);
    return { success: true, invoiceNo, saleId };
  },

  // FIXED: frontend sends {ID} (sale ID)
  reverseSale: async (p, db) => {
    const saleId = p.ID || p.SaleID;
    const sale = await one(db, 'SELECT * FROM Sales WHERE ID=?', saleId);
    if (!sale) return { success: false, error: 'Sale not found' };
    if (sale.Status === 'Cancelled') return { success: false, error: 'Already cancelled' };

    const items = await all(db, 'SELECT * FROM SalesItems WHERE InvoiceNo=?', sale.InvoiceNo);
    for (const it of items) {
      if (it.ProductID) {
        const prod = await one(db, 'SELECT Stock,ProductName FROM Inventory WHERE ID=?', it.ProductID);
        if (prod) {
          const oldS = num(prod.Stock);
          const newS = oldS + num(it.Quantity);
          await run(
            db,
            'UPDATE Inventory SET Stock=?, LastUpdated=? WHERE ID=?',
            newS,
            now(),
            it.ProductID
          );

          // Restore consumed FIFO batches
          await restoreBatchesFIFO(
            db,
            it.ProductID,
            num(it.Quantity)
          );
          await insertRow(db, 'StockLog', {
            ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'), Date: now(), ProductID: it.ProductID, ProductName: prod.ProductName,
            Action: 'Reverse', Quantity: num(it.Quantity), OldStock: oldS, NewStock: newS,
            Reference: sale.InvoiceNo, Notes: 'Sale cancelled'
          });
        }
      }
    }

    const newNotes = (sale.Notes ? sale.Notes + ' | ' : '') + `Cancelled on ${now()}`;
    await run(db, 'UPDATE Sales SET Status=?, Notes=? WHERE ID=?', 'Cancelled', newNotes, saleId);

    if (num(sale.PaidAmount) > 0) {
      await insertRow(db, 'Payments', {
        ID: uid('PAY-'), Date: now(), InvoiceNo: sale.InvoiceNo,
        CustomerID: sale.CustomerID, CustomerName: sale.CustomerName,
        Amount: -num(sale.PaidAmount), PaymentMethod: 'Adjustment',
        ReceivedBy: '', Notes: 'REFUND — Sale cancelled'
      });
    }

    if (sale.CustomerID) await recomputeCustomer(db, sale.CustomerID);

    // Reverse affiliate commissions
    if (sale.AffiliateCode) {
      const comms = await all(db, 'SELECT * FROM AffiliateCommissions WHERE InvoiceNo=?', sale.InvoiceNo);
      for (const cm of comms) {
        if (cm.Status === 'Paid') {
          await run(db, 'UPDATE AffiliateCommissions SET Status=?, Notes=? WHERE ID=?',
            'Reversed', (cm.Notes || '') + ' | Sale cancelled', cm.ID);
        } else {
          await run(db, 'DELETE FROM AffiliateCommissions WHERE ID=?', cm.ID);
          await run(db,
            'UPDATE Affiliates SET PendingCommission=PendingCommission-? WHERE Code=?',
            num(cm.CommissionAmount), sale.AffiliateCode);
        }
      }
    }
    return { success: true, message: 'Sale reversed successfully' };
  },
  
  deleteSale: async (p, db) => {
    const sale = await one(db, 'SELECT * FROM Sales WHERE ID=?', p.ID);
    if (!sale) return { success: false, error: 'Sale not found' };
    
    // Cascade delete: items, payments, deliveries, commissions
    await run(db, 'DELETE FROM SalesItems WHERE InvoiceNo=?', sale.InvoiceNo);
    await run(db, 'DELETE FROM Payments WHERE InvoiceNo=?', sale.InvoiceNo);
    await run(db, 'DELETE FROM Deliveries WHERE InvoiceNo=?', sale.InvoiceNo);
    await run(db, 'DELETE FROM AffiliateCommissions WHERE InvoiceNo=?', sale.InvoiceNo);
    await run(db, 'DELETE FROM Sales WHERE ID=?', p.ID);
    
    if (sale.CustomerID) await recomputeCustomer(db, sale.CustomerID);
    return { success: true };
  },
  
  // FIXED: matches frontend params
  adjustSaleAmount: async (p, db) => {
    const saleId = p.ID || p.SaleID;
    const sale = await one(db, 'SELECT * FROM Sales WHERE ID=?', saleId);
    if (!sale) return { success: false, error: 'Sale not found' };

    const oldAmount = num(sale.FinalAmount);
    const newAmount = num(p.NewFinalAmount ?? p.NewAmount);
    const diff = oldAmount - newAmount;
    const type = newAmount < oldAmount ? 'Reduced' : 'Increased';
    const ts = now();
    const method = p.AdjustmentMethod || p.Mode || 'Price Reduction';

    const adjNote = `[ADJUSTMENT ${ts}] ${type} ${sale.InvoiceNo} by ${Math.abs(diff)}. Old: ${oldAmount} → New: ${newAmount}. Reason: ${p.Reason || 'N/A'}`;
    const newNotes = (sale.Notes ? sale.Notes + ' | ' : '') + adjNote;

    let newPaid = num(sale.PaidAmount);

    // If "Refund" method and reduced → create refund payment
    if ((method === 'Refund' || p.Mode === 'refund') && type === 'Reduced' && diff > 0) {
      await insertRow(db, 'Payments', {
        ID: uid('PAY-'), Date: ts, InvoiceNo: sale.InvoiceNo,
        CustomerID: sale.CustomerID, CustomerName: sale.CustomerName,
        Amount: -diff, PaymentMethod: p.RefundMethod || 'Cash',
        ReceivedBy: '', Notes: `REFUND — Sale adjustment. Reason: ${p.Reason || ''}`
      });
      newPaid -= diff;
    }

    // Cap newPaid if it exceeds new total
    if (newPaid > newAmount) newPaid = newAmount;
    const newDue = Math.max(0, newAmount - newPaid);

    let newStatus = 'Unpaid';
    if (newPaid >= newAmount) newStatus = 'Paid';
    else if (newPaid > 0) newStatus = 'Partial';

    const newProfit = newAmount - num(sale.CostTotal) - num(sale.DeliveryCharge);

    await run(db,
      'UPDATE Sales SET FinalAmount=?, PaidAmount=?, DueAmount=?, Profit=?, Status=?, Notes=? WHERE ID=?',
      newAmount, newPaid, newDue, newProfit, newStatus, newNotes, saleId);

    if (sale.CustomerID) await recomputeCustomer(db, sale.CustomerID);

    // Optional: proportional commission adjustment
    if (p.AdjustCommission && sale.AffiliateCode && oldAmount > 0) {
      const ratio = newAmount / oldAmount;
      const comms = await all(db, 'SELECT * FROM AffiliateCommissions WHERE InvoiceNo=? AND Status=?', sale.InvoiceNo, 'Unpaid');
      for (const cm of comms) {
        const newCommAmt = num(cm.CommissionAmount) * ratio;
        const delta = num(cm.CommissionAmount) - newCommAmt;
        await run(db, 'UPDATE AffiliateCommissions SET CommissionAmount=? WHERE ID=?', newCommAmt, cm.ID);
        await run(db,
          'UPDATE Affiliates SET PendingCommission=PendingCommission-? WHERE Code=?',
          delta, sale.AffiliateCode);
      }
    }

    return { success: true, message: `Sale adjusted: ${type} by ${Math.abs(diff)}` };
  },

  // NEW: partial return
  partialReverseSale: async (p, db) => {
  const sale = await one(db, 'SELECT * FROM Sales WHERE ID=?', p.SaleID);
  if (!sale) return { success: false, error: 'Sale not found' };
  const items = p.Items || [];
  if (!items.length) return { success: false, error: 'No items to return' };

  // Load customization costs once
  const patchCost = num(await getSetting(db, 'patchDefaultCost', 130));
  const namePrintCost = num(await getSetting(db, 'namePrintDefaultCost', 130));

  let refundAmount = 0;
  let itemsReturned = 0;
  let totalCostReturned = 0;

  for (const ret of items) {
    const retQty = num(ret.ReturnQty);
    if (retQty <= 0) continue;
    itemsReturned += retQty;

    const lineRefund = (num(ret.UnitPrice) + num(ret.ModPerUnit || 0)) * retQty;
    refundAmount += lineRefund;
    
    // Cost reversal must mirror what was charged at sale time
    const hadPatch = ret.Patch === 'Yes' || num(ret.PatchPrice) > 0;
    const hadNamePrint = !!ret.NamePrint || num(ret.NamePrintPrice) > 0;
    const patchC = hadPatch ? patchCost : 0;
    const nameC = hadNamePrint ? namePrintCost : 0;
    totalCostReturned += (num(ret.CostPrice) + patchC + nameC) * retQty;

      // Restore stock
      if (ret.ProductID) {
        const prod = await one(db, 'SELECT Stock,ProductName FROM Inventory WHERE ID=?', ret.ProductID);
        if (prod) {
          const oldS = num(prod.Stock);
          const newS = oldS + retQty;
          await run(
            db,
            'UPDATE Inventory SET Stock=?, LastUpdated=? WHERE ID=?',
            newS,
            now(),
            ret.ProductID
          );

          // Restore FIFO batch quantities
          await restoreBatchesFIFO(
            db,
            ret.ProductID,
            retQty
          );
          await insertRow(db, 'StockLog', {
            ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'),
            Date: now(),
            ProductID: ret.ProductID,
            ProductName: prod.ProductName,
            Action: p.IsExchange ? 'Exchange' : 'Return',
            Quantity: retQty,
            OldStock: oldS,
            NewStock: newS,
            Reference: sale.InvoiceNo,
            Notes: p.Notes || ''
          });
        }
      }

      // Reduce the SalesItem qty
      const si = await one(db, 'SELECT * FROM SalesItems WHERE ID=?', ret.SalesItemID);
      if (si) {
        const newQty = num(si.Quantity) - retQty;
        if (newQty <= 0) {
          await run(db, 'DELETE FROM SalesItems WHERE ID=?', ret.SalesItemID);
        } else {
          const newTotal = newQty * num(si.UnitPrice);
          const newMod = newQty * (num(si.PatchPrice) + num(si.NamePrintPrice));
          await run(db, 'UPDATE SalesItems SET Quantity=?, TotalPrice=?, ModTotal=?, LineTotal=? WHERE ID=?',
            newQty, newTotal, newMod, newTotal + newMod, ret.SalesItemID);
        }
      }
    }

    // Update Sales totals
    const newFinal = num(sale.FinalAmount) - refundAmount;
    const newCost = num(sale.CostTotal) - totalCostReturned;
    let newPaid = num(sale.PaidAmount);
    if (!p.IsExchange) {
      // Issue refund only if not exchange
      if (refundAmount > 0 && newPaid > 0) {
        const refundPay = Math.min(refundAmount, newPaid);
        await insertRow(db, 'Payments', {
          ID: uid('PAY-'), Date: now(), InvoiceNo: sale.InvoiceNo,
          CustomerID: sale.CustomerID, CustomerName: sale.CustomerName,
          Amount: -refundPay, PaymentMethod: p.RefundMethod || 'Cash',
          ReceivedBy: '', Notes: `REFUND — Partial return. ${p.Notes || ''}`
        });
        newPaid -= refundPay;
      }
    }
    const newDue = Math.max(0, newFinal - newPaid);
    let newStatus = 'Unpaid';
    if (newPaid >= newFinal && newFinal > 0) newStatus = 'Paid';
    else if (newPaid > 0) newStatus = 'Partial';
    if (newFinal <= 0) newStatus = 'Cancelled';

    const notesUpd = (sale.Notes ? sale.Notes + ' | ' : '') + `Partial return on ${now()}: ${itemsReturned} item(s), refund ${refundAmount}. ${p.Notes || ''}`;

    await run(db,
      'UPDATE Sales SET FinalAmount=?, CostTotal=?, Profit=?, PaidAmount=?, DueAmount=?, Status=?, Notes=? WHERE ID=?',
      newFinal, newCost, newFinal - newCost - num(sale.DeliveryCharge),
      newPaid, newDue, newStatus, notesUpd, p.SaleID);

    if (sale.CustomerID) await recomputeCustomer(db, sale.CustomerID);

    return { success: true, itemsReturned, refundAmount };
  },

  // NEW: exchange & replace — edits ORIGINAL invoice in place, no new sale created
  exchangeAndReplace: async (p, db) => {
    const sale = await one(db, 'SELECT * FROM Sales WHERE ID=?', p.SaleID);
    if (!sale) return { success: false, error: 'Sale not found' };
    if (sale.Status === 'Cancelled') return { success: false, error: 'Cannot modify cancelled sale' };

    const returnItems = p.ReturnItems || [];
    const replaceItems = p.ReplaceItems || [];
    
    if (!returnItems.length) return { success: false, error: 'No items to return' };
    if (!replaceItems.length) return { success: false, error: 'No replacement items selected' };

    // Load customization costs (used for both returns and replacements)
    const patchCost = num(await getSetting(db, 'patchDefaultCost', 130));
    const namePrintCost = num(await getSetting(db, 'namePrintDefaultCost', 130));

    // ========================================================
    // STEP 1: Process RETURNED items
    // - Restore stock
    // - Remove/reduce SalesItems rows
    // - Track return credit & returned cost
    // ========================================================
    let returnCredit = 0;        // money credited back from returned items
    let returnedCost = 0;        // cost of returned items (to subtract from CostTotal)
    let returnedQty = 0;
    const returnLog = [];        // for notes

    for (const ret of returnItems) {
      const retQty = num(ret.ReturnQty);
      if (retQty <= 0) continue;
      returnedQty += retQty;

      // Calculate credit (selling price + customization fees) per unit, times qty
      const lineCredit = (num(ret.UnitPrice) + num(ret.ModPerUnit || 0)) * retQty;
      returnCredit += lineCredit;

      // Cost reversal: product cost + (patch cost if applied) + (name print cost if applied)
      const hadPatch = ret.Patch === 'Yes' || num(ret.PatchPrice) > 0;
      const hadNamePrint = !!ret.NamePrint || num(ret.NamePrintPrice) > 0;
      const patchC = hadPatch ? patchCost : 0;
      const nameC = hadNamePrint ? namePrintCost : 0;
      returnedCost += (num(ret.CostPrice) + patchC + nameC) * retQty;

      // Restore stock
      if (ret.ProductID) {
        const prod = await one(db, 'SELECT Stock, ProductName, CostPrice FROM Inventory WHERE ID=?', ret.ProductID);
        if (prod) {
          const oldS = num(prod.Stock);
          const newS = oldS + retQty;
          await run(
            db,
            'UPDATE Inventory SET Stock=?, LastUpdated=? WHERE ID=?',
            newS,
            now(),
            ret.ProductID
          );

          // Restore consumed FIFO quantities
          await restoreBatchesFIFO(
            db,
            ret.ProductID,
            retQty
          );
          await insertRow(db, 'StockLog', {
            ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'),
            Date: now(),
            ProductID: ret.ProductID,
            ProductName: prod.ProductName,
            Action: 'Exchange Return',
            Quantity: retQty,
            OldStock: oldS,
            NewStock: newS,
            Reference: sale.InvoiceNo,
            Notes: 'Exchange & Replace: ' + (p.Notes || '')
          });
        }
      }

      // Reduce or remove the SalesItem row
      if (ret.SalesItemID) {
        const si = await one(db, 'SELECT * FROM SalesItems WHERE ID=?', ret.SalesItemID);
        if (si) {
          const newQty = num(si.Quantity) - retQty;
          if (newQty <= 0) {
            await run(db, 'DELETE FROM SalesItems WHERE ID=?', ret.SalesItemID);
          } else {
            const newTotalPrice = newQty * num(si.UnitPrice);
            const newModTotal = newQty * (num(si.PatchPrice) + num(si.NamePrintPrice));
            await run(db, 
              'UPDATE SalesItems SET Quantity=?, TotalPrice=?, ModTotal=?, LineTotal=? WHERE ID=?',
              newQty, newTotalPrice, newModTotal, newTotalPrice + newModTotal, ret.SalesItemID);
          }
        }
      }

      returnLog.push(`${ret.ProductName || 'Item'} x${retQty}`);
    }

    // ========================================================
    // STEP 2: Process REPLACEMENT items
    // - Add new SalesItems rows
    // - Decrement stock
    // - Track replacement total & cost
    // ========================================================
    let replaceTotal = 0;
    let replaceCost = 0;
    const replaceLog = [];

    for (const rep of replaceItems) {
      const qty = num(rep.Quantity);
      if (qty <= 0) continue;

      const unitPrice = num(rep.UnitPrice);
      const productCost = num(rep.CostPrice);
      const hasPatch = rep.Patch === true || rep.Patch === 'Yes';
      const hasNamePrint = !!rep.NamePrint;
      const patchP = hasPatch ? num(rep.PatchPrice) : 0;
      const nameP = hasNamePrint ? num(rep.NamePrintPrice) : 0;
      const modPerUnit = patchP + nameP;
      const modTotal = modPerUnit * qty;
      const lineTotal = (unitPrice * qty) + modTotal;

      replaceTotal += lineTotal;

      // Cost = product cost + (patch material cost) + (name print material cost), per unit, × qty
      const patchC = hasPatch ? patchCost : 0;
      const nameC = hasNamePrint ? namePrintCost : 0;
      replaceCost += ((rep._fifoUnitCost ?? productCost) + patchC + nameC) * qty;

      // Insert new SalesItem
      await insertRow(db, 'SalesItems', {
        ID: uid('SI-'),
        InvoiceNo: sale.InvoiceNo,
        ProductID: (rep.ProductID === '__CUSTOMIZATION__') ? '' : (rep.ProductID || ''),
        ProductName: rep.ProductName || '',
        Club_Country: rep.Club_Country || '',
        Edition: rep.Edition || '',
        Type: rep.Type || '',
        Size: rep.Size || '',
        Quantity: qty,
        UnitPrice: unitPrice,
        CostPrice: productCost,
        TotalPrice: unitPrice * qty,
        Patch: hasPatch ? 'Yes' : '',
        NamePrint: rep.NamePrint || '',
        PatchPrice: patchP,
        NamePrintPrice: nameP,
        ModTotal: modTotal,
        LineTotal: lineTotal
      });

      // Decrement stock (skip for customization service items)
      const isService = rep.IsCustomizationService === true 
        || rep.ProductID === '__CUSTOMIZATION__' 
        || !rep.ProductID;
        
      if (!isService) {
        const prod = await one(db, 'SELECT Stock, ProductName FROM Inventory WHERE ID=?', rep.ProductID);
        if (prod) {
          const oldS = num(prod.Stock);
          const newS = oldS - qty;
          await run(
            db,
            'UPDATE Inventory SET Stock=?, LastUpdated=? WHERE ID=?',
            newS,
            now(),
            rep.ProductID
          );

          // Consume FIFO layers for replacement item
          const { actualCost } = await consumeBatchesFIFO(db, rep.ProductID, qty, prod.CostPrice);
          // Optional: make exchange profit exact
          rep._fifoUnitCost =
            qty > 0
              ? actualCost / qty
              : num(prod.CostPrice);
          await insertRow(db, 'StockLog', {
            ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'),
            Date: now(),
            ProductID: rep.ProductID,
            ProductName: prod.ProductName,
            Action: 'Exchange Replace',
            Quantity: -qty,
            OldStock: oldS,
            NewStock: newS,
            Reference: sale.InvoiceNo,
            Notes: 'Exchange & Replace: ' + (p.Notes || '')
          });
        }
      }

      replaceLog.push(`${rep.ProductName || 'Item'} x${qty}`);
    }

    // ========================================================
    // STEP 3: Apply EXTRA DISCOUNT (if any)
    // ========================================================
    const extraDiscount = num(p.ExtraDiscount);

    // ========================================================
    // STEP 4: Recalculate Sale totals
    // ========================================================
    // New TotalAmount (subtotal before any discount) = old total - returned + replaced
    const newTotal = num(sale.TotalAmount) - returnCredit + replaceTotal;
    
    // New Discount: original discount + any extra exchange discount
    const newDiscount = num(sale.Discount) + extraDiscount;
    
    // New CostTotal = old cost - returned cost + replacement cost
    const newCostTotal = num(sale.CostTotal) - returnedCost + replaceCost;

    // FinalAmount = newTotal - newDiscount - affiliate disc - voucher disc + delivery
    const newFinalAmount = newTotal 
      - newDiscount 
      - num(sale.AffiliateDiscount) 
      - num(sale.VoucherDiscount) 
      + num(sale.DeliveryCharge);

    // Settlement = how much MORE the customer owes (positive) or is owed (negative)
    // After this exchange, the net change in price is: (replaceTotal - returnCredit - extraDiscount)
    const settlement = replaceTotal - returnCredit - extraDiscount;

    // Track payments
    let newPaid = num(sale.PaidAmount);
    
    // If customer pays extra at the moment of exchange
    if (settlement > 0 && p.SettleMethod && p.SettleMethod !== 'No Settlement') {
      // Record the extra payment received
      await insertRow(db, 'Payments', {
        ID: uid('PAY-'),
        Date: now(),
        InvoiceNo: sale.InvoiceNo,
        CustomerID: sale.CustomerID || '',
        CustomerName: sale.CustomerName,
        Amount: settlement,
        PaymentMethod: p.SettleMethod || 'Cash',
        ReceivedBy: '',
        Notes: `Exchange settlement — customer paid extra. Returned: ${returnLog.join(', ')}. Replaced with: ${replaceLog.join(', ')}.`
      });
      newPaid += settlement;
    } 
    // If we owe the customer (refund)
    else if (settlement < 0 && p.SettleMethod && p.SettleMethod !== 'No Settlement') {
      // Record refund as negative payment
      await insertRow(db, 'Payments', {
        ID: uid('PAY-'),
        Date: now(),
        InvoiceNo: sale.InvoiceNo,
        CustomerID: sale.CustomerID || '',
        CustomerName: sale.CustomerName,
        Amount: settlement, // negative
        PaymentMethod: p.SettleMethod || 'Cash',
        ReceivedBy: '',
        Notes: `Exchange refund — customer refunded ${Math.abs(settlement)}. Returned: ${returnLog.join(', ')}. Replaced with: ${replaceLog.join(', ')}.`
      });
      newPaid += settlement; // adding a negative reduces newPaid
    }
    // If even exchange or No Settlement — no payment record needed

    // Compute new due & status
    const newDue = Math.max(0, newFinalAmount - newPaid);
    let newStatus = 'Unpaid';
    if (newPaid >= newFinalAmount && newFinalAmount > 0) newStatus = 'Paid';
    else if (newPaid > 0 && newDue > 0) newStatus = 'Partial';
    else if (newFinalAmount <= 0) newStatus = 'Paid';

    // Compute new profit
    const newProfit = newFinalAmount - newCostTotal - num(sale.DeliveryCharge);

    // ========================================================
    // STEP 5: Build exchange note for invoice
    // ========================================================
    const ts = now();
    const exchangeNote = `[EXCHANGE ${ts}] Returned: ${returnLog.join(', ')} → Replaced with: ${replaceLog.join(', ')}. Net settlement: ${settlement >= 0 ? '+' : ''}${settlement.toFixed(2)}. ${p.Notes ? 'Reason: ' + p.Notes : ''}`;
    const updatedNotes = sale.Notes ? sale.Notes + ' | ' + exchangeNote : exchangeNote;

    // ========================================================
    // STEP 6: Rebuild Items summary (the JSON blob in Sales.Items column)
    // ========================================================
    const remainingItems = await all(db, 'SELECT * FROM SalesItems WHERE InvoiceNo=?', sale.InvoiceNo);
    const itemsBlob = JSON.stringify(remainingItems.map(it => ({
      n: it.ProductName, q: it.Quantity, p: it.UnitPrice
    })));
    const modsBlob = JSON.stringify(remainingItems.filter(it => it.Patch === 'Yes' || it.NamePrint).map(it => ({
      n: it.ProductName, patch: it.Patch, np: it.NamePrint
    })));

    // ========================================================
    // STEP 7: UPDATE the Sales row in-place
    // ========================================================
    await run(db, `
      UPDATE Sales SET 
        Items = ?,
        Modifications = ?,
        TotalAmount = ?,
        Discount = ?,
        FinalAmount = ?,
        PaidAmount = ?,
        DueAmount = ?,
        Status = ?,
        CostTotal = ?,
        Profit = ?,
        Notes = ?
      WHERE ID = ?
    `, itemsBlob, modsBlob, newTotal, newDiscount, newFinalAmount, newPaid, newDue, newStatus, newCostTotal, newProfit, updatedNotes, p.SaleID);

    // ========================================================
    // STEP 8: Recompute customer totals
    // ========================================================
    if (sale.CustomerID) await recomputeCustomer(db, sale.CustomerID);

    // ========================================================
    // STEP 9: Done! Return summary
    // ========================================================
    return {
      success: true,
      invoiceNo: sale.InvoiceNo,
      returnedQty,
      returnCredit,
      replaceTotal,
      settlement,
      newFinalAmount,
      newDue,
      newStatus,
      message: settlement > 0 
        ? `Exchange complete. Customer pays ${settlement.toFixed(2)} extra.`
        : settlement < 0 
          ? `Exchange complete. Refund ${Math.abs(settlement).toFixed(2)} to customer.`
          : `Even exchange complete. No money changed hands.`
    };
  },

  // NEW: updateSaleProfit
  updateSaleProfit: async (p, db) => {
    await run(db, 'UPDATE Sales SET Profit=?, ProfitOverride=? WHERE ID=?',
      num(p.Profit), num(p.ProfitOverride), p.ID);
    return { success: true };
  },

  // ============== PAYMENTS ==============
  getPayments: async (p, db) => await all(db, 'SELECT * FROM Payments ORDER BY Date DESC'),

  // FIXED: frontend sends InvoiceNo (not SaleID)
  addPayment: async (p, db) => {
    const amt = num(p.Amount);
    const row = {
      ID: uid('PAY-'), Date: p.Date || now(), InvoiceNo: p.InvoiceNo || '',
      CustomerID: p.CustomerID || '', CustomerName: p.CustomerName || '',
      Amount: amt, PaymentMethod: p.PaymentMethod || 'Cash',
      ReceivedBy: p.ReceivedBy || '', Notes: p.Notes || ''
    };
    await insertRow(db, 'Payments', row);

    if (p.InvoiceNo) {
      const sale = await one(db, 'SELECT * FROM Sales WHERE InvoiceNo=?', p.InvoiceNo);
      if (sale) {
        const newPaid = num(sale.PaidAmount) + amt;
        const newDue = Math.max(0, num(sale.FinalAmount) - newPaid);
        let newStatus = 'Unpaid';
        if (newPaid >= num(sale.FinalAmount)) newStatus = 'Paid';
        else if (newPaid > 0) newStatus = 'Partial';
        await run(db, 'UPDATE Sales SET PaidAmount=?, DueAmount=?, Status=? WHERE InvoiceNo=?',
          newPaid, newDue, newStatus, p.InvoiceNo);
        if (sale.CustomerID) await recomputeCustomer(db, sale.CustomerID);
      }
    }
    return { success: true, id: row.ID };
  },

  // ============== SUPPLIERS ==============
  getSuppliers: async (p, db) => await all(db, 'SELECT * FROM Suppliers ORDER BY Name'),
  getSupplierLedger: async (p, db) => p.supplierId
    ? await all(db, 'SELECT * FROM SupplierLedger WHERE SupplierID=? ORDER BY Date DESC', p.supplierId)
    : await all(db, 'SELECT * FROM SupplierLedger ORDER BY Date DESC'),

  addSupplier: async (p, db) => {
    const id = uid('SUP-');
    await insertRow(db, 'Suppliers', {
      ID: id, Name: p.Name, Phone: fmtPhone(p.Phone || ''), Email: p.Email || '',
      Address: p.Address || '', TotalOrders: 0, TotalAmount: 0, PendingAmount: 0,
      Notes: p.Notes || ''
    });
    return { success: true, id };
  },

  updateSupplier: async (p, db) => {
    const cur = await one(db, 'SELECT * FROM Suppliers WHERE ID=?', p.ID);
    if (!cur) return { success: false, error: 'Supplier not found' };
    await updateRow(db, 'Suppliers', p.ID, {
      Name: p.Name ?? cur.Name,
      Phone: fmtPhone(p.Phone ?? cur.Phone),
      Email: p.Email ?? cur.Email,
      Address: p.Address ?? cur.Address,
      Notes: p.Notes ?? cur.Notes
    });
    return { success: true };
  },

  deleteSupplier: async (p, db) => {
    await run(db, 'DELETE FROM Suppliers WHERE ID=?', p.ID);
    return { success: true };
  },

  // FIXED: action name matches frontend
  addSupplierLedgerEntry: async (p, db) => {
    const sup = await one(db, 'SELECT * FROM Suppliers WHERE ID=?', p.SupplierID);
    if (!sup) return { success: false, error: 'Supplier not found' };
    const debit = num(p.Debit);
    const credit = num(p.Credit);
    const newBalance = num(sup.PendingAmount) + debit - credit;
    await insertRow(db, 'SupplierLedger', {
      ID: uid('SLG-'), Date: p.Date || now(),
      SupplierID: p.SupplierID, SupplierName: sup.Name,
      Description: p.Description || '', Debit: debit, Credit: credit,
      Balance: newBalance, Notes: p.Notes || ''
    });
    await run(db,
      'UPDATE Suppliers SET TotalOrders=TotalOrders+?, TotalAmount=TotalAmount+?, PendingAmount=? WHERE ID=?',
      debit > 0 ? 1 : 0, debit, newBalance, p.SupplierID);
    return { success: true };
  },

  // ============== EXPENSES ==============
  getExpenses: async (p, db) => await all(db, 'SELECT * FROM Expenses ORDER BY Date DESC'),

  addExpense: async (p, db) => {
    const id = uid('EXP-');
    await insertRow(db, 'Expenses', {
      ID: id, Date: p.Date || now(), Category: p.Category || 'Other',
      Description: p.Description || '', Amount: num(p.Amount),
      PaidTo: p.PaidTo || '', PaymentMethod: p.PaymentMethod || 'Cash',
      Notes: p.Notes || ''
    });
    return { success: true, id };
  },

  deleteExpense: async (p, db) => {
    await run(db, 'DELETE FROM Expenses WHERE ID=?', p.ID);
    return { success: true };
  },

  // ============== AFFILIATES ==============
  // FIXED: compute TotalCommission on-the-fly
  getAffiliates: async (p, db) => {
    const rows = await all(db, 'SELECT * FROM Affiliates ORDER BY Code');
    return rows.map(r => ({
      ...r,
      TotalCommission: num(r.PaidCommission) + num(r.PendingCommission)
    }));
  },

  getAffiliateCommissions: async (p, db) => p.affiliateId
    ? await all(db, 'SELECT * FROM AffiliateCommissions WHERE AffiliateID=? ORDER BY Date DESC', p.affiliateId)
    : await all(db, 'SELECT * FROM AffiliateCommissions ORDER BY Date DESC'),

  addAffiliate: async (p, db) => {
    const id = uid('AFF-');
    await insertRow(db, 'Affiliates', {
      ID: id, Code: p.Code, Name: p.Name || '',
      DiscountType: p.DiscountType || 'Percentage', DiscountValue: num(p.DiscountValue),
      TotalUsed: 0, TotalDiscountGiven: 0,
      Active: (p.Active === false || p.Active === 'No') ? 'No' : 'Yes',
      CommissionFanEdition: num(p.CommFan),
      CommissionPlayerEdition: num(p.CommPlayer),
      CommissionSpecialEdition: num(p.CommSpecial),
      CommissionBDPremium: num(p.CommBD),
      PaidCommission: 0, PendingCommission: 0
    });
    return { success: true, id };
  },

  updateAffiliate: async (p, db) => {
    const cur = await one(db, 'SELECT * FROM Affiliates WHERE ID=?', p.ID);
    if (!cur) return { success: false, error: 'Affiliate not found' };
    await updateRow(db, 'Affiliates', p.ID, {
      Code: p.Code ?? cur.Code,
      Name: p.Name ?? cur.Name,
      DiscountType: p.DiscountType ?? cur.DiscountType,
      DiscountValue: num(p.DiscountValue ?? cur.DiscountValue),
      Active: (p.Active === false || p.Active === 'No') ? 'No' : 'Yes',
      CommissionFanEdition: num(p.CommFan ?? cur.CommissionFanEdition),
      CommissionPlayerEdition: num(p.CommPlayer ?? cur.CommissionPlayerEdition),
      CommissionSpecialEdition: num(p.CommSpecial ?? cur.CommissionSpecialEdition),
      CommissionBDPremium: num(p.CommBD ?? cur.CommissionBDPremium)
    });
    return { success: true };
  },

  deleteAffiliate: async (p, db) => {
    await run(db, 'DELETE FROM Affiliates WHERE ID=?', p.ID);
    return { success: true };
  },

  payAffiliateCommission: async (p, db) => {
    const amt = num(p.Amount);
    if (amt <= 0) return { success: false, error: 'Invalid amount' };
    const expId = uid('EXP-');
    await insertRow(db, 'Expenses', {
      ID: expId, Date: now(), Category: 'Affiliate Payout',
      Description: `Commission paid to ${p.AffiliateName} (${p.AffiliateCode})`,
      Amount: amt, PaidTo: p.AffiliateName || '',
      PaymentMethod: p.PaymentMethod || 'Cash', Notes: `Affiliate: ${p.AffiliateID}`
    });

    let remaining = amt;
    const unpaid = await all(db,
      'SELECT * FROM AffiliateCommissions WHERE AffiliateID=? AND Status=? ORDER BY Date',
      p.AffiliateID, 'Unpaid');
    for (const cm of unpaid) {
      if (remaining <= 0.01) break;
      const cAmt = num(cm.CommissionAmount);
      if (cAmt <= remaining + 0.01) {
        await run(db,
          'UPDATE AffiliateCommissions SET Status=?, PaidDate=?, ExpenseID=? WHERE ID=?',
          'Paid', now(), expId, cm.ID);
        remaining -= cAmt;
      }
    }
    await run(db,
      'UPDATE Affiliates SET PaidCommission=PaidCommission+?, PendingCommission=PendingCommission-? WHERE ID=?',
      amt, amt, p.AffiliateID);
    return { success: true, expenseId: expId };
  },

  reverseAffiliatePayment: async (p, db) => {
    const cm = await one(db, 'SELECT * FROM AffiliateCommissions WHERE ID=?', p.CommissionID);
    if (!cm) return { success: false, error: 'Commission not found' };
    await run(db, 'UPDATE AffiliateCommissions SET Status=?, PaidDate=NULL, ExpenseID=NULL WHERE ID=?', 'Unpaid', p.CommissionID);
    await run(db,
      'UPDATE Affiliates SET PaidCommission=PaidCommission-?, PendingCommission=PendingCommission+? WHERE ID=?',
      num(cm.CommissionAmount), num(cm.CommissionAmount), cm.AffiliateID);
    if (cm.ExpenseID) await run(db, 'DELETE FROM Expenses WHERE ID=?', cm.ExpenseID);
    return { success: true };
  },
  
  // ============== VOUCHERS ==============
  getVouchers: async (p, db) => await all(db, 'SELECT * FROM Vouchers ORDER BY Code'),

  addVoucher: async (p, db) => {
    const id = uid('VCH-');
    await insertRow(db, 'Vouchers', {
      ID: id, Code: p.Code, Name: p.Name || '',
      DiscountType: p.DiscountType || 'Fixed', DiscountValue: num(p.DiscountValue),
      MinOrder: num(p.MinOrder), MaxUses: num(p.MaxUses),
      TimesUsed: 0, TotalDiscountGiven: 0,
      ValidFrom: p.ValidFrom || '', ValidTo: p.ValidTo || '',
      Active: (p.Active === false || p.Active === 'No') ? 'No' : 'Yes'
    });
    return { success: true, id };
  },

  updateVoucher: async (p, db) => {
    const cur = await one(db, 'SELECT * FROM Vouchers WHERE ID=?', p.ID);
    if (!cur) return { success: false, error: 'Voucher not found' };
    await updateRow(db, 'Vouchers', p.ID, {
      Code: p.Code ?? cur.Code,
      Name: p.Name ?? cur.Name,
      DiscountType: p.DiscountType ?? cur.DiscountType,
      DiscountValue: num(p.DiscountValue ?? cur.DiscountValue),
      MinOrder: num(p.MinOrder ?? cur.MinOrder),
      MaxUses: num(p.MaxUses ?? cur.MaxUses),
      Active: (p.Active === false || p.Active === 'No') ? 'No' : 'Yes'
    });
    return { success: true };
  },

  deleteVoucher: async (p, db) => {
    await run(db, 'DELETE FROM Vouchers WHERE ID=?', p.ID);
    return { success: true };
  },

  // ============== DELIVERIES ==============
  getDeliveries: async (p, db) => await all(db, 'SELECT * FROM Deliveries ORDER BY Date DESC'),

  // FIXED: action name matches frontend
  updateDeliveryStatus: async (p, db) => {
    const cur = await one(db, 'SELECT * FROM Deliveries WHERE ID=?', p.ID);
    if (!cur) return { success: false, error: 'Delivery not found' };
    await updateRow(db, 'Deliveries', p.ID, {
      CourierService: p.CourierService ?? cur.CourierService,
      TrackingNo: p.TrackingNo ?? cur.TrackingNo,
      DeliveryStatus: p.DeliveryStatus ?? cur.DeliveryStatus,
      Notes: p.Notes ?? cur.Notes
    });
    // Mirror to Sales
    if (p.InvoiceNo) {
      await run(db,
        'UPDATE Sales SET CourierService=?, TrackingNo=?, DeliveryStatus=? WHERE InvoiceNo=?',
        p.CourierService ?? cur.CourierService,
        p.TrackingNo ?? cur.TrackingNo,
        p.DeliveryStatus ?? cur.DeliveryStatus,
        p.InvoiceNo);
    }
    return { success: true };
  },
  updateSaleDelivery: async (p, db) => {
    // Alias that only updates the Sales table
    if (!p.InvoiceNo) return { success: false, error: 'InvoiceNo required' };
    const fields = {};
    if (p.DeliveryStatus !== undefined) fields.DeliveryStatus = p.DeliveryStatus;
    if (p.CourierService !== undefined) fields.CourierService = p.CourierService;
    if (p.TrackingNo !== undefined) fields.TrackingNo = p.TrackingNo;
    
    const keys = Object.keys(fields);
    if (!keys.length) return { success: true };
    
    const sets = keys.map(k => `"${k}"=?`).join(',');
    await db.prepare(`UPDATE Sales SET ${sets} WHERE InvoiceNo=?`)
      .bind(...keys.map(k => fields[k]), p.InvoiceNo).run();
    return { success: true };
  },

  // ============== STOCK LOG ==============
  getStockLog: async (p, db) => {
    const limit = num(p.limit) || 5000;
    return p.productId
      ? await all(db, 'SELECT * FROM StockLog WHERE ProductID=? ORDER BY Date DESC LIMIT ?', p.productId, limit)
      : await all(db, 'SELECT * FROM StockLog ORDER BY Date DESC LIMIT ?', limit);
  },

  
  // ============== DEAD STOCK ==============
  getDeadStockHistory: async (p, db) => await all(db, 'SELECT * FROM DeadStockHistory ORDER BY SnapshotDate DESC'),

  saveDeadStockSnapshot: async (p, db) => {
    const date = today();
    const month = date.slice(0, 7);
    const items = p.items || [];

    // Replace existing snapshot for this month
    const existing = await all(db, 'SELECT SnapshotID FROM DeadStockHistory WHERE Month=?', month);
    const replaced = existing.length > 0;
    if (replaced) {
      await run(db, 'DELETE FROM DeadStockHistory WHERE Month=?', month);
    }

    for (const it of items) {
      await insertRow(db, 'DeadStockHistory', {
        SnapshotID: uid('DS-'), SnapshotDate: date, Month: month,
        ProductID: it.ProductID || '', ProductName: it.ProductName || '',
        Club_Country: it.Club_Country || '', Edition: it.Edition || '',
        Type: it.Type || '', Size: it.Size || '',
        Stock: num(it.Stock), CostPrice: num(it.CostPrice),
        SellPrice: num(it.SellPrice), CapitalTied: num(it.CapitalTied),
        DaysInactive: num(it.DaysInactive), Urgency: it.Urgency || ''
      });
    }
    return { success: true, itemCount: items.length, replaced };
  },

  // NEW: delete a month's snapshot
  deleteDeadStockSnapshot: async (p, db) => {
    const r = await run(db, 'DELETE FROM DeadStockHistory WHERE Month=?', p.monthKey);
    return { success: true, deleted: r.meta?.changes || 0 };
  },

  // NEW: chronic dead stock analysis
  getChronicDeadStock: async (p, db) => {
    const all_rows = await all(db, 'SELECT * FROM DeadStockHistory ORDER BY Month DESC');

    // Group by ProductID
    const byProduct = {};
    for (const r of all_rows) {
      const pid = r.ProductID || '';
      if (!byProduct[pid]) {
        byProduct[pid] = {
          ProductID: pid, ProductName: r.ProductName,
          Club_Country: r.Club_Country, Edition: r.Edition,
          Type: r.Type, Size: r.Size,
          appearances: 0, months: [],
          totalCapitalAcrossMonths: 0,
          latestStock: 0, latestDays: 0,
          firstSeen: r.Month, lastSeen: r.Month
        };
      }
      byProduct[pid].appearances++;
      if (!byProduct[pid].months.includes(r.Month)) byProduct[pid].months.push(r.Month);
      byProduct[pid].totalCapitalAcrossMonths += num(r.CapitalTied);
      if (r.Month > byProduct[pid].lastSeen) {
        byProduct[pid].lastSeen = r.Month;
        byProduct[pid].latestStock = num(r.Stock);
        byProduct[pid].latestDays = num(r.DaysInactive);
      }
      if (r.Month < byProduct[pid].firstSeen) byProduct[pid].firstSeen = r.Month;
    }

    const items = Object.values(byProduct)
      .filter(x => x.appearances >= 2)
      .sort((a, b) => b.totalCapitalAcrossMonths - a.totalCapitalAcrossMonths);

    // Category trends
    const aggregateBy = (key) => {
      const map = {};
      for (const r of all_rows) {
        const k = r[key] || 'Unknown';
        if (!map[k]) map[k] = { name: k, appearances: 0, uniqueProducts: new Set(), totalCapital: 0 };
        map[k].appearances++;
        map[k].uniqueProducts.add(r.ProductID);
        map[k].totalCapital += num(r.CapitalTied);
      }
      return Object.values(map)
        .map(x => ({ name: x.name, appearances: x.appearances, uniqueProducts: x.uniqueProducts.size, totalCapital: x.totalCapital }))
        .sort((a, b) => b.totalCapital - a.totalCapital)
        .slice(0, 10);
    };

    return {
      items,
      categoryTrends: {
        clubs: aggregateBy('Club_Country'),
        editions: aggregateBy('Edition'),
        types: aggregateBy('Type')
      }
    };
  },

  // ============== DASHBOARD ==============
  getDashboard: async (p, db) => {
    const todayStr = today();
    const monthStart = todayStr.slice(0, 7) + '-01';

    // Today's sales
    const todayR = await one(db,
      `SELECT COALESCE(SUM(FinalAmount),0) as total, COUNT(*) as cnt, COALESCE(SUM(Profit),0) as profit
       FROM Sales WHERE Date >= ? AND Status != 'Cancelled'`, todayStr);

    // Monthly
    const monthR = await one(db,
      `SELECT COALESCE(SUM(FinalAmount),0) as total, COUNT(*) as cnt, COALESCE(SUM(Profit),0) as profit
       FROM Sales WHERE Date >= ? AND Status != 'Cancelled'`, monthStart);

    // Lifetime
    const totalR = await one(db,
      `SELECT 
        COALESCE(SUM(FinalAmount),0) as revenue, 
        COALESCE(SUM(Profit),0) as profit,
        COALESCE(SUM(DueAmount),0) as pending
       FROM Sales WHERE Status != 'Cancelled'`);

    // Inventory
    const invR = await one(db,
      `SELECT COUNT(*) as products, COALESCE(SUM(Stock),0) as items,
        COALESCE(SUM(Stock * CostPrice),0) as value
       FROM Inventory`);

    const lowStock = await all(db, 'SELECT * FROM Inventory WHERE Stock > 0 AND Stock <= MinStock ORDER BY Stock');
    const outOfStock = await all(db, 'SELECT * FROM Inventory WHERE Stock = 0');

    const unitsSoldR = await one(db, 'SELECT COALESCE(SUM(Quantity),0) as units FROM SalesItems');

    const expensesR = await one(db, 'SELECT COALESCE(SUM(Amount),0) as total FROM Expenses');
    const custCount = await one(db, 'SELECT COUNT(*) as cnt FROM Customers');

    // Recent sales
    const recentSales = await all(db, "SELECT * FROM Sales WHERE Status != 'Cancelled' ORDER BY Date DESC LIMIT 10");
    // Top products by revenue (aggregated from SalesItems)
    const topProducts = await all(db, `
      SELECT 
        si.ProductName as name,
        SUM(si.Quantity) as qty,
        SUM(si.LineTotal) as revenue
      FROM SalesItems si
      INNER JOIN Sales s ON s.InvoiceNo = si.InvoiceNo
      WHERE s.Status != 'Cancelled' AND si.ProductName IS NOT NULL AND si.ProductName != ''
      GROUP BY si.ProductName
      ORDER BY revenue DESC
      LIMIT 10
    `);
    // Monthly sales chart (last 6 months)
    const months = [];
    const monthlyChart = [];
    const now_d = new Date();
    for (let i = 5; i >= 0; i--) {
      const d = new Date(now_d.getFullYear(), now_d.getMonth() - i, 1);
      const mStart = d.toISOString().slice(0, 10);
      const mEnd = new Date(d.getFullYear(), d.getMonth() + 1, 1).toISOString().slice(0, 10);
      const r = await one(db,
        "SELECT COALESCE(SUM(FinalAmount),0) as s FROM Sales WHERE Date >= ? AND Date < ? AND Status != 'Cancelled'",
        mStart, mEnd);
      monthlyChart.push({
        month: d.toLocaleString('en-US', { month: 'short' }),
        sales: num(r.s)
      });
    }

    return {
      todaySales: num(todayR.total),
      todaySaleCount: num(todayR.cnt),
      todayProfit: num(todayR.profit),
      monthlySales: num(monthR.total),
      monthlySaleCount: num(monthR.cnt),
      monthlyProfit: num(monthR.profit),
      totalRevenue: num(totalR.revenue),
      totalProfit: num(totalR.profit),
      totalPending: num(totalR.pending),
      totalProducts: num(invR.products),
      totalItems: num(invR.items),
      totalStockValue: num(invR.value),
      lowStockCount: lowStock.length,
      outOfStockCount: outOfStock.length,
      lowStockItems: lowStock.slice(0, 20),
      outOfStockItems: outOfStock.slice(0, 20),
      totalUnitsSold: num(unitsSoldR.units),
      totalExpenses: num(expensesR.total),
      totalCustomers: num(custCount.cnt),
      recentSales,
      monthlySalesChart: monthlyChart,
      topProducts
    };
  },

  // ============== SETTINGS ==============
  getSettings: async (p, db) => {
    const rows = await all(db, 'SELECT * FROM Settings');
    const obj = {};
    for (const r of rows) obj[r.Key] = r.Value;
    return obj;
  },

  updateSettings: async (p, db) => {
    const s = p.settings || {};
    for (const [k, v] of Object.entries(s)) {
      await setSetting(db, k, v);
    }
    return { success: true };
  },

  // ============== PHONE MIGRATION ==============
  migratePhoneFormats: async (p, db) => {
    let cu = 0, su = 0, sa = 0;
    const customers = await all(db, "SELECT ID, Phone FROM Customers WHERE Phone IS NOT NULL AND Phone != ''");
    for (const c of customers) {
      const f = fmtPhone(c.Phone);
      if (f !== c.Phone) { await run(db, 'UPDATE Customers SET Phone=? WHERE ID=?', f, c.ID); cu++; }
    }
    const suppliers = await all(db, "SELECT ID, Phone FROM Suppliers WHERE Phone IS NOT NULL AND Phone != ''");
    for (const s of suppliers) {
      const f = fmtPhone(s.Phone);
      if (f !== s.Phone) { await run(db, 'UPDATE Suppliers SET Phone=? WHERE ID=?', f, s.ID); su++; }
    }
    const sales = await all(db, "SELECT ID, CustomerPhone FROM Sales WHERE CustomerPhone IS NOT NULL AND CustomerPhone != ''");
    for (const s of sales) {
      const f = fmtPhone(s.CustomerPhone);
      if (f !== s.CustomerPhone) { await run(db, 'UPDATE Sales SET CustomerPhone=? WHERE ID=?', f, s.ID); sa++; }
    }
    return { success: true, customersUpdated: cu, suppliersUpdated: su, salesUpdated: sa };
  },

  // ============================================================
  // BOOKINGS
  // ============================================================

  // Inventory now returns AvailableQty alongside Stock & ReservedQty
  getInventoryWithReserved: async (p, db) => {
    const rows = await all(db, 'SELECT * FROM Inventory ORDER BY Club_Country, Type, Edition, Size');
    return rows.map(r => ({
      ...r,
      ReservedQty: num(r.ReservedQty),
      AvailableQty: num(r.Stock) - num(r.ReservedQty)
    }));
  },

  getBookings: async (p, db) => {
    const bookings = await all(db, 'SELECT * FROM Bookings ORDER BY CreatedAt DESC');
    // Attach item counts / backorder flag for list rendering
    for (const b of bookings) {
      const items = await all(db, 'SELECT * FROM BookingItems WHERE BookingID=?', b.BookingID);
      b.itemCount = items.length;
      b.totalRequested = items.reduce((s, i) => s + num(i.RequestedQty), 0);
      b.totalReserved = items.reduce((s, i) => s + num(i.ReservedQty), 0);
      b.totalBackordered = items.reduce((s, i) => s + num(i.BackorderedQty), 0);
      b.totalValue = items.reduce((s, i) => s + num(i.TotalPrice), 0);
    }
    return bookings;
  },

  getBooking: async (p, db) => {
    const booking = await one(db, 'SELECT * FROM Bookings WHERE BookingID=?', p.bookingId);
    if (!booking) return { success: false, error: 'Booking not found' };
    const items = await all(db, 'SELECT * FROM BookingItems WHERE BookingID=?', p.bookingId);
    const timeline = await all(db,
      'SELECT * FROM BookingTimeline WHERE BookingID=? ORDER BY CreatedAt DESC', p.bookingId);
    return { success: true, booking, items, timeline };
  },

  // Create booking with items. Reserves stock per item atomically.
  createBooking: async (p, db) => {
    const bookingId = uid('BKG-');
    const bookingNumber = await nextBookingNumber(db);
    const date = p.BookingDate || now();
    let items = p.Items || [];
    if (typeof items === 'string') { try { items = JSON.parse(items); } catch { items = []; } }
    if (!Array.isArray(items)) items = [];

    const header = {
      BookingID: bookingId,
      BookingNumber: bookingNumber,
      CustomerID: p.CustomerID || '',
      CustomerName: p.CustomerName || 'Walk-in Customer',
      CustomerPhone: fmtPhone(p.CustomerPhone || ''),
      Status: 'Draft',
      BookingDate: date,
      Priority: p.Priority || 'Normal',
      DepositAmount: num(p.DepositAmount),
      Notes: p.Notes || '',
      InternalNotes: p.InternalNotes || '',
      LinkedSaleID: '',
      LinkedInvoiceNo: '',
      CreatedBy: p.CreatedBy || '',
      CreatedAt: now(),
      UpdatedAt: now()
    };
    await insertRow(db, 'Bookings', header);
    await addTimeline(db, bookingId, 'BOOKING_CREATED', `Booking ${bookingNumber} created`, header.CreatedBy);

    for (const it of items) {
      const requested = num(it.RequestedQty || it.Quantity);
      if (requested <= 0) continue;
      const { reserved, backordered } = await reserveStock(
        db, it.ProductID, requested, bookingNumber, it.ProductName);

      const unitPrice = num(it.UnitPrice);
      const itemRow = {
        BookingItemID: uid('BKI-'),
        BookingID: bookingId,
        ProductID: it.ProductID || '',
        ProductName: it.ProductName || '',
        Club_Country: it.Club_Country || '',
        Edition: it.Edition || '',
        Type: it.Type || '',
        Size: it.Size || '',
        RequestedQty: requested,
        ReservedQty: reserved,
        BackorderedQty: backordered,
        UnitPrice: unitPrice,
        CostPrice: num(it.CostPrice),
        TotalPrice: unitPrice * requested,
        Status: computeItemStatus({ RequestedQty: requested, ReservedQty: reserved }),
        CreatedAt: now(),
        UpdatedAt: now()
      };
      await insertRow(db, 'BookingItems', itemRow);

      if (reserved > 0)
        await addTimeline(db, bookingId, 'BOOK_RESERVED',
          `Reserved ${reserved} × ${itemRow.ProductName}`, header.CreatedBy);
      if (backordered > 0)
        await addTimeline(db, bookingId, 'BOOK_PARTIAL',
          `Backordered ${backordered} × ${itemRow.ProductName}`, header.CreatedBy);
    }

    const status = await refreshBookingStatus(db, bookingId);
    return { success: true, bookingId, bookingNumber, status };
  },

  // Update booking header fields (notes, priority, customer, etc.)
  updateBooking: async (p, db) => {
    const b = await one(db, 'SELECT * FROM Bookings WHERE BookingID=?', p.BookingID);
    if (!b) return { success: false, error: 'Booking not found' };
    if (b.Status === 'Completed' || b.Status === 'Cancelled')
      return { success: false, error: 'Cannot edit a completed/cancelled booking' };
    await updateRow(db, 'Bookings', p.BookingID, {
      CustomerID: p.CustomerID ?? b.CustomerID,
      CustomerName: p.CustomerName ?? b.CustomerName,
      CustomerPhone: fmtPhone(p.CustomerPhone ?? b.CustomerPhone),
      Priority: p.Priority ?? b.Priority,
      DepositAmount: num(p.DepositAmount ?? b.DepositAmount),
      Notes: p.Notes ?? b.Notes,
      InternalNotes: p.InternalNotes ?? b.InternalNotes,
      UpdatedAt: now()
    }, 'BookingID');
    return { success: true };
  },

  // Add a new product to an existing booking (reserves stock)
  addBookingItem: async (p, db) => {
    const b = await one(db, 'SELECT * FROM Bookings WHERE BookingID=?', p.BookingID);
    if (!b) return { success: false, error: 'Booking not found' };
    if (b.Status === 'Completed' || b.Status === 'Cancelled')
      return { success: false, error: 'Cannot edit a completed/cancelled booking' };

    const requested = num(p.RequestedQty);
    if (requested <= 0) return { success: false, error: 'Quantity must be > 0' };

    const { reserved, backordered } = await reserveStock(
      db, p.ProductID, requested, b.BookingNumber, p.ProductName);
    const unitPrice = num(p.UnitPrice);
    const itemRow = {
      BookingItemID: uid('BKI-'),
      BookingID: p.BookingID,
      ProductID: p.ProductID || '',
      ProductName: p.ProductName || '',
      Club_Country: p.Club_Country || '',
      Edition: p.Edition || '',
      Type: p.Type || '',
      Size: p.Size || '',
      RequestedQty: requested,
      ReservedQty: reserved,
      BackorderedQty: backordered,
      UnitPrice: unitPrice,
      CostPrice: num(p.CostPrice),
      TotalPrice: unitPrice * requested,
      Status: computeItemStatus({ RequestedQty: requested, ReservedQty: reserved }),
      CreatedAt: now(),
      UpdatedAt: now()
    };
    await insertRow(db, 'BookingItems', itemRow);
    await addTimeline(db, p.BookingID, 'BOOK_RESERVED',
      `Added ${requested} × ${itemRow.ProductName} (reserved ${reserved}, backordered ${backordered})`);
    const status = await refreshBookingStatus(db, p.BookingID);
    return { success: true, status, itemId: itemRow.BookingItemID };
  },

  // Change quantity of a booking item (reserves more or releases excess)
  updateBookingItemQty: async (p, db) => {
    const item = await one(db, 'SELECT * FROM BookingItems WHERE BookingItemID=?', p.BookingItemID);
    if (!item) return { success: false, error: 'Item not found' };
    const b = await one(db, 'SELECT * FROM Bookings WHERE BookingID=?', item.BookingID);
    if (!b || b.Status === 'Completed' || b.Status === 'Cancelled')
      return { success: false, error: 'Cannot edit this booking' };

    const newQty = num(p.NewQty);
    if (newQty <= 0) return { success: false, error: 'Quantity must be > 0 (use remove instead)' };

    const oldReserved = num(item.ReservedQty);
    const oldRequested = num(item.RequestedQty);

    let newReserved = oldReserved;
    let newBackordered = num(item.BackorderedQty);

    if (newQty > oldRequested) {
      // Increasing — try to reserve the extra
      const extra = newQty - oldRequested;
      const { reserved, backordered } = await reserveStock(
        db, item.ProductID, extra, b.BookingNumber, item.ProductName);
      newReserved = oldReserved + reserved;
      newBackordered = num(item.BackorderedQty) + backordered;
      await addTimeline(db, item.BookingID, 'BOOK_RESERVED',
        `Increased ${item.ProductName} to ${newQty} (reserved +${reserved})`);
    } else if (newQty < oldRequested) {
      // Decreasing — release from backorder first, then reserved
      let reduceBy = oldRequested - newQty;
      const fromBackorder = Math.min(reduceBy, num(item.BackorderedQty));
      newBackordered = num(item.BackorderedQty) - fromBackorder;
      reduceBy -= fromBackorder;
      if (reduceBy > 0) {
        await releaseStock(db, item.ProductID, reduceBy, b.BookingNumber);
        newReserved = oldReserved - reduceBy;
      }
      await addTimeline(db, item.BookingID, 'BOOK_RELEASED',
        `Decreased ${item.ProductName} to ${newQty} (released reserved)`);
    }

    await updateRow(db, 'BookingItems', p.BookingItemID, {
      RequestedQty: newQty,
      ReservedQty: newReserved,
      BackorderedQty: newBackordered,
      TotalPrice: num(item.UnitPrice) * newQty,
      Status: computeItemStatus({ RequestedQty: newQty, ReservedQty: newReserved }),
      UpdatedAt: now()
    }, 'BookingItemID');

    const status = await refreshBookingStatus(db, item.BookingID);
    return { success: true, status };
  },

  // Remove a single item from a booking (releases its reserved stock)
  removeBookingItem: async (p, db) => {
    const item = await one(db, 'SELECT * FROM BookingItems WHERE BookingItemID=?', p.BookingItemID);
    if (!item) return { success: false, error: 'Item not found' };
    const b = await one(db, 'SELECT * FROM Bookings WHERE BookingID=?', item.BookingID);
    if (!b || b.Status === 'Completed' || b.Status === 'Cancelled')
      return { success: false, error: 'Cannot edit this booking' };

    if (num(item.ReservedQty) > 0)
      await releaseStock(db, item.ProductID, num(item.ReservedQty), b.BookingNumber);
    await run(db, 'DELETE FROM BookingItems WHERE BookingItemID=?', p.BookingItemID);
    await addTimeline(db, item.BookingID, 'BOOK_RELEASED',
      `Removed ${item.ProductName} from booking`);
    const status = await refreshBookingStatus(db, item.BookingID);
    return { success: true, status };
  },

  // Cancel a booking — release ALL reserved stock, keep record
  cancelBooking: async (p, db) => {
    const b = await one(db, 'SELECT * FROM Bookings WHERE BookingID=?', p.BookingID);
    if (!b) return { success: false, error: 'Booking not found' };
    if (b.Status === 'Cancelled') return { success: false, error: 'Already cancelled' };
    if (b.Status === 'Completed') return { success: false, error: 'Cannot cancel a completed booking' };

    const items = await all(db, 'SELECT * FROM BookingItems WHERE BookingID=?', p.BookingID);
    for (const it of items) {
      if (num(it.ReservedQty) > 0)
        await releaseStock(db, it.ProductID, num(it.ReservedQty), b.BookingNumber, 'BOOK_CANCELLED');
      await updateRow(db, 'BookingItems', it.BookingItemID, {
        ReservedQty: 0, BackorderedQty: 0, Status: 'Cancelled', UpdatedAt: now()
      }, 'BookingItemID');
    }
    await run(db, 'UPDATE Bookings SET Status=?, UpdatedAt=? WHERE BookingID=?',
      'Cancelled', now(), p.BookingID);
    await addTimeline(db, p.BookingID, 'BOOK_CANCELLED',
      `Booking cancelled. ${p.Reason ? 'Reason: ' + p.Reason : ''}`);
    return { success: true };
  },

  // Convert a booking into a real Sale. Reuses createSale logic by calling it.
  convertBookingToSale: async (p, db) => {
    const b = await one(db, 'SELECT * FROM Bookings WHERE BookingID=?', p.BookingID);
    if (!b) return { success: false, error: 'Booking not found' };
    if (b.Status === 'Completed') return { success: false, error: 'Already converted' };
    if (b.Status === 'Cancelled') return { success: false, error: 'Booking is cancelled' };

    const items = await all(db, 'SELECT * FROM BookingItems WHERE BookingID=?', p.BookingID);
    const reservedItems = items.filter(i => num(i.ReservedQty) > 0);
    if (!reservedItems.length)
      return { success: false, error: 'No reserved items to sell. Restock first.' };

    // Build the sale payload from RESERVED quantities only.
    // We release reserved BEFORE createSale so its normal Stock decrement is correct.
    for (const it of reservedItems) {
      await releaseStock(db, it.ProductID, num(it.ReservedQty), b.BookingNumber, 'BOOK_CONVERTED');
    }

    const saleItems = reservedItems.map(it => ({
      ProductID: it.ProductID,
      ProductName: it.ProductName,
      Club_Country: it.Club_Country,
      Edition: it.Edition,
      Type: it.Type,
      Size: it.Size,
      Quantity: num(it.ReservedQty),
      UnitPrice: num(it.UnitPrice),
      CostPrice: num(it.CostPrice)
    }));

    const salePayload = {
      CustomerID: b.CustomerID,
      CustomerName: b.CustomerName,
      CustomerPhone: b.CustomerPhone,
      Items: saleItems,
      Discount: num(p.Discount),
      AffiliateCode: p.AffiliateCode || '',
      AffiliateDiscount: num(p.AffiliateDiscount),
      VoucherCode: p.VoucherCode || '',
      VoucherDiscount: num(p.VoucherDiscount),
      DeliveryType: p.DeliveryType || 'None',
      DeliveryCharge: num(p.DeliveryCharge),
      DeliveryAddress: p.DeliveryAddress || '',
      PaidAmount: num(p.PaidAmount),
      PaymentMethod: p.PaymentMethod || 'Cash',
      Notes: (p.Notes || b.Notes || '') + ` [From booking ${b.BookingNumber}]`
    };

    // Call the existing createSale handler directly (it decrements physical Stock).
    const saleResult = await ACTIONS.createSale(salePayload, db);
    if (!saleResult.success) {
      // Re-reserve to keep things consistent if sale failed
      for (const it of reservedItems) {
        await reserveStock(db, it.ProductID, num(it.ReservedQty), b.BookingNumber, it.ProductName);
      }
      return { success: false, error: saleResult.error || 'Sale creation failed' };
    }

    // Mark booking completed & link it
    await run(db,
      'UPDATE Bookings SET Status=?, LinkedSaleID=?, LinkedInvoiceNo=?, UpdatedAt=? WHERE BookingID=?',
      'Completed', saleResult.saleId, saleResult.invoiceNo, now(), p.BookingID);

    for (const it of reservedItems) {
      await updateRow(db, 'BookingItems', it.BookingItemID, {
        Status: 'Sold', ReservedQty: 0, UpdatedAt: now()
      }, 'BookingItemID');
    }

    await addTimeline(db, p.BookingID, 'BOOK_CONVERTED',
      `Converted to sale ${saleResult.invoiceNo}`);
    await addTimeline(db, p.BookingID, 'SALE_COMPLETED',
      `Sale completed. Invoice ${saleResult.invoiceNo}`);

    return { success: true, invoiceNo: saleResult.invoiceNo, saleId: saleResult.saleId };
  },

  

  // Auto-allocate newly restocked inventory to waiting bookings (FIFO).
  // Call this after any restock (manual/batch). Pass {productId} to limit scope.
  allocateRestock: async (p, db) => {
    let productIds = [];
    if (p.productId) productIds = [p.productId];
    else {
      // Find all products that have any backordered booking items
      const rows = await all(db,
        `SELECT DISTINCT bi.ProductID FROM BookingItems bi
         INNER JOIN Bookings b ON b.BookingID = bi.BookingID
         WHERE bi.BackorderedQty > 0 AND b.Status NOT IN ('Cancelled','Completed')`);
      productIds = rows.map(r => r.ProductID).filter(Boolean);
    }

    let allocated = 0;
    const touchedBookings = new Set();

    for (const pid of productIds) {
      // Oldest bookings first (FIFO by CreatedAt)
      const waiting = await all(db,
        `SELECT bi.*, b.BookingNumber, b.CreatedAt as bCreated, b.Priority
         FROM BookingItems bi
         INNER JOIN Bookings b ON b.BookingID = bi.BookingID
         WHERE bi.ProductID = ? AND bi.BackorderedQty > 0
           AND b.Status NOT IN ('Cancelled','Completed')
         ORDER BY 
           CASE b.Priority WHEN 'Urgent' THEN 1 WHEN 'VIP' THEN 2 WHEN 'High' THEN 3 ELSE 4 END,
           b.CreatedAt ASC`, pid);

      for (const w of waiting) {
        const prod = await one(db, 'SELECT * FROM Inventory WHERE ID=?', pid);
        if (!prod) break;
        const avail = num(prod.Stock) - num(prod.ReservedQty);
        if (avail <= 0) break;

        const need = num(w.BackorderedQty);
        const give = Math.min(need, avail);
        if (give <= 0) continue;

        const { reserved } = await reserveStock(db, pid, give, w.BookingNumber, w.ProductName);
        if (reserved > 0) {
          await updateRow(db, 'BookingItems', w.BookingItemID, {
            ReservedQty: num(w.ReservedQty) + reserved,
            BackorderedQty: need - reserved,
            Status: computeItemStatus({
              RequestedQty: num(w.RequestedQty),
              ReservedQty: num(w.ReservedQty) + reserved
            }),
            UpdatedAt: now()
          }, 'BookingItemID');
          await auditStock(db, {
            ProductID: pid, ProductName: w.ProductName, Action: 'RESTOCK_ALLOCATED',
            ReferenceID: w.BookingNumber, QtyChange: reserved
          });
          await addTimeline(db, w.BookingID, 'RESTOCK_ALLOCATED',
            `Restocked & reserved ${reserved} × ${w.ProductName}`);
          allocated += reserved;
          touchedBookings.add(w.BookingID);
        }
      }
    }

    for (const bid of touchedBookings) await refreshBookingStatus(db, bid);
    return { success: true, allocated, bookingsUpdated: touchedBookings.size };
  },

  getStockAudit: async (p, db) => p.productId
    ? await all(db, 'SELECT * FROM StockAudit WHERE ProductID=? ORDER BY Timestamp DESC LIMIT 1000', p.productId)
    : await all(db, 'SELECT * FROM StockAudit ORDER BY Timestamp DESC LIMIT 1000'),

  getBookingStats: async (p, db) => {
    const active = await one(db, "SELECT COUNT(*) c FROM Bookings WHERE Status='Active'");
    const partial = await one(db, "SELECT COUNT(*) c FROM Bookings WHERE Status='Partial'");
    const restock = await one(db, "SELECT COUNT(*) c FROM Bookings WHERE Status='Need Restocking'");
    const cancelled = await one(db, "SELECT COUNT(*) c FROM Bookings WHERE Status='Cancelled'");
    const todayStr = today();
    const convertedToday = await one(db,
      "SELECT COUNT(*) c FROM Bookings WHERE Status='Completed' AND UpdatedAt >= ?", todayStr);
    const reservedValue = await one(db,
      `SELECT COALESCE(SUM(bi.ReservedQty * bi.UnitPrice),0) v
       FROM BookingItems bi INNER JOIN Bookings b ON b.BookingID=bi.BookingID
       WHERE b.Status NOT IN ('Cancelled','Completed')`);
    return {
      activeBookings: num(active.c),
      partialBookings: num(partial.c),
      needRestocking: num(restock.c),
      cancelledBookings: num(cancelled.c),
      convertedToday: num(convertedToday.c),
      reservedStockValue: num(reservedValue.v)
    };
  },
  // ONE-TIME MIGRATION: consolidate all variant suffixes (PRD-0043A/B/C…)
  // back into their base product (PRD-0043). Moves stock, creates batch
  // layers, deletes variants. Safe-guards reserved variants.
  migrateConsolidateVariants: async (p, db) => {
    const allProducts = await all(db, 'SELECT * FROM Inventory ORDER BY ID');
    const byId = {};
    allProducts.forEach(pr => byId[pr.ID] = pr);

    // A "variant" = base ID + trailing letters, where the base ID exists.
    const results = [];
    let basesProcessed = 0, variantsMerged = 0, skippedReserved = 0;

    // Group variants by their base
    const groups = {}; // baseId -> [variants]
    for (const pr of allProducts) {
      const m = String(pr.ID).match(/^(.*?)([A-Z]+)$/);
      if (!m) continue;
      const baseId = m[1].replace(/[-]?$/, m[1].endsWith('-') ? '-' : ''); // keep as-is
      // Only treat as variant if the trailing part is letters AND base exists
      const candidateBase = m[1];
      if (byId[candidateBase] && /^[A-Z]+$/.test(m[2])) {
        if (!groups[candidateBase]) groups[candidateBase] = [];
        groups[candidateBase].push(pr);
      }
    }

    for (const [baseId, variants] of Object.entries(groups)) {
      const base = byId[baseId];
      if (!base) continue;

      let totalUnits = num(base.Stock);
      let weightedCostSum = num(base.Stock) * num(base.CostPrice);
      let moved = 0;

      // Seed a batch for the base's existing stock (so FIFO has a layer for it)
      if (num(base.Stock) > 0) {
        const existingBatches = await all(db, 'SELECT 1 FROM InventoryBatches WHERE ProductID=? LIMIT 1', baseId);
        if (!existingBatches.length) {
          await insertRow(db, 'InventoryBatches', {
            ID: uid('BT-'), ProductID: baseId, ProductName: base.ProductName,
            SupplierID: base.SupplierID || '', SupplierName: base.SupplierName || '',
            BatchDate: base.DateAdded || now(), Quantity: num(base.Stock),
            CostPrice: num(base.CostPrice), RemainingQty: num(base.Stock),
            Notes: 'Seeded from base stock during consolidation'
          });
        }
      }

      for (const v of variants) {
        // Don't merge if variant has reserved stock (active bookings)
        if (num(v.ReservedQty) > 0) { skippedReserved++; continue; }

        const vStock = num(v.Stock);
        if (vStock > 0) {
          await insertRow(db, 'InventoryBatches', {
            ID: uid('BT-'), ProductID: baseId, ProductName: base.ProductName,
            SupplierID: v.SupplierID || '', SupplierName: v.SupplierName || '',
            BatchDate: v.DateAdded || now(), Quantity: vStock,
            CostPrice: num(v.CostPrice), RemainingQty: vStock,
            Notes: `Consolidated from variant ${v.ID}`
          });
          totalUnits += vStock;
          weightedCostSum += vStock * num(v.CostPrice);
          moved += vStock;
        }

        // Re-point any past SalesItems from the variant to the base (keeps history linked)
        await run(db, 'UPDATE SalesItems SET ProductID=? WHERE ProductID=?', baseId, v.ID);

        // Delete the variant product
        await run(db, 'DELETE FROM Inventory WHERE ID=?', v.ID);
        variantsMerged++;
      }

      const newStock = num(base.Stock) + moved;
      const newCost = totalUnits > 0 ? weightedCostSum / totalUnits : num(base.CostPrice);
      await run(db, 'UPDATE Inventory SET Stock=?, CostPrice=?, LastUpdated=? WHERE ID=?',
        newStock, Math.round(newCost * 100) / 100, now(), baseId);

      await insertRow(db, 'StockLog', {
        ID: await seqId(db, 'StockLog', 'ID', 'SLOG-'), Date: now(),
        ProductID: baseId, ProductName: base.ProductName,
        Action: 'Variants Consolidated', Quantity: moved,
        OldStock: num(base.Stock), NewStock: newStock,
        Reference: 'Migration', Notes: `Merged ${variants.length} variant(s)`
      });

      results.push({ baseId, newStock, newAvgCost: Math.round(newCost*100)/100, mergedUnits: moved });
      basesProcessed++;
    }

    return { success: true, basesProcessed, variantsMerged, skippedReserved, results };
  },
  
};

// ============================================================
  // SYNC TO GOOGLE SHEETS (runs every 5 min via cron)
  // ============================================================
  async function syncToGoogleSheets(env) {
    const SYNC_URL = env.SHEETS_SYNC_URL;
    const SYNC_KEY = env.SHEETS_SYNC_KEY;
    
    if (!SYNC_URL) {
      console.log('Sync skipped: SHEETS_SYNC_URL not configured');
      return;
    }
    
    const db = env.DB;
    const startTime = Date.now();
    
    try {
      // Fetch all tables (use last-modified watermark for incremental sync)
      const lastSync = await getSetting(db, 'lastSheetSyncAt', '2000-01-01T00:00:00.000Z');
      
      const payload = {
        syncedAt: new Date().toISOString(),
        lastSyncAt: lastSync,
        key: SYNC_KEY,
        tables: {
          Inventory: await all(db, 'SELECT * FROM Inventory'),
          Customers: await all(db, 'SELECT * FROM Customers'),
          Suppliers: await all(db, 'SELECT * FROM Suppliers'),
          Sales: await all(db, 'SELECT * FROM Sales WHERE Date >= ?', lastSync),
          SalesItems: await all(db, `
            SELECT si.* FROM SalesItems si 
            INNER JOIN Sales s ON s.InvoiceNo = si.InvoiceNo 
            WHERE s.Date >= ?
          `, lastSync),
          Payments: await all(db, 'SELECT * FROM Payments WHERE Date >= ?', lastSync),
          Expenses: await all(db, 'SELECT * FROM Expenses WHERE Date >= ?', lastSync),
          StockLog: await all(db, 'SELECT * FROM StockLog WHERE Date >= ?', lastSync),
          Affiliates: await all(db, 'SELECT * FROM Affiliates'),
          AffiliateCommissions: await all(db, 'SELECT * FROM AffiliateCommissions WHERE Date >= ?', lastSync),
          Vouchers: await all(db, 'SELECT * FROM Vouchers'),
          Deliveries: await all(db, 'SELECT * FROM Deliveries WHERE Date >= ?', lastSync),
          SupplierLedger: await all(db, 'SELECT * FROM SupplierLedger WHERE Date >= ?', lastSync),
          InventoryBatches: await all(db, 'SELECT * FROM InventoryBatches WHERE BatchDate >= ?', lastSync),
          Settings: await all(db, 'SELECT * FROM Settings'),
          DeadStockHistory: await all(db, 'SELECT * FROM DeadStockHistory'),
          Bookings: await all(db, 'SELECT * FROM Bookings'),
          BookingItems: await all(db, 'SELECT * FROM BookingItems'),
          BookingTimeline: await all(db, 'SELECT * FROM BookingTimeline'),
          StockAudit: await all(db, 'SELECT * FROM StockAudit WHERE Timestamp >= ?', lastSync),
        }
      };
      
      const totalRows = Object.values(payload.tables).reduce((s, t) => s + t.length, 0);
      console.log(`Syncing ${totalRows} rows to Sheets...`);
      
      const res = await fetch(SYNC_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      
      if (!res.ok) {
        throw new Error(`Sheets sync failed: ${res.status} ${await res.text()}`);
      }
      
      const result = await res.json();
      
      // Save watermark only if sync succeeded
      await setSetting(db, 'lastSheetSyncAt', payload.syncedAt);
      await setSetting(db, 'lastSheetSyncStatus', `OK: ${totalRows} rows in ${Date.now() - startTime}ms`);
      
      console.log('Sync complete:', result);
    } catch (err) {
      console.error('Sync error:', err);
      await setSetting(db, 'lastSheetSyncStatus', `ERROR: ${err.message}`);
    }
  }