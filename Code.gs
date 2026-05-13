// ============================================================
// Code.gs — Google Apps Script Backend for LumpiaPOS
// ============================================================

const BRANCHES = {
  'Lumpia Pusat': { password: 'lumpia123', sheetId: 'YOUR_SPREADSHEET_ID_HERE' },
  'Lumpia Cabang 1': { password: 'cabang1', sheetId: 'YOUR_SPREADSHEET_ID_2' },
};

const SHEET_SCHEMAS = {
  Menu:        ['id','name','category','icon','image','price_cash','price_dana','price_gofood','price_shopee','price_grab','active'],
  Ingredients: ['id','name','unit','current_stock','min_stock','pack_size','pack_price','cost_per_unit'],
  Recipes:     ['menu_id','ingredient_id','ingredient_name','amount','unit'],
  Orders:      ['id','date','time','payment','total','branch','status','note','cancel_reason','cancelled_at'],
  OrderItems:  ['order_id','menu_id','menu_name','qty','price','payment_type'],
  Ledger:      ['id','date','time','type','category','description','amount','payment','order_id'],
  StockLog:    ['id','date','ingredient_id','ingredient_name','change','reason','order_id'],
  Settings:    ['key','value'],
};

function doGet(e) {
  try {
    const raw = e.parameter.payload;
    if (!raw) return jsonResp({ ok: false, error: 'No payload' });
    const data = JSON.parse(decodeURIComponent(raw));
    const result = route(data);
    return jsonResp(result);
  } catch (err) {
    return jsonResp({ ok: false, error: err.message });
  }
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function route(data) {
  const { action, branch, password } = data;
  if (action === 'login') return handleLogin(branch, password);
  const ss = authenticate(branch, password);
  if (!ss) return { ok: false, error: 'Auth failed' };

  switch (action) {
    case 'syncAll':        return handleSyncAll(ss, branch);
    case 'placeOrder':     return handlePlaceOrder(ss, data, branch);
    case 'cancelOrder':    return handleCancelOrder(ss, data);
    case 'saveMenu':       return handleSaveMenu(ss, data);
    case 'deleteMenu':     return handleDeleteMenu(ss, data);
    case 'saveIngredient':  return handleSaveIngredient(ss, data);
    case 'deleteIngredient':return handleDeleteIngredient(ss, data);
    case 'saveRecipe':     return handleSaveRecipe(ss, data);
    case 'restockIngredient': return handleRestock(ss, data);
    case 'addLedgerEntry': return handleAddLedger(ss, data);
    case 'deleteLedgerEntry': return handleDeleteLedger(ss, data);
    case 'saveSettings':   return handleSaveSettings(ss, data);
    case 'pushData':       return handlePushData(ss, data, branch);
    default: return { ok: false, error: 'Unknown action: ' + action };
  }
}

function authenticate(branch, password) {
  const b = BRANCHES[branch];
  if (!b || b.password !== password) return null;
  const ss = SpreadsheetApp.openById(b.sheetId);
  ensureSheets(ss);
  return ss;
}

function handleLogin(branch, password) {
  const b = BRANCHES[branch];
  if (!b || b.password !== password) return { ok: false, error: 'Invalid branch or password' };
  try {
    const ss = SpreadsheetApp.openById(b.sheetId);
    ensureSheets(ss);
    return { ok: true, branch: branch };
  } catch (err) {
    return { ok: false, error: 'Cannot open spreadsheet: ' + err.message };
  }
}

// ---- Sheet Management ----

function ensureSheets(ss) {
  for (const [name, cols] of Object.entries(SHEET_SCHEMAS)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
      sheet.setFrozenRows(1);
    } else {
      const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
      for (let i = 0; i < cols.length; i++) {
        if (existing.indexOf(cols[i]) === -1) {
          const insertAt = Math.min(i + 1, sheet.getLastColumn() + 1);
          if (insertAt <= sheet.getLastColumn()) {
            sheet.insertColumnBefore(insertAt);
          }
          sheet.getRange(1, insertAt).setValue(cols[i]);
        }
      }
    }
  }
}

function sheetToObjects(sheet) {
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const tz = Session.getScriptTimeZone();
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
      }
      obj[h] = val;
    });
    return obj;
  });
}

function appendRow(sheet, schema, obj) {
  const row = schema.map(col => obj[col] !== undefined ? obj[col] : '');
  sheet.appendRow(row);
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  if (idCol === -1) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return i + 1;
  }
  return -1;
}

function generateId(prefix) {
  return prefix + '-' + Date.now() + '-' + Math.random().toString(36).substr(2, 5);
}

// ---- Sync ----

function handleSyncAll(ss, branch) {
  const menu = sheetToObjects(ss.getSheetByName('Menu'));
  const ingredients = sheetToObjects(ss.getSheetByName('Ingredients'));
  const recipes = sheetToObjects(ss.getSheetByName('Recipes'));
  const orders = sheetToObjects(ss.getSheetByName('Orders'));
  const orderItems = sheetToObjects(ss.getSheetByName('OrderItems'));
  const ledger = sheetToObjects(ss.getSheetByName('Ledger'));
  const stockLog = sheetToObjects(ss.getSheetByName('StockLog'));

  const settingsSheet = ss.getSheetByName('Settings');
  const settingsRows = sheetToObjects(settingsSheet);
  const settings = {};
  settingsRows.forEach(r => { settings[r.key] = r.value; });

  return {
    ok: true,
    data: { menu, ingredients, recipes, orders, orderItems, ledger, stockLog, settings }
  };
}

// ---- Orders ----

function handlePlaceOrder(ss, data, branch) {
  const orderId = data.orderId || generateId('ORD');
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const dateStr = data.date || Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const timeStr = data.time || Utilities.formatDate(now, tz, 'HH:mm:ss');

  const orderSheet = ss.getSheetByName('Orders');
  appendRow(orderSheet, SHEET_SCHEMAS.Orders, {
    id: orderId, date: dateStr, time: timeStr,
    payment: data.payment, total: data.total,
    branch: branch, status: 'Completed', note: data.note || '',
    cancel_reason: '', cancelled_at: ''
  });

  const itemsSheet = ss.getSheetByName('OrderItems');
  (data.items || []).forEach(item => {
    appendRow(itemsSheet, SHEET_SCHEMAS.OrderItems, {
      order_id: orderId, menu_id: item.menu_id,
      menu_name: item.menu_name, qty: item.qty,
      price: item.price, payment_type: data.payment
    });
  });

  const ledgerSheet = ss.getSheetByName('Ledger');
  appendRow(ledgerSheet, SHEET_SCHEMAS.Ledger, {
    id: generateId('LED'), date: dateStr, time: timeStr,
    type: 'Income', category: 'Sales',
    description: 'Order ' + orderId,
    amount: data.total, payment: data.payment, order_id: orderId
  });

  deductStock(ss, data.items, orderId, dateStr);

  return { ok: true, orderId: orderId };
}

function deductStock(ss, items, orderId, dateStr) {
  const recipeSheet = ss.getSheetByName('Recipes');
  const recipes = sheetToObjects(recipeSheet);
  const ingSheet = ss.getSheetByName('Ingredients');
  const ingredients = sheetToObjects(ingSheet);
  const stockLogSheet = ss.getSheetByName('StockLog');
  const headers = ingSheet.getRange(1, 1, 1, ingSheet.getLastColumn()).getValues()[0];
  const stockCol = headers.indexOf('current_stock') + 1;

  (items || []).forEach(item => {
    const menuRecipes = recipes.filter(r => String(r.menu_id) === String(item.menu_id));
    menuRecipes.forEach(rec => {
      const ingRow = findIngredientRow(ingSheet, rec.ingredient_id);
      if (ingRow === -1) return;
      const totalDeduct = Number(rec.amount) * Number(item.qty);
      const currentStock = Number(ingSheet.getRange(ingRow, stockCol).getValue());
      ingSheet.getRange(ingRow, stockCol).setValue(currentStock - totalDeduct);

      appendRow(stockLogSheet, SHEET_SCHEMAS.StockLog, {
        id: generateId('STK'), date: dateStr,
        ingredient_id: rec.ingredient_id, ingredient_name: rec.ingredient_name,
        change: -totalDeduct, reason: 'Order ' + orderId, order_id: orderId
      });
    });
  });
}

function findIngredientRow(sheet, ingId) {
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(ingId)) return i + 1;
  }
  return -1;
}

function handleCancelOrder(ss, data) {
  const orderSheet = ss.getSheetByName('Orders');
  const row = findRowById(orderSheet, data.orderId);
  if (row === -1) return { ok: false, error: 'Order not found' };

  const headers = orderSheet.getRange(1, 1, 1, orderSheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('status') + 1;
  const reasonCol = headers.indexOf('cancel_reason') + 1;
  const cancelledCol = headers.indexOf('cancelled_at') + 1;

  const currentStatus = orderSheet.getRange(row, statusCol).getValue();
  if (currentStatus === 'Cancelled') return { ok: false, error: 'Already cancelled' };

  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  orderSheet.getRange(row, statusCol).setValue('Cancelled');
  if (reasonCol > 0) orderSheet.getRange(row, reasonCol).setValue(data.reason || '');
  if (cancelledCol > 0) orderSheet.getRange(row, cancelledCol).setValue(now);

  const orderData = sheetToObjects(orderSheet).find(o => String(o.id) === String(data.orderId));
  if (!orderData) return { ok: false, error: 'Order data not found' };

  const ledgerSheet = ss.getSheetByName('Ledger');
  appendRow(ledgerSheet, SHEET_SCHEMAS.Ledger, {
    id: generateId('LED'), date: orderData.date, time: orderData.time,
    type: 'Income', category: 'Sales Reversal',
    description: 'Cancel ' + data.orderId,
    amount: -Number(orderData.total), payment: orderData.payment,
    order_id: data.orderId
  });

  const itemsSheet = ss.getSheetByName('OrderItems');
  const allItems = sheetToObjects(itemsSheet);
  const orderItems = allItems.filter(i => String(i.order_id) === String(data.orderId));

  restoreStock(ss, orderItems, data.orderId, orderData.date);

  return { ok: true };
}

function restoreStock(ss, items, orderId, dateStr) {
  const recipeSheet = ss.getSheetByName('Recipes');
  const recipes = sheetToObjects(recipeSheet);
  const ingSheet = ss.getSheetByName('Ingredients');
  const stockLogSheet = ss.getSheetByName('StockLog');
  const headers = ingSheet.getRange(1, 1, 1, ingSheet.getLastColumn()).getValues()[0];
  const stockCol = headers.indexOf('current_stock') + 1;

  items.forEach(item => {
    const menuRecipes = recipes.filter(r => String(r.menu_id) === String(item.menu_id));
    menuRecipes.forEach(rec => {
      const ingRow = findIngredientRow(ingSheet, rec.ingredient_id);
      if (ingRow === -1) return;
      const totalRestore = Number(rec.amount) * Number(item.qty);
      const currentStock = Number(ingSheet.getRange(ingRow, stockCol).getValue());
      ingSheet.getRange(ingRow, stockCol).setValue(currentStock + totalRestore);

      appendRow(stockLogSheet, SHEET_SCHEMAS.StockLog, {
        id: generateId('STK'), date: dateStr,
        ingredient_id: rec.ingredient_id, ingredient_name: rec.ingredient_name,
        change: totalRestore, reason: 'Cancel ' + orderId, order_id: orderId
      });
    });
  });
}

// ---- Menu ----

function handleSaveMenu(ss, data) {
  const sheet = ss.getSheetByName('Menu');
  const item = data.item;
  if (!item.id) item.id = generateId('MNU');
  const row = findRowById(sheet, item.id);
  if (row > 0) {
    const vals = SHEET_SCHEMAS.Menu.map(col => item[col] !== undefined ? item[col] : '');
    sheet.getRange(row, 1, 1, vals.length).setValues([vals]);
  } else {
    appendRow(sheet, SHEET_SCHEMAS.Menu, item);
  }
  return { ok: true, id: item.id };
}

function handleDeleteMenu(ss, data) {
  const sheet = ss.getSheetByName('Menu');
  const row = findRowById(sheet, data.menuId);
  if (row > 0) sheet.deleteRow(row);

  const recipeSheet = ss.getSheetByName('Recipes');
  const recipes = sheetToObjects(recipeSheet);
  for (let i = recipes.length - 1; i >= 0; i--) {
    if (String(recipes[i].menu_id) === String(data.menuId)) {
      recipeSheet.deleteRow(i + 2);
    }
  }
  return { ok: true };
}

// ---- Ingredients ----

function handleSaveIngredient(ss, data) {
  const sheet = ss.getSheetByName('Ingredients');
  const item = data.item;
  if (!item.id) item.id = generateId('ING');
  if (item.pack_size && item.pack_price) {
    item.cost_per_unit = Number(item.pack_price) / Number(item.pack_size);
  }
  const row = findRowById(sheet, item.id);
  if (row > 0) {
    const vals = SHEET_SCHEMAS.Ingredients.map(col => item[col] !== undefined ? item[col] : '');
    sheet.getRange(row, 1, 1, vals.length).setValues([vals]);
  } else {
    appendRow(sheet, SHEET_SCHEMAS.Ingredients, item);
  }
  return { ok: true, id: item.id };
}

function handleDeleteIngredient(ss, data) {
  const sheet = ss.getSheetByName('Ingredients');
  const row = findRowById(sheet, data.ingredientId);
  if (row > 0) sheet.deleteRow(row);
  return { ok: true };
}

function handleRestock(ss, data) {
  const sheet = ss.getSheetByName('Ingredients');
  const row = findRowById(sheet, data.ingredientId);
  if (row === -1) return { ok: false, error: 'Ingredient not found' };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const stockCol = headers.indexOf('current_stock') + 1;
  const current = Number(sheet.getRange(row, stockCol).getValue());
  sheet.getRange(row, stockCol).setValue(current + Number(data.amount));

  const stockLogSheet = ss.getSheetByName('StockLog');
  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  appendRow(stockLogSheet, SHEET_SCHEMAS.StockLog, {
    id: generateId('STK'), date: dateStr,
    ingredient_id: data.ingredientId, ingredient_name: data.ingredientName || '',
    change: Number(data.amount), reason: 'Restock', order_id: ''
  });

  if (data.cost && Number(data.cost) > 0) {
    const ledgerSheet = ss.getSheetByName('Ledger');
    const timeStr = Utilities.formatDate(new Date(), tz, 'HH:mm:ss');
    appendRow(ledgerSheet, SHEET_SCHEMAS.Ledger, {
      id: generateId('LED'), date: dateStr, time: timeStr,
      type: 'Expense', category: 'Restock',
      description: 'Restock ' + (data.ingredientName || data.ingredientId),
      amount: -Math.abs(Number(data.cost)),
      payment: data.payment || 'Cash', order_id: ''
    });
  }

  return { ok: true };
}

// ---- Recipes ----

function handleSaveRecipe(ss, data) {
  const sheet = ss.getSheetByName('Recipes');
  const existing = sheetToObjects(sheet);
  for (let i = existing.length - 1; i >= 0; i--) {
    if (String(existing[i].menu_id) === String(data.menuId)) {
      sheet.deleteRow(i + 2);
    }
  }
  (data.recipes || []).forEach(r => {
    appendRow(sheet, SHEET_SCHEMAS.Recipes, {
      menu_id: data.menuId,
      ingredient_id: r.ingredient_id,
      ingredient_name: r.ingredient_name,
      amount: r.amount,
      unit: r.unit
    });
  });
  return { ok: true };
}

// ---- Ledger ----

function handleAddLedger(ss, data) {
  const sheet = ss.getSheetByName('Ledger');
  const entry = data.entry;
  if (!entry.id) entry.id = generateId('LED');
  appendRow(sheet, SHEET_SCHEMAS.Ledger, entry);
  return { ok: true, id: entry.id };
}

function handleDeleteLedger(ss, data) {
  const sheet = ss.getSheetByName('Ledger');
  const row = findRowById(sheet, data.ledgerId);
  if (row > 0) sheet.deleteRow(row);
  return { ok: true };
}

// ---- Settings ----

function handleSaveSettings(ss, data) {
  const sheet = ss.getSheetByName('Settings');
  const rows = sheet.getDataRange().getValues();
  const entries = data.settings || {};

  for (const [key, value] of Object.entries(entries)) {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, value]);
      rows.push([key, value]);
    }
  }
  return { ok: true };
}

// ---- Push Data (bulk upload from local) ----

function handlePushData(ss, data, branch) {
  let counts = { ingredients: 0, menu: 0, recipes: 0, orders: 0, orderItems: 0, ledger: 0 };

  if (data.ingredients) {
    const sheet = ss.getSheetByName('Ingredients');
    const existing = sheetToObjects(sheet).map(r => String(r.id));
    data.ingredients.forEach(item => {
      if (!existing.includes(String(item.id))) {
        appendRow(sheet, SHEET_SCHEMAS.Ingredients, item);
        counts.ingredients++;
      }
    });
  }

  if (data.menu) {
    const sheet = ss.getSheetByName('Menu');
    const existing = sheetToObjects(sheet).map(r => String(r.id));
    data.menu.forEach(item => {
      if (!existing.includes(String(item.id))) {
        appendRow(sheet, SHEET_SCHEMAS.Menu, item);
        counts.menu++;
      }
    });
  }

  if (data.recipes) {
    const sheet = ss.getSheetByName('Recipes');
    data.recipes.forEach(r => {
      appendRow(sheet, SHEET_SCHEMAS.Recipes, r);
      counts.recipes++;
    });
  }

  if (data.orders) {
    const sheet = ss.getSheetByName('Orders');
    const existing = sheetToObjects(sheet).map(r => String(r.id));
    data.orders.forEach(item => {
      if (!existing.includes(String(item.id))) {
        item.branch = item.branch || branch;
        appendRow(sheet, SHEET_SCHEMAS.Orders, item);
        counts.orders++;
      }
    });
  }

  if (data.orderItems) {
    const sheet = ss.getSheetByName('OrderItems');
    data.orderItems.forEach(item => {
      appendRow(sheet, SHEET_SCHEMAS.OrderItems, item);
      counts.orderItems++;
    });
  }

  if (data.ledger) {
    const sheet = ss.getSheetByName('Ledger');
    const existing = sheetToObjects(sheet).map(r => String(r.id));
    data.ledger.forEach(item => {
      if (!existing.includes(String(item.id))) {
        appendRow(sheet, SHEET_SCHEMAS.Ledger, item);
        counts.ledger++;
      }
    });
  }

  return { ok: true, counts: counts };
}
