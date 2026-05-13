// ================================================
// SUKIDO JASTIP — Google Apps Script Backend v6
// ================================================

const SHEET_NAME_ORDERS    = "Orders";
const SHEET_NAME_KLOTERS   = "Kloters";
const SHEET_NAME_ADDRESSES = "Addresses";
const SHEET_NAME_KIRIM     = "KirimStatus";
const SHEET_NAME_KATALOG   = "Katalog";

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    let result;
    switch (action) {
      case "getAll":           result = getAllData();                              break;
      case "saveOrder":        result = saveOrder(JSON.parse(params.data));       break;
      case "updateOrder":      result = updateOrder(JSON.parse(params.data));     break;
      case "deleteOrder":      result = deleteOrder(params.id);                   break;
      case "saveKloter":       result = saveKloter(JSON.parse(params.data));      break;
      case "updateKloter":     result = updateKloter(JSON.parse(params.data));    break;
      case "deleteKloter":     result = deleteKloter(params.id);                  break;
      case "saveAddress":      result = saveAddress(JSON.parse(params.data));     break;
      case "getAddresses":     result = getAddresses();                           break;
      case "saveKirimStatus":  result = saveKirimStatus(JSON.parse(params.data)); break;
      case "getKirimStatus":   result = getKirimStatus();                         break;
      case "saveKatalog":      result = saveKatalogItem(JSON.parse(params.data)); break;
      case "updateKatalog":    result = updateKatalogItem(JSON.parse(params.data)); break;
      case "deleteKatalog":    result = deleteKatalogItem(params.id);             break;
      case "getKatalog":       result = getKatalog();                             break;
      case "getPin":           result = getPin();                                 break;
      case "setPin":           result = setPin(params.hash);                      break;
      default:
        result = { ping: "ok", message: "Sukido Jastip API v5 running!" };
    }
    output.setContent(JSON.stringify({ ok: true, data: result }));
  } catch (err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
  }
  return output;
}

// ── SHEETS SETUP ──────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_NAME_ORDERS)) {
    const s = ss.insertSheet(SHEET_NAME_ORDERS);
    s.appendRow(["id","date","cust","stype","pic","ship","item","qty","jpy","rate","jual","admin","dp","sbayar","ostatus","sudahBeli","batch","sudahDipesan","sudahReady","ongkirPayer"]);
    s.getRange(1,1,1,20).setFontWeight("bold").setBackground("#e91e8c").setFontColor("#ffffff");
  }

  if (!ss.getSheetByName(SHEET_NAME_KLOTERS)) {
    const s = ss.insertSheet(SHEET_NAME_KLOTERS);
    s.appendRow(["id","name","rate","bagasi","eta","status","pengeluaran"]);
    s.getRange(1,1,1,7).setFontWeight("bold").setBackground("#7c3aed").setFontColor("#ffffff");
  }

  if (!ss.getSheetByName(SHEET_NAME_ADDRESSES)) {
    const s = ss.insertSheet(SHEET_NAME_ADDRESSES);
    s.appendRow(["custName","nama","hp","alamat","kota","provinsi","kodepos","catatan","ship","timestamp"]);
    s.getRange(1,1,1,10).setFontWeight("bold").setBackground("#0891b2").setFontColor("#ffffff");
  }

  if (!ss.getSheetByName(SHEET_NAME_KIRIM)) {
    const s = ss.insertSheet(SHEET_NAME_KIRIM);
    s.appendRow(["custName","packed","sent","updatedAt"]);
    s.getRange(1,1,1,4).setFontWeight("bold").setBackground("#16a34a").setFontColor("#ffffff");
  }

  if (!ss.getSheetByName(SHEET_NAME_KATALOG)) {
    const s = ss.insertSheet(SHEET_NAME_KATALOG);
    s.appendRow(["id","nama","jpy","jual","updatedAt"]);
    s.getRange(1,1,1,5).setFontWeight("bold").setBackground("#f59e0b").setFontColor("#ffffff");
  }
}

// ── ORDERS ────────────────────────────────────────────────
function getAllData() {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const oData = ss.getSheetByName(SHEET_NAME_ORDERS).getDataRange().getValues();
  const oHeaders = oData[0];
  const orders = oData.slice(1).map(row => {
    const obj = {};
    oHeaders.forEach((h,i) => { obj[h] = row[i]; });
    obj.id        = String(obj.id);
    obj.qty       = Number(obj.qty)   || 1;
    obj.jpy       = Number(obj.jpy)   || 0;
    obj.rate      = Number(obj.rate)  || 0;
    obj.jual      = Number(obj.jual)  || 0;
    obj.admin     = Number(obj.admin) || 0;
    obj.dp        = Number(obj.dp)    || 0;
    obj.sudahBeli     = (obj.sudahBeli === true     || String(obj.sudahBeli).toLowerCase()     === "true");
    obj.sudahDipesan  = (obj.sudahDipesan === true  || String(obj.sudahDipesan).toLowerCase()  === "true");
    obj.sudahReady    = (obj.sudahReady === true     || String(obj.sudahReady).toLowerCase()    === "true");
    obj.ongkirPayer   = obj.ongkirPayer || "seller";
    return obj;
  }).filter(o => o.id && o.id !== 'undefined' && o.id !== '');

  const kData = ss.getSheetByName(SHEET_NAME_KLOTERS).getDataRange().getValues();
  const kHeaders = kData[0];
  const kloters = kData.slice(1).map(row => {
    const obj = {};
    kHeaders.forEach((h,i) => { obj[h] = row[i]; });
    obj.id   = String(obj.id);
    obj.rate = Number(obj.rate) || 0;
    try { obj.pengeluaran = obj.pengeluaran ? JSON.parse(obj.pengeluaran) : []; }
    catch(e) { obj.pengeluaran = []; }
    return obj;
  }).filter(k => k.id && k.id !== 'undefined' && k.id !== '');

  const addresses   = getAddresses();
  const kirimStatus = getKirimStatus();
  const katalog     = getKatalog();

  return { orders, kloters, addresses, kirimStatus, katalog };
}

function saveOrder(order) {
  setupSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_ORDERS);
  sheet.appendRow([
    String(order.id), order.date, order.cust, order.stype, order.pic,
    order.ship, order.item, order.qty, order.jpy, order.rate,
    order.jual, order.admin||0, order.dp||0, order.sbayar,
    order.ostatus, order.sudahBeli||false, order.batch,
    order.sudahDipesan||false, order.sudahReady||false, order.ongkirPayer||"seller"
  ]);
  return { saved: order.id };
}

function updateOrder(order) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_ORDERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(order.id)) {
      sheet.getRange(i+1,1,1,20).setValues([[
        String(order.id), order.date, order.cust, order.stype, order.pic,
        order.ship, order.item, order.qty, order.jpy, order.rate,
        order.jual, order.admin||0, order.dp||0, order.sbayar,
        order.ostatus, order.sudahBeli||false, order.batch,
        order.sudahDipesan||false, order.sudahReady||false, order.ongkirPayer||"seller"
      ]]);
      return { updated: order.id };
    }
  }
  return saveOrder(order);
}

function deleteOrder(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_ORDERS);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length-1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { deleted: id }; }
  }
  return { notFound: id };
}

// ── KLOTERS ───────────────────────────────────────────────
function saveKloter(kloter) {
  setupSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_KLOTERS);
  sheet.appendRow([
    String(kloter.id), kloter.name, kloter.rate||0,
    kloter.bagasi||"", kloter.eta||"", kloter.status||"open",
    JSON.stringify(kloter.pengeluaran||[])
  ]);
  return { saved: kloter.id };
}

function updateKloter(kloter) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_KLOTERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(kloter.id)) {
      sheet.getRange(i+1,1,1,7).setValues([[
        String(kloter.id), kloter.name, kloter.rate||0,
        kloter.bagasi||"", kloter.eta||"", kloter.status||"open",
        JSON.stringify(kloter.pengeluaran||[])
      ]]);
      return { updated: kloter.id };
    }
  }
  return saveKloter(kloter);
}

function deleteKloter(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_KLOTERS);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length-1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { deleted: id }; }
  }
  return { notFound: id };
}

// ── ADDRESSES ─────────────────────────────────────────────
function saveAddress(addr) {
  setupSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_ADDRESSES);
  const data  = sheet.getDataRange().getValues();
  const custName = String(addr.custName || addr.nama || '').trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === custName.toLowerCase()) {
      sheet.getRange(i+1,1,1,10).setValues([[
        custName, addr.nama||"", addr.hp||"", addr.alamat||"",
        addr.kota||"", addr.provinsi||"", addr.kodepos||"",
        addr.catatan||"", addr.ship||"", addr.timestamp||new Date().toISOString()
      ]]);
      return { updated: custName };
    }
  }
  sheet.appendRow([
    custName, addr.nama||"", addr.hp||"", addr.alamat||"",
    addr.kota||"", addr.provinsi||"", addr.kodepos||"",
    addr.catatan||"", addr.ship||"", addr.timestamp||new Date().toISOString()
  ]);
  return { saved: custName };
}

function getAddresses() {
  setupSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_ADDRESSES);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i]; });
    return obj;
  }).filter(a => a.custName && a.custName !== '');
}

// ── KIRIM STATUS ──────────────────────────────────────────
function saveKirimStatus(data) {
  setupSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_KIRIM);
  const rows = sheet.getDataRange().getValues();
  const custName = String(data.custName || '').trim();
  const now = new Date().toISOString();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === custName.toLowerCase()) {
      sheet.getRange(i+1,1,1,4).setValues([[
        custName, data.packed ? "true" : "false",
        data.sent ? "true" : "false", now
      ]]);
      return { updated: custName };
    }
  }
  sheet.appendRow([custName, data.packed ? "true" : "false", data.sent ? "true" : "false", now]);
  return { saved: custName };
}

function getKirimStatus() {
  setupSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_KIRIM);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};
  const result = {};
  data.slice(1).forEach(row => {
    const custName = String(row[0] || '').trim();
    if (!custName) return;
    result[custName] = {
      packed: String(row[1]).toLowerCase() === 'true',
      sent:   String(row[2]).toLowerCase() === 'true'
    };
  });
  return result;
}

// ── KATALOG BARANG ────────────────────────────────────────
function getKatalog() {
  setupSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_KATALOG);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    obj.id   = String(obj.id);
    obj.jpy  = Number(obj.jpy)  || 0;
    obj.jual = Number(obj.jual) || 0;
    return obj;
  }).filter(k => k.id && k.id !== '' && k.nama);
}

function saveKatalogItem(item) {
  setupSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_KATALOG);
  sheet.appendRow([
    String(item.id), item.nama, item.jpy||0, item.jual||0,
    new Date().toISOString()
  ]);
  return { saved: item.id };
}

function updateKatalogItem(item) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_KATALOG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(item.id)) {
      sheet.getRange(i+1,1,1,5).setValues([[
        String(item.id), item.nama, item.jpy||0, item.jual||0,
        new Date().toISOString()
      ]]);
      return { updated: item.id };
    }
  }
  return saveKatalogItem(item);
}

function deleteKatalogItem(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_KATALOG);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length-1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { deleted: id }; }
  }
  return { notFound: id };
}

// ── TEST ──────────────────────────────────────────────────
function testGetAll() {
  const result = getAllData();
  Logger.log(JSON.stringify(result));
}

// ── PIN SECURITY ──────────────────────────────────────────
// PIN hash is stored in Script Properties (not visible in sheet cells).
// Only a SHA-256 hash is ever stored — never the raw PIN.
function getPin() {
  const hash = PropertiesService.getScriptProperties().getProperty('sk_pin_hash') || '';
  return { hash: hash };
}

function setPin(hash) {
  if (!hash || hash.length !== 64) throw new Error('Invalid PIN hash');
  PropertiesService.getScriptProperties().setProperty('sk_pin_hash', hash);
  return { saved: true };
}
