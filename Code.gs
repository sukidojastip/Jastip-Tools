// ================================================
// SUKIDO JASTIP — Google Apps Script Backend v3
// ================================================

const SHEET_NAME_ORDERS   = "Orders";
const SHEET_NAME_KLOTERS  = "Kloters";
const SHEET_NAME_ADDRESSES = "Addresses";

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  // Allow CORS
  try {
    let result;
    switch (action) {
      case "getAll":         result = getAllData();                            break;
      case "saveOrder":      result = saveOrder(JSON.parse(params.data));     break;
      case "updateOrder":    result = updateOrder(JSON.parse(params.data));   break;
      case "deleteOrder":    result = deleteOrder(params.id);                 break;
      case "saveKloter":     result = saveKloter(JSON.parse(params.data));    break;
      case "updateKloter":   result = updateKloter(JSON.parse(params.data));  break;
      case "deleteKloter":   result = deleteKloter(params.id);                break;
      case "saveAddress":    result = saveAddress(JSON.parse(params.data));   break;
      case "getAddresses":   result = getAddresses();                         break;
      default:
        result = { ping: "ok", message: "Sukido Jastip API v3 running!" };
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
    s.appendRow(["id","date","cust","stype","pic","ship","item","qty","jpy","rate","jual","admin","dp","sbayar","ostatus","sudahBeli","batch"]);
    s.getRange(1,1,1,17).setFontWeight("bold").setBackground("#e91e8c").setFontColor("#ffffff");
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
    obj.sudahBeli = (obj.sudahBeli === true || String(obj.sudahBeli).toLowerCase() === "true");
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

  // Also fetch addresses
  const addresses = getAddresses();

  return { orders, kloters, addresses };
}

function saveOrder(order) {
  setupSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_ORDERS);
  sheet.appendRow([
    String(order.id), order.date, order.cust, order.stype, order.pic,
    order.ship, order.item, order.qty, order.jpy, order.rate,
    order.jual, order.admin||0, order.dp||0, order.sbayar,
    order.ostatus, order.sudahBeli||false, order.batch
  ]);
  return { saved: order.id };
}

function updateOrder(order) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_ORDERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(order.id)) {
      sheet.getRange(i+1,1,1,17).setValues([[
        String(order.id), order.date, order.cust, order.stype, order.pic,
        order.ship, order.item, order.qty, order.jpy, order.rate,
        order.jual, order.admin||0, order.dp||0, order.sbayar,
        order.ostatus, order.sudahBeli||false, order.batch
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

  // Update existing row if custName matches
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
  // Insert new row
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

// ── TEST ──────────────────────────────────────────────────
function testGetAll() {
  const result = getAllData();
  Logger.log(JSON.stringify(result));
}

function testSaveAddress() {
  saveAddress({
    custName: "Test Customer",
    nama: "Test Nama",
    hp: "08123456789",
    alamat: "Jl. Test No. 1",
    kota: "Jakarta",
    provinsi: "DKI Jakarta",
    kodepos: "12345",
    catatan: "",
    ship: "Regular JNT",
    timestamp: new Date().toISOString()
  });
}
