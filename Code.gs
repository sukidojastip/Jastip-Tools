// ================================================
// SUKIDO JASTIP — Google Apps Script Backend v2
// ================================================

const SHEET_NAME_ORDERS = "Orders";
const SHEET_NAME_KLOTERS = "Kloters";

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // FIX: safely handle missing event object
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';

  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    let result;
    switch (action) {
      case "getAll":       result = getAllData(); break;
      case "saveOrder":    result = saveOrder(JSON.parse(params.data)); break;
      case "updateOrder":  result = updateOrder(JSON.parse(params.data)); break;
      case "deleteOrder":  result = deleteOrder(params.id); break;
      case "saveKloter":   result = saveKloter(JSON.parse(params.data)); break;
      case "updateKloter": result = updateKloter(JSON.parse(params.data)); break;
      case "deleteKloter": result = deleteKloter(params.id); break;
      default:
        result = { ping: "ok", message: "Sukido Jastip API is running!" };
    }
    output.setContent(JSON.stringify({ ok: true, data: result }));
  } catch (err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
  }

  return output;
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let ordersSheet = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!ordersSheet) {
    ordersSheet = ss.insertSheet(SHEET_NAME_ORDERS);
    ordersSheet.appendRow(["id","date","cust","stype","pic","ship","item","qty","jpy","rate","jual","admin","dp","sbayar","ostatus","sudahBeli","batch"]);
    ordersSheet.getRange(1,1,1,17).setFontWeight("bold").setBackground("#e91e8c").setFontColor("#ffffff");
  }

  let klotersSheet = ss.getSheetByName(SHEET_NAME_KLOTERS);
  if (!klotersSheet) {
    klotersSheet = ss.insertSheet(SHEET_NAME_KLOTERS);
    klotersSheet.appendRow(["id","name","rate","bagasi","eta","status","pengeluaran"]);
    klotersSheet.getRange(1,1,1,7).setFontWeight("bold").setBackground("#7c3aed").setFontColor("#ffffff");
  }
}

function getAllData() {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const oSheet = ss.getSheetByName(SHEET_NAME_ORDERS);
  const oData = oSheet.getDataRange().getValues();
  const oHeaders = oData[0];
  const orders = oData.slice(1).map(row => {
    const obj = {};
    oHeaders.forEach((h, i) => { obj[h] = row[i]; });
    obj.id       = String(obj.id);
    obj.qty      = Number(obj.qty)   || 1;
    obj.jpy      = Number(obj.jpy)   || 0;
    obj.rate     = Number(obj.rate)  || 0;
    obj.jual     = Number(obj.jual)  || 0;
    obj.admin    = Number(obj.admin) || 0;
    obj.dp       = Number(obj.dp)    || 0;
    obj.sudahBeli = (obj.sudahBeli === true || obj.sudahBeli === "TRUE" || obj.sudahBeli === "true");
    return obj;
  }).filter(o => o.id && o.id !== 'undefined' && o.id !== '');

  const kSheet = ss.getSheetByName(SHEET_NAME_KLOTERS);
  const kData = kSheet.getDataRange().getValues();
  const kHeaders = kData[0];
  const kloters = kData.slice(1).map(row => {
    const obj = {};
    kHeaders.forEach((h, i) => { obj[h] = row[i]; });
    obj.id   = String(obj.id);
    obj.rate = Number(obj.rate) || 0;
    try { obj.pengeluaran = obj.pengeluaran ? JSON.parse(obj.pengeluaran) : []; }
    catch(e) { obj.pengeluaran = []; }
    return obj;
  }).filter(k => k.id && k.id !== 'undefined' && k.id !== '');

  return { orders, kloters };
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

// Jalankan fungsi ini dulu untuk test apakah bisa baca Sheet
function testGetAll() {
  const result = getAllData();
  Logger.log(JSON.stringify(result));
}
