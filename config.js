// ================================================================
// SOUDSADA SHOP — Code.gs (Backend API v5 — GitHub Pages Edition)
// ================================================================
// ວິທີ Deploy:
//   1. Extensions > Apps Script
//   2. Deploy > New deployment > Web App
//   3. Execute as: Me | Who has access: Anyone
//   4. Copy the Web App URL ໄວ້ໃຊ້ໃນ Frontend (config.js)
// ================================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();
const LAO_TZ = "Asia/Vientiane";
const SLACK_WEBHOOK_URL = "https://hooks.slack.com/services/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";

// ================================================================
// CORS Helper — ຕ້ອງໃສ່ທຸກ Response
// ================================================================
function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
}

function jsonResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function errorResponse(msg, code) {
  return jsonResponse({ success: false, error: msg, code: code || 400 });
}

function okResponse(data) {
  return jsonResponse(Object.assign({ success: true }, data));
}

// ================================================================
// doGet — ຈັດການ GET requests
// action = ?action=xxx&param=yyy
// ================================================================
function doGet(e) {
  var action = (e.parameter && e.parameter.action) ? e.parameter.action : "";

  try {
    switch (action) {
      case "getProducts":     return okResponse({ data: getProductData() });
      case "getDashboard":    return okResponse({ data: getDashboardData() });
      case "getReport":       return okResponse({ data: getReportData() });
      case "getHistory":      return okResponse({ data: getHistoryData() });
      case "getWeeklyTop10":  return okResponse({ data: getWeeklyTop10() });
      case "getBannerSettings": return okResponse({ data: getBannerSettings() });
      case "ping":            return okResponse({ message: "SOUDSADA SHOP API ✅", time: new Date().toISOString() });
      default:
        return errorResponse("Unknown action: " + action, 404);
    }
  } catch (err) {
    Logger.log("doGet error: " + err.message);
    return errorResponse(err.message, 500);
  }
}

// ================================================================
// doPost — ຈັດການ POST requests (JSON body)
// { "action": "xxx", ...params }
// ================================================================
function doPost(e) {
  var body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return errorResponse("Invalid JSON body", 400);
  }

  var action = body.action || "";

  try {
    switch (action) {
      case "saveProduct":
        var res1 = saveProduct(body.name, body.cost, body.price, body.qty, body.imgUrl || "");
        return okResponse({ message: res1 });

      case "editProductFull":
        var res2 = editProductFull(body.oldName, body.newName, body.cost, body.price, body.qty, body.imgUrl || "");
        return okResponse({ message: res2 });

      case "quickAddStock":
        var res3 = quickAddStock(body.name, body.qty);
        return okResponse({ message: res3 });

      case "deleteProduct":
        var res4 = deleteProduct(body.name);
        return okResponse({ message: res4 });

      case "recordBulkSale":
        var res5 = recordBulkSale(body.cart);
        return okResponse({ message: res5 });

      case "cancelSaleData":
        var res6 = cancelSaleData(body.dateStr, body.name, body.qty);
        return okResponse({ message: res6 });

      case "editSaleRow":
        var res7 = editSaleRow(body.dateStr, body.name, body.oldQty, body.newQty);
        return okResponse({ message: res7 });

      case "cancelStockRow":
        var res8 = cancelStockRow(body.dateStr, body.name, body.qty);
        return okResponse({ message: res8 });

      case "editStockRow":
        var res9 = editStockRow(body.dateStr, body.name, body.oldQty, body.newQty);
        return okResponse({ message: res9 });

      case "saveBannerSettings":
        var res10 = saveBannerSettings(body.settings);
        return okResponse({ message: res10 });

      default:
        return errorResponse("Unknown action: " + action, 404);
    }
  } catch (err) {
    Logger.log("doPost error [" + action + "]: " + err.message);
    return errorResponse(err.message, 500);
  }
}

// ================================================================
// ຟັງຊັນ Sheet helpers
// ================================================================
function getSheetSafe(name) {
  var s = SS.getSheetByName(name);
  if (!s) throw new Error("ຫາ Sheet '" + name + "' ບໍ່ເຫັນ!");
  return s;
}

function getOrderCode(phone) {
  var p = String(phone || "").replace(/\D/g, "");
  var last4 = p.length >= 4 ? p.slice(-4) : p.padStart(4, "0");
  return "P" + last4;
}

// ================================================================
// SLACK
// ================================================================
function sendSlackAlert(message) {
  if (!message || String(message).trim() === "") return;
  try {
    var payload = JSON.stringify({ text: String(message).trim() });
    var options = { method: "post", contentType: "application/json", payload: payload, muteHttpExceptions: true };
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
  } catch (e) {
    Logger.log("Slack error: " + e.message);
  }
}

// ================================================================
// ຜະລິດຕະພັນ — col A=ຊື່, B=ທຶນ, C=ຂາຍ, D=ຄົງເຫຼືອ, E=ຮູບ
// ================================================================
function getProductData() {
  var sheet = getSheetSafe("ຜະລິດຕະພັນ");
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      name:  data[i][0],
      cost:  data[i][1] || 0,
      price: data[i][2] || 0,
      qty:   data[i][3] || 0,
      img:   data[i].length > 4 ? (data[i][4] || "") : ""
    });
  }
  return result;
}

function saveProduct(name, cost, price, qty, imgUrl) {
  var sheet = getSheetSafe("ຜະລິດຕະພັນ");
  var log   = SS.getSheetByName("ປະຫວັດສະຕ໋ອກ");
  var data  = sheet.getDataRange().getValues();
  var img   = imgUrl || "";

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == name) {
      var newQty = (Number(data[i][3]) || 0) + Number(qty);
      var existingImg = data[i].length > 4 ? (data[i][4] || "") : "";
      var finalImg = img || existingImg;
      sheet.getRange(i + 1, 1, 1, 5).setValues([[name, cost, price, newQty, finalImg]]);
      if (log) log.appendRow([new Date(), name, "+" + qty, "ເຕີມສະຕ໋ອກ", newQty]);
      sendSlackAlert("📦 ເຕີມ: " + name + " +" + qty + " (ຄົງ " + newQty + ")");
      return "ອັບເດດສຳເລັດ!";
    }
  }
  sheet.appendRow([name, cost, price, qty, img]);
  if (log) log.appendRow([new Date(), name, "+" + qty, "ເພີ່ມໃໝ່", qty]);
  sendSlackAlert("🆕 ເພີ່ມ: " + name + " (" + qty + ")");
  return "ບັນທຶກໃໝ່ສຳເລັດ!";
}

function editProductFull(oldName, newName, cost, price, qty, imgUrl) {
  var sheet = getSheetSafe("ຜະລິດຕະພັນ");
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == oldName) {
      var existingImg = data[i].length > 4 ? (data[i][4] || "") : "";
      sheet.getRange(i + 1, 1, 1, 5).setValues([[newName, cost, price, qty, imgUrl || existingImg]]);
      return "ແກ້ໄຂສຳເລັດ!";
    }
  }
  return "ບໍ່ພົບ!";
}

function quickAddStock(name, qty) {
  var sheet = getSheetSafe("ຜະລິດຕະພັນ");
  var log   = SS.getSheetByName("ປະຫວັດສະຕ໋ອກ");
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == name) {
      var newQty = (Number(data[i][3]) || 0) + Number(qty);
      sheet.getRange(i + 1, 4).setValue(newQty);
      if (log) log.appendRow([new Date(), name, "+" + qty, "ເຕີມດ່ວນ", newQty]);
      checkLowStockAlert(name, newQty);
      return "ເຕີມສຳເລັດ! (" + name + " ຄົງ " + newQty + ")";
    }
  }
  return "ບໍ່ພົບ!";
}

function deleteProduct(name) {
  var sheet = getSheetSafe("ຜະລິດຕະພັນ");
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == name) { sheet.deleteRow(i + 1); return "ລຶບສຳເລັດ!"; }
  }
  return "ບໍ່ພົບ!";
}

// ================================================================
// ການຂາຍ
// ================================================================
function recordBulkSale(cart) {
  var sheetSale = getSheetSafe("ການຂາຍ");
  var sheetProd = getSheetSafe("ຜະລິດຕະພັນ");
  var sheetLog  = SS.getSheetByName("ປະຫວັດສະຕ໋ອກ");
  var pData     = sheetProd.getDataRange().getValues();

  var fi     = cart[0] || {};
  var cName  = fi.customerName  || "";
  var cPhone = fi.customerPhone || "";
  var cAddr  = fi.customerAddr  || "";
  var cShip  = fi.shipping      || "";
  var cCod   = fi.cod           || "";

  var orderCode  = getOrderCode(cPhone);
  var dateLabel  = Utilities.formatDate(new Date(), LAO_TZ, "dd-MM-yyyy");
  var timeStr    = Utilities.formatDate(new Date(), LAO_TZ, "dd/MM/yyyy HH:mm");
  var grandTotal = 0;

  var msgLines = ["🛒 ລາຍການຂາຍໃໝ່ — " + dateLabel, "════════════════════"];
  if (cName)  msgLines.push("👤 " + cName + (cPhone ? " | " + cPhone : ""));
  if (cAddr)  msgLines.push("🏠 " + cAddr);
  if (cShip)  msgLines.push("🚚 " + cShip);
  if (cCod)   msgLines.push("💵 COD: " + cCod);
  msgLines.push("────────────────────");

  cart.forEach(function (item) {
    var cost = 0;
    for (var i = 1; i < pData.length; i++) {
      if (pData[i][0] == item.name) {
        cost = Number(pData[i][1]) || 0;
        var curStock = Number(pData[i][3]) || 0;
        var newStock = curStock - Number(item.qty);
        sheetProd.getRange(i + 1, 4).setValue(newStock);
        if (sheetLog) sheetLog.appendRow([new Date(), item.name, "-" + item.qty, "ຂາຍ", newStock]);
        checkLowStockAlert(item.name, newStock);
        break;
      }
    }
    var tCost  = cost * item.qty;
    var tSales = (item.price * item.qty) - (item.discount || 0);
    grandTotal += tSales;
    sheetSale.appendRow([
      new Date(), item.name, item.qty, tSales, tCost, tSales - tCost,
      item.discount || 0, item.note || "",
      cName, cPhone, cAddr, cShip, cCod, orderCode
    ]);
    msgLines.push("📦 " + item.name + " x" + item.qty + " = " + tSales.toLocaleString() + " ₭");
  });

  msgLines.push("════════════════════");
  msgLines.push("💰 ລວມ: " + grandTotal.toLocaleString() + " ₭ | " + timeStr);
  sendSlackAlert(msgLines.join("\n"));
  return "ບັນທຶກສຳເລັດ! (" + orderCode + ")";
}

// ================================================================
// Dashboard
// ================================================================
function getDashboardData() {
  var sheet = SS.getSheetByName("ການຂາຍ");
  if (!sheet) return {};
  var data   = sheet.getDataRange().getValues();
  var result = {};

  for (var i = 1; i < data.length; i++) {
    if (!(data[i][0] instanceof Date)) continue;
    var dateStr = Utilities.formatDate(data[i][0], LAO_TZ, "yyyy-MM-dd");
    var rowDate = new Date(data[i][0]);
    var diff = (new Date() - rowDate) / (1000 * 60 * 60 * 24);
    if (diff > 30) continue;

    if (!result[dateStr]) result[dateStr] = { totalSales: 0, profit: 0, items: [], rawRows: [] };
    var sales  = Number(data[i][3]) || 0;
    var profit = Number(data[i][5]) || 0;
    result[dateStr].totalSales += sales;
    result[dateStr].profit     += profit;
    result[dateStr].items.push({ name: data[i][1], qty: data[i][2], total: sales });

    var row = [];
    for (var c = 0; c < 14; c++) row.push(data[i][c] !== undefined ? String(data[i][c]) : "");
    result[dateStr].rawRows.push(row);
  }
  return result;
}

// ================================================================
// Report
// ================================================================
function getReportData() {
  var saleSheet = SS.getSheetByName("ການຂາຍ");
  if (!saleSheet) return [];
  var data = saleSheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var row = [];
    var d = data[i][0];
    row.push(d instanceof Date ? Utilities.formatDate(d, LAO_TZ, "yyyy-MM-dd HH:mm:ss") : String(d));
    for (var c = 1; c < 14; c++) row.push(data[i][c] !== undefined ? data[i][c] : "");
    result.push(row);
  }
  return result;
}

// ================================================================
// History
// ================================================================
function getHistoryData() {
  var saleSheet = SS.getSheetByName("ການຂາຍ");
  var logSheet  = SS.getSheetByName("ປະຫວັດສະຕ໋ອກ");

  var sales = [];
  if (saleSheet) {
    var sData = saleSheet.getDataRange().getValues();
    for (var i = 1; i < sData.length; i++) {
      if (!sData[i][0]) continue;
      var r    = [];
      var d    = sData[i][0];
      var dStr = d instanceof Date ? Utilities.formatDate(d, LAO_TZ, "yyyy-MM-dd HH:mm:ss") : String(d);
      r.push(dStr);
      for (var c = 1; c < 14; c++) r.push(sData[i][c] !== undefined ? sData[i][c] : "");
      r.push(dStr);
      sales.push(r);
    }
  }

  var logs = [];
  if (logSheet) {
    var lData = logSheet.getDataRange().getValues();
    for (var j = 1; j < lData.length; j++) {
      if (!lData[j][0]) continue;
      var lr   = [];
      var ld   = lData[j][0];
      var ldStr = ld instanceof Date ? Utilities.formatDate(ld, LAO_TZ, "yyyy-MM-dd HH:mm:ss") : String(ld);
      lr.push(ldStr);
      for (var lc = 1; lc < 5; lc++) lr.push(lData[j][lc] !== undefined ? lData[j][lc] : "");
      lr.push(ldStr);
      logs.push(lr);
    }
  }
  return { sales: sales, logs: logs };
}

// ================================================================
// Edit / Delete History
// ================================================================
function cancelSaleData(dateStr, name, qty) {
  var sheet = getSheetSafe("ການຂາຍ");
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var d = data[i][0];
    var rowDateStr = d instanceof Date ? Utilities.formatDate(d, LAO_TZ, "yyyy-MM-dd HH:mm:ss") : String(d);
    if (rowDateStr == dateStr && data[i][1] == name && String(data[i][2]) == String(qty)) {
      sheet.deleteRow(i + 1); return "ລຶບສຳເລັດ!";
    }
  }
  return "ບໍ່ພົບ!";
}

function editSaleRow(dateStr, name, oldQty, newQty) {
  var sheet = getSheetSafe("ການຂາຍ");
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var d = data[i][0];
    var rowDateStr = d instanceof Date ? Utilities.formatDate(d, LAO_TZ, "yyyy-MM-dd HH:mm:ss") : String(d);
    if (rowDateStr == dateStr && data[i][1] == name && String(data[i][2]) == String(oldQty)) {
      sheet.getRange(i + 1, 3).setValue(Number(newQty)); return "ແກ້ໄຂສຳເລັດ!";
    }
  }
  return "ບໍ່ພົບ!";
}

function cancelStockRow(dateStr, name, qty) {
  var sheet = SS.getSheetByName("ປະຫວັດສະຕ໋ອກ");
  if (!sheet) return "ບໍ່ພົບ!";
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var d = data[i][0];
    var rowDateStr = d instanceof Date ? Utilities.formatDate(d, LAO_TZ, "yyyy-MM-dd HH:mm:ss") : String(d);
    if (rowDateStr == dateStr && data[i][1] == name && String(data[i][2]) == String(qty)) {
      sheet.deleteRow(i + 1); return "ລຶບສຳເລັດ!";
    }
  }
  return "ບໍ່ພົບ!";
}

function editStockRow(dateStr, name, oldQty, newQty) {
  var sheet = SS.getSheetByName("ປະຫວັດສະຕ໋ອກ");
  if (!sheet) return "ບໍ່ພົບ!";
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var d = data[i][0];
    var rowDateStr = d instanceof Date ? Utilities.formatDate(d, LAO_TZ, "yyyy-MM-dd HH:mm:ss") : String(d);
    if (rowDateStr == dateStr && data[i][1] == name && String(data[i][2]) == String(oldQty)) {
      sheet.getRange(i + 1, 3).setValue(Number(newQty)); return "ແກ້ໄຂສຳເລັດ!";
    }
  }
  return "ບໍ່ພົບ!";
}

// ================================================================
// ແຈ້ງເຕືອນ
// ================================================================
function checkLowStockAlert(productName, currentQty) {
  var qty = Number(currentQty);
  if (qty <= 0)      sendSlackAlert("🚨 ໝົດ: " + productName);
  else if (qty <= 5) sendSlackAlert("⚠️ ໃກ້ໝົດ: " + productName + " (" + qty + ")");
}

// ================================================================
// Top 10 ສິນຄ້າຂາຍດີ ອາທິດ
// ================================================================
function getWeeklyTop10() {
  var saleSheet = SS.getSheetByName("ການຂາຍ");
  var prodSheet = getSheetSafe("ຜະລິດຕະພັນ");
  if (!saleSheet) return [];

  var sData = saleSheet.getDataRange().getValues();
  var pData = prodSheet.getDataRange().getValues();

  var imgMap = {};
  for (var p = 1; p < pData.length; p++) {
    if (pData[p][0]) imgMap[pData[p][0]] = pData[p][4] || "";
  }

  var now     = new Date();
  var weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  var salesMap = {};

  for (var i = 1; i < sData.length; i++) {
    if (!(sData[i][0] instanceof Date)) continue;
    if (sData[i][0] < weekAgo) continue;
    var nm = String(sData[i][1] || "").trim();
    if (!nm) continue;
    salesMap[nm] = (salesMap[nm] || 0) + (Number(sData[i][2]) || 0);
  }

  return Object.keys(salesMap)
    .map(function (n) { return { name: n, qty: salesMap[n], img: imgMap[n] || "" }; })
    .sort(function (a, b) { return b.qty - a.qty; })
    .slice(0, 10);
}

// ================================================================
// Banner Settings — Sheet "ຕັ້ງຄ່າ"
// ================================================================
function ensureSettingsSheet() {
  var sh = SS.getSheetByName("ຕັ້ງຄ່າ");
  if (!sh) {
    sh = SS.insertSheet("ຕັ້ງຄ່າ");
    sh.getRange(1, 1, 1, 2).setValues([["key", "value"]]);
    sh.getRange(2, 1, 1, 2).setValues([["tickerText", "🔥 Soudsada Shop · ສິນຄ້າຂາຍດີປະຈຳອາທິດ · ຍິນດີຕ້ອນຮັບ! 🛒"]]);
  }
  return sh;
}

function getBannerSettings() {
  var sh   = ensureSettingsSheet();
  var data = sh.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) result[String(data[i][0])] = String(data[i][1] || "");
  }
  if (!result.tickerText) result.tickerText = "🔥 Soudsada Shop · ສິນຄ້າຂາຍດີປະຈຳອາທິດ";
  return result;
}

function saveBannerSettings(settings) {
  var sh   = ensureSettingsSheet();
  var data = sh.getDataRange().getValues();
  Object.keys(settings).forEach(function (key) {
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === key) {
        sh.getRange(i + 1, 2).setValue(settings[key]);
        found = true;
        break;
      }
    }
    if (!found) sh.appendRow([key, settings[key]]);
  });
  return "ບັນທຶກສຳເລັດ!";
}