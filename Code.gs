/**
 * Moon Star Tailors — Google Apps Script API Backend (Secured)
 * 
 * SECURITY NOTE:
 * Set your SECRET_KEY below. Any request without this key will be REJECTED with 403 Unauthorized.
 */

// Set your secret key/PIN here. Set to "" if you want to disable authentication.
var SECRET_KEY = "MST8898850"; 

function isAuthorized(e, contents) {
  if (!SECRET_KEY || SECRET_KEY === "") return true;
  
  // Check URL parameter e.g. ?key=...
  if (e && e.parameter && e.parameter.key === SECRET_KEY) return true;
  
  // Check POST JSON payload body e.g. { key: "..." }
  if (contents && contents.key === SECRET_KEY) return true;
  
  return false;
}

function doGet(e) {
  try {
    if (!isAuthorized(e, null)) {
      return jsonResponse({ status: "error", message: "403 Unauthorized: Invalid or missing Security Key." });
    }

    var sheet = getOrCreateSheet();
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return jsonResponse({ status: "success", orders: [] });
    }
    
    var headers = data[0];
    var orders = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      
      var order = {
        id: row[0].toString(),
        custName: row[1] ? row[1].toString() : "",
        custPhone: row[2] ? row[2].toString() : "",
        garmentType: row[3] ? row[3].toString() : "",
        dueDate: row[4] ? formatDate(row[4]) : "",
        status: row[5] ? row[5].toString() : "Received",
        total: parseFloat(row[6]) || 0,
        advance: parseFloat(row[7]) || 0,
        notes: row[8] ? row[8].toString() : "",
        measurements: row[9] ? safeParseJSON(row[9]) : {},
        created: row[10] ? row[10].toString() : new Date().toISOString()
      };
      orders.push(order);
    }
    
    return jsonResponse({ status: "success", orders: orders });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
    
    if (!isAuthorized(e, contents)) {
      return jsonResponse({ status: "error", message: "403 Unauthorized: Invalid or missing Security Key." });
    }

    var action = contents.action;
    var sheet = getOrCreateSheet();
    
    if (action === "save") {
      var order = contents.order;
      saveOrUpdateOrder(sheet, order);
      return jsonResponse({ status: "success", message: "Order saved successfully", id: order.id });
    } 
    else if (action === "delete") {
      var orderId = contents.id;
      deleteOrderById(sheet, orderId);
      return jsonResponse({ status: "success", message: "Order deleted", id: orderId });
    }
    else if (action === "sync_all") {
      var allOrders = contents.orders || [];
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 11).clearContent();
      }
      for (var i = 0; i < allOrders.length; i++) {
        saveOrUpdateOrder(sheet, allOrders[i]);
      }
      return jsonResponse({ status: "success", message: "All orders synced", count: allOrders.length });
    }
    else {
      return jsonResponse({ status: "error", message: "Invalid action" });
    }
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Orders");
  if (!sheet) {
    sheet = ss.insertSheet("Orders");
    sheet.appendRow([
      "Order ID",
      "Customer Name",
      "Phone",
      "Garment Type",
      "Due Date",
      "Status",
      "Total (₹)",
      "Advance (₹)",
      "Notes",
      "Measurements (JSON)",
      "Created At"
    ]);
    var headerRange = sheet.getRange(1, 1, 1, 11);
    headerRange.setBackground("#1a1a2e").setFontColor("#c9a84c").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveOrUpdateOrder(sheet, order) {
  var data = sheet.getDataRange().getValues();
  var targetRow = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === order.id.toString()) {
      targetRow = i + 1;
      break;
    }
  }
  
  var rowValues = [
    order.id,
    order.custName,
    order.custPhone,
    order.garmentType,
    order.dueDate,
    order.status,
    order.total || 0,
    order.advance || 0,
    order.notes || "",
    JSON.stringify(order.measurements || {}),
    order.created || new Date().toISOString()
  ];
  
  if (targetRow > -1) {
    sheet.getRange(targetRow, 1, 1, 11).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
}

function deleteOrderById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeParseJSON(str) {
  try {
    return typeof str === "string" ? JSON.parse(str) : str;
  } catch (e) {
    return {};
  }
}

function formatDate(val) {
  if (val instanceof Date) {
    var year = val.getFullYear();
    var month = ("0" + (val.getMonth() + 1)).slice(-2);
    var day = ("0" + val.getDate()).slice(-2);
    return year + "-" + month + "-" + day;
  }
  return val ? val.toString() : "";
}
