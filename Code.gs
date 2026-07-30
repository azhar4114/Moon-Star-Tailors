/**
 * Moon Star Tailors — Google Apps Script API Backend
 * 
 * Instructions:
 * 1. Create a new Google Sheet named "Moon Star Tailors Orders"
 * 2. In Google Sheets menu: Go to Extensions > Apps Script
 * 3. Delete existing code in Code.gs and paste this ENTIRE file
 * 4. Click "Save" (disk icon)
 * 5. Click "Deploy" > "New deployment"
 * 6. Select type: "Web app"
 * 7. Description: "Moon Star Tailors API"
 * 8. Execute as: "Me"
 * 9. Who has access: "Anyone"  <-- IMPORTANT!
 * 10. Click "Deploy" -> Authorize access
 * 11. Copy the "Web app URL" (starts with https://script.google.com/macros/s/...)
 * 12. Paste the Web app URL in your Tracker app's "Cloud Sync Setup" modal!
 */

function doGet(e) {
  try {
    var sheet = getOrCreateSheet();
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return jsonResponse({ status: "success", orders: [] });
    }
    
    var headers = data[0];
    var orders = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // Skip empty rows
      
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
    var action = contents.action; // "save", "delete", "sync_all"
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
      // Clear sheet body and re-populate
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
    // Add header row
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
    // Format header
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
      targetRow = i + 1; // 1-indexed
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
