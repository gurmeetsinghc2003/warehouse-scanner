/**
 * WAREHOUSE RACK SCANNER — Code.gs (Final)
 */

var SHEET_PRODUCTS = 'ProductLocations';
var SHEET_SCANLOG  = 'ScanLog';
var COL_BARCODE     = 1;
var COL_RACKS       = 2;
var COL_LASTUPDATED = 3;

function doGet(e) {
  // No action = serve the HTML app
  if (!e.parameter.action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Rack Scanner')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var result = {};
  try {
    var params  = e.parameter;
    var barcode = params.barcode || '';
    var rack    = params.rack    || '';

    var action = params.action || '';

    if (action === 'undo') {
      if (!barcode || !rack) {
        result = { success: false, result: 'Error', message: 'Missing barcode or rack for undo.' };
      } else {
        result = processUndo(barcode, rack);
      }
    } else if (action === 'save') {
      if (!barcode || !rack) {
        result = { success: false, result: 'Error', message: 'Missing barcode or rack parameter.' };
      } else {
        result = processScan(barcode, rack);
      }
    } else {
      result = { success: false, result: 'Error', message: 'Unknown action: ' + action };
    }
  } catch (err) {
    Logger.log('doGet error: ' + err.toString());
    result = { success: false, result: 'Error', message: 'Server error: ' + err.message };
  }

  // JSONP support — required for GitHub Pages cross-origin requests
  var callback = e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function processScan(barcode, rack) {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var sheetProducts = getOrCreateSheet(ss, SHEET_PRODUCTS, ['ProductBarcode', 'RackNumbers', 'LastUpdated']);
  var sheetScanLog  = getOrCreateSheet(ss, SHEET_SCANLOG,  ['Timestamp', 'ProductBarcode', 'RackScanned', 'Result']);

  var now  = formatTimestamp(new Date());
  var data = sheetProducts.getDataRange().getValues();

  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_BARCODE - 1]).trim() === barcode) {
      rowIndex = i;
      break;
    }
  }

  var result;
  var message;

  if (rowIndex === -1) {
    sheetProducts.appendRow([barcode, rack, now]);
    result  = 'Added';
    message = 'Product added with rack ' + rack;
  } else {
    var existingRacks = String(data[rowIndex][COL_RACKS - 1]).trim();
    var rackList      = existingRacks.split(',').map(function(r) { return r.trim(); });

    if (rackList.indexOf(rack) !== -1) {
      result  = 'Duplicate';
      message = 'Rack ' + rack + ' already assigned to ' + barcode;
    } else {
      var updatedRacks = existingRacks + ',' + rack;
      var sheetRow     = rowIndex + 1;
      sheetProducts.getRange(sheetRow, COL_RACKS).setValue(updatedRacks);
      sheetProducts.getRange(sheetRow, COL_LASTUPDATED).setValue(now);
      result  = 'Appended';
      message = 'Rack ' + rack + ' added to ' + barcode + '. Total: ' + (rackList.length + 1);
    }
  }

  writeLog(sheetScanLog, now, barcode, rack, result);
  return { success: true, result: result, message: message };
}

function processUndo(barcode, rack) {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var sheetProducts = getOrCreateSheet(ss, SHEET_PRODUCTS, ['ProductBarcode', 'RackNumbers', 'LastUpdated']);
  var sheetScanLog  = getOrCreateSheet(ss, SHEET_SCANLOG,  ['Timestamp', 'ProductBarcode', 'RackScanned', 'Result']);
  var now           = formatTimestamp(new Date());
  var data          = sheetProducts.getDataRange().getValues();

  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_BARCODE - 1]).trim() === barcode) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, result: 'Error', message: 'Product not found in sheet.' };
  }

  var existingRacks = String(data[rowIndex][COL_RACKS - 1]).trim();
  var rackList      = existingRacks.split(',').map(function(r) { return r.trim(); });
  var rackIndex     = rackList.indexOf(rack);

  if (rackIndex === -1) {
    return { success: false, result: 'Error', message: 'Rack not found for this product.' };
  }

  // Remove the rack from the list
  rackList.splice(rackIndex, 1);
  var sheetRow = rowIndex + 1;

  if (rackList.length === 0) {
    // No racks left — delete the entire row
    sheetProducts.deleteRow(sheetRow);
  } else {
    // Update with remaining racks
    sheetProducts.getRange(sheetRow, COL_RACKS).setValue(rackList.join(','));
    sheetProducts.getRange(sheetRow, COL_LASTUPDATED).setValue(now);
  }

  writeLog(sheetScanLog, now, barcode, rack, 'Undone');
  return { success: true, result: 'Undone', message: 'Removed ' + rack + ' from ' + barcode };
}

function writeLog(sheet, timestamp, barcode, rack, result) {
  try {
    sheet.appendRow([timestamp, barcode, rack, result]);
  } catch (err) {
    Logger.log('ScanLog write error: ' + err.toString());
  }
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a1a1a');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatTimestamp(date) {
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var d  = date.getDate();
  var m  = months[date.getMonth()];
  var y  = date.getFullYear();
  var hh = String(date.getHours()).padStart(2, '0');
  var mm = String(date.getMinutes()).padStart(2, '0');
  return d + '-' + m + '-' + y + ' ' + hh + ':' + mm;
}
