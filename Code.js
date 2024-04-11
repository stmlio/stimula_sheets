function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu("Data Access")
    .addItem("Instant Data Access", "openSidebar")
    .addToUi();
}

function openSidebar() {
  const version = '0.1'
  var htmlOutput = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Instant Data Access (' + version + ')')
    .setWidth(100);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function doStandaloneConnect(url, params) {
  const options = {
    "method": "post",
    "payload":
      'host=' + encodeURIComponent(params['host']) +
      '&port=' + encodeURIComponent(params['port']) +
      '&user=' + encodeURIComponent(params['user']) +
      '&password=' + encodeURIComponent(params['password']) +
      '&database=' + encodeURIComponent(params['database'])
  }
  const response = _makeHttpRequest(url + '/auth', options)
  return JSON.parse(response);
}
function doOdooConnect(url, database, user, password) {
  const options = {
    "method": "post",
    "payload":
      'database=' + encodeURIComponent(database) +
      '&username=' + encodeURIComponent(user) +
      '&password=' + encodeURIComponent(password)
  }
  const response = _makeHttpRequest(url + '/auth', options)
  return JSON.parse(response);
}

function getTables(url, token, filter) {
  var response = _makeHttpRequest(url + '/tables' + '?q=' + encodeURIComponent(filter), {}, token);
  var data = JSON.parse(response);
  return data;
}

function openTab(name) {
  _activateSheet(name)
}

function getDefaultHeaders(url, token) {
  // get table name from active sheet
  const tableName = _getActiveTableName(url, token)
  // get default headers, ignore current headers
  const headers = _makeHttpRequest(url + '/tables/' + tableName + '/header', {}, token)
  return JSON.parse(headers)
}

function getHeaders(url, token) {
  // get table name from active sheet
  const tableName = _getActiveTableName(url, token)
  const headerFromSpreadsheet = _getCurrentHeader(tableName) ?? ''
  // get default headers, merged with current headers
  const headers = _makeHttpRequest(url + '/tables/' + tableName + '/header?h=' + headerFromSpreadsheet, {}, token)
  return JSON.parse(headers)
}

function applyHeaders(headers) {
  var data = [];
  for (var header of headers) {
    if (header['enabled']) {
      data.push(header['key']);
    }
  }
  Logger.log(data)
  // get active sheet
  const sheet = SpreadsheetApp.getActiveSheet()
  // clear first row
  if (sheet.getLastColumn() > 0)
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).clearContent();
  Logger.log(data.length)
  // set headers
  sheet.getRange(1, 1, 1, data.length).setValues([data]);
}

function getRowCount(base_url, token, whereClause) {
  // get table name from active sheet, don't test if it's a table, the count call is a good test by itself
  tableName = SpreadsheetApp.getActiveSheet().getName()
  // get current header, because that defines the query we want to restrict
  const header = _getCurrentHeader(tableName)
  // URL encode header
  const safe_header = encodeURIComponent(header)
  // URL encode where clause
  const safe_whereClause = encodeURIComponent(whereClause)
  // create URL
  const url = base_url + '/tables/' + tableName + '/count?h=' + safe_header + '&q=' + safe_whereClause
  // get count of rows that match where clause
  const response = _makeHttpRequest(url, {}, token)
  // parse json
  const data = JSON.parse(response)
  // return count
  return data.count
}

function getTable(url, token, whereClause) {
  // get table name from active sheet
  const tableName = _getActiveTableName(url, token)

  // get desired header from first row in spreadsheet
  const headerFromSpreadsheet = _getCurrentHeader(tableName) ?? ''

  // request headers with types from server
  const headerWithTypes = _makeHttpRequest(url + '/tables/' + tableName + '/header?h=' + headerFromSpreadsheet, {}, token)

  // set cell format in spreadsheet.
  _setCellFormat(headerWithTypes)

  // request table data from server
  response = _makeHttpRequest(url + '/tables/' + tableName + '?h=' + headerFromSpreadsheet + '&q=' + whereClause, {}, token)
  _displaySheet(tableName, response)
}

function _activateSheet(name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName(name);
  if (sheet) {
    // activate existing
    spreadsheet.setActiveSheet(sheet);
  } else {
    // insert and activate new sheet
    sheet = spreadsheet.insertSheet(name);
  }
  return sheet
}

function _setCellFormat(headerWithTypes) {
  // get active sheet
  const sheet = SpreadsheetApp.getActiveSheet()
  // parse header with types
  const header = JSON.parse(headerWithTypes)
  // get columns from header
  const columns = header.columns
  // filter enabled columns
    const enabledColumns = columns.filter(column => column.enabled)
  // iterate over enabled columns
  for (var i = 0; i < enabledColumns.length; i++) {
    const column = enabledColumns[i]
    const type = column.type
    // get number of rows in the sheet
    const rows = sheet.getMaxRows() - 1
    Logger.log('column: ' + (i+1) + ', rows: ' + rows + ', type: ' + type)
    // get range of column
    const range = sheet.getRange(2, i+1, rows, 1)
    if (type == 'number') {
      range.setNumberFormat('#,##0.00')
    } else if (type == 'integer') {
        range.setNumberFormat('#,##0')
    } else if (type == 'text' || type == 'varchar') {
      range.setNumberFormat('@')
    } else if (type == 'boolean') {
      range.setNumberFormat('@')
    } else if (type == 'date') {
      range.setNumberFormat('yyyy-mm-dd')
    } else if (type == 'timestamp') {
      range.setNumberFormat('yyyy-mm-dd hh:mm:ss')
    } else {
      Logger.log('Unknown type: ' + type)
    }
  }
}

function _displaySheet(name, content) {
  const data = Utilities.parseCsv(content);
  const sheet = _activateSheet(name)
  sheet.clear();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

function _getCurrentHeader() {
  // Access the sheet by name
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet) {
    // sheet doesn't exist yet, return empty header
    return ''
  }

  // get number of columns
  const columns = sheet.getLastColumn()

  // Get the data from the first row (assuming headers are in the first row)
  let firstRowData = ''
  if (columns > 0) {
    const values = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // values may not be an array, so we need to check
    if (values.length > 0) {
      firstRowData = values.join(',')
    }
  }

  // convert to URL safe string
  encodedString = encodeURIComponent(firstRowData)

  return encodedString
}

function postTable(baseUrl, token, whereClause, isInsert, isUpdate, isDelete, isExecute, isCommit) {
  // get table name from active sheet
  const tableName = _getActiveTableName(baseUrl, token)
  const header = _getCurrentHeader(tableName)
  const url = baseUrl + '/tables/' + tableName + '?style=sql&skiprows=1'
    + (isInsert ? '&insert=true' : '')
    + (isUpdate ? '&update=true' : '')
    + (isDelete ? '&delete=true' : '')
    + (isExecute ? '&execute=true' : '')
    + (isCommit ? '&commit=true' : '')
    + (header ? '&h=' + header : '')
    + (whereClause ? '&q=' + whereClause : '')
  const csv = _getActiveSheetAsCsv()
  const options = {
    "method": "post",
    "contentType": "text/csv",
    "payload": csv,
  };

  response = _makeHttpRequest(url, options, token)

  // store last posted table name so we can move back to this sheet after hiding the result sheet
  PropertiesService.getUserProperties().setProperty('last-posted-table-name', tableName)

  // show result sheet
  _displaySheet('result', response)
}

function _getActiveSheetAsCsv() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange().getValues();

  var csv = '';

  // Loop through the rows in the data range
  for (var i = 0; i < dataRange.length; i++) {
    var row = dataRange[i];

    // Loop through the cells in each row
    for (var j = 0; j < row.length; j++) {
      var cell = row[j];

      // Enclose cell value in double quotes if it contains a comma, newline, or double quote
      if (cell.toString().match(/["\n,]/)) {
        cell = '"' + cell.replace(/"/g, '""') + '"';
      }

      // Append the cell value to the CSV string
      csv += cell;

      // Add a comma to separate cells (except for the last cell in a row)
      if (j < row.length - 1) {
        csv += ',';
      }
    }

    // Add a newline character to separate rows (except for the last row)
    if (i < dataRange.length - 1) {
      csv += '\n';
    }
  }
  return csv
}

function viewResult() {
  _activateSheet("result")
}

function hideResult() {
  // delete the result sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const resultSheet = spreadsheet.getSheetByName('result')
  if (resultSheet) {
    // first active the last posted sheet
    const previousSheetName = PropertiesService.getUserProperties().getProperty('last-posted-table-name')
    if (previousSheetName) {
      const previousSheet = spreadsheet.getSheetByName(previousSheetName)
      if (previousSheet) {
        SpreadsheetApp.setActiveSheet(previousSheet);
      }
    }
    // the hide the result sheet (for some reason it flashes the result-1 sheet first)
    resultSheet.hideSheet()
  }
}

function _getActiveTableName(url, token) {
  // get active sheet's name
  name = SpreadsheetApp.getActiveSheet().getName()

  // check if it's a table
  filter = '^' + name + '$'
  tables = getTables(url, token, filter)

  if (tables.length == 0) {
    throw "Sheet '" + name + "' does not match any table."
  }

  return name
}

function _makeHttpRequest(url, options, token = null) {
    // Log the method, URL and optional payload
  Logger.log(options.method + ' ' + url)
  if (options.payload) {
    Logger.log(options.payload)
  }

  // Set default options
  options.method = options.method || 'GET'
  // Set the muteHttpExceptions option to true to prevent HTTP 4xx and 5xx errors from being treated as exceptions.
  options.muteHttpExceptions = true
  if (token) {
    options.headers = {
      Authorization: "Bearer " + token,
    }
  }

  // Make the HTTP request
  var response = UrlFetchApp.fetch(url, options);

  // Check the HTTP response code
  var statusCode = response.getResponseCode();

  // Check if the response code indicates an error (e.g., 4xx or 5xx)
  if (statusCode >= 400) {
    // Get the response body
    var responseBody = response.getContentText();

    // Throw an exception with the response body as the error message
    throw responseBody;
  }

  // If the response code is not an error, return the response
  return response;
}


