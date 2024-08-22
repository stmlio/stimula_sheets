function postMultiTable(baseUrl, token, sheetNames, whereClause, isInsert, isUpdate, isDelete, isExecute, isCommit, isDeduplicate) {
    // log sheet names
    Logger.log('Exporting sheets: ' + sheetNames.join(','))

    // get sheets
    const sheets = _getSheets(sheetNames)

    // convert source sheets to csv files
    const files = _exportSheets(sheets)

    // resolve table names
    const tables = _getTableNames(sheets)


    const url = baseUrl + '/tables?style=full&t=' + tables.join(',') + (isInsert ? '&insert=true' : '') + (isUpdate ? '&update=true' : '') + (isDelete ? '&delete=true' : '') + (isExecute ? '&execute=true' : '') + (isCommit ? '&commit=true' : '')
    // create multipart request
    const multipartData = createMultipartBody(files);
    const options = {
        method: 'POST',
        contentType: `multipart/form-data; boundary=${multipartData.boundary}`,
        payload: multipartData.body,
        muteHttpExceptions: true
    };

    // get source sheets to update results in
    const sourceSheets = _getSourceSheets(sheets)

    // clear formatting
    _clearMultiPostResult(sourceSheets)

    response = _makeHttpRequest(url, options, token)

    // parse response
    result = JSON.parse(response)

    // log rows
    Logger.log(result['rows'])

    // display line-by-line feedback in sheets
    _displayMultiPostFullReport(result['rows'], sourceSheets, isExecute)

    // return summary for the front-end to display
    return result['summary']
}

function _getSheets(sheetNames) {
    // get sheets by name
    return sheetNames.map(sheetName => {
        return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
    })
}



function _exportSheets(sheets) {
    // Prepare the files to be sent
    const files = sheets.map(sheet => {
        // log sheet name
        Logger.log('Exporting sheet: ' + sheet.getName())

        // check if it's an STML sheet, then cell A1 starts with '@'
        if (sheet.getRange(1, 1).getValue().startsWith('@')) {
            // get source sheet
            const sourceSheet = _getSourceSheet(sheet);
            // get STML map and list
            const stml = _getSheetAsStml(sheet);
            // replace header line with STML
            const contentWithHeader = _getSheetAsCsvWithStml(sourceSheet, stml);
            return {
                name: `${sourceSheet.getName()}.csv`, mimeType: 'text/csv', content: contentWithHeader
            };
        } else {
            const content = _getSheetAsCsv(sheet);
            return {
                name: `${sheet.getName()}.csv`, mimeType: 'text/csv', content: content
            };
        }
    });

    return files
}
function _getSourceSheets(sheets) {
    // get list of source sheets that contain the actual data to display results in
    const sourceSheets = sheets.map(sheet => {
        // check if it's an STML sheet, then cell A1 starts with '@'
        if (sheet.getRange(1, 1).getValue().startsWith('@')) {
            // get source sheet
            return _getSourceSheet(sheet);
        } else {
            return sheet
        }
    });

    return sourceSheets
}


function _getSourceSheet(sheet) {
    // verify cell A1 equals '@source'
    assert(sheet.getRange(1, 1).getValue() === '@source', 'Cell A1 in an STML sheet must be @source')
    // get source name from cell B1
    const sourceName = sheet.getRange(1, 2).getValue()
    // assert source name is not empty
    assert(sourceName, 'Source name in B2 cannot be empty')
    // get source sheet by source name
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceName)
    // assert source sheet exists
    assert(sourceSheet, 'Source sheet ' + sourceName + ' does not exist')
    return sourceSheet
}

function _getTableNames(sheets) {
    const tableNames = sheets.map(sheet => {
        const firstCell = sheet.getRange(1, 1).getValue()
        // check if the first cell starts with '@'
        if (firstCell.startsWith('@')) {
            return _getTargetTableName(sheet)
        } else {
            // only keep characters before the first non-alphanumeric character
            const tableName = sheet.getName().replace(/[^a-zA-Z0-9_].*/, '')
            // assert table name is not empty
            assert(tableName, 'Table name from sheet cannot be empty or start with a non-alphanumeric character')
            return sheet.getName()
        }
    })
    return tableNames
}

function _getTargetTableName(sheet) {
    // verify cell A2 equals '@target'
    assert(sheet.getRange(2, 1).getValue() === '@target', 'Cell A2 in an STML sheet must be @target')
    // get table name from cell B2
    const tableName = sheet.getRange(2, 2).getValue()
    // assert table name is not empty
    assert(tableName, 'Table name in B2 cannot be empty')
    return tableName;
}


function _getSheetAsStml(sheet) {
    //     get STML as map from sheet. Result is a tuple with two elements:
    //     - map with source column names as keys and target column names as values.
    //     - array of additional columns: target column names that do not have a source column name.

    // find row with 'source_column' in A column
    const headerRow = _findRow(sheet, 'source_column')
    // assert header row is found
    assert(headerRow >= 0, 'Missing header row with source_column in STML sheet')
    // assert column B is 'target_column'
    let targetColumnKey = sheet.getRange(headerRow, 2).getValue();
    assert(targetColumnKey === 'target_column', 'Missing target_column in STML sheet, found: ' + targetColumnKey)
    // get modifier column names
    const modifiers = _getModifierColumnNames(sheet, headerRow);

    const columnMap = {}
    const additionalColumns = []
    // iterate over rows
    for (let i = headerRow + 1; i <= sheet.getLastRow(); i++) {
        const sourceColumn = sheet.getRange(i, 1).getValue()
        const targetColumn = _createTargetColumnName(sheet, i, modifiers)

        if (sourceColumn && targetColumn) {
            //     if source and target are not empty, add to map
            columnMap[sourceColumn] = targetColumn
        } else if (!sourceColumn && targetColumn) {
            // if source column is empty, add target to additional columns
            additionalColumns.push(targetColumn)
        }

    }
    return [columnMap, additionalColumns]
}

function _findRow(sheet, text) {
    // get all values in column A
    const values = sheet.getRange('A:A').getValues();
    // find row with text
    for (let i = 0; i < values.length; i++) {
        if (values[i][0] === text) {
            // sheets are 1-indexed
            return i + 1;
        }
    }
    return -1
}


function _getModifierColumnNames(sheet, headerRow) {
    // get list of additional non-empty headers to treat as modifiers from the first row
    const modifiers = sheet.getRange(headerRow, 3, 1, sheet.getLastColumn() - 2).getValues()[0]
    // supported modifiers
    const knownModifiers = ['unique', 'skip', 'default-value', 'exp']
    // list unknown modifiers
    const unknownModifiers = modifiers.filter(modifier => modifier && !knownModifiers.includes(modifier))
    // assert no unknown modifiers
    assert(unknownModifiers.length === 0, 'Unknown modifiers: ' + unknownModifiers.join(', '))
    return modifiers;
}

function _createTargetColumnName(sheet, rowIndex, modifiers) {
    // get base target column name
    const targetColumn = sheet.getRange(rowIndex, 2).getValue()
    const modifierList = []
    // iterate over modifiers
    for (let i = 0; i < modifiers.length; i++) {
        // if modifier name is not empty
        if (modifiers[i]) {
            // get modifier value
            const modifierValue = sheet.getRange(rowIndex, i + 3).getValue()
            // if modifier value is not empty
            if (modifierValue !== '') {
                // add to list
                modifierList.push(modifiers[i] + '=' + modifierValue)
            }
        }
    }
    // append modifiers to target column name. STML supports multiple sets of modifiers
    // check if there are any modifiers
    if (modifierList.length > 0) {
        return targetColumn + '[' + modifierList.join(': ') + ']'
    } else {
        return targetColumn
    }

}


function _getSheetAsCsvWithStml(sheet, stml) {
    // get header values
    const headerNames = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    // get map from STML
    const columnMap = stml[0]
    // get keys from columnMap that don't have a matching header value
    const missingColumns = Object.keys(columnMap).filter(key => !headerNames.includes(key))
    // assert no missing columns
    assert(missingColumns.length === 0, 'The following column exist in STML but not in the source file: ' + missingColumns.join(', '))
    // get additional columns from STML
    const additionalColumns = stml[1]
    // substitute header values with target column names, empty if not found
    const substitutedHeaderNames = headerNames.map(header => columnMap[header] || '')
    // append additional columns
    substitutedHeaderNames.push(...additionalColumns)
    // get content as csv
    const headerLine = _convertDataRangeToCsv([substitutedHeaderNames])
    // get rows 2 and below
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()
    // pad with empty values for additional columns
    const paddedRows = rows.map(row => row.concat(Array(additionalColumns.length).fill('')))
    // convert rows to csv
    const rowsCsv = _convertDataRangeToCsv(paddedRows)
    // return header line and rows
    return headerLine + '\n' + rowsCsv
}

function _clearMultiPostResult(sheets) {
    sheets.forEach(sheet => {
        _clearPostResult(sheet)
    })
}

function _clearPostResult(sheet) {
    // remove background color of complete sheet
    if (sheet.getLastRow() > 0) {
        sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground(null)
        // clear notes of first column
        sheet.getRange(1, 1, sheet.getMaxRows(), 1).clearNote()
    }
}


function _displayMultiPostFullReport(rows, sheets, isExecute) {
    //     iterate over sheets
    //     iterate rows and log context attribute
    rows.forEach(row => {
        Logger.log('Context: ' + row.context)
    })

    sheets.forEach(sheet => {
        // log sheetname
        Logger.log('Displaying sheet: ' + sheet.getName())
        // filter rows for this sheet
        const sheetRows = rows.filter(row => row.context == (sheet.getName() + '.csv'))
        // logger row count
        Logger.log('Rows: ' + sheetRows.length)
        // display rows
        _displayPostFullReport(sheetRows, sheet, isExecute)
    })

}

function _displayPostFullReport(rows, sheet, isExecute) {
//   find and activate sheet based on context
    if (!sheet) {
        Logger.log('Sheet to display updates not found: ' + context)
        return
    }

    // iterate rows
    for (var i = 0; i < rows.length; i++) {
        const row = rows[i]

        // convert from string to integer

        const lineNumber = parseInt(row.line_number)
        // undefined also means success
        const success = row.success || row.success === undefined

        // log line and success
        Logger.log('Line: ' + lineNumber + ', Success: ' + success)

        // if line is NaN, skip
        // TODO, figure out how to display deletes
        if (isNaN(lineNumber)) {
            continue
        }
        // get line number
        // select full row
        const range = sheet.getRange(lineNumber + 2, 1, 1, sheet.getLastColumn())

        // set background color based on success
        if (success) {
            if (isExecute) {
                // bright green for execute
                range.setBackground('#44FF44')
            } else {
                // light green for not execute
                range.setBackground('#AAFFAA')
            }
        } else {
            range.setBackground('#FF4444')
        }

        //   select first column of row
        const cell = sheet.getRange(lineNumber + 2, 1)

        // get error
        const error = row.error
        // get query and parameters
        const query = row.query
        const params = row.params
        // format error
        errorMessage = error ? 'Error: ' + error + '\n\n' : ''
        // add query
        errorMessage += query + '\n\n'
        // add parameters as json string
        errorMessage += JSON.stringify(params, null, 2)

        // set error message as note
        cell.setNote(errorMessage)
    }
}


function assert(condition, message) {
    if (!condition) {
        throw new Error(message || "Assertion failed");
    }
}