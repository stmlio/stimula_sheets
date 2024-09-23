var startTime;

function tick(message) {
    if (!startTime) {
        Logger.log(message + " - Timer started.");
        startTime = new Date();
        return;
    }

    var currentTime = new Date();
    var elapsedTime = (currentTime - startTime)
    startTime = currentTime;
    Logger.log(message + " - Elapsed time: " + elapsedTime + " ms");
}

function getStmlMappings() {
    // skip if this is already an STML sheet
    if (SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, 1).getValue().startsWith('@')) {
        return
    }

    // get current active sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

    // get header line values
    const values = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()
    const header = _convertDataRangeToCsv(values)

    const options = {
        method: 'post',
        contentType: 'text/csv',
        payload: header
    }

    response = _makeHttpRequest('https://api.stml.io/1.0/mappings', options)
    Logger.log('mappings: ' + response)

    return JSON.parse(response)
}

function createStmlSheet(stmlTemplateId, sampleDataRows) {
    tick('createStmlSheet')
    // get current active sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    tick('Got active sheet')

    // check that this is not an STML sheet
    assert(!sheet.getRange(1, 1).getValue().startsWith('@'), 'This is already an STML sheet')
    tick('Checked STML sheet')

    // get mapping from api.stml.io
    const mapping = getStmlMapping(stmlTemplateId)
    tick('Got mapping')

    // create blank STML sheet and make it active
    const stmlSheet = createBlankStmlSheet(sheet);
    tick('Created blank STML sheet')

    // fill table with data
    fillStmlValues(stmlSheet, sheet, mapping, sampleDataRows)
    tick('Filled STML sheet')

    // set auto resize columns
    stmlSheet.autoResizeColumns(1, 2);
}

function getStmlMapping(id) {

    if (!id) {
        Logger.log('No STML template provided, creating blank STML sheet')
        return {}
    }
    try {
        // retrieve mapping from api.stml.io
        response = _makeHttpRequest('https://api.stml.io/1.0/mappings/' + id)

        return JSON.parse(response)
    } catch (e) {
        Logger.log('Error fetching mapping, ID: ' + id + ', error: ' + e)
        return {}
    }
}

function createBlankStmlSheet(sheet) {
    const sheetName = sheet.getName()
    // start with empty suffix
    let suffix = ''
    // iterate over suffixes until free sheetname found
    while (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName + suffix + '.stml')) {
        // increment suffix
        suffix = (parseInt(suffix) || 0) + 1
    }
    // create new sheet with suffix
    const stmlSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName + suffix + '.stml')

    // freeze rows
    stmlSheet.setFrozenRows(4)

    return stmlSheet;
}

function fillStmlValues(stmlSheet, sheet, mapping, sampleDataRows) {

    const numberOfColumns = sheet.getLastColumn()

    // get additional target columns without source column
    const additionalColumns = mapping['columns'].filter(column => !column['source']).map(column => column['target'])

    // get number of additional columns
    const additionalColumnsCount = additionalColumns.length

    // initialize table to hold the data
    const values = Array.from({length: numberOfColumns + additionalColumnsCount + 4}, () => Array.from({length: sampleDataRows + 4}, () => ''));

    // meta data
    values[0][0] = '@source'
    values[0][1] = sheet.getName()
    values[1][0] = '@target'
    values[1][1] = mapping['target'] || ''
    values[3][0] = 'source_column'
    values[3][1] = 'target_column'

    // read source data into array
    const sourceData = sheet.getRange(1, 1, 1 + sampleDataRows, numberOfColumns).getValues()

    // map source columns to target columns
    if (mapping && mapping['columns']) {
        column_map = mapping['columns'].reduce((acc, column) => {
            acc[column['source']] = column['target'];
            return acc
        }, {})

    } else {
        column_map = {}
    }

    // fill source, target and sample data columns
    for (let i = 0; i < numberOfColumns; i++) {
        values[4 + i][0] = sourceData[0][i]
        values[4 + i][1] = column_map[sourceData[0][i]] || ''

        // copy sample data
        for (let j = 0; j < sampleDataRows; j++) {
            values[4 + i][4 + j] = sourceData[1 + j][i]
        }
    }

    // fill additional columns at the bottom of the target columns
    additionalColumns.forEach((column, index) => {
        values[4 + numberOfColumns + index][1] = column
    })

    // copy target column names to STML sheet, in column B, starting at B5
    stmlSheet.getRange(1, 1, values.length, values[0].length).setValues(values)
}

function postMultiTable(baseUrl, token, sheetNames, whereClause, isInsert, isUpdate, isDelete, isExecute, isCommit) {
    // log sheet names
    Logger.log('Exporting sheets: ' + sheetNames.join(','))

    tick('Start')

    // get sheets
    const sheets = _getSheets(sheetNames)

    tick('Got sheets')

    // convert source sheets to csv files
    const files = _exportSheets(sheets)

    tick('Exported sheets')

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

    Logger.log('Clear source sheets: ' + sourceSheets)

    // clear formatting
    _clearMultiPostResult(sourceSheets)

    response = _makeHttpRequest(url, options, token)

    tick('Posted tables')

    // parse response
    result = JSON.parse(response)

    // log max 3 rows
    Logger.log(result['rows'].slice(0, 3))

    tick('Parsed response')

    // display line-by-line feedback in sheets
    _displayMultiPostFullReport(result['rows'], sourceSheets, isExecute)

    tick('Displayed results')

    // return summary for the front-end to display
    return result['summary']
}

function _getSheets(sheetNames) {
    // get sheets by name
    return sheetNames.map(sheetName => {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
        // assert sheet exists
        assert(sheet, 'Sheet ' + sheetName + ' does not exist')
        return sheet
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

            tick('Got source sheet')

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

    // Retrieve all the data from the sheet at once
    const allData = sheet.getDataRange().getValues();

    tick('Got all data')

    // find row with 'source_column' in A column
    const headerRow = _findRow(allData, 'source_column')
    assert(headerRow >= 0, 'Missing header row with source_column in STML sheet')

    // assert column B is 'target_column'
    const targetColumnKey = allData[headerRow][1];
    assert(targetColumnKey === 'target_column', 'Missing target_column header in STML sheet, found: ' + targetColumnKey)

    // get modifier column names
    const modifiers = _getModifierColumnNames(allData, headerRow);
    tick('Got modifiers')

    const columnMap = {}
    const additionalColumns = []

    // Iterate over rows, starting from the row after the header
    for (let i = headerRow + 1; i < allData.length; i++) {
        const sourceColumn = allData[i][0]
        const targetColumn = _createTargetColumnName(allData, i, modifiers)

        if (sourceColumn && targetColumn) {
            // if source and target are not empty, add to map
            columnMap[sourceColumn] = targetColumn
        } else if (!sourceColumn && targetColumn) {
            // if source column is empty, add target to additional columns
            additionalColumns.push(targetColumn)
        }
    }
    return [columnMap, additionalColumns]
}

function _findRow(allData, text) {
    // find row with text in first column
    for (let i = 0; i < allData.length; i++) {
        if (allData[i][0] === text) {
            // return header row number (0-indexed)
            return i;
        }
    }
    return -1
}


function _getModifierColumnNames(allData, headerRow) {
    // get list of additional non-empty headers to treat as modifiers from the first row
    const modifiers = allData[headerRow].slice(2)
    // supported modifiers
    const knownModifiers = ['unique', 'skip', 'default-value', 'exp', 'deduplicate', 'table', 'name', 'qualifier', 'key']
    // list unknown modifiers
    const unknownModifiers = modifiers.filter(modifier => modifier && !knownModifiers.includes(modifier))
    // assert no unknown modifiers
    assert(unknownModifiers.length === 0, 'Unknown modifiers: ' + unknownModifiers.join(', '))
    return modifiers;
}

function _createTargetColumnName(allData, rowIndex, modifiers) {
    // Get base target column name from column B (index 1) in the current row
    const targetColumn = allData[rowIndex][1];  // Column B (target_column)
    const modifierList = []

    // iterate over modifiers
    for (let i = 0; i < modifiers.length; i++) {
        // if modifier name is not empty
        if (modifiers[i]) {
            // Get modifier value from the respective column in the current row (starting from column C)
            const modifierValue = allData[rowIndex][i + 2];  // Column C onwards (index 2 and beyond)

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

    tick('Got rows')

    // pad with empty values for additional columns
    const paddedRows = rows.map(row => row.concat(Array(additionalColumns.length).fill('')))

    // convert rows to csv
    const rowsCsv = _convertDataRangeToCsv(paddedRows)

    tick('Converted rows to csv')

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
    sheets.forEach(sheet => {
        // log sheetname
        Logger.log('Displaying sheet: ' + sheet.getName())
        // filter rows for this sheet
        const sheetName = sheet.getName() + '.csv';
        const sheetRows = rows.filter(row => row.context == sheetName);

        tick('Filtered rows')

        // logger row count
        Logger.log('Rows: ' + sheetRows.length)
        // display background color in rows
        _displayBackgroundColor(sheetRows, sheet, isExecute)

        tick('Displayed background color')

        // display notes in first column
        _displayNotes(sheetRows, sheet)

        tick('Displayed notes')

    })

}

function _displayBackgroundColor(rows, sheet, isExecute) {

    // get range of successful rows. Undefined also means success.
    let successRange = rows.filter(row => row.success === undefined || row.success).map(r => r.line_number).filter(l => !isNaN(l)).map(l => parseInt(l) + 2).map(l => l + ':' + l)

    // get range of failed rows
    let failedRange = rows.filter(row => row.success !== undefined && !row.success).map(r => r.line_number).filter(l => !isNaN(l)).map(l => parseInt(l) + 2).map(l => l + ':' + l)

    // set background color. For some odd reason a rangelist must not be empty.
    if (successRange.length > 0) {
        // bright green for execute, light green for evaluate
        const color = isExecute ? '#44FF44' : '#AAFFAA'
        sheet.getRangeList(successRange).setBackground(color)
    }

    // red for failed
    if (failedRange.length > 0) {
        sheet.getRangeList(failedRange).setBackground('#FF4444')
    }
}

function _displayNotes(rows, sheet) {

    // get single range for column A. We can't use a range list to set notes.
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)

    // create an empty array with the same length as range
    const notesArray = Array.from({length: sheet.getLastRow() - 1}, (v, i) => '')

    // iterate rows to set notes in array
    rows.filter(row => !isNaN(row.line_number)).forEach(row => {
        // parse line number to get 0-indexed line number
        const lineNumber = parseInt(row.line_number)

        // format error
        notesArray[lineNumber] = (row.error ? 'Error: ' + row.error + '\n\n' : '') + row.query + '\n\n' + JSON.stringify(row.params, null, 2)
    })

    // set all notes in sheet at once
    range.setNotes(notesArray.map(note => [note]))
}


function assert(condition, message) {
    if (!condition) {
        throw new Error(message || "Assertion failed");
    }
}