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

function getStmlMappingsForActiveSheet() {
    // get current active sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

    // skip if this is not an STML sheet
    if (!sheet.getRange(1, 1).getValue().startsWith('@')) {
        return
    }

    // get first column
    const stmLData = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues()

    // get mappings from api.stml.io
    return getStmlMappings(stmLData)
}

function getStmlMappings(stmlData) {
    // convert first column into a csv string
    const header = stmlData.map(row => row[0]).join(', ')

    const options = {
        method: 'post',
        contentType: 'text/csv',
        payload: header
    }

    response = _makeHttpRequest('https://api.stml.io/1.0/mappings', options)
    Logger.log('mappings: ' + response)

    return JSON.parse(response)
}

function createStmlSheet(sampleDataRows) {
    tick('createStmlSheet')
    // get current active sheet
    let sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    tick('Got active sheet')

    // check if this is an STML sheet
    if (sourceSheet.getRange(1, 1).getValue().startsWith('@')) {
        // get sheet from source name
        const sourceName = sourceSheet.getRange(1, 2).getValue()

        // change source sheet to the one with the source name
        sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceName)
    }
    tick('Checked STML sheet')

    // create blank STML sheet and make it active
    const stmlSheet = createBlankStmlSheet(sourceSheet);
    tick('Created blank STML sheet')

    // fill table with data
    stmlData = fillStmlValues(stmlSheet, sourceSheet, sampleDataRows)
    tick('Filled STML sheet')

    // set auto resize columns
    stmlSheet.autoResizeColumns(1, 2);

    // get mappings from api.stml.io
    return getStmlMappings(stmlData)
}


function updateStmlSheet(stmlTemplateId) {
    tick('createStmlSheet')
    // get current active sheet
    const stmlSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    tick('Got active sheet')

    // assert that this is an STML sheet
    assert(stmlSheet.getRange(1, 1).getValue().startsWith('@'), 'This is not STML sheet. First select an STML sheet.')
    tick('Checked STML sheet')

    // assert that STML template ID is provided
    assert(stmlTemplateId, 'STML template ID is required to update existing STML sheet')

    // get mapping from api.stml.io
    const mapping = getStmlMapping(stmlTemplateId)
    tick('Got mapping')

    // assert that mapping is not empty
    assert(Object.keys(mapping).length > 0, 'No mapping found for ID: ' + stmlTemplateId)

    // update STML with mapping
    updateStmlValues(stmlSheet, mapping)

    // set auto resize columns
    stmlSheet.autoResizeColumns(1, 2);

    // we may have renamed the sheet already, so skip if it already looks fine
    if (!stmlSheet.getName().startsWith(mapping['short'])) {
        // create free sheet name
        sheetName = createSheetName(mapping['short'], 'stml')
        // set sheet name
        stmlSheet.setName(sheetName)
    }
}

function createBlankStmlSheet(sheet) {
    // create new sheet
    const stmlSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet()

    // freeze rows
    stmlSheet.setFrozenRows(4)

    // activate sheet
    SpreadsheetApp.setActiveSheet(stmlSheet)

    // find free sheet name based on source sheet name
    const sheetName = createSheetName(sheet.getName(), 'stml')

    // set sheet name
    stmlSheet.setName(sheetName)

    return stmlSheet;
}

function fillStmlValues(stmlSheet, sourceSheet, sampleDataRows) {

    const numberOfColumns = sourceSheet.getLastColumn()

    // initialize table to hold the data
    const values = Array.from({length: numberOfColumns + 4}, () => Array.from({length: sampleDataRows + 2}, () => ''));

    // meta data
    values[0][0] = '@source'
    values[0][1] = sourceSheet.getName()
    values[1][0] = '@target'
    values[3][0] = 'source_column'
    values[3][1] = 'target_column'

    // read source data into array
    const sourceData = sourceSheet.getRange(1, 1, 1 + sampleDataRows, numberOfColumns).getValues()

    // fill source, target and sample data columns
    for (let i = 0; i < numberOfColumns; i++) {
        values[4 + i][0] = sourceData[0][i]

        // copy sample data
        for (let j = 0; j < sampleDataRows; j++) {
            values[4 + i][2 + j] = sourceData[1 + j][i]
        }
    }

    // copy all data to STML sheet
    stmlSheet.getRange(1, 1, values.length, values[0].length).setValues(values)

    // set text wrapping to clip for all sample data columns
    if (sampleDataRows > 0) {
        stmlSheet.getRange(5, 3, numberOfColumns, sampleDataRows).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }

    // return values, so we can use them to retrieve mappings
    return values
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

function updateStmlValues(stmlSheet, mapping) {

    // read source and target columns from STML into array
    const stmlData = stmlSheet.getRange(1, 1, stmlSheet.getLastRow(), 2).getValues()

    // find last row with data in column A, subtract 4 to get the last column index, add one to get the number of columns
    const numberOfColumns = Math.max(...Array.from(stmlData, (e, i) => e[0] ? i : 0)) - 4 + 1

    // get additional target columns without source column
    const additionalColumns = mapping['columns'].filter(column => !column['source']).map(column => column['target'])

    // get number of additional columns
    const additionalColumnsCount = additionalColumns.length

    // take sufficient length to overwrite existing values
    const max_length = Math.max(stmlData.length, 4 + numberOfColumns + additionalColumnsCount)

    // initialize table to hold the data
    const values = Array.from({length: max_length}, () => Array.from({length: 1}, () => ''));

    // copy and set source and target etc.
    values[0][0] = stmlData[0][1]
    values[1][0] = mapping['target'] || ''
    values[3][0] = 'target_column'

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
        // map source column to target column
        values[4 + i][0] = column_map[stmlData[4+i][0]] || ''
    }

    // fill additional columns at the bottom of the target columns
    additionalColumns.forEach((column, index) => {
        values[4 + numberOfColumns + index][0] = column
    })

    // copy target column names to STML sheet, in column B, starting at B1
    stmlSheet.getRange(1, 2, values.length, 1).setValues(values)
}

function createSheetName(name, extension) {
    // start with empty suffix
    let suffix = ''
    // get active spreadsheet
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // iterate over suffixes until free sheetname found
    while (activeSpreadsheet.getSheetByName(name + suffix + '.' + extension)) {
        // increment suffix
        suffix = (parseInt(suffix) || 0) + 1
    }
    return name + suffix + '.' + extension
}

function postMultiTable(baseUrl, token, sheetNames, whereClause, isInsert, isUpdate, isDelete, isExecute, isCommit) {
    // raise error if no sheets are selected
    assert(sheetNames.length > 0, 'No sheets selected')

    // log sheet names
    Logger.log('Exporting sheets: ' + sheetNames.join(','))
    tick('Start')

    // get sheets
    const sheets = _getSheets(sheetNames)
    tick('Got sheets')
    Logger.log('Got sheets: ' + sheets.map(sheet => sheet.getName()))

    // find STML sheets
    const stmlSheets = _getStmlSheets(sheets)
    Logger.log('STML sheets: ' + stmlSheets.map(sheet => sheet.getName()))

    // map to source sheets
    const sourceSheets = _getSourceSheets(stmlSheets)
    Logger.log('Source sheets: ' + sourceSheets.map(sheet => sheet.getName()))

    // map to substitution sheets
    const substitutionSheets = _getSubstitutionSheets(stmlSheets)
    Logger.log('Substitution sheets: ' + substitutionSheets.map(sheet => sheet.getName()))

    // only support zero or one substitution sheet. List sheets by name.
    assert(substitutionSheets.length <= 1, 'Only zero or one substitution sheet is supported. Found: ' + substitutionSheets.map(sheet => sheet.getName()))

    // get csv files by removing source sheets from original list
    const csvSheets = _getCsvSheet(sheets, stmlSheets, sourceSheets)
    Logger.log('CSV sheets: ' + csvSheets.map(sheet => sheet.getName()))

    // create mime objects for stml sheets
    const stmlFiles = _exportStmlSheets(stmlSheets, sourceSheets)
    tick('Exported sheets')
    Logger.log('Exported STML sheets: ' + stmlFiles.map(file => file.name))

    // create mime objects for csv sheets
    const csvFiles = _exportCsvSheets(csvSheets)
    Logger.log('Exported CSV sheets: ' + csvFiles.map(file => file.name))

    // create mime object for substitution sheets (zero or one)
    const substitutionFiles = _exportSubstitutionsSheets(substitutionSheets)
    Logger.log('Exported substitution sheets: ' + substitutionFiles.map(file => file.name))

    // resolve table names for all sheets
    const tables = _getTableNames(stmlSheets.concat(csvSheets))
    Logger.log('Tables: ' + tables)

    const url = baseUrl + '/tables?style=full&t=' + tables.join(',') + (isInsert ? '&insert=true' : '') + (isUpdate ? '&update=true' : '') + (isDelete ? '&delete=true' : '') + (isExecute ? '&execute=true' : '') + (isCommit ? '&commit=true' : '')
    // create multipart request
    const multipartData = createMultipartBody(stmlFiles.concat(csvFiles).concat(substitutionFiles));
    const options = {
        method: 'POST',
        contentType: `multipart/form-data; boundary=${multipartData.boundary}`,
        payload: multipartData.body,
        muteHttpExceptions: true
    };

    // get source and CSV sheets to update results in
    const sheetsToFormat = sourceSheets.concat(csvSheets)
    Logger.log('Sheets to format: ' + sheetsToFormat.map(sheet => sheet.getName()))

    // clear formatting
    _clearMultiPostResult(sheetsToFormat)

    response = _makeHttpRequest(url, options, token)

    tick('Posted tables')

    // parse response
    result = JSON.parse(response)

    // log max 3 rows
    Logger.log(result['rows'].slice(0, 3))
    tick('Parsed response')

    // display line-by-line feedback in sheets
    _displayMultiPostFullReport(result, sourceSheets)
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

function _getStmlSheets(sheets) {
    return sheets.filter(sheet => sheet.getRange(1, 1).getValue().startsWith('@'))
}

function _getCsvSheet(sheets, stmlSheets, sourceSheets) {
    // return sheets that are not STML sheets and not source sheets. Compare sheets by ID, not by object reference.
    return sheets.filter(sheet =>
        !stmlSheets.some(s => s.id === sheet.id) &&
        !sourceSheets.some(s => s.id === sheet.id)
    );}
function _exportStmlSheets(stmlSheets, sourceSheets) {

    // get mime object for STML sheets
    return stmlSheets.map((stmlSheet, index)=> {
        // get source sheet
        const sourceSheet = sourceSheets[index];

        // get STML map and list
        const stml = _getSheetAsStml(stmlSheet);

        // replace header line with STML
        const contentWithHeader = _getSheetAsCsvWithStml(sourceSheet, stml);

        // return mime object
        return {
            name: `${sourceSheet.getName()}.csv`, mimeType: 'text/csv', content: contentWithHeader
        }
    })
}
function _exportCsvSheets(csvSheets) {
    // get mime object for csv sheets
    return csvSheets.map(sheet => {
        const content = _getSheetAsCsv(sheet);
        return {
            name: `${sheet.getName()}.csv`, mimeType: 'text/csv', content: content
        };
    })
}
function _exportSubstitutionsSheets(csvSheets) {
    // can only send zero or one substitution sheets, for now
    assert (csvSheets.length <= 1, 'Only zero or one substitution sheet is supported. Found: ' + csvSheets.map(sheet => sheet.getName()))
    // get mime object for csv sheets
    return csvSheets.map(sheet => {
        const content = _getSheetAsCsv(sheet);
        return {
            // name must be substitutions.csv
            name: `substitutions.csv`, mimeType: 'text/csv', content: content
        };
    })
}

function _getSourceSheets(stmlSheets) {
    // get list of source sheets that contain the actual data to display results in
    return  stmlSheets.map(sheet => _getSourceSheet(sheet))
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

function _getSubstitutionSheets(stmlSheets) {
    // return the list of non-empty substitution sheets
    return stmlSheets.map(sheet => _getSubstitutionSheet(sheet)).filter(sheet => sheet)
}

function _getSubstitutionSheet(sheet) {
    // if A3 equals '@substitutions', and B3 is not empty, then return the sheet with that name
    if (sheet.getRange(3, 1).getValue() === '@substitutions') {
        const substitutionName = sheet.getRange(3, 2).getValue()
        if (substitutionName) {
            // find the sheet
            const substitutionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(substitutionName)
            // assert it exists
            assert(substitutionSheet, 'Source sheet ' + substitutionName + ' does not exist')
            return substitutionSheet
        }
    }
    // no substitutions specified, return null
    return null
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


function _displayMultiPostFullReport(report, sheets) {
    // test or committed?
    const isCommit = report['summary']['commit']
    const rows = report['rows']

    // the same sheet may appear multiple times, so we need to remove duplicates. First create a map from sheet ID to sheet.
    const sheetMap = sheets.reduce((acc, sheet) => {
        acc[sheet.id] = sheet
        return acc
    }, {})

    // then get the map's values
    const uniqueSheets = Object.values(sheetMap)

    //     iterate over sheets
    uniqueSheets.forEach(sheet => {
        // log sheetname
        Logger.log('Displaying sheet: ' + sheet.getName())
        // filter rows for this sheet
        const sheetName = sheet.getName() + '.csv';
        const sheetRows = rows.filter(row => row.context == sheetName);

        tick('Filtered rows')

        // logger row count
        Logger.log('Rows: ' + sheetRows.length)
        // display background color in rows
        _displayBackgroundColor(sheetRows, sheet, isCommit)

        tick('Displayed background color')

        // display notes in first column
        _displayNotes(sheetRows, sheet)

        tick('Displayed notes')
    })

}

function _displayBackgroundColor(rows, sheet, isCommit) {

    // get range of successful rows. Undefined also means success.
    let successRange = rows.filter(row => row.success === undefined || row.success).map(r => r.line_number).filter(l => !isNaN(l)).map(l => parseInt(l) + 2).map(l => l + ':' + l)

    // get range of failed rows
    let failedRange = rows.filter(row => row.success !== undefined && !row.success).map(r => r.line_number).filter(l => !isNaN(l)).map(l => parseInt(l) + 2).map(l => l + ':' + l)

    // same row may appear multiple times, so we need to remove duplicates. Also remove failed rows from success range
    successRange = Array.from(new Set(successRange.filter(l => !failedRange.includes(l))) )
    failedRange = Array.from(new Set(failedRange))

    // set background color. For some odd reason a rangelist must not be empty.
    if (successRange.length > 0) {
        // bright green for execute, light green for evaluate
        const color = isCommit ? '#9F9' : '#FFC'
        sheet.getRangeList(successRange).setBackground(color)
    }

    // red for failed
    if (failedRange.length > 0) {
        sheet.getRangeList(failedRange).setBackground('#F99')
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
        const newLines = notesArray[lineNumber] ? '\n\n' : ''
        const error = row.error ? 'Error: ' + row.error + '\n\n' : ''
        const query = row.query ? 'Query: ' + row.query + '\n\n': ''
        const params = row.params ? 'Params: ' + JSON.stringify(row.params, null, 2) : ''

        // Append to existing note. There may be multiple notes for the same line.
        notesArray[lineNumber] += newLines + error + query + params
    })

    // set all notes in sheet at once
    range.setNotes(notesArray.map(note => [note]))
}


function assert(condition, message) {
    if (!condition) {
        throw new Error(message || "Assertion failed");
    }
}