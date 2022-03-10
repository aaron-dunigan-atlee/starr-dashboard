/**
 * updateRows
 * Version 1.3: added .upsert option 
 *              support for compound primary key (passed as array)
 */
/**
 * Update rows data for existing rows, with a lock to avoid collisions or changes to the sheet while we are processing.
 * @param {SpreadsheetApp.Sheet} sheet 
 * @param {Object[]} rows Rows data
 * @param {Object} setOptions Options to be passed to setRowsData
 * @param {string} primaryKey  Unique key to use to identify the row to update.  If not present, we will use the .sheetRow property in the rows data.
 * @param {Object} updateOptions .onlyPresentColumns {boolean} If true, only set columns that are passed as properties.  i.e. preserve values for all columns not present.
 *                               .upsert {boolean} If true, insert rows if no row is found with the primary key value.
 */
function updateRows(sheet, rows, setOptions, primaryKey, passedLock, updateOptions)
{
  // Default options
  setOptions = setOptions || {}
  // setOptions.preserveArrayFormulas = true;

  updateOptions = updateOptions || {}

  // If one row is passed, make it an array
  if (!(rows instanceof Array)) rows = [rows]

  if (rows.length === 0)
  {
    console.warn("updateRows: No rows to update")
    return []
  }

  // Set a lock
  var lock = passedLock;
  if (!lock)
  {
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
  }

  // Update rows via primaryKey
  var updatedRows
  if (primaryKey)
  {
    // Transfer set options to get options, but don't include start/end headers b/c we don't know if they include the primary key column
    var getOptions = { getMetadata: true }
    Object.assign(getOptions, setOptions)
    getOptions.startHeader = null;
    getOptions.endHeader = null;
    var dataByKey = hashObjects(
      getRowsData(sheet, null, getOptions),
      primaryKey
    )
    if (primaryKey instanceof Array)
    {
      var key = primaryKey.join('.')
      updatedRows = rows.map(function (row)
      {
        if (!primaryKey.every(x => { return row[x] }))
        {
          notifyError("Unable to update row: primary key " + key + " not found in this row: " + JSON.stringify(row) + " on sheet " + sheet.getName() + " of " + sheet.getParent().getName())
          return null
        }
        var rowKey = primaryKey.map(x => { return row[x] }).join('.')
        var rowToUpdate = dataByKey[rowKey]
        if (!rowToUpdate)
        {
          if (updateOptions.upsert)
          {
            var insertOptions = {}
            Object.assign(insertOptions, setOptions)
            insertOptions.writeMethod = 'append';
            setRowsData(sheet, [row], insertOptions)
            return row
          }
          else
          {
            notifyError("Unable to update row: no row found with " + key + " = " + rowKey + " on sheet " + sheet.getName() + " of " + sheet.getParent().getName())
            return null
          }

        }
        // Write only indicated columns if instructed
        if (updateOptions.onlyPresentColumns)
        {
          for (var prop in rowToUpdate)
          {
            if (!row[prop])
            {
              row[prop] = rowToUpdate[prop]
            }
          }
        }
        setOptions.firstRowIndex = rowToUpdate.sheetRow
        setRowsData(
          sheet,
          [row],
          setOptions
        )
        if (setOptions.log) console.log("Updated data for %s=%s on row %s", key, rowKey, rowToUpdate.sheetRow)
        return row
      })
    }
    else
    {
      updatedRows = rows.map(function (row)
      {
        if (!row[primaryKey])
        {
          notifyError("Unable to update row: primary key " + primaryKey + " not found in this row: " + JSON.stringify(row) + " on sheet " + sheet.getName() + " of " + sheet.getParent().getName())
          return null
        }
        var rowToUpdate = dataByKey[row[primaryKey]]
        if (!rowToUpdate)
        {
          if (updateOptions.upsert)
          {
            var insertOptions = {}
            Object.assign(insertOptions, setOptions)
            insertOptions.writeMethod = 'append';
            setRowsData(sheet, [row], insertOptions)
            return row
          }
          else
          {
            notifyError("Unable to update row: no row found with " + primaryKey + " = " + row[primaryKey] + " on sheet " + sheet.getName() + " of " + sheet.getParent().getName())
            return null
          }
        }
        // Write only indicated columns if instructed
        if (updateOptions.onlyPresentColumns)
        {
          for (var prop in rowToUpdate)
          {
            if (!row[prop])
            {
              row[prop] = rowToUpdate[prop]
            }
          }
        }
        setOptions.firstRowIndex = rowToUpdate.sheetRow
        setRowsData(
          sheet,
          [row],
          setOptions
        )
        if (setOptions.log) console.log("Updated data for %s=%s on row %s", primaryKey, row[primaryKey], rowToUpdate.sheetRow)
        return row
      })
    }
  } else
  {
    // Update via sheetRow
    // We don't getRowsData so throw an error if .onlyPresentColumns is passed
    if (updateOptions.onlyPresentColumns) throw new Error("updateRows does not support option 'onlyPresentColumns' without a primary key")
    updatedRows = rows.map(function (row)
    {
      if (!row.sheetRow) throw new Error("Unable to update row: no metadata attached. On sheet " + sheet.getName() + " of " + sheet.getParent().getName())
      setOptions.firstRowIndex = row.sheetRow
      setRowsData(
        sheet,
        [row],
        setOptions
      )
    })
    if (setOptions.log) console.log("Updated rows at these indices:\n%s", rows.map(function (x) { return x.sheetRow }))
    return row
  }

  // Flush before releasing lock
  SpreadsheetApp.flush()
  // Don't release the lock if it was passed.  Allow the calling function to release it.
  if (!passedLock) lock.releaseLock()

  return updatedRows
}

/**
 * Delete rows data for existing rows, with a lock to avoid collisions or changes to the sheet while we are processing.
 * @param {SpreadsheetApp.Sheet} sheet 
 * @param {Object[]} rows Rows data objects
 * @param {Object} options Options to be passed to getRowsData
 * @param {string} primaryKey  Unique key to use to identify the row to update.  If not present, we will use the .sheetRow property in the rows data.
 */
function removeRows(sheet, rows, options, primaryKey, passedLock)
{
  // Default options
  options = options || {}


  // If one row is passed, make it an array
  if (!(rows instanceof Array)) rows = [rows]
  if (rows.length === 0)
  {
    console.warn("removeRows: No rows to remove")
    return
  }

  // Set a lock
  var lock = passedLock;
  if (!lock)
  {
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
  }

  // Update rows via primaryKey
  if (primaryKey)
  {
    var dataByKey = hashObjects(
      getRowsData(sheet, null, Object.assign(options, { getMetadata: true })),
      primaryKey
    )
    var rowsToDelete = rows.map(function (row)
    {
      if (!row[primaryKey]) throw new Error("Unable to remove row: primary key " + primaryKey + " not found in this row: " + JSON.stringify(row) + " on sheet " + sheet.getName() + " of " + sheet.getParent().getName())
      var rowToUpdate = dataByKey[row[primaryKey]]
      if (!rowToUpdate) throw new Error("Unable to update row: no row found with " + primaryKey + " = " + row[primaryKey] + " on sheet " + sheet.getName() + " of " + sheet.getParent().getName())
      if (options.log) console.log("Removing row %s where %s=%s", rowToUpdate.sheetRow, primaryKey, row[primaryKey])
      return rowToUpdate.sheetRow
    })
    deleteSheetRows(sheet, rowsToDelete)
  } else
  {
    // Remove via sheetRow
    deleteSheetRows(sheet, rows.map(function (x) { return x.sheetRow }))
    if (options.log) console.log("Removed rows at these indices:\n%s", rows.map(function (x) { return x.sheetRow }))
  }

  // Flush before releasing lock
  SpreadsheetApp.flush()
  // Don't release the lock if it was passed.  Allow the calling function to release it.
  if (!passedLock) lock.releaseLock()
}

/**
 * Use the batchUpdate method to delete multiple rows from a sheet.
 * @param {integer[]} rowsToDelete Array of 1-based row indices to delete
 * @requires Service Advanced Sheets service
 */
function deleteSheetRows(sheet, rowsToDelete)
{
  // Each time we remove a row, the indices change, so pre-calculate the adjusted indices.
  var adjustedRowsToDelete = rowsToDelete.sort(function (a, b) { return a - b }).map(function (x, i) { return x - i })
  var sheetId = sheet.getSheetId()
  var requests = adjustedRowsToDelete.map(function (x)
  {
    console.log("Deleting row %s", x)
    return {
      'deleteDimension': {
        'range': {
          "sheetId": sheetId,
          "dimension": 'ROWS',
          // Half-open range: start at x-1 b/c API uses 0-based indices
          "startIndex": x - 1,
          "endIndex": x
        }
      }
    }
  })
  try
  {
    Sheets.Spreadsheets.batchUpdate({ 'requests': requests }, sheet.getParent().getId())
  } catch (err)
  {
    // If it's the last row we'll get this error:
    if (err.message.includes('Sorry, it is not possible to delete all non-frozen rows.'))
    {
      sheet.insertRowAfter(sheet.getLastRow())
      Sheets.Spreadsheets.batchUpdate({ 'requests': requests }, sheet.getParent().getId())
    } else
    {
      throw err
    }
  }
}
