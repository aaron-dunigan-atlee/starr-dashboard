
/**
 * Get the range object corresponding to the data range minus a header row.
 * @param {Sheet} sheet 
 */
function getDataRangeMinusHeaders(sheet, optHeaderRowCount)
{
  var headerRowCount = optHeaderRowCount || 1;
  var dataRange = sheet.getDataRange();
  var height = dataRange.getHeight();
  if (header <= headerRowCount)
  {
    return null;
  }
  var width = dataRange.getWidth();
  return sheet.getRange(1, 1, height - headerRowCount, width);
}

/**
 * Hash an array of objects by a key
 * @param {Object[]} array 
 * @param {string} key 
 * @param {Object} options
 *    strict {boolean} If true, throw error if key is absent;
 *    keyCase {string} Convert case of key before hashing.  'lower' or 'upper';
 *    verbose {boolean} Log a warning if key is absent;
 *    toString {boolean} Explicitly convert keys to strings.  Default false.
 * @return {Object} Object of form {key: Object from array}
 */
function hashObjects(array, key, options)
{
  if (key instanceof Array) return multihashObjects(array, key, options)
  options = options || {}
  var hash = {};
  array.forEach(function (object)
  {
    if (object[key])
    {
      var thisKey = object[key];
      if (options.toString) thisKey = thisKey.toString();
      if (options.keyCase == 'upper') thisKey = thisKey.toLocaleUpperCase();
      if (options.keyCase == 'lower') thisKey = thisKey.toLocaleLowerCase();
      hash[thisKey] = object;
    } else
    {
      if (options.strict) throw new Error("Can't hash object because it doesn't have key " + key)
      if (options.verbose) console.warn("Can't hash object because it doesn't have key " + key + ": " + JSON.stringify(object))
    }
  })
  return hash
}


/**
 * Hash an array of objects by a key, where there may be multiple elements sharing the same key
 * @param {Object[]} array 
 * @param {string} key 
 * @param {Object} options
 *    strict {boolean} If true, throw error if key is absent;
 *    keyCase {string} Convert case of key before hashing.  'lower' or 'upper';
 *    verbose {boolean} Log a warning if key is absent;
 * @return {Object} Object of form {key: [Objects from array]}
 */
function hashObjectsManyToOne(array, key, options)
{
  options = options || {}
  var hash = {};
  array.forEach(function (object)
  {
    if (object[key])
    {
      var thisKey = object[key];
      if (options.keyCase == 'upper') thisKey = thisKey.toLocaleUpperCase();
      if (options.keyCase == 'lower') thisKey = thisKey.toLocaleLowerCase()
      if (hash[thisKey])
      {
        hash[thisKey].push(object);
      } else
      {
        hash[thisKey] = [object];
      }

    } else
    {
      if (options.strict) throw new Error("Can't hash object because it doesn't have key " + key)
      if (options.verbose) console.warn("Can't hash object because it doesn't have key " + key + ": " + JSON.stringify(object))
    }
  })
  return hash
}

/**
 * Hash an array of objects by several keys, which will be joined
 * @param {Object[]} array 
 * @param {string[]} keys
 * @param {Object} options
 *    strict {boolean} If true, throw error if key is absent;
 *    keyCase {string} Convert case of key before hashing.  'lower' or 'upper';
 *    verbose {boolean} Log a warning if key is absent;
 *    separator {string} Used to separate keys.  Default is '.'
 * @return {Object} Object of form {key: Object from array}
 */
function multihashObjects(array, keys, options)
{
  options = options || {}
  var separator = options.separator || '.'
  var hash = {};
  array.forEach(function (object)
  {

    var thisKey = keys.map(function (key) { return object[key] }).join(separator);
    if (options.keyCase === 'upper') thisKey = thisKey.toLocaleUpperCase();
    if (options.keyCase === 'lower') thisKey = thisKey.toLocaleLowerCase();

    hash[thisKey] = object;

  })
  return hash
}


/**
 * Hash an array of objects by a compound key, where there may be multiple elements sharing the same key
 * @param {Object[]} array 
 * @param {string[]} keys 
 * @param {Object} options
 * @return {Object} Object of form {key: [Objects from array]}
 */
function multihashObjectsManyToOne(array, keys, options)
{
  options = options || {}
  var separator = options.separator || '.'
  var hash = {};
  array.forEach(function (object)
  {
    var thisKey = keys.map(function (key) { return object[key] }).join(separator);
    if (hash[thisKey])
    {
      hash[thisKey].push(object);
    } else
    {
      hash[thisKey] = [object];
    }
  })
  return hash
}

// TODO: nested hash

/**
 * Get the sheet index of the first empty row on a leadsheet.  
 * @param {SpreadsheetApp.Sheet} sheet 
 * @param {integer} startRow            Optional row to start looking on.  Defaults to 2 (i.e. assuming there is 1 header row)
 * Sometimes we can't use .getLastRow() because checkboxes and other stuff count as data.
 */
function getFirstEmptyRow(sheet, startRow)
{
  startRow = startRow || 2;
  var emptyRowIndex = sheet
    .getRange('A:A')
    .getValues()
    .findIndex(function (row, index) { return index >= startRow - 1 && !row[0] })
  if (emptyRowIndex > -1)
  {
    // console.log("First empty row on " + leadsheet.getName() + " is row " + (emptyRowIndex + 1))
    return emptyRowIndex + 1;
  } else
  {
    // Sheet is full, so insert a row at the bottom.
    // console.log("No empty rows on " + leadsheet.getName() + ", so we'll insert one at the bottom.")
    var lastRow = sheet.getMaxRows();
    sheet.insertRowAfter(lastRow);
    return lastRow + 1;
  }
}


/**
 * Update the values in a single column of an MSP table, for a subset of the rows.
 * Makes a batchupdate to Sheets API, with a separate request for each value.
 * @param {string} abbrev MSP2, MSP3, etc.
 * @param {string} columnName Normalized column header
 * @param {string} data   Rows data.  Must include metadata (sheetRow).
 * @requires Array.prototype.findIndex() (Polyfill if running on Rhino)
 * @requires Sheets Advanced service must be turned on
 */
function updateColumn(abbrev, columnName, data)
{
  console.log("Updating column '%s' on %s", columnName, abbrev)
  var id = SOURCE_SPREADSHEETS[abbrev].id;
  var ss = SpreadsheetApp.openById(id);
  var sheetName = SOURCE_SPREADSHEETS[abbrev].sheetName;
  var sheet = ss.getSheetByName(sheetName);

  // Get column index to update
  var columnIndex = sheet.getRange("2:2").getValues()[0].findIndex(function (header)
  {
    return normalizeHeader(header) === columnName
  }) + 1;
  if (columnIndex === 0) throw new Error(Utilities.formatString("updateColumn: Failed to find column '%s' on %s", columnName, abbrev))
  console.log("'%s' is on column %s", columnName, columnIndex.toString())

  // Create batch update requests
  var requestData = data.map(function (row)
  {
    var range = "'" + sheetName + "'!" + sheet.getRange(row.sheetRow, columnIndex).getA1Notation()
    return {
      // ValueRange object: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values#ValueRange
      "range": range,
      // "majorDimension": enum (Dimension),
      "values": [
        [row[columnName]]
      ]
    }
  }) // requestsData = data.map()

  // See https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/batchUpdate
  var request = {
    "valueInputOption": "USER_ENTERED",
    "data": requestData,
    // "includeValuesInResponse": boolean,
    "responseValueRenderOption": "FORMATTED_VALUE"
    // "responseDateTimeRenderOption": enum (DateTimeRenderOption)
  }

  var response = Sheets.Spreadsheets.Values.batchUpdate(request, id);
  console.log("updateColumn response: %s", JSON.stringify(response))
}