var SHEETS_TABLES = {}

function objectMatchesFilter(object, filter)
{
  for (var property in filter)
  {
    if (filter[property].indexOf(object[property]) === -1) { return false; }
  }
  return true;
}

function SheetsTable(config)
{
  for (var prop in config)
  {
    this[prop] = config[prop]
  }
  // Set project-wide defaults here
  this.defaultOptions = this.defaultOptions || {}
  this.propertyStore = PropertiesService.getScriptProperties();
  SHEETS_TABLES[config.name] = this
}

// Getter for the sheet; we don't want to call openById on initialization in the global scope
// This is a Smart / self-overwriting / lazy getter, see https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Functions/get#examples (but the example there doesn't seem to work here in apps script)
Object.defineProperty(SheetsTable.prototype, 'sheet', {
  get: function ()
  {
    if (this.sheet_)
    {
      return this.sheet_
    }
    var spreadsheet;
    if (this.spreadsheetId === 'active')
    {
      spreadsheet = SpreadsheetApp.getActive()
      this.spreadsheetId = spreadsheet.getId();
    }
    else if (this.hasOwnProperty('spreadsheetId'))
    {
      spreadsheet = SpreadsheetApp.openById(this.spreadsheetId)
    }
    else if (this.hasOwnProperty('spreadsheetUrl'))
    {
      spreadsheet = SpreadsheetApp.openByUrl(this.spreadsheetUrl)
    }
    if (!spreadsheet)
    {
      throw new Error("Spreadsheet not specified: please provide an id or url");
    }
    if (this.hasOwnProperty('sheetName'))
    {
      this.sheet_ = spreadsheet.getSheetByName(this.sheetName);
    }
    else if (this.hasOwnProperty('sheetIndex'))
    {
      this.sheet_ = spreadsheet.getSheets()[this.sheetIndex];
    }
    else
    {
      throw new Error("Sheet not specified: please provide a sheet name or index");
    }
    return this.sheet_
  }
});

// Getter for the sheet headers; 
// This is a Smart / self-overwriting / lazy getter, see https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Functions/get#examples (but the example there doesn't seem to work here in apps script)
Object.defineProperty(SheetsTable.prototype, 'headers', {
  get: function ()
  {
    if (this.headers_)
    {
      return this.headers_
    }
    var sheet = this.sheet
    // TODO: this will always use camel case.  Create a normalization function that uses the case option
    this.headers_ = normalizeHeaders(sheet.getRange(this.defaultOptions.headersRowIndex || 1, 1, 1, sheet.getLastColumn()).getValues()[0])
    return this.headers_
  }
});

SheetsTable.prototype.getRows = function (getOptions, rangeA1Notation, config)
{
  config = config || {}
  getOptions = getOptions || {}
  // In some cases we want to get the data once per execution and cache it; in this case, set options.refresh = false
  if (getOptions.refresh !== false || config.refresh !== false || !this.data_) // First condition is for backward compatibility
  {
    var currentOptions = new Object(this.defaultOptions)
    Object.assign(currentOptions, getOptions || {})
    this.data_ = getRowsData(
      this.sheet,
      rangeA1Notation ? this.sheet.getRange(rangeA1Notation) : null,
      currentOptions
    )
  }
  if (config.filter)
  {
    return this.data_.filter(function (row)
    {
      return objectMatchesFilter(row, config.filter)
    });
  }
  else
  {
    return this.data_
  }
}

SheetsTable.prototype.getRow = function (primaryKey, options, range)
{
  options = options || {}
  if (options.log !== false) console.log("Searching for row with %s==%s on table %s", this.primaryKey, primaryKey, this.name)
  var table = this
  var row = this.getRows(options, range).find(function (row)
  {
    return row[table.primaryKey] == primaryKey
  })
  if (row)
  {
    if (options.log !== false) console.log("Found row with %s==%s on table %s", this.primaryKey, primaryKey, this.name)
  }
  else
  {
    var message = "No row found with " + this.primaryKey + "==" + primaryKey + " on table " + this.name
    if (options.strict) throw new Error(message)
    if (options.log !== false) console.warn(message)
  }
  return row
}

SheetsTable.prototype.incrementPrimaryKey = function ()
{
  var propertyStoreKey = this.name + '.' + this.primaryKey
  var nextKey = parseInt((this.propertyStore.getProperty(propertyStoreKey) || 0), 10) + 1
  this.propertyStore.setProperty(propertyStoreKey, nextKey.toString())
  return nextKey
}

/**
 * 
 * @param {Object[] | Object} rows 
 * @param {Object} setRowsOptions Options for setRowsData/getRowsData
 * @param {Lock} lock Can be passed if we already have a lock. Otherwise one will be requested.
 * @param {Object} updateOptions .onlyPresentColumns {boolean} If true, only set columns that are passed as properties.  i.e. preserve values for all columns not present.
 * @returns 
 */
SheetsTable.prototype.updateRows = function (rows, setRowsOptions, lock, updateOptions)
{
  var currentOptions = new Object(this.defaultOptions)
  Object.assign(currentOptions, setRowsOptions || {})
  return updateRows(this.sheet, rows, currentOptions, this.primaryKey, lock, updateOptions)
}

SheetsTable.prototype.removeRows = function (rows, options, lock)
{
  var currentOptions = new Object(this.defaultOptions)
  Object.assign(currentOptions, options || {})
  return removeRows(this.sheet, rows, currentOptions, this.primaryKey, lock)
}

SheetsTable.prototype.insertRows = function (rows, options)
{
  if (!(rows instanceof Array))
  {
    rows = [rows]
  }
  var currentOptions = new Object(this.defaultOptions)
  Object.assign(currentOptions, options || {})
  currentOptions.writeMethod = 'appendRow'
  var table = this;
  rows.forEach(function (row)
  {
    if (!row[table.primaryKey]) row[table.primaryKey] = table.incrementPrimaryKey()
  })
  setRowsData(this.sheet, rows, currentOptions)
  return rows
}

SheetsTable.prototype.getRowsHashedBy = function (key, hashOptions, getRowsOptions, range)
{
  var data = this.getRows(getRowsOptions, range)
  hashOptions = hashOptions || {}
  if (hashOptions.manyToOne)
  {
    return hashObjectsManyToOne(data, key, hashOptions)
  }
  else
  {
    return hashObjects(data, key, hashOptions)
  }
}

var DASHBOARD_TABLE = new SheetsTable({
  name: 'Dashboard',
  spreadsheetId: 'active',
  sheetName: 'GRE4T STARR Dashboard',
  primaryKey: ['districtName', 'schoolName'],
  defaultOptions: {
    headersRowIndex: 2,
    headersCase: 'camel',
  }
})

var DIRECTORY_TABLE = new SheetsTable({
  name: 'Directory',
  spreadsheetId: 'active',
  sheetName: 'School Directory',
  primaryKey: ['districtName', 'schoolName'],
  defaultOptions: {
    headersRowIndex: 1,
    headersCase: 'camel',
    get: 'hyperlinks'
  }
})
