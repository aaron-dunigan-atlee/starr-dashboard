/**
 * @OnlyCurrentDoc
 */

/**
 * Scripts by Aaron Dunigan AtLee
 * aaron.dunigan.atlee [at gmail]
 * March 2021
 */



/**
 * Get the sheet with the given id (why isn't this a built in method?)
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet Optional, defaults to the active spreadsheet 
 * @param {string || integer} id  Sheet ID 
 * @returns {SpreadsheetApp.Sheet}  The sheet, or null if it doesn't exist.
 */
function getSheetById(id, spreadsheet)
{
  spreadsheet = spreadsheet || SpreadsheetApp.getActive()
  return spreadsheet.getSheets().find(
    // Soft equals here because id may be passed as string
    function (s) { return s.getSheetId() == id; }
  );
}


/**
 * Create a new spreadsheet in the folder, and copy the sheet to it.
 * @param {string} filename
 * @param {Sheet} sheet 
 * @param {Folder} folder Optional destination folder.  Defaults to root folder.
 */
function createSpreadsheetFromSheet(filename, sheet, folder)
{
  folder = folder || DriveApp.getRootFolder();

  // Create a new blank file
  var spreadsheet = SpreadsheetApp.create(filename)

  // Move to the folder
  var file = DriveApp.getFileById(spreadsheet.getId())
  file.moveTo(folder)

  // Copy the sheet
  sheet.copyTo(spreadsheet).setName(sheet.getName())

  // Remove the default blank Sheet1
  var blankSheet = spreadsheet.getSheetByName('Sheet1')
  if (blankSheet) spreadsheet.deleteSheet(blankSheet)

  console.log("Created new spreadsheet based on sheet " + sheet.getName())
  return spreadsheet
}


/**
 * Get the values in a named range as a single array, filtering out empty values.
 * @param {string} rangeName 
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet  Optional.  Defaults to active spreadsheet.
 * @param {SpreadsheetApp.Sheet} sheet              Optional.  If not given, get the range from the spreadsheet object.
 */
function getFlattenedValues(rangeName, spreadsheet, sheet)
{
  // Get the range if it exists.
  var range;
  if (sheet)
  {
    try
    {
      range = sheet.getRange(rangeName)
    } catch (err)
    {
      console.warn("No range named " + rangeName)
      return []
    }
  } else
  {
    spreadsheet = spreadsheet || SpreadsheetApp.getActive();
    range = spreadsheet.getRangeByName(rangeName);
    if (!range)
    {
      console.warn("No range named " + rangeName)
      return []
    }
  }

  // Get the values
  var values = range.getValues();
  var flatValues = []

  for (var i = 0; i < values.length; i++)
  {
    flatValues = flatValues.concat(values[i])
  }

  // Remove empty values
  return flatValues.filter(function (x) { return x !== undefined && x !== '' })
}

/**
 * Set a sheet's name, but if the name is already taken, append a parenthetical number (2), (3), etc.
 * @param {SpreadsheetApp.Sheet} sheet 
 * @param {string} desiredName 
 * @returns {string} The actual name assigned to the sheet.
 */
function setSheetNameWithoutCollisions(sheet, desiredName)
{
  console.log("Attempting to rename sheet to " + desiredName)
  var ss = sheet.getParent()
  var sheetName = desiredName;
  var n = 2;
  // Check for duplicate sheet name
  while (ss.getSheetByName(sheetName))
  {
    console.log("There was already a sheet called '" + sheetName + ".'")
    sheetName = desiredName + ' (' + n + ')';
    n++;
  }
  sheet.setName(sheetName)
  console.log("Renamed sheet to " + sheetName)
  return sheetName;
}

/**
 * "Make JSON Pretty"
 * Show a dialog with any json in the active cell stringified.
 */
function showJsonInActiveCell()
{
  var text = SpreadsheetApp.getActive().getActiveCell().getValue();
  try
  {
    var json = JSON.parse(text);
  } catch (err)
  {
    SpreadsheetApp.getActive().toast('No JSON in that cell');
    return;
  }

  var html = HtmlService.createHtmlOutput('<pre>' + JSON.stringify(json, null, 2) + '</pre>');

  return SpreadsheetApp.getUi().showModalDialog(html, 'Detail');
}

/**
 * Sort the order of the tabs in a spreadsheet.
 * @param {Function} sortFunction 
 */
function sortSheetTabs(sortFunction)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets().sort(sortFunction);

  sheets.forEach(function (sheet, index)
  {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index + 1);
  });

}

/**
 * Sort function for sorting sheets by name
 */
function bySheetName(a, b)
{
  if (a.getName() < b.getName()) return -1
  if (a.getName() > b.getName()) return 1
  return 0;
}

/**
 * Update the number of rows in a named range.
 * @param {Sheet} sheet 
 * @param {string} rangeName 
 * @param {integer} rowsToAdd 
 */
function addRowsToNamedRange(sheet, rangeName, rowsToAdd)
{
  // Update the size of the named range if needed.
  var namedRange = sheet.getNamedRanges().find(function (namedRange) { return namedRange.getName() === rangeName })
  if (namedRange)
  {
    var existingRange = namedRange.getRange();
    var targetRange = existingRange.offset(0, 0, existingRange.getHeight() + rowsToAdd)
    namedRange.setRange(targetRange);
    console.log("Changed range '%s' to %s", rangeName, targetRange.getA1Notation());
  }
}

/**
 * Get a named range from a sheet.  Return null if this sheet doesn't have that named range 
 * (or throw an error if we are in strict mode)
 * Not equivalent to sheet.getRange(name) because that throws an exception if the range doesn't exist 
 * (and worse, can return a range from another sheet if the range does exist on another sheet but not on this one)
 * Also, this method handles ranges whose name gets a sheetname prepended when copying the sheet, for example.
 * @param {Sheet} sheet 
 * @param {string} rangeName 
 * @param {Object} options 
 * options include:
 *  -- strict {boolean}: Default is true.  If true, throw a (hopefully helpful) error if the range doesn't exist.
 *  -- getNamedRangeObject {boolean} Default false.  If true, return the entire NamedRange object; otherwise just the Range
 */
function getRangeByName(sheet, rangeName, options)
{
  options = options || {}
  if (options.strict !== false) options.strict = true;
  if (!sheet)
  {
    if (options.strict) throw new Error(`Missing named range "${rangeName}"`)
    return null;
  }
  var namedRanges = sheet.getNamedRanges();
  var match = namedRanges.find(function (namedRange)
  {
    // console.log("Named range: " + namedRange.getName())
    // Range names may have sheet names, e.g. 'Sheet1'!RangeName, so we want to strip the sheet name:
    return namedRange.getName().replace(/^.*!/, '') === rangeName
  }) // find

  if (match)
  {
    if (options.getNamedRangeObject)
    {
      return match
    } else
    {
      return match.getRange();
    }
  } else
  {
    if (options.strict) throw new Error(
      `Missing named range "${rangeName}" on sheet "${sheet.getName()}" of spreadsheet "${sheet.getParent().getName()}"`
    )
    return null;
  }

}

/**
 * Sort a sheet by a column specified by the column header
 * @param {Sheet} sheet       Sheet to be sorted
 * @param {string} header     Text of header cell in the column to sort by
 * @param {Object} options    {headersRowIndex: {integer} Row where the headers are found.
 *                             strict: {boolean} Whether to throw an exception if header is not found.
 *                              ascending: {boolean} If false, sort descending
 *                            }
 */
function sortSheetByHeader(sheet, header, options)
{
  options = options || {};
  var headerRow = options.headersRowIndex || 1;
  var sheetName = sheet.getName();
  var headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var sortColumn = headers.indexOf(header) + 1;
  if (sortColumn === 0)
  {
    var errorMessage = Utilities.formatString("Failed to sort sheet: Can't find header '%s' on row %s of sheet '%s'", header, headerRow, sheetName);
    if (options.strict) throw new Error(errorMessage);
    console.warn(errorMessage)
    return;
  }
  sheet.sort(sortColumn, Boolean(options.ascending))
  console.log("Sorted by sheet '%s' by column '%s'", sheetName, header.header)
}


/**
 * Sort a sheet by (one or more) columns, using column headers to identify the sort columns.
 * @param {Sheet}     sheet       The sheet to sort
 * @param {Object[]}  sortHeaders Of the form [{header: {string}, ascending: {boolean}}]
 * @param {Object}    options     headersRowIndex: {integer} Row where the headers are found.
 *                                strict: {boolean} Whether to throw an exception if header is not found.
 *                                
 */
function sortSheetByHeaders(sheet, sortHeaders, options)
{
  options = options || {};
  var headerRow = options.headersRowIndex || 1;
  var sheetName = sheet.getName();
  var headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0]
  sortHeaders.forEach(function (header)
  {
    var column = headers.indexOf(header.header) + 1 // +1 to map 0-based js array to 1-based sheet index
    if (column === 0)
    {
      var errorMessage = Utilities.formatString("sortSheetByHeaders: Can't find header '%s' on row %s of sheet '%s'", header.header, headerRow, sheetName);
      if (options.strict) throw new Error(errorMessage);
      console.warn(errorMessage)
      return;
    }
    sheet.sort(column, Boolean(header.ascending))
    console.log("Sorted by sheet '%s' by column '%s'", sheetName, header.header)
  })

}


/**
 * Get a link directly to this sheet.
 * @param {Sheet} sheet 
 * @param {Object} options  .noTools: boolean, if true, link to a minimal interface with no Sheets header or toolbar
 */
function getSheetUrl(sheet, options)
{
  var baseUrl = sheet.getParent().getUrl()
  options = options || {}
  // If the url has parameters, remove them, and append #gid=...
  var url = baseUrl.replace(/\?.*$/, '');
  if (options.noTools) url += '?rm=minimal'
  url += '#gid=' + sheet.getSheetId();
  return url
}

/**
 * Get a link directly to this range.
 * @param {Range} range 
 */
function getRangeUrl(range)
{
  return getSheetUrl(range.getSheet()) + '&range=' + range.getA1Notation()
}


/**
 * Convert range to blob (useful for converting to pdf)
 * https://xfanatical.com/blog/print-google-sheet-as-pdf-using-apps-script/
 * @param {*} sheet 
 * @param {*} range 
 */
function getRangeAsBlob(sheet, range, fullSheet)
{
  var spreadsheetUrl = sheet.getParent().getUrl();
  var rangeParam = ''
  if (range)
  {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  var sheetParam = '&gid=' + sheet.getSheetId()

  var exportUrl = spreadsheetUrl.replace(/\/edit.*$/, '')
    + '/export?exportFormat=pdf&format=pdf'
    + '&size=LETTER'
    + (fullSheet ? '&portrait=false' : '&portrait=true')
    + '&scale=4'   // 1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page
    + '&top_margin=0.75'
    + '&bottom_margin=0.75'
    + '&left_margin=0.7'
    + '&right_margin=0.7'
    + '&sheetnames=false&printtitle=false'
    + '&pagenum=false'
    + '&gridlines=false'
    + '&fzr=FALSE' // Whether to repeat frozen rows on each page
    + sheetParam
    + rangeParam

  // console.log('exportUrl=' + exportUrl)
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  })

  return response.getBlob()
}

/*
Can't find documentation for the params above, but this post from stack overflow claims the following
https://stackoverflow.com/a/60653901/10332984
and see also https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
and this one might help when translated from russian: https://kandiral.ru/googlescript/eksport_tablic_google_sheets_v_pdf_fajl.html

function getBlob(){
  var url = 'https://docs.google.com/spreadsheets/d/';
  var id = '<YOUR-FILE-ID>';
  var url_ext = '/export?'
  +'format=pdf'
  +'&size=a4'                      //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
  +'&portrait=true'                //true= Potrait / false= Landscape
  +'&scale=1'                      //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page
  +'&top_margin=0.00'              //All four margins must be set!
  +'&bottom_margin=0.00'           //All four margins must be set!
  +'&left_margin=0.00'             //All four margins must be set!
  +'&right_margin=0.00'            //All four margins must be set!
  +'&gridlines=true'               //true/false
  +'&printnotes=false'             //true/false
  +'&pageorder=2'                  //1= Down, then over / 2= Over, then down
  +'&horizontal_alignment=LEFT'  //LEFT/CENTER/RIGHT
  +'&vertical_alignment=TOP'       //TOP/MIDDLE/BOTTOM
  +'&printtitle=false'             //true/false
  +'&sheetnames=false'             //true/false
  +'&fzr=false'                    //true/false
  +'&fzc=false'                    //true/false
  +'&attachment=false'
  +'&gid=0';
  // console.log(url+id+url_ext);
  var blob = UrlFetchApp.fetch(url+id+url_ext).getBlob().getAs('application/pdf');
  return blob;
}
*/

/**
 * Generate a pdf from a sheet and save it in Drive.
 * @param {Sheet} sheet 
 * @param {Folder} destinationFolder Defaults to root folder
 * @param {string} filename          Defaults to name of sheet
 */
function saveSheetAsPdf(sheet, destinationFolder, filename)
{
  filename = filename || sheet.getName();
  destinationFolder = destinationFolder || DriveApp.getRootFolder();
  // Empirically, changes need to be flushed in order to appear in the PDF.  
  // We probably should do this before calling saveSheetAsPdf, but I always forget.
  SpreadsheetApp.flush();
  var blob = getRangeAsBlob(sheet);
  var file = destinationFolder.createFile(blob).setName(filename)
  return file.getId();
}

/**
 * Generate a pdf from a sheet and save it in Drive.
 * @param {SpreadsheetApp.Range} range 
 * @param {Folder} destinationFolder Defaults to root folder
 * @param {string} filename          Defaults to name of sheet
 */
function saveRangeAsPdf(range, destinationFolder, filename)
{
  var sheet = range.getSheet()
  filename = filename || sheet.getName();
  destinationFolder = destinationFolder || DriveApp.getRootFolder();
  // Empirically, changes need to be flushed in order to appear in the PDF.  
  // We probably should do this before calling saveSheetAsPdf, but I always forget.
  SpreadsheetApp.flush();
  var blob = getRangeAsBlob(sheet, range);
  var file = destinationFolder.createFile(blob).setName(filename)
  return file.getId();
}

/**
 * Test whether a named range and a Range object have the same upper left cell.
 * @param {string} rangeName 
 * @param {Range} cellRange 
 */
function namedRangeStartsAtCell(rangeName, cellRange)
{
  var range = SS.getRangeByName(rangeName);
  if (!range) return false;
  return (
    range.getRow() === cellRange.getRow()
    && range.getColumn() === cellRange.getColumn()
  )
}



/**
 * Programmatically "Allow Access" for an IMPORTRANGE formula.  Assumes the formula is already present on the sheet.
 * See discussion: https://stackoverflow.com/a/64121004
 * Empirically, this only seems to work temporarily
 * @param {SpreadsheetApp.Range} targetCell Cell with the importrange formula
 * @param {string} sourceId  Id of the Spreadsheet that the importrange points to
 */
function allowAccessImportRange(targetCell, sourceId)
{
  var source = DriveApp.getFileById(sourceId)
  source.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  var targetSpreadsheet = targetCell.getSheet().getParent()
  var formula = targetCell.getFormula()
  targetCell.clearContent()
  SpreadsheetApp.flush()
  targetCell.setFormula(formula)
  SpreadsheetApp.flush()
  source.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
}

/**
 * Count the cells in a whole spreadsheet
 */
function countCells()
{
  console.log("Cells in spreadsheet: %s",
    SS.getSheets().reduce(function (cellCount, sheet)
    {
      const sheetCellCount = (sheet.getMaxRows() * sheet.getMaxColumns())
      console.log("Cells in sheet %s: %s", sheet.getName(), sheetCellCount)
      return cellCount + sheetCellCount
    }, 0)
  )
}


/**
 * Retrieve the first spreadsheet named spreadsheetName, from the parentFolder
 */
function getOrCreateSpreadsheetByName(parentFolder, spreadsheetName)
{
  var iterator = parentFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (iterator.hasNext())
  {
    var file = iterator.next()
    if (file.getName() === spreadsheetName)
    {
      return SpreadsheetApp.open(file);
    }
  }
  var newSpreadsheet = SpreadsheetApp.create(spreadsheetName)

  // Move to the folder
  var file = DriveApp.getFileById(newSpreadsheet.getId())
  file.moveTo(parentFolder)

  console.log("Created spreadsheet '%s' in '%s'", spreadsheetName, parentFolder.getName());
  return newSpreadsheet
}


/**
 * Get all sheets in the active spreadsheet, that are form response sheets
 * @returns {Sheet[]}
 */
function getFormResponseSheets()
{
  return SS.getSheets().filter(sheet => sheet.getFormUrl())
}
