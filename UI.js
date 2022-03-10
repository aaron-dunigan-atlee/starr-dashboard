/**
 * Create custom menu(s)
 * @param {Event} e 
 */
function onOpen(e)
{
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📊 Dashboard')
    .addItem('🔄 Update dashboard', 'updateDashboard')
    .addItem('❔ View documentation', 'showDocumentation')
    .addSeparator()
    .addItem('🆗 Authorize automation', 'authorize')
  menu.addToUi();
  // addDebugMenu(ui, true) 
}

/**
 * Just here to force (re-)authorization
 */
function authorize()
{
  SpreadsheetApp.getActive().toast('🆗 The script is authorized')
  console.log('🆗 The script is authorized')
}

/**
 * Show a spinner in a modal dialog while a function runs.
 * @param {string} title        Title of dialog
 * @param {string} message      Message to show user while the function runs.
 * @param {string} functionName Function to run.
 * @param {string} successMessage Message to be displayed in the modal if the function runs successfully.
 * @param {string} failureMessage Message to be displayed on failure.  The error message will be appended.
 * @requires spinner-modal.html
 */
function showSpinnerModal(message, functionName, successMessage, failureMessage, title, initArgs = { continuationToken: 0 })
{
  title = title || "One moment..."
  var template = HtmlService.createTemplateFromFile('html/spinner-modal-batch');
  template.message = message;
  template.functionName = functionName;
  template.successMessage = successMessage;
  template.failureMessage = failureMessage;
  template.initArgs = JSON.stringify(initArgs);
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(300).setHeight(250),
    title
  )
}

function onFinishSpinnerModal(message)
{
  SS.toast(message)
}


/**
 * Show a dialog with a brief message and a link
 * @param {string} title
 * @param {string} message
 * @param {string} linkText 
 * @param {string} url 
 * @requires link-dialog.html
 */
function showLinkDialog(title, message, linkText, url, height)
{
  var template = HtmlService.createTemplateFromFile('html/link-dialog');
  template.url = url;
  template.message = message;
  template.linkText = linkText;
  var dialog = template.evaluate().setHeight(height || 100).setWidth(300);
  SpreadsheetApp.getUi().showModalDialog(dialog, title);
}

function showDocumentation()
{
  var file = DriveApp.getFileById(DOCUMENTATION_ID)
  showLinkDialog(
    'Open Documentation',
    '',
    'Click here to open the dashboard documentation',
    file.getUrl()
  )
}

/**
 * Show a dialog with a link to a folder/file.
 * @param {string} fileId 
 * @requires link-dialog.html
 */
function showFileLink(fileId)
{
  var file = DriveApp.getFileById(fileId);
  var message = "Open '" + file.getName() + "':";
  var linkText = 'Click here to open this file or folder';
  var url = file.getUrl();
  showLinkDialog('Open a File or Folder', message, linkText, url);
}