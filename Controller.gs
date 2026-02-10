function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Generator')
    .addItem('Subscriber Table', 'openGeneratorDialog')
    .addToUi();
}

function openGeneratorDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(400)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Subscriber Table Generator');
}

/**
 * Exposed to HTML: Gets list of available months
 */
function apiGetAvailableMonths(showArchived) {
  return Service.getAvailableMonths(showArchived);
}

/**
 * Exposed to HTML: Triggers table generation
 */
function apiGenerateTable(monthYear) {
  try {
    return Service.generateTable(monthYear);
  } catch (e) {
    throw new Error(e.message);
  }
}