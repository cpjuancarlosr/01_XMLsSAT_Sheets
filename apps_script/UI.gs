/**
 * @OnlyCurrentDoc
 *
 * The onOpen function runs automatically when the spreadsheet is opened.
 * It creates a custom menu for the user to interact with the system.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Contabilidad Automatizada')
    .addItem('Cargar XML desde Drive', 'showPicker')
    .addToUi();
}

/**
 * Displays an HTML sidebar interface that allows the user to select XML files from Google Drive.
 */
function showPicker() {
  const html = HtmlService.createHtmlOutputFromFile('Picker')
    .setWidth(600)
    .setHeight(425)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Selecciona los archivos XML');
}
