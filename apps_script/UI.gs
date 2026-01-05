/**
 * @OnlyCurrentDoc
 *
 * Se ejecuta automáticamente al abrir la hoja de cálculo.
 * Configura el espacio de trabajo y crea los menús personalizados.
 */
function onOpen() {
  // Asegura que todas las hojas necesarias existan antes de configurar la UI
  setupWorkspace();

  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Contabilidad Automatizada')
    .addItem('Cargar XML desde PC', 'showPicker')
    .addToUi();

  ui.createMenu('Estilo')
    .addItem('Aplicar Tema Industrial Moderno', 'applyModernIndustrialTheme')
    .addToUi();
}

/**
 * Muestra el cuadro de diálogo para seleccionar archivos XML locales.
 */
function showPicker() {
  const html = HtmlService.createHtmlOutputFromFile('Picker')
    .setWidth(600)
    .setHeight(425)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Selecciona los archivos XML');
}
