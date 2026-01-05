/**
 * @OnlyCurrentDoc
 *
 * Se ejecuta automáticamente al abrir la hoja de cálculo.
 * Crea los menús personalizados para interactuar con el sistema.
 * Esta función se ejecuta en un modo que no requiere autorización, por lo que solo debe contener
 * la creación de la UI (menús), no llamadas a funciones que modifiquen la hoja.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('CONTABILIDAD')
    .addItem('Cargar XML desde PC', 'showPicker')
    .addToUi();

  ui.createMenu('SISTEMA')
    .addItem('1. Preparar Espacio de Trabajo', 'setupWorkspace')
    .addSeparator()
    .addItem('2. Aplicar Tema Visual', 'applyModernIndustrialTheme')
    .addToUi();
}

/**
 * Muestra el cuadro de diálogo para seleccionar archivos XML locales.
 * Esta función será llamada desde el menú, por lo que tendrá los permisos necesarios.
 */
function showPicker() {
  const html = HtmlService.createHtmlOutputFromFile('Picker')
    .setWidth(600)
    .setHeight(425)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Selecciona los archivos XML');
}
