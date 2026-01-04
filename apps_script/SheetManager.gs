// Nombres de las hojas de cálculo para centralizar la configuración
const SHEETS = {
  REGISTRO_INGRESOS: 'CFDI Ingresos',
  REGISTRO_EGRESOS: 'CFDI Egresos',
  LIBRO_DIARIO: 'LibroDiario',
  CATALOGO_PROVEEDORES: 'CatalogoProveedores',
  CATALOGO_CUENTAS: 'CatalogoCuentas',
};

const HEADERS_CFDI = [
  'UUID', 'Fecha', 'Tipo Comprobante', 'Emisor RFC', 'Emisor Nombre',
  'Receptor RFC', 'Receptor Nombre', 'Moneda', 'Subtotal', 'IVA Trasladado', 'Total',
  'Forma de Pago', 'Método de Pago', 'Cuenta Contable (Override)'
];

const HEADERS_DIARIO = [
  'Folio', 'Fecha', 'Cuenta', 'Descripción', 'Debe', 'Haber', 'UUID Ref'
];

/**
 * Busca una fila en una hoja de cálculo dado un UUID en la primera columna.
 * @param {string} uuid El UUID a buscar.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja en la que se buscará.
 * @returns {number|null} El número de fila (1-indexed) si se encuentra, de lo contrario null.
 */
function findRowByUuid(uuid, sheet) {
  if (!sheet) return null;
  const range = sheet.getRange("A:A");
  const textFinder = range.createTextFinder(uuid);
  const found = textFinder.findNext();
  return found ? found.getRow() : null;
}

/**
 * Lee un catálogo de una hoja y lo convierte en un mapa para fácil acceso.
 * Asume un formato de dos columnas: Col A (Clave), Col B (Valor).
 * @param {string} catalogName El nombre de la hoja que contiene el catálogo.
 * @returns {Map<string, string>} Un mapa con los datos del catálogo.
 */
function getCatalog(catalogName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(catalogName);
  const catalogMap = new Map();

  if (sheet) {
    const data = sheet.getDataRange().getValues();
    // Empezar en 1 para saltar los encabezados
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { // Asegurarse de que la clave no esté vacía
        catalogMap.set(data[i][0].toString().trim(), data[i][1].toString().trim());
      }
    }
  } else {
    Logger.log(`ADVERTENCIA: No se encontró la hoja de catálogo "${catalogName}".`);
  }
  return catalogMap;
}

/**
 * Escribe los datos de un CFDI procesado en la hoja de registro correspondiente.
 * @param {object} cfdiData El objeto con los datos del CFDI parseado.
 */
function writeCfdiData(cfdiData) {
  const sheetName = cfdiData.tipoContable === 'Ingreso' ? SHEETS.REGISTRO_INGRESOS : SHEETS.REGISTRO_EGRESOS;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(HEADERS_CFDI);
    sheet.setFrozenRows(1);
  }

  // Formatear los datos para la fila
  const rowData = [
    cfdiData.uuid,
    cfdiData.fecha,
    cfdiData.tipoComprobante,
    cfdiData.emisorRfc,
    cfdiData.emisorNombre,
    cfdiData.receptorRfc,
    cfdiData.receptorNombre,
    cfdiData.moneda,
    cfdiData.subtotal,
    cfdiData.impuestosTrasladados,
    cfdiData.total,
    cfdiData.formaPago,
    cfdiData.metodoPago,
    '' // Celda vacía para el override manual
  ];

  sheet.appendRow(rowData);
}

/**
 * Escribe las líneas de una póliza contable en el Libro Diario.
 * @param {Array<Array<any>>} polizaItems Un array de filas, donde cada fila representa un movimiento.
 */
function writePoliza(polizaItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.LIBRO_DIARIO);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.LIBRO_DIARIO);
    sheet.appendRow(HEADERS_DIARIO);
    sheet.setFrozenRows(1);
  }

  // Obtener el último folio para generar el siguiente
  const lastRow = sheet.getLastRow();
  const lastFolio = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : 0;
  const newFolio = Number(lastFolio) + 1;

  const rowsToWrite = polizaItems.map(item => [newFolio, ...item]);

  // Escribir todas las filas de la póliza de una vez
  sheet.getRange(sheet.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
}
