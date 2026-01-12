// Nombres de las hojas de cálculo para centralizar la configuración
const SHEETS = {
  REGISTRO_INGRESOS: 'CFDI Ingresos',
  REGISTRO_EGRESOS: 'CFDI Egresos',
  REGISTRO_PAGOS: 'CFDI Pagos',
  LIBRO_DIARIO: 'LibroDiario',
  CATALOGO_PROVEEDORES: 'CatalogoProveedores',
  CATALOGO_CUENTAS: 'CatalogoCuentas',
  CATALOGO_PROD_SERV: 'CatalogoProdServ',
  CXC_CXP: 'CuentasPorCobrarPagar',
};

// ... (Definición de HEADERS)

const HEADERS_CFDI_IE = [
  'UUID', 'Fecha', 'Tipo Comprobante', 'Emisor RFC', 'Emisor Nombre',
  'Receptor RFC', 'Receptor Nombre', 'Moneda', 'Subtotal', 'IVA Trasladado', 'Total',
  'Forma de Pago', 'Método de Pago'
];

const HEADERS_CFDI_P = [
  'UUID Pago', 'Fecha Pago', 'Monto Pago', 'Moneda Pago', 'UUID Factura Relacionada', 'Parcialidad'
];

const HEADERS_CXC_CXP = [
  'UUID Factura', 'Fecha Factura', 'Tipo', 'RFC Cliente/Proveedor', 'Nombre Cliente/Proveedor',
  'Total Factura', 'Saldo Pendiente', 'IVA Pendiente', 'Estado'
];

const HEADERS_DIARIO = [
  'Folio', 'Fecha', 'Cuenta', 'Descripción', 'Debe', 'Haber', 'UUID Ref'
];

const HEADERS_CATALOGO_CUENTAS = ['Número de Cuenta', 'Nombre de Cuenta'];
const HEADERS_CATALOGO_PROVEEDORES = ['RFC del Proveedor', 'Cuenta Contable Asignada'];
const HEADERS_CATALOGO_PROD_SERV = ['ClaveProdServ SAT', 'Cuenta Contable Asignada'];


// --- FUNCIONES DE ESCRITURA Y LECTURA ---

/**
 * Orquestador para escribir los datos de cualquier tipo de CFDI en su hoja correspondiente.
 */
function writeCfdiData(cfdiData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet, headers, rowData, sheetName;

  switch (cfdiData.tipoComprobante) {
    case 'I':
      sheetName = SHEETS.REGISTRO_INGRESOS;
      sheet = ss.getSheetByName(sheetName);
      headers = HEADERS_CFDI_IE;
      rowData = [[
        cfdiData.uuid, cfdiData.fecha, cfdiData.tipoComprobante, cfdiData.emisorRfc, cfdiData.emisorNombre,
        cfdiData.receptorRfc, cfdiData.receptorNombre, cfdiData.moneda, cfdiData.subtotal,
        cfdiData.impuestosTrasladados, cfdiData.total, cfdiData.formaPago, cfdiData.metodoPago
      ]];
      break;
    case 'E':
      sheetName = SHEETS.REGISTRO_EGRESOS;
      sheet = ss.getSheetByName(sheetName);
      headers = HEADERS_CFDI_IE;
      rowData = [[
        cfdiData.uuid, cfdiData.fecha, cfdiData.tipoComprobante, cfdiData.emisorRfc, cfdiData.emisorNombre,
        cfdiData.receptorRfc, cfdiData.receptorNombre, cfdiData.moneda, cfdiData.subtotal,
        cfdiData.impuestosTrasladados, cfdiData.total, cfdiData.formaPago, cfdiData.metodoPago
      ]];
      break;
    case 'P':
      sheetName = SHEETS.REGISTRO_PAGOS;
      sheet = ss.getSheetByName(sheetName);
      headers = HEADERS_CFDI_P;
      rowData = cfdiData.documentosRelacionados.map(doc => [
        cfdiData.uuid, cfdiData.fecha, doc.montoPagado, cfdiData.moneda, doc.uuidRelacionado, doc.parcialidad
      ]);
      break;
    default:
      return; // No hacer nada si el tipo no es soportado
  }

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rowData.length, rowData[0].length).setValues(rowData);
}

/**
 * Añade una nueva factura PPD a la hoja de Cuentas por Cobrar/Pagar.
 */
function agregarCxC_CxP(cfdiData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.CXC_CXP);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.CXC_CXP);
    sheet.appendRow(HEADERS_CXC_CXP);
    sheet.setFrozenRows(1);
  }

  const partnerRfc = cfdiData.tipoContable === 'Ingreso' ? cfdiData.receptorRfc : cfdiData.emisorRfc;
  const partnerName = cfdiData.tipoContable === 'Ingreso' ? cfdiData.receptorNombre : cfdiData.emisorNombre;

  sheet.appendRow([
    cfdiData.uuid,
    cfdiData.fecha,
    cfdiData.tipoContable,
    partnerRfc,
    partnerName,
    cfdiData.total,
    cfdiData.total, // Saldo inicial es el total
    cfdiData.impuestosTrasladados,
    'Pendiente'
  ]);
}

/**
 * Busca una factura en la hoja de CxC/CxP y devuelve sus datos.
 * @param {string} uuid - El UUID de la factura a buscar.
 * @returns {object|null} Un objeto con los datos de la factura o null si no se encuentra.
 */
function buscarFacturaEnCxC(uuid) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CXC_CXP);
  if (!sheet) return null;
  const row = findRowByUuid(uuid, sheet);
  if (!row) return null;

  const data = sheet.getRange(row, 1, 1, HEADERS_CXC_CXP.length).getValues()[0];
  return {
    uuid: data[0],
    tipo: data[2],
    totalFactura: parseFloat(data[5]),
    saldoPendiente: parseFloat(data[6]),
    ivaPendiente: parseFloat(data[7]),
  };
}

/**
 * Actualiza el saldo de una factura en la hoja de CxC/CxP.
 * @param {string} uuid - El UUID de la factura a liquidar.
 * @param {number} montoPagado - El monto del pago recibido.
 */
function liquidarCxC_CxP(uuid, montoPagado) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CXC_CXP);
  if (!sheet) return;
  const row = findRowByUuid(uuid, sheet);
  if (!row) return;

  const saldoCell = sheet.getRange(row, 7); // Columna G: Saldo Pendiente
  const estadoCell = sheet.getRange(row, 9); // Columna I: Estado

  const saldoActual = parseFloat(saldoCell.getValue());
  const nuevoSaldo = saldoActual - montoPagado;

  saldoCell.setValue(nuevoSaldo.toFixed(2));

  if (nuevoSaldo <= 0.01) { // Margen de centavos
    estadoCell.setValue('Pagada');
  }
}


/**
 * Escribe las líneas de una póliza contable en el Libro Diario.
 * (Función sin cambios, pero necesaria aquí)
 */
function writePoliza(polizaItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.LIBRO_DIARIO);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.LIBRO_DIARIO);
    sheet.appendRow(HEADERS_DIARIO);
    sheet.setFrozenRows(1);
  }

  const lastRow = sheet.getLastRow();
  const lastFolio = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : 0;
  const newFolio = Number(lastFolio) + 1;

  const rowsToWrite = polizaItems.map(item => [newFolio, ...item]);

  sheet.getRange(sheet.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
}

/**
 * Busca una fila por UUID en la columna A.
 * (Función sin cambios, pero necesaria aquí)
 */
function findRowByUuid(uuid, sheet) {
  if (!sheet || !uuid) return null;
  const range = sheet.getRange("A:A");
  const textFinder = range.createTextFinder(uuid);
  const found = textFinder.findNext();
  return found ? found.getRow() : null;
}

/**
 * Lee un catálogo de una hoja.
 * (Función sin cambios, pero necesaria aquí)
 */
function getCatalog(catalogName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(catalogName);
  const catalogMap = new Map();

  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        catalogMap.set(data[i][0].toString().trim(), data[i][1].toString().trim());
      }
    }
  }
  return catalogMap;
}

/**
 * Verifica y crea todas las hojas de cálculo necesarias para el sistema si no existen.
 * Esta función asegura que el espacio de trabajo esté correctamente configurado al abrir el archivo.
 */
function setupWorkspace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = {
    [SHEETS.CATALOGO_CUENTAS]: HEADERS_CATALOGO_CUENTAS,
    [SHEETS.CATALOGO_PROVEEDORES]: HEADERS_CATALOGO_PROVEEDORES,
    [SHEETS.CATALOGO_PROD_SERV]: HEADERS_CATALOGO_PROD_SERV,
    [SHEETS.CXC_CXP]: HEADERS_CXC_CXP,
    [SHEETS.LIBRO_DIARIO]: HEADERS_DIARIO,
    [SHEETS.REGISTRO_INGRESOS]: HEADERS_CFDI_IE,
    [SHEETS.REGISTRO_EGRESOS]: HEADERS_CFDI_IE,
    [SHEETS.REGISTRO_PAGOS]: HEADERS_CFDI_P,
  };

  for (const sheetName in requiredSheets) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = requiredSheets[sheetName];
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      Logger.log(`Hoja "${sheetName}" creada.`);
    }
  }
}
