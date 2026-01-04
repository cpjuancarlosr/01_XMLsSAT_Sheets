/**
 * Suite de pruebas para el sistema de contabilidad avanzado.
 * Para ejecutar, selecciona la función de prueba y haz clic en "Ejecutar".
 */

// --- CONFIGURACIÓN Y MOCKS ---

// Mock de la hoja de CxC/P para simular la búsqueda de facturas
const mockSheetCxC = {
  data: [
    ['UUID Factura', 'Fecha Factura', 'Tipo', 'RFC', 'Nombre', 'Total Factura', 'Saldo Pendiente', 'IVA Pendiente', 'Estado'],
    ['EGRESO-PPD-UUID', new Date(), 'Egreso', 'PROV123456RFC', 'PROVEEDOR PPD', 1160.00, 1160.00, 160.00, 'Pendiente']
  ],
  getSheetByName: function(name) { return this; },
  getRange: function(row, col, numRows, numCols) { return this; },
  getValues: function() { return [this.data[1]]; },
};

// Mock de funciones globales de Apps Script para aislamiento
const SpreadsheetApp = {
  getActiveSpreadsheet: function() { return mockSheetCxC; }
};

const CUENTAS_PREDETERMINADAS = {
  BANCOS: '102-Bancos', CLIENTES: '105-Clientes', PROVEEDORES: '201-Proveedores',
  VENTAS: '401-Ventas', GASTOS_GENERALES: '601-Gastos Generales', IVA_ACREDITABLE: '118-IVA Acreditable',
  IVA_POR_ACREDITAR: '119-IVA por Acreditar', IVA_TRASLADADO: '208-IVA Trasladado', IVA_POR_TRASLADAR: '209-IVA por Trasladar',
};

// --- DATOS DE PRUEBA ---

const mockCfdiEgresoPUE = {
  fecha: new Date(), uuid: 'EGRESO-PUE-UUID', total: 1160.00, subtotal: 1000.00, impuestosTrasladados: 160.00,
  tipoContable: 'Egreso', emisorNombre: 'PROVEEDOR PUE', conceptos: [{claveProdServ: '001'}], emisorRfc: 'PROVPUE123'
};

const mockCfdiEgresoPPD = {
  fecha: new Date(), uuid: 'EGRESO-PPD-UUID', total: 1160.00, subtotal: 1000.00, impuestosTrasladados: 160.00,
  tipoContable: 'Egreso', emisorNombre: 'PROVEEDOR PPD', conceptos: [{claveProdServ: '002'}], emisorRfc: 'PROVPPD456'
};

const mockCfdiPagoEgreso = {
  fecha: new Date(), uuid: 'PAGO-EGRESO-UUID', tipoComprobante: 'P',
  documentosRelacionados: [{ uuidRelacionado: 'EGRESO-PPD-UUID', montoPagado: 1160.00 }]
};

// --- SUITE DE PRUEBAS ---

function ejecutarTodasLasPruebas() {
  testGeneracionPolizaPUE_Egreso();
  testGeneracionPolizaPPD_Egreso();
  testGeneracionPolizaPago_Egreso();
}

function testGeneracionPolizaPUE_Egreso() {
  Logger.log('--- Probando Póliza PUE (Egreso) ---');
  const poliza = generarPolizaPUE(mockCfdiEgresoPUE, new Map(), new Map());
  validarPartidaDoble(poliza, 'PUE Egreso');
}

function testGeneracionPolizaPPD_Egreso() {
  Logger.log('--- Probando Póliza PPD (Egreso) ---');
  const poliza = generarPolizaPPD(mockCfdiEgresoPPD, new Map(), new Map());
  validarPartidaDoble(poliza, 'PPD Egreso');

  const ivaMov = poliza.find(mov => mov[1] === CUENTAS_PREDETERMINADAS.IVA_POR_ACREDITAR);
  if (ivaMov) {
    Logger.log('✅ PRUEBA SUPERADA: Se utilizó correctamente la cuenta puente de IVA por Acreditar.');
  } else {
    Logger.log('❌ PRUEBA FALLIDA: No se encontró el movimiento a la cuenta puente de IVA.');
  }
}

function testGeneracionPolizaPago_Egreso() {
  Logger.log('--- Probando Póliza de Pago (Egreso) ---');

  // Sobrescribir la función global para esta prueba específica
  buscarFacturaEnCxC = function(uuid) {
    return { tipo: 'Egreso', totalFactura: 1160.00, ivaPendiente: 160.00 };
  }
  // Simular que no hace nada para no depender de la hoja real
  liquidarCxC_CxP = function(uuid, monto) {};

  const poliza = generarPolizaPago(mockCfdiPagoEgreso);
  validarPartidaDoble(poliza, 'Pago Egreso');

  const ivaReclasificado = poliza.find(mov => mov[1] === CUENTAS_PREDETERMINADAS.IVA_ACREDITABLE && mov[2].includes('Reclasificación'));
  if (ivaReclasificado) {
    Logger.log('✅ PRUEBA SUPERADA: Se reclasificó correctamente el IVA de puente a definitiva.');
  } else {
    Logger.log('❌ PRUEBA FALLIDA: No se encontró la reclasificación del IVA.');
  }
}

// --- FUNCIÓN DE VALIDACIÓN AUXILIAR ---
function validarPartidaDoble(poliza, nombrePrueba) {
  const totalDebe = poliza.reduce((sum, mov) => sum + mov[3], 0);
  const totalHaber = poliza.reduce((sum, mov) => sum + mov[4], 0);

  Logger.log(`[${nombrePrueba}] Total Debe: ${totalDebe.toFixed(2)}, Total Haber: ${totalHaber.toFixed(2)}`);

  if (Math.abs(totalDebe - totalHaber) < 0.01 && totalDebe > 0) {
    Logger.log(`✅ PRUEBA SUPERADA: La póliza [${nombrePrueba}] está balanceada.`);
  } else {
    Logger.log(`❌ PRUEBA FALLIDA: La póliza [${nombrePrueba}] NO está balanceada.`);
  }
}
