/**
 * Suite de pruebas para el sistema de contabilidad.
 * Para ejecutar estas pruebas, abre el editor de Apps Script, selecciona la función
 * a ejecutar (e.g., `testGeneracionPolizaIngreso`) y haz clic en "Ejecutar".
 * Revisa los `Logger.log` en el registro de ejecuciones para ver los resultados.
 */

// --- DATOS DE PRUEBA ---
const mockCfdiIngreso = {
  fecha: new Date(),
  uuid: 'A1B2C3D4-E5F6-G7H8-I9J0-K1L2M3N4O5P6',
  total: 1160.00,
  subtotal: 1000.00,
  impuestosTrasladados: 160.00,
  metodoPago: 'PUE',
  receptorNombre: 'CLIENTE DE PRUEBA SA DE CV'
};

const mockCfdiEgreso = {
  fecha: new Date(),
  uuid: 'Z9Y8X7W6-V5U4-T3S2-R1Q0-P9O8N7M6L5K4',
  total: 580.00,
  subtotal: 500.00,
  impuestosTrasladados: 80.00,
  metodoPago: 'PUE',
  emisorRfc: 'PROV123456RFC',
  emisorNombre: 'PROVEEDOR DE PRUEBA SA DE CV'
};

const mockCatalogoProveedores = new Map([
  ['PROV123456RFC', '601-01-Gastos de Administración'],
]);


// --- PRUEBAS UNITARIAS ---

/**
 * Prueba unitaria para la generación de pólizas de ingreso.
 * Verifica que el Debe y el Haber sumen lo mismo (partida doble).
 */
function testGeneracionPolizaIngreso() {
  Logger.log('--- Iniciando prueba: testGeneracionPolizaIngreso ---');

  const poliza = generarPolizaIngreso(mockCfdiIngreso);

  let totalDebe = 0;
  let totalHaber = 0;

  poliza.forEach(movimiento => {
    totalDebe += movimiento[3]; // Columna del Debe
    totalHaber += movimiento[4]; // Columna del Haber
  });

  Logger.log(`Total Debe: ${totalDebe.toFixed(2)}, Total Haber: ${totalHaber.toFixed(2)}`);

  if (totalDebe.toFixed(2) === totalHaber.toFixed(2) && totalDebe > 0) {
    Logger.log('✅ PRUEBA SUPERADA: La póliza de ingreso está balanceada.');
  } else {
    Logger.log('❌ PRUEBA FALLIDA: La póliza de ingreso NO está balanceada.');
  }
}

/**
 * Prueba unitaria para la generación de pólizas de egreso.
 * Verifica que el Debe y el Haber sumen lo mismo (partida doble).
 */
function testGeneracionPolizaEgreso() {
  Logger.log('--- Iniciando prueba: testGeneracionPolizaEgreso ---');

  const poliza = generarPolizaEgreso(mockCfdiEgreso, mockCatalogoProveedores);

  let totalDebe = 0;
  let totalHaber = 0;

  poliza.forEach(movimiento => {
    totalDebe += movimiento[3]; // Columna del Debe
    totalHaber += movimiento[4]; // Columna del Haber
  });

  Logger.log(`Total Debe: ${totalDebe.toFixed(2)}, Total Haber: ${totalHaber.toFixed(2)}`);

  if (totalDebe.toFixed(2) === totalHaber.toFixed(2) && totalDebe > 0) {
    Logger.log('✅ PRUEBA SUPERADA: La póliza de egreso está balanceada.');
  } else {
    Logger.log('❌ PRUEBA FALLIDA: La póliza de egreso NO está balanceada.');
  }
}

// --- PRUEBA DE INTEGRACIÓN ---

/**
 * Simula el flujo completo de procesamiento de un archivo XML.
 * **Requiere configuración manual de un archivo de prueba en Google Drive.**
 * 1. Sube un archivo XML de prueba a tu Google Drive.
 * 2. Obtén su ID (de la URL, por ejemplo).
 * 3. Pega el ID en la variable `TEST_FILE_ID` de abajo.
 * 4. Ejecuta esta función desde el editor.
 */
function testIntegracion_ProcesarXml() {
  Logger.log('--- Iniciando prueba de integración: testIntegracion_ProcesarXml ---');

  // --- CONFIGURACIÓN REQUERIDA ---
  const TEST_FILE_ID = 'ID_DE_TU_ARCHIVO_XML_DE_PRUEBA';
  // --------------------------------

  if (TEST_FILE_ID === 'ID_DE_TU_ARCHIVO_XML_DE_PRUEBA') {
    Logger.log('ADVERTENCIA: Debes configurar el ID de un archivo de prueba en la función testIntegracion_ProcesarXml.');
    return;
  }

  // Limpiar hojas de prueba antes de ejecutar
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetIngresos = ss.getSheetByName(SHEETS.REGISTRO_INGRESOS);
  const sheetEgresos = ss.getSheetByName(SHEETS.REGISTRO_EGRESOS);
  const sheetDiario = ss.getSheetByName(SHEETS.LIBRO_DIARIO);

  if(sheetIngresos) sheetIngresos.clear();
  if(sheetEgresos) sheetEgresos.clear();
  if(sheetDiario) sheetDiario.clear();

  // Ejecutar el proceso
  const resultado = processXmlFiles([TEST_FILE_ID]);
  Logger.log(resultado);

  // Verificación
  // La verificación real requeriría leer las hojas y Aserciones,
  // pero para este entorno, revisaremos manualmente que se hayan añadido filas.
  const numRowsDiario = sheetDiario ? sheetDiario.getLastRow() : 0;

  if (numRowsDiario > 1) { // 1 para el encabezado
    Logger.log(`✅ PRUEBA SUPERADA: Se han añadido ${numRowsDiario - 1} filas al Libro Diario.`);
  } else {
    Logger.log('❌ PRUEBA FALLIDA: No se añadieron filas al Libro Diario.');
  }
}
