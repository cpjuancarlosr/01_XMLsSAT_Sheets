/**
 * Suite de pruebas para el sistema de contabilidad.
 * Para ejecutar estas pruebas, abre el editor de Apps Script, selecciona la función
 * a ejecutar y haz clic en "Ejecutar". Revisa los `Logger.log` para ver los resultados.
 */

// --- DATOS DE PRUEBA UNITARIA ---
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

function testGeneracionPolizaIngreso() {
  Logger.log('--- Iniciando prueba: testGeneracionPolizaIngreso ---');
  const poliza = generarPolizaIngreso(mockCfdiIngreso);
  let totalDebe = poliza.reduce((sum, mov) => sum + mov[3], 0);
  let totalHaber = poliza.reduce((sum, mov) => sum + mov[4], 0);
  Logger.log(`Total Debe: ${totalDebe.toFixed(2)}, Total Haber: ${totalHaber.toFixed(2)}`);
  if (totalDebe.toFixed(2) === totalHaber.toFixed(2) && totalDebe > 0) {
    Logger.log('✅ PRUEBA SUPERADA: La póliza de ingreso está balanceada.');
  } else {
    Logger.log('❌ PRUEBA FALLIDA: La póliza de ingreso NO está balanceada.');
  }
}

function testGeneracionPolizaEgreso() {
  Logger.log('--- Iniciando prueba: testGeneracionPolizaEgreso ---');
  const poliza = generarPolizaEgreso(mockCfdiEgreso, mockCatalogoProveedores);
  let totalDebe = poliza.reduce((sum, mov) => sum + mov[3], 0);
  let totalHaber = poliza.reduce((sum, mov) => sum + mov[4], 0);
  Logger.log(`Total Debe: ${totalDebe.toFixed(2)}, Total Haber: ${totalHaber.toFixed(2)}`);
  if (totalDebe.toFixed(2) === totalHaber.toFixed(2) && totalDebe > 0) {
    Logger.log('✅ PRUEBA SUPERADA: La póliza de egreso está balanceada.');
  } else {
    Logger.log('❌ PRUEBA FALLIDA: La póliza de egreso NO está balanceada.');
  }
}

// --- PRUEBA DE INTEGRACIÓN ---

/**
 * Simula el flujo completo de procesamiento de un archivo XML local.
 * Utiliza un string de un CFDI de prueba para simular la carga.
 */
function testIntegracion_ProcessLocalXml() {
  Logger.log('--- Iniciando prueba de integración: testIntegracion_ProcessLocalXml ---');

  // Establecer un RFC de prueba para que el parser funcione
  PropertiesService.getScriptProperties().setProperty('RFC_PROPIO', 'XEXX010101000');

  const mockXmlContent = `
  <cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/3"
    Fecha="2023-01-15T10:00:00"
    TipoDeComprobante="I"
    SubTotal="100.00"
    Total="116.00"
    Moneda="MXN"
    MetodoPago="PUE"
    FormaPago="01">
    <cfdi:Emisor Rfc="EKU9003173C9" Nombre="EMPRESA EMISORA SA DE CV"/>
    <cfdi:Receptor Rfc="XEXX010101000" Nombre="NUESTRA EMPRESA SA DE CV"/>
    <cfdi:Conceptos>
      <cfdi:Concepto Cantidad="1" Descripcion="PRODUCTO" ValorUnitario="100.00" Importe="100.00"/>
    </cfdi:Conceptos>
    <cfdi:Impuestos TotalImpuestosTrasladados="16.00">
      <cfdi:Traslados>
        <cfdi:Traslado Impuesto="002" TasaOCuota="0.160000" Importe="16.00"/>
      </cfdi:Traslados>
    </cfdi:Impuestos>
    <cfdi:Complemento>
      <tfd:TimbreFiscalDigital xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" UUID="TEST-UUID-12345" />
    </cfdi:Complemento>
  </cfdi:Comprobante>`;

  // Limpiar hojas de prueba antes de ejecutar
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEgresos = ss.getSheetByName(SHEETS.REGISTRO_EGRESOS);
  const sheetDiario = ss.getSheetByName(SHEETS.LIBRO_DIARIO);

  if(sheetEgresos) {
    // No borrar encabezados
    if (sheetEgresos.getLastRow() > 1) {
      sheetEgresos.deleteRows(2, sheetEgresos.getLastRow() - 1);
    }
  }
  if(sheetDiario) {
    if (sheetDiario.getLastRow() > 1) {
      sheetDiario.deleteRows(2, sheetDiario.getLastRow() - 1);
    }
  }

  // Ejecutar el proceso con el contenido del XML de prueba
  const resultado = processLocalXmlFiles([mockXmlContent]);
  Logger.log(resultado);

  // Verificación
  const numRowsDiario = sheetDiario ? sheetDiario.getLastRow() : 0;

  if (numRowsDiario > 1) { // 1 para el encabezado
    Logger.log(`✅ PRUEBA SUPERADA: Se han añadido ${numRowsDiario - 1} filas al Libro Diario.`);
  } else {
    Logger.log('❌ PRUEBA FALLIDA: No se añadieron filas al Libro Diario.');
  }
}
