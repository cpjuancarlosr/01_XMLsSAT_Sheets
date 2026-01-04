/**
 * @OnlyCurrentDoc
 *
 * Función principal que se llama desde la interfaz de usuario del selector de archivos.
 * Orquesta el proceso de leer, parsear, validar y registrar cada CFDI XML.
 *
 * @param {string[]} fileIds Array de IDs de los archivos de Google Drive seleccionados.
 * @returns {string} Un mensaje de resumen para el usuario.
 */
function processXmlFiles(fileIds) {
  if (!fileIds || fileIds.length === 0) {
    return "No se seleccionaron archivos.";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetIngresos = ss.getSheetByName(SHEETS.REGISTRO_INGRESOS);
  const sheetEgresos = ss.getSheetByName(SHEETS.REGISTRO_EGRESOS);

  // Cargar catálogos una sola vez para eficiencia
  const catalogoProveedores = getCatalog(SHEETS.CATALOGO_PROVEEDORES);

  let procesados = 0;
  let duplicados = 0;
  let errores = 0;

  fileIds.forEach(id => {
    try {
      const file = DriveApp.getFileById(id);
      const xmlTexto = file.getBlob().getDataAsString('UTF-8');

      const cfdiData = parseCfdi(xmlTexto);

      // 1. Validación de duplicados por UUID
      const sheetDestino = cfdiData.tipoContable === 'Ingreso' ? sheetIngresos : sheetEgresos;
      if (findRowByUuid(cfdiData.uuid, sheetDestino)) {
        Logger.log(`CFDI duplicado omitido: ${cfdiData.uuid}`);
        duplicados++;
        return; // Continuar con el siguiente archivo
      }

      // 2. Generación de la póliza contable
      let poliza;
      if (cfdiData.tipoContable === 'Ingreso') {
        poliza = generarPolizaIngreso(cfdiData);
      } else {
        poliza = generarPolizaEgreso(cfdiData, catalogoProveedores);
      }

      // 3. Escritura en Hojas de Cálculo
      writeCfdiData(cfdiData);
      writePoliza(poliza);

      procesados++;

    } catch (e) {
      Logger.log(`Error procesando el archivo con ID ${id}: ${e.message} \n ${e.stack}`);
      errores++;
    }
  });

  // Devolver un resumen para la UI
  let resumen = `Proceso completado. <br/>`;
  resumen += `- ${procesados} XML procesados correctamente. <br/>`;
  resumen += `- ${duplicados} duplicados omitidos. <br/>`;
  resumen += `- ${errores} errores encontrados.`;

  return resumen;
}
