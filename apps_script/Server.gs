/**
 * @OnlyCurrentDoc
 *
 * Función principal llamada desde la UI. Procesa el contenido de los archivos XML.
 * Orquesta el proceso de parsear, validar y registrar cada CFDI, ahora con lógica avanzada.
 *
 * @param {string[]} xmlContents Array de strings, donde cada string es el contenido de un archivo XML.
 * @returns {string} Un mensaje de resumen para el usuario.
 */
function processLocalXmlFiles(xmlContents) {
  if (!xmlContents || xmlContents.length === 0) {
    return "No se procesaron archivos.";
  }

  // Cargar todos los catálogos necesarios al inicio
  const catalogoProveedores = getCatalog(SHEETS.CATALOGO_PROVEEDORES);
  const catalogoProdServ = getCatalog(SHEETS.CATALOGO_PROD_SERV);

  let cfdis = [];
  try {
    cfdis = xmlContents.map(xml => parseCfdi(xml));
  } catch (e) {
    return `Error crítico durante el parseo inicial: ${e.message}. Verifique los archivos.`;
  }

  // Ordenar los CFDI por fecha para procesar facturas antes que sus pagos
  cfdis.sort((a, b) => a.fecha - b.fecha);

  let procesados = 0;
  let duplicados = 0;
  let errores = 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  cfdis.forEach((cfdiData, index) => {
    try {
      // Validación de duplicados por UUID en la hoja correspondiente
      let sheetDestino;
      if (cfdiData.tipoComprobante === 'I') sheetDestino = ss.getSheetByName(SHEETS.REGISTRO_INGRESOS);
      else if (cfdiData.tipoComprobante === 'E') sheetDestino = ss.getSheetByName(SHEETS.REGISTRO_EGRESOS);
      else if (cfdiData.tipoComprobante === 'P') sheetDestino = ss.getSheetByName(SHEETS.REGISTRO_PAGOS);

      if (findRowByUuid(cfdiData.uuid, sheetDestino)) {
        Logger.log(`CFDI duplicado omitido: ${cfdiData.uuid}`);
        duplicados++;
        return;
      }

      // Escritura del registro del CFDI
      writeCfdiData(cfdiData);

      // Creación de la póliza contable
      crearPolizaDesdeCfdi(cfdiData, catalogoProveedores, catalogoProdServ);

      procesados++;

    } catch (e) {
      Logger.log(`Error procesando el CFDI con UUID ${cfdiData.uuid || `(archivo #${index + 1})`}: ${e.message} \n ${e.stack}`);
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
