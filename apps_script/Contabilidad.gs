// Configuración de cuentas contables. Se añaden las cuentas puente para IVA de PPD.
const CUENTAS_PREDETERMINADAS = {
  BANCOS: '102-Bancos',
  CLIENTES: '105-Clientes',
  PROVEEDORES: '201-Proveedores',
  VENTAS: '401-Ventas',
  GASTOS_GENERALES: '601-Gastos Generales',
  IVA_ACREDITABLE: '118-IVA Acreditable',      // Definitiva (PUE y Pagos)
  IVA_POR_ACREDITAR: '119-IVA por Acreditar',    // Puente (PPD)
  IVA_TRASLADADO: '208-IVA Trasladado',        // Definitiva (PUE y Pagos)
  IVA_POR_TRASLADAR: '209-IVA por Trasladar',      // Puente (PPD)
};

/**
 * Orquestador principal. Analiza el CFDI y llama a la función de póliza correspondiente.
 * @param {object} cfdiData - Datos del CFDI parseado.
 * @param {Map<string, string>} catalogoProveedores - Mapa de RFC de proveedor a cuenta de gasto.
 * @param {Map<string, string>} catalogoProdServ - Mapa de ClaveProdServ a cuenta de resultado.
 * @returns {void}
 */
function crearPolizaDesdeCfdi(cfdiData, catalogoProveedores, catalogoProdServ) {
  let poliza;

  switch (cfdiData.tipoComprobante) {
    case 'I': // Ingreso
    case 'E': // Egreso
      if (cfdiData.metodoPago === 'PPD') {
        poliza = generarPolizaPPD(cfdiData, catalogoProveedores, catalogoProdServ);
        agregarCxC_CxP(cfdiData); // Registrar la cuenta por cobrar/pagar
      } else {
        poliza = generarPolizaPUE(cfdiData, catalogoProveedores, catalogoProdServ);
      }
      break;
    case 'P': // Pago
      poliza = generarPolizaPago(cfdiData);
      // La liquidación de la CxC/CxP se maneja dentro de generarPolizaPago
      break;
    default:
      Logger.log(`Tipo de comprobante ${cfdiData.tipoComprobante} no soportado para póliza automática.`);
      return;
  }

  if (poliza && poliza.length > 0) {
    writePoliza(poliza);
  }
}

/**
 * Genera la póliza para una factura PUE (Pago en Una Sola Exhibición).
 * El IVA se considera acreditable/trasladado inmediatamente.
 */
function generarPolizaPUE(cfdiData, catalogoProveedores, catalogoProdServ) {
  const { fecha, uuid, total, subtotal, impuestosTrasladados, tipoContable, emisorNombre, receptorNombre } = cfdiData;
  const poliza = [];

  if (tipoContable === 'Ingreso') {
    const cuentaIngreso = obtenerCuentaDeResultado(cfdiData, catalogoProdServ);
    poliza.push(
      [fecha, CUENTAS_PREDETERMINADAS.BANCOS, `Ingreso PUE de ${receptorNombre}`, total, 0, uuid],
      [fecha, cuentaIngreso, `Ingreso PUE de ${receptorNombre}`, 0, subtotal, uuid],
      [fecha, CUENTAS_PREDETERMINADAS.IVA_TRASLADADO, `IVA de Ingreso PUE de ${receptorNombre}`, 0, impuestosTrasladados, uuid]
    );
  } else { // Egreso
    const cuentaGasto = obtenerCuentaDeResultado(cfdiData, catalogoProdServ, catalogoProveedores);
    poliza.push(
      [fecha, cuentaGasto, `Egreso PUE de ${emisorNombre}`, subtotal, 0, uuid],
      [fecha, CUENTAS_PREDETERMINADAS.IVA_ACREDITABLE, `IVA de Egreso PUE de ${emisorNombre}`, impuestosTrasladados, 0, uuid],
      [fecha, CUENTAS_PREDETERMINADAS.BANCOS, `Egreso PUE de ${emisorNombre}`, 0, total, uuid]
    );
  }
  return poliza;
}

/**
 * Genera la póliza para una factura PPD (Pago en Parcialidades o Diferido).
 * Utiliza cuentas puente para el IVA, ya que aún no es efectivamente cobrado/pagado.
 */
function generarPolizaPPD(cfdiData, catalogoProveedores, catalogoProdServ) {
  const { fecha, uuid, total, subtotal, impuestosTrasladados, tipoContable, emisorNombre, receptorNombre } = cfdiData;
  const poliza = [];

  if (tipoContable === 'Ingreso') {
    const cuentaIngreso = obtenerCuentaDeResultado(cfdiData, catalogoProdServ);
    poliza.push(
      [fecha, CUENTAS_PREDETERMINADAS.CLIENTES, `Ingreso PPD de ${receptorNombre}`, total, 0, uuid],
      [fecha, cuentaIngreso, `Ingreso PPD de ${receptorNombre}`, 0, subtotal, uuid],
      [fecha, CUENTAS_PREDETERMINADAS.IVA_POR_TRASLADAR, `IVA PPD de Ingreso de ${receptorNombre}`, 0, impuestosTrasladados, uuid]
    );
  } else { // Egreso
    const cuentaGasto = obtenerCuentaDeResultado(cfdiData, catalogoProdServ, catalogoProveedores);
    poliza.push(
      [fecha, cuentaGasto, `Egreso PPD de ${emisorNombre}`, subtotal, 0, uuid],
      [fecha, CUENTAS_PREDETERMINADAS.IVA_POR_ACREDITAR, `IVA PPD de Egreso de ${emisorNombre}`, impuestosTrasladados, 0, uuid],
      [fecha, CUENTAS_PREDETERMINADAS.PROVEEDORES, `Egreso PPD de ${emisorNombre}`, 0, total, uuid]
    );
  }
  return poliza;
}

/**
 * Genera la póliza para un Complemento de Pago.
 * Liquida la cuenta por cobrar/pagar y reclasifica el IVA de las cuentas puente a las definitivas.
 */
function generarPolizaPago(cfdiData) {
  const { fecha, uuid, montoTotalPago, documentosRelacionados } = cfdiData;
  const poliza = [];

  for (const doc of documentosRelacionados) {
    const infoFactura = buscarFacturaEnCxC(doc.uuidRelacionado);
    if (!infoFactura) {
      Logger.log(`ADVERTENCIA: No se encontró la factura ${doc.uuidRelacionado} en Cuentas por Cobrar/Pagar.`);
      continue;
    }

    const proporcionPago = doc.montoPagado / infoFactura.totalFactura;
    const ivaReclasificar = infoFactura.ivaPendiente * proporcionPago;

    liquidarCxC_CxP(doc.uuidRelacionado, doc.montoPagado);

    if (infoFactura.tipo === 'Ingreso') {
      poliza.push(
        [fecha, CUENTAS_PREDETERMINADAS.BANCOS, `Cobro factura ${doc.uuidRelacionado}`, doc.montoPagado, 0, uuid],
        [fecha, CUENTAS_PREDETERMINADAS.CLIENTES, `Cobro factura ${doc.uuidRelacionado}`, 0, doc.montoPagado, uuid],
        [fecha, CUENTAS_PREDETERMINADAS.IVA_POR_TRASLADAR, `Reclasificación IVA cobrado de ${doc.uuidRelacionado}`, ivaReclasificar, 0, uuid],
        [fecha, CUENTAS_PREDETERMINADAS.IVA_TRASLADADO, `Reclasificación IVA cobrado de ${doc.uuidRelacionado}`, 0, ivaReclasificar, uuid]
      );
    } else { // Egreso
      poliza.push(
        [fecha, CUENTAS_PREDETERMINADAS.PROVEEDORES, `Pago factura ${doc.uuidRelacionado}`, doc.montoPagado, 0, uuid],
        [fecha, CUENTAS_PREDETERMINADAS.BANCOS, `Pago factura ${doc.uuidRelacionado}`, 0, doc.montoPagado, uuid],
        [fecha, CUENTAS_PREDETERMINADAS.IVA_ACREDITABLE, `Reclasificación IVA pagado de ${doc.uuidRelacionado}`, ivaReclasificar, 0, uuid],
        [fecha, CUENTAS_PREDETERMINADAS.IVA_POR_ACREDITAR, `Reclasificación IVA pagado de ${doc.uuidRelacionado}`, 0, ivaReclasificar, uuid]
      );
    }
  }
  return poliza;
}

/**
 * Determina la cuenta de resultado (Ingreso/Gasto) a utilizar.
 * Prioriza el catálogo de productos/servicios, y si no encuentra, usa el de proveedores.
 */
function obtenerCuentaDeResultado(cfdiData, catalogoProdServ, catalogoProveedores) {
  // Intenta encontrar una cuenta por clave de producto/servicio primero
  for (const concepto of cfdiData.conceptos) {
    if (catalogoProdServ.has(concepto.claveProdServ)) {
      return catalogoProdServ.get(concepto.claveProdServ);
    }
  }
  // Si no, si es un egreso, busca por proveedor
  if (cfdiData.tipoContable === 'Egreso' && catalogoProveedores.has(cfdiData.emisorRfc)) {
    return catalogoProveedores.get(cfdiData.emisorRfc);
  }
  // Como último recurso, devuelve la cuenta por defecto
  return cfdiData.tipoContable === 'Ingreso' ? CUENTAS_PREDETERMINADAS.VENTAS : CUENTAS_PREDETERMINADAS.GASTOS_GENERALES;
}
