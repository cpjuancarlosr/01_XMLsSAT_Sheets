// Objeto para centralizar la configuración de cuentas contables predeterminadas.
// Facilita la personalización sin tener que modificar la lógica de las funciones.
const CUENTAS_PREDETERMINADAS = {
  BANCOS: '102-Bancos',
  CLIENTES: '105-Clientes',
  PROVEEDORES: '201-Proveedores',
  VENTAS: '401-Ventas',
  IVA_TRASLADADO: '208-IVA Trasladado',
  IVA_ACREDITABLE: '118-IVA Acreditable',
  GASTOS_GENERALES: '601-Gastos Generales',
};

/**
 * Genera una póliza de diario para un CFDI de tipo Ingreso.
 * La lógica de negocio determina las cuentas a afectar (Bancos/Clientes vs Ventas/IVA).
 * @param {object} cfdiData Datos del CFDI parseado.
 * @returns {Array<Array<any>>} Un array de filas que representan los movimientos de la póliza.
 */
function generarPolizaIngreso(cfdiData) {
  const { fecha, uuid, total, subtotal, impuestosTrasladados, metodoPago, receptorNombre } = cfdiData;
  const descripcion = `Ingreso según factura a ${receptorNombre}`;

  // Se usan las cuentas del objeto de configuración.
  const cuentaPrincipal = metodoPago === 'PUE' ? CUENTAS_PREDETERMINADAS.BANCOS : CUENTAS_PREDETERMINADAS.CLIENTES;
  const cuentaVentas = CUENTAS_PREDETERMINADAS.VENTAS;
  const cuentaIvaTrasladado = CUENTAS_PREDETERMINADAS.IVA_TRASLADADO;

  const poliza = [];

  // 1. Cargo a Bancos o Clientes por el total
  poliza.push([
    fecha,
    cuentaPrincipal,
    descripcion,
    total, // Debe
    0,     // Haber
    uuid
  ]);

  // 2. Abono a Ventas por el subtotal
  poliza.push([
    fecha,
    cuentaVentas,
    descripcion,
    0,        // Debe
    subtotal, // Haber
    uuid
  ]);

  // 3. Abono a IVA Trasladado (si existe)
  if (impuestosTrasladados > 0) {
    poliza.push([
      fecha,
      cuentaIvaTrasladado,
      descripcion,
      0,                       // Debe
      impuestosTrasladados,    // Haber
      uuid
    ]);
  }

  return poliza;
}

/**
 * Genera una póliza de diario para un CFDI de tipo Egreso.
 * Utiliza el catálogo de proveedores para asignar la cuenta de gasto correcta.
 * @param {object} cfdiData Datos del CFDI parseado.
 * @param {Map<string, string>} catalogoProveedores Mapa del catálogo de proveedores (RFC -> Cuenta Contable).
 * @returns {Array<Array<any>>} Un array de filas que representan los movimientos de la póliza.
 */
function generarPolizaEgreso(cfdiData, catalogoProveedores) {
  const { fecha, uuid, total, subtotal, impuestosTrasladados, metodoPago, emisorRfc, emisorNombre } = cfdiData;
  const descripcion = `Egreso según factura de ${emisorNombre}`;

  // Buscar la cuenta del proveedor en el catálogo. Si no se encuentra, usar la de gastos generales.
  const cuentaGasto = catalogoProveedores.get(emisorRfc) || CUENTAS_PREDETERMINADAS.GASTOS_GENERALES;

  // Se usan las cuentas del objeto de configuración.
  const cuentaPrincipal = metodoPago === 'PUE' ? CUENTAS_PREDETERMINADAS.BANCOS : CUENTAS_PREDETERMINADAS.PROVEEDORES;
  const cuentaIvaAcreditable = CUENTAS_PREDETERMINADAS.IVA_ACREDITABLE;

  const poliza = [];

  // 1. Cargo a la cuenta de Gasto/Costo por el subtotal
  poliza.push([
    fecha,
    cuentaGasto,
    descripcion,
    subtotal, // Debe
    0,        // Haber
    uuid
  ]);

  // 2. Cargo a IVA Acreditable (si existe)
  if (impuestosTrasladados > 0) {
    poliza.push([
      fecha,
      cuentaIvaAcreditable,
      descripcion,
      impuestosTrasladados, // Debe
      0,                    // Haber
      uuid
    ]);
  }

  // 3. Abono a Bancos o Proveedores por el total
  poliza.push([
    fecha,
    cuentaPrincipal,
    descripcion,
    0,     // Debe
    total, // Haber
    uuid
  ]);

  return poliza;
}
