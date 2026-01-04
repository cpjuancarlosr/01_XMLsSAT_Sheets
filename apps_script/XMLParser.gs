/**
 * Parsea el contenido de un string XML de un CFDI y extrae los datos más importantes.
 * @param {string} xmlTexto El contenido del archivo XML.
 * @returns {object} Un objeto con los datos extraídos del CFDI.
 */
function parseCfdi(xmlTexto) {
  const documento = XmlService.parse(xmlTexto);
  const root = documento.getRootElement();

  // Namespace para el Timbre Fiscal Digital
  const nsCfdi = root.getNamespace();
  const nsTfd = XmlService.getNamespace('tfd', 'http://www.sat.gob.mx/TimbreFiscalDigital');

  const tipoRaw = obtenerAtributo(root, ['TipoDeComprobante']);
  const tipo = normalizarTipo(tipoRaw);

  const fechaStr = obtenerAtributo(root, ['Fecha']);
  if (!fechaStr) {
    throw new Error('El CFDI no tiene atributo Fecha');
  }
  const fecha = new Date(fechaStr);

  const emisor = extraerEmisor(root, nsCfdi);
  const receptor = extraerReceptor(root, nsCfdi);
  const impuestos = extraerImpuestos(root, nsCfdi);
  const totales = extraerTotales(root);
  const uuid = extraerUuid(root, nsTfd);

  // Determinar si es Ingreso o Egreso para nuestra contabilidad
  // Esto se basa en quién es el receptor. Se debe configurar el RFC de la propia empresa.
  const props = PropertiesService.getScriptProperties();
  const rfcPropio = props.getProperty('RFC_PROPIO');
  if (!rfcPropio) {
    throw new Error('La propiedad "RFC_PROPIO" no está configurada en las propiedades del script.');
  }
  const tipoContable = (receptor.rfc.toUpperCase() === rfcPropio.toUpperCase()) ? 'Egreso' : 'Ingreso';

  return {
    uuid: uuid,
    fecha: fecha,
    tipoComprobante: tipo, // I, E, T, P, N
    tipoContable: tipoContable, // Ingreso, Egreso
    emisorRfc: emisor.rfc,
    emisorNombre: emisor.nombre,
    receptorRfc: receptor.rfc,
    receptorNombre: receptor.nombre,
    moneda: obtenerAtributo(root, ['Moneda']),
    subtotal: parseFloat(totales.subtotal || 0),
    total: parseFloat(totales.total || 0),
    formaPago: obtenerAtributo(root, ['FormaPago']),
    metodoPago: obtenerAtributo(root, ['MetodoPago']),
    impuestosTrasladados: parseFloat(impuestos.trasladados || 0),
    impuestosRetenidos: parseFloat(impuestos.retenidos || 0),
  };
}

/**
 * Normaliza el tipo de comprobante a un valor conocido (I, E, T, P, N).
 * @param {string} tipo El valor del atributo TipoDeComprobante.
 * @returns {string} El tipo normalizado.
 */
function normalizarTipo(tipo) {
  if (!tipo) return 'N/A';
  const upper = tipo.toUpperCase().trim();
  switch (upper) {
    case 'I':
    case 'INGRESO':
      return 'I';
    case 'E':
    case 'EGRESO':
      return 'E';
    case 'T':
    case 'TRASLADO':
      return 'T';
    case 'P':
    case 'PAGO':
      return 'P';
    case 'N':
    case 'NOMINA':
      return 'N';
    default:
      return upper;
  }
}

// --- FUNCIONES AUXILIARES DE PARSEO ---

function obtenerAtributo(elemento, nombres) {
  if (!elemento) return null;
  for (let i = 0; i < nombres.length; i++) {
    const attr = elemento.getAttribute(nombres[i]);
    if (attr) return attr.getValue();
  }
  return null;
}

function extraerEmisor(root, ns) {
  const emisor = root.getChild('Emisor', ns);
  if (!emisor) return { rfc: null, nombre: null };
  return {
    rfc: obtenerAtributo(emisor, ['Rfc']),
    nombre: obtenerAtributo(emisor, ['Nombre']),
  };
}

function extraerReceptor(root, ns) {
  const receptor = root.getChild('Receptor', ns);
  if (!receptor) return { rfc: null, nombre: null };
  return {
    rfc: obtenerAtributo(receptor, ['Rfc']),
    nombre: obtenerAtributo(receptor, ['Nombre']),
  };
}

function extraerImpuestos(root, ns) {
  const impuestosNode = root.getChild('Impuestos', ns);
  if (!impuestosNode) return { trasladados: 0, retenidos: 0 };

  // Ahora, en lugar de buscar una tasa específica, usamos el total de impuestos trasladados
  // que el propio emisor del CFDI ha calculado. Esto es más robusto.
  const totalTraslados = parseFloat(obtenerAtributo(impuestosNode, ['TotalImpuestosTrasladados']) || 0);

  return {
    trasladados: totalTraslados,
    retenidos: parseFloat(obtenerAtributo(impuestosNode, ['TotalImpuestosRetenidos']) || 0),
  };
}


function extraerTotales(root) {
  return {
    subtotal: obtenerAtributo(root, ['SubTotal']),
    total: obtenerAtributo(root, ['Total']),
  };
}

function extraerUuid(root, ns) {
  const complemento = root.getChild('Complemento', root.getNamespace());
  if (!complemento) return null;
  const timbre = complemento.getChild('TimbreFiscalDigital', ns);
  if (!timbre) return null;
  return obtenerAtributo(timbre, ['UUID']);
}
