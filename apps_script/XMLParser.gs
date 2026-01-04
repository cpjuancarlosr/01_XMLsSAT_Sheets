/**
 * Parsea el contenido de un string XML de un CFDI 4.0 y extrae los datos más importantes.
 * Distingue entre facturas (Ingreso/Egreso) y complementos de pago.
 * @param {string} xmlTexto El contenido del archivo XML.
 * @returns {object} Un objeto con los datos extraídos del CFDI.
 */
function parseCfdi(xmlTexto) {
  const documento = XmlService.parse(xmlTexto);
  const root = documento.getRootElement();

  const nsCfdi = root.getNamespace();
  const nsTfd = XmlService.getNamespace('tfd', 'http://www.sat.gob.mx/TimbreFiscalDigital');

  const tipoComprobante = obtenerAtributo(root, ['TipoDeComprobante']);

  if (tipoComprobante === 'P') {
    return parseComplementoPago(root, nsCfdi, nsTfd);
  } else {
    return parseFactura(root, nsCfdi, nsTfd);
  }
}

/**
 * Parsea un CFDI de tipo Ingreso o Egreso.
 */
function parseFactura(root, nsCfdi, nsTfd) {
  const props = PropertiesService.getScriptProperties();
  const rfcPropio = props.getProperty('RFC_PROPIO');
  if (!rfcPropio) {
    throw new Error('La propiedad "RFC_PROPIO" no está configurada.');
  }

  const emisor = extraerEmisor(root, nsCfdi);
  const receptor = extraerReceptor(root, nsCfdi);

  const tipoContable = (receptor.rfc.toUpperCase() === rfcPropio.toUpperCase()) ? 'Egreso' : 'Ingreso';

  return {
    uuid: extraerUuid(root, nsTfd),
    fecha: new Date(obtenerAtributo(root, ['Fecha'])),
    tipoComprobante: obtenerAtributo(root, ['TipoDeComprobante']),
    tipoContable: tipoContable,
    metodoPago: obtenerAtributo(root, ['MetodoPago']),
    formaPago: obtenerAtributo(root, ['FormaPago']),
    emisorRfc: emisor.rfc,
    emisorNombre: emisor.nombre,
    receptorRfc: receptor.rfc,
    receptorNombre: receptor.nombre,
    subtotal: parseFloat(obtenerAtributo(root, ['SubTotal']) || 0),
    total: parseFloat(obtenerAtributo(root, ['Total']) || 0),
    impuestosTrasladados: parseFloat(extraerTotalImpuestos(root, nsCfdi, 'Traslados') || 0),
    impuestosRetenidos: parseFloat(extraerTotalImpuestos(root, nsCfdi, 'Retenciones') || 0),
    conceptos: extraerConceptos(root, nsCfdi),
  };
}

/**
 * Parsea un CFDI de tipo Pago (Complemento de Pago).
 */
function parseComplementoPago(root, nsCfdi, nsTfd) {
  const nsPago = XmlService.getNamespace('pago20', 'http://www.sat.gob.mx/Pagos20');
  const complemento = root.getChild('Complemento', nsCfdi);
  const pago = complemento.getChild('Pago', nsPago);

  const documentosRelacionados = [];
  const doctoRelacionadoNodes = pago.getChildren('DoctoRelacionado', nsPago);

  doctoRelacionadoNodes.forEach(nodo => {
    documentosRelacionados.push({
      uuidRelacionado: obtenerAtributo(nodo, ['IdDocumento']),
      montoPagado: parseFloat(obtenerAtributo(nodo, ['ImpPagado']) || 0),
      parcialidad: parseInt(obtenerAtributo(nodo, ['NumParcialidad']) || 1),
    });
  });

  return {
    uuid: extraerUuid(root, nsTfd),
    fecha: new Date(obtenerAtributo(pago, ['FechaPago'])),
    tipoComprobante: 'P',
    montoTotalPago: parseFloat(obtenerAtributo(pago, ['Monto']) || 0),
    documentosRelacionados: documentosRelacionados,
  };
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
  return { rfc: obtenerAtributo(emisor, ['Rfc']), nombre: obtenerAtributo(emisor, ['Nombre']) };
}

function extraerReceptor(root, ns) {
  const receptor = root.getChild('Receptor', ns);
  return { rfc: obtenerAtributo(receptor, ['Rfc']), nombre: obtenerAtributo(receptor, ['Nombre']) };
}

function extraerUuid(root, ns) {
  const complemento = root.getChild('Complemento', root.getNamespace());
  if (!complemento) return null;
  const timbre = complemento.getChild('TimbreFiscalDigital', ns);
  if (!timbre) return null;
  return obtenerAtributo(timbre, ['UUID']);
}

function extraerTotalImpuestos(root, ns, tipoImpuesto) { // tipoImpuesto: "Traslados" o "Retenciones"
  const impuestosNode = root.getChild('Impuestos', ns);
  if (!impuestosNode) return 0;
  return obtenerAtributo(impuestosNode, [`TotalImpuestos${tipoImpuesto}`]);
}

function extraerConceptos(root, ns) {
  const conceptos = [];
  const conceptosNode = root.getChild('Conceptos', ns);
  if (conceptosNode) {
    const listaConceptos = conceptosNode.getChildren('Concepto', ns);
    listaConceptos.forEach(concepto => {
      conceptos.push({
        claveProdServ: obtenerAtributo(concepto, ['ClaveProdServ']),
        importe: parseFloat(obtenerAtributo(concepto, ['Importe']) || 0),
      });
    });
  }
  return conceptos;
}
