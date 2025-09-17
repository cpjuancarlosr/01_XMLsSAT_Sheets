const SAT_NS = {
  SOAP: 'http://schemas.xmlsoap.org/soap/envelope/',
  DES: 'http://DescargaMasivaTerceros.sat.gob.mx',
  WSSE: 'http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd',
  WSU: 'http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd',
  DS: 'http://www.w3.org/2000/09/xmldsig#',
};

const SAT_CFDI_HEADERS = [
  'UUID',
  'Fecha',
  'Tipo',
  'Origen',
  'Estado Comprobante',
  'RFC Emisor',
  'Nombre Emisor',
  'RFC Receptor',
  'Nombre Receptor',
  'Moneda',
  'SubTotal',
  'Total',
  'Forma Pago',
  'Método Pago',
  'Uso CFDI',
  'Serie',
  'Folio',
  'Total Impuestos Trasladados',
  'Total Impuestos Retenidos',
  'Id Paquete',
  'Archivo XML',
  'Código Descarga',
  'Mensaje Descarga',
];

const SAT_ENDPOINTS = {
  PRODUCTION: {
    autenticacion: 'https://cfdidescargamasivasolicitud.clouda.sat.gob.mx/Autenticacion/Autenticacion.svc',
    solicita: 'https://cfdidescargamasivasolicitud.clouda.sat.gob.mx/SolicitaDescargaService.svc',
    verifica: 'https://cfdidescargamasivasolicitud.clouda.sat.gob.mx/VerificaSolicitudDescargaService.svc',
    descarga: 'https://cfdidescargamasiva.clouda.sat.gob.mx/DescargaMasivaService.svc',
    soapAction: {
      autentica: 'http://DescargaMasivaTerceros.gob.mx/IAutenticacion/Autentica',
      solicitaEmitidos: 'http://DescargaMasivaTerceros.sat.gob.mx/ISolicitaDescargaService/SolicitaDescargaEmitidos',
      solicitaRecibidos: 'http://DescargaMasivaTerceros.sat.gob.mx/ISolicitaDescargaService/SolicitaDescargaRecibidos',
      verifica: 'http://DescargaMasivaTerceros.sat.gob.mx/IVerificaSolicitudDescargaService/VerificaSolicitudDescarga',
      descarga: 'http://DescargaMasivaTerceros.sat.gob.mx/IDescargaMasivaTercerosService/Descargar',
    },
  },
  TEST: {
    autenticacion: 'https://pruebassolicituddescargamasivasolicitud.clouda.sat.gob.mx/Autenticacion/Autenticacion.svc',
    solicita: 'https://pruebassolicituddescargamasivasolicitud.clouda.sat.gob.mx/SolicitaDescargaService.svc',
    verifica: 'https://pruebassolicituddescargamasivasolicitud.clouda.sat.gob.mx/VerificaSolicitudDescargaService.svc',
    descarga: 'https://pruebassolicituddescargamasiva.clouda.sat.gob.mx/DescargaMasivaService.svc',
    soapAction: {
      autentica: 'http://DescargaMasivaTerceros.gob.mx/IAutenticacion/Autentica',
      solicitaEmitidos: 'http://DescargaMasivaTerceros.sat.gob.mx/ISolicitaDescargaService/SolicitaDescargaEmitidos',
      solicitaRecibidos: 'http://DescargaMasivaTerceros.sat.gob.mx/ISolicitaDescargaService/SolicitaDescargaRecibidos',
      verifica: 'http://DescargaMasivaTerceros.sat.gob.mx/IVerificaSolicitudDescargaService/VerificaSolicitudDescarga',
      descarga: 'http://DescargaMasivaTerceros.sat.gob.mx/IDescargaMasivaTercerosService/Descargar',
    },
  },
};

const DEFAULT_OPTIONS = {
  startDate: '',
  endDate: '',
  tipoConsulta: 'emitidos',
  estadoComprobante: 'Vigente',
  tipoSolicitud: 'CFDI',
  tipoComprobante: '',
  rfcEmisor: '',
  rfcReceptor: '',
  rfcReceptores: [],
  rfcACuentaTerceros: '',
  complemento: '',
  skipDriveUpload: false,
};

function getScriptConfig() {
  const props = PropertiesService.getScriptProperties();
  const cfg = {
    certificatePem: props.getProperty('FIEL_CERT_PEM') || '',
    privateKeyPem: props.getProperty('FIEL_PRIVATE_KEY_PEM') || '',
    certificateIssuer: props.getProperty('FIEL_ISSUER_NAME') || '',
    certificateSerial: props.getProperty('FIEL_CERT_SERIAL') || '',
    rfc: props.getProperty('FIEL_RFC') || '',
    spreadsheetId: props.getProperty('GOOGLE_SHEETS_ID') || '',
    driveFolderId: props.getProperty('GOOGLE_DRIVE_FOLDER_ID') || '',
    environment: (props.getProperty('SAT_ENVIRONMENT') || 'PRODUCTION').toUpperCase(),
    pollIntervalSeconds: parseInt(props.getProperty('SAT_POLL_INTERVAL_SECONDS') || '60', 10),
    maxWaitMinutes: parseInt(props.getProperty('SAT_MAX_WAIT_MINUTES') || '30', 10),
    defaultOptions: props.getProperty('SAT_DEFAULT_OPTIONS'),
  };
  if (cfg.defaultOptions) {
    try {
      const parsed = JSON.parse(cfg.defaultOptions);
      cfg.defaultOptions = Object.assign({}, DEFAULT_OPTIONS, parsed);
    } catch (err) {
      throw new Error('SAT_DEFAULT_OPTIONS tiene JSON inválido: ' + err.message);
    }
  } else {
    cfg.defaultOptions = DEFAULT_OPTIONS;
  }
  return cfg;
}

function assertConfig(config) {
  const missing = [];
  if (!config.certificatePem) missing.push('FIEL_CERT_PEM');
  if (!config.privateKeyPem) missing.push('FIEL_PRIVATE_KEY_PEM');
  if (!config.certificateIssuer) missing.push('FIEL_ISSUER_NAME');
  if (!config.certificateSerial) missing.push('FIEL_CERT_SERIAL');
  if (!config.rfc) missing.push('FIEL_RFC');
  if (!config.spreadsheetId) missing.push('GOOGLE_SHEETS_ID');
  if (missing.length) {
    throw new Error('Faltan propiedades de script obligatorias: ' + missing.join(', '));
  }
  if (!SAT_ENDPOINTS[config.environment]) {
    throw new Error('SAT_ENVIRONMENT debe ser PRODUCTION o TEST');
  }
}

function runDescargaMasivaCfdi() {
  const config = getScriptConfig();
  assertConfig(config);
  descargarCfdiMasivo(config.defaultOptions);
}

function descargarCfdiMasivo(userOptions) {
  const config = getScriptConfig();
  assertConfig(config);
  const options = mergeOptions(config.defaultOptions, userOptions || {});
  validarOpciones(options);

  const token = obtenerToken(config);
  const solicitud = solicitarDescarga(config, options, token);
  const requestId = solicitud.IdSolicitud;
  const status = esperarPaquetes(config, requestId, token);
  const paquetes = [];
  const folder = options.skipDriveUpload ? null : prepararCarpetaDrive(config, options, requestId);

  status.IdsPaquetes.forEach(function (paqueteId) {
    const paquete = descargarPaquete(config, paqueteId, token);
    if (folder) {
      guardarZipEnDrive(folder, paqueteId, paquete.zipBytes);
    }
    paquetes.push({
      id: paqueteId,
      zipBytes: paquete.zipBytes,
      metadata: paquete.metadata,
    });
  });

  const registros = [];
  paquetes.forEach(function (paquete) {
    const paqueteRegistros = parsearPaquete(paquete, options, solicitud.EstadoComprobante);
    Array.prototype.push.apply(registros, paqueteRegistros);
  });

  const agrupados = agruparPorHoja(registros);
  escribirEnSheets(config.spreadsheetId, agrupados);
  return {
    requestId: requestId,
    paquetes: paquetes.map(function (p) { return p.id; }),
    hojasActualizadas: Object.keys(agrupados),
  };
}

function mergeOptions(base, override) {
  const resultado = Object.assign({}, base);
  Object.keys(override).forEach(function (key) {
    if (override[key] !== undefined && override[key] !== null && override[key] !== '') {
      resultado[key] = override[key];
    }
  });
  if (override && Array.isArray(override.rfcReceptores)) {
    resultado.rfcReceptores = override.rfcReceptores.slice();
  }
  return resultado;
}

function validarOpciones(options) {
  if (!options.startDate || !options.endDate) {
    throw new Error('Debes proporcionar startDate y endDate en formato YYYY-MM-DD');
  }
  if (!['emitidos', 'recibidos'].includes(options.tipoConsulta.toLowerCase())) {
    throw new Error('tipoConsulta debe ser "emitidos" o "recibidos"');
  }
  if (options.rfcReceptores && options.rfcReceptores.length > 5) {
    throw new Error('El SAT sólo permite hasta 5 RFC receptores por solicitud');
  }
}

function obtenerToken(config) {
  const cacheKey = 'sat_token_' + config.environment;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    const data = JSON.parse(cached);
    if (Date.now() < data.expiresAt - 30000) {
      return data.token;
    }
  }
  const envelope = construirSolicitudAutenticacion(config);
  const response = llamarSat(config, {
    url: SAT_ENDPOINTS[config.environment].autenticacion,
    soapAction: SAT_ENDPOINTS[config.environment].soapAction.autentica,
    payload: envelope,
    authorization: null,
  });
  const parsed = XmlService.parse(response);
  const namespace = XmlService.getNamespace('s', SAT_NS.SOAP);
  const body = parsed.getRootElement().getChild('Body', namespace);
  const header = parsed.getRootElement().getChild('Header', namespace);
  if (!body) {
    throw new Error('Respuesta de Autenticación inválida');
  }
  const authResult = buscarDescendiente(body, 'AutenticaResult');
  if (!authResult) {
    throw new Error('No se encontró AutenticaResult en la respuesta de autenticación');
  }
  const token = authResult.getText();
  const timestampNode = header ? buscarDescendiente(header, 'Timestamp') : null;
  var expiresAt = Date.now() + 5 * 60 * 1000;
  if (timestampNode) {
    const expiresNode = buscarDescendiente(timestampNode, 'Expires');
    if (expiresNode) {
      expiresAt = Date.parse(expiresNode.getText()) || expiresAt;
    }
  }
  cache.put(cacheKey, JSON.stringify({ token: token, expiresAt: expiresAt }), 280);
  return token;
}

function solicitarDescarga(config, options, token) {
  const payload = construirSolicitudDescarga(config, options);
  const response = llamarSat(config, {
    url: SAT_ENDPOINTS[config.environment].solicita,
    soapAction: options.tipoConsulta.toLowerCase() === 'emitidos'
      ? SAT_ENDPOINTS[config.environment].soapAction.solicitaEmitidos
      : SAT_ENDPOINTS[config.environment].soapAction.solicitaRecibidos,
    payload: payload,
    authorization: token,
  });
  const parsed = XmlService.parse(response);
  const resultado = buscarDescendiente(parsed.getRootElement(), 'SolicitaDescargaEmitidosResult')
    || buscarDescendiente(parsed.getRootElement(), 'SolicitaDescargaRecibidosResult');
  if (!resultado) {
    throw new Error('No se pudo interpretar la respuesta de solicitud de descarga');
  }
  const atributos = extraerAtributos(resultado);
  if (atributos.CodEstatus && atributos.CodEstatus !== '5000') {
    throw new Error('SAT rechazó la solicitud: ' + atributos.CodEstatus + ' ' + (atributos.Mensaje || ''));
  }
  if (!atributos.IdSolicitud) {
    throw new Error('La respuesta de solicitud no incluye IdSolicitud');
  }
  atributos.EstadoComprobante = options.estadoComprobante || null;
  return atributos;
}

function esperarPaquetes(config, requestId, token) {
  const deadline = Date.now() + config.maxWaitMinutes * 60 * 1000;
  while (true) {
    const status = verificarSolicitud(config, requestId, token);
    const estadoSolicitud = parseInt(status.EstadoSolicitud, 10);
    if (estadoSolicitud === 3) {
      if (!status.IdsPaquetes || status.IdsPaquetes.length === 0) {
        throw new Error('La solicitud se marcó terminada pero sin paquetes disponibles');
      }
      return status;
    }
    if ([4, 5, 6].indexOf(estadoSolicitud) !== -1) {
      throw new Error('La solicitud ' + requestId + ' falló con estado ' + estadoSolicitud + ': ' + (status.Mensaje || ''));
    }
    if (Date.now() > deadline) {
      throw new Error('La solicitud ' + requestId + ' no concluyó en el tiempo máximo configurado');
    }
    Utilities.sleep(Math.max(1, config.pollIntervalSeconds) * 1000);
  }
}

function verificarSolicitud(config, requestId, token) {
  const payload = construirVerificacion(config, requestId);
  const response = llamarSat(config, {
    url: SAT_ENDPOINTS[config.environment].verifica,
    soapAction: SAT_ENDPOINTS[config.environment].soapAction.verifica,
    payload: payload,
    authorization: token,
  });
  const parsed = XmlService.parse(response);
  const resultado = buscarDescendiente(parsed.getRootElement(), 'VerificaSolicitudDescargaResult');
  if (!resultado) {
    throw new Error('Respuesta de verificación inválida');
  }
  const atributos = extraerAtributos(resultado);
  const ids = resultado.getChildren('IdsPaquetes', resultado.getNamespace()).map(function (node) {
    return node.getText();
  });
  atributos.IdsPaquetes = ids;
  return atributos;
}

function descargarPaquete(config, packageId, token) {
  const payload = construirDescarga(config, packageId);
  const response = llamarSat(config, {
    url: SAT_ENDPOINTS[config.environment].descarga,
    soapAction: SAT_ENDPOINTS[config.environment].soapAction.descarga,
    payload: payload,
    authorization: token,
  });
  const parsed = XmlService.parse(response);
  const paqueteNode = buscarDescendiente(parsed.getRootElement(), 'Paquete');
  if (!paqueteNode) {
    throw new Error('Respuesta de descarga sin elemento Paquete');
  }
  const headerRespuesta = buscarDescendiente(parsed.getRootElement(), 'respuesta');
  const metadata = headerRespuesta ? extraerAtributos(headerRespuesta) : {};
  const zipBytes = Utilities.base64Decode(paqueteNode.getText());
  return { zipBytes: zipBytes, metadata: metadata };
}

function llamarSat(config, request) {
  const headers = {
    'Content-Type': 'text/xml; charset=utf-8',
    'Accept': 'text/xml',
    'Cache-Control': 'no-cache',
    'SOAPAction': request.soapAction,
  };
  if (request.authorization) {
    headers.Authorization = 'WRAP access_token="' + request.authorization + '"';
  }
  const respuesta = UrlFetchApp.fetch(request.url, {
    method: 'post',
    headers: headers,
    muteHttpExceptions: true,
    payload: request.payload,
  });
  const codigo = respuesta.getResponseCode();
  if (codigo < 200 || codigo >= 300) {
    throw new Error('Error HTTP ' + codigo + ' al invocar ' + request.url + ': ' + respuesta.getContentText());
  }
  return respuesta.getContentText();
}

function construirSolicitudAutenticacion(config) {
  const created = formatoFechaIso(new Date());
  const expires = formatoFechaIso(new Date(Date.now() + 5 * 60 * 1000));
  const certBase64 = limpiarCertificado(config.certificatePem);
  const timestampC14n = [
    '<u:Timestamp xmlns:u="', SAT_NS.WSU, '" u:Id="_0">',
    '<u:Created>', created, '</u:Created>',
    '<u:Expires>', expires, '</u:Expires>',
    '</u:Timestamp>'
  ].join('');
  const digestValue = sha1DigestBase64(timestampC14n);
  const signedInfo = construirSignedInfoAutenticacion(digestValue);
  const signatureValue = firmarSha1(signedInfo, config.privateKeyPem);
  const signatureXml = [
    '<Signature xmlns="', SAT_NS.DS, '">',
    signedInfo,
    '<SignatureValue>', signatureValue, '</SignatureValue>',
    '<KeyInfo><o:SecurityTokenReference xmlns:o="', SAT_NS.WSSE, '">',
    '<o:Reference ValueType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-x509-token-profile-1.0#X509v3" URI="#BinarySecurityToken"/>',
    '</o:SecurityTokenReference></KeyInfo>',
    '</Signature>'
  ].join('');
  return [
    '<s:Envelope xmlns:s="', SAT_NS.SOAP, '" xmlns:o="', SAT_NS.WSSE, '" xmlns:u="', SAT_NS.WSU, '">',
    '<s:Header>',
    '<o:Security s:mustUnderstand="1">',
    timestampC14n,
    '<o:BinarySecurityToken u:Id="BinarySecurityToken" ValueType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-x509-token-profile-1.0#X509v3" EncodingType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-soap-message-security-1.0#Base64Binary">',
    certBase64,
    '</o:BinarySecurityToken>',
    signatureXml,
    '</o:Security>',
    '</s:Header>',
    '<s:Body><Autentica xmlns="http://DescargaMasivaTerceros.gob.mx"/></s:Body>',
    '</s:Envelope>'
  ].join('');
}

function construirSignedInfoAutenticacion(digestValue) {
  return [
    '<SignedInfo xmlns="', SAT_NS.DS, '">',
    '<CanonicalizationMethod Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"/>',
    '<SignatureMethod Algorithm="http://www.w3.org/2000/09/xmldsig#rsa-sha1"/>',
    '<Reference URI="#_0">',
    '<Transforms><Transform Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"/></Transforms>',
    '<DigestMethod Algorithm="http://www.w3.org/2000/09/xmldsig#sha1"/>',
    '<DigestValue>', digestValue, '</DigestValue>',
    '</Reference>',
    '</SignedInfo>'
  ].join('');
}

function construirSolicitudDescarga(config, options) {
  const nsDes = ' xmlns:des="' + SAT_NS.DES + '"';
  const atributos = construirAtributosSolicitud(config, options);
  const hijos = construirElementosHijoSolicitud(options);
  const solicitudXml = hijos
    ? ['<des:solicitud', atributos, '>', hijos, '</des:solicitud>'].join('')
    : ['<des:solicitud', atributos, '/>'].join('');
  const operacion = options.tipoConsulta.toLowerCase() === 'emitidos'
    ? 'SolicitaDescargaEmitidos'
    : 'SolicitaDescargaRecibidos';
  const contenidoSinFirma = ['<des:', operacion, nsDes, '>', solicitudXml, '</des:', operacion, '>'].join('');
  const digestValue = sha1DigestBase64(contenidoSinFirma);
  const signedInfo = construirSignedInfoPeticion(digestValue);
  const signatureValue = firmarSha1(signedInfo, config.privateKeyPem);
  const keyInfo = construirKeyInfo(config);
  const signatureXml = ['<Signature xmlns="', SAT_NS.DS, '">', signedInfo, '<SignatureValue>', signatureValue, '</SignatureValue>', keyInfo, '</Signature>'].join('');
  const contenidoConFirma = ['<des:', operacion, nsDes, '>', solicitudXml, signatureXml, '</des:', operacion, '>'].join('');
  return ['<s:Envelope xmlns:s="', SAT_NS.SOAP, '"><s:Header/>', '<s:Body>', contenidoConFirma, '</s:Body></s:Envelope>'].join('');
}

function construirAtributosSolicitud(config, options) {
  const attrs = {
    FechaInicial: options.startDate,
    FechaFinal: options.endDate,
    RfcSolicitante: config.rfc,
    TipoSolicitud: options.tipoSolicitud,
    EstadoComprobante: options.estadoComprobante || null,
    TipoComprobante: options.tipoComprobante || null,
    RfcACuentaTerceros: options.rfcACuentaTerceros || null,
    Complemento: options.complemento || null,
  };
  const tipoConsulta = (options.tipoConsulta || '').toLowerCase();
  if (tipoConsulta === 'emitidos') {
    attrs.RfcEmisor = options.rfcEmisor || config.rfc;
    if (options.rfcReceptor) {
      attrs.RfcReceptor = options.rfcReceptor;
    }
    if (options.rfcReceptores && options.rfcReceptores.length) {
      attrs.RfcReceptor = null;
    }
  } else {
    attrs.RfcReceptor = options.rfcReceptor || config.rfc;
    if (options.rfcEmisor) {
      attrs.RfcEmisor = options.rfcEmisor;
    }
    if (options.rfcReceptores && options.rfcReceptores.length) {
      attrs.RfcReceptor = null;
    }
  }
  const entries = Object.keys(attrs).filter(function (key) {
    return attrs[key] !== undefined && attrs[key] !== null && attrs[key] !== '';
  }).map(function (key) {
    return [key, attrs[key]];
  });
  entries.sort(function (a, b) {
    return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0;
  });
  return entries.map(function (entry) {
    return ' ' + entry[0] + '="' + escaparXml(entry[1]) + '"';
  }).join('');
}

function construirElementosHijoSolicitud(options) {
  const partes = [];
  const receptores = Array.isArray(options.rfcReceptores) ? options.rfcReceptores.filter(function (r) { return r; }) : [];
  if (receptores.length) {
    const nodos = receptores.map(function (rfc) {
      return '<des:RfcReceptor>' + escaparXml(rfc) + '</des:RfcReceptor>';
    }).join('');
    partes.push('<des:RfcReceptores>' + nodos + '</des:RfcReceptores>');
  }
  return partes.join('');
}

function construirSignedInfoPeticion(digestValue) {
  return [
    '<SignedInfo xmlns="', SAT_NS.DS, '">',
    '<CanonicalizationMethod Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>',
    '<SignatureMethod Algorithm="http://www.w3.org/2000/09/xmldsig#rsa-sha1"/>',
    '<Reference URI="">',
    '<Transforms><Transform Algorithm="http://www.w3.org/2000/09/xmldsig#enveloped-signature"/></Transforms>',
    '<DigestMethod Algorithm="http://www.w3.org/2000/09/xmldsig#sha1"/>',
    '<DigestValue>', digestValue, '</DigestValue>',
    '</Reference>',
    '</SignedInfo>'
  ].join('');
}

function construirKeyInfo(config) {
  const certificado = limpiarCertificado(config.certificatePem);
  const issuer = escaparXml(config.certificateIssuer);
  const serial = normalizarSerial(config.certificateSerial);
  return [
    '<KeyInfo><X509Data>',
    '<X509IssuerSerial>',
    '<X509IssuerName>', issuer, '</X509IssuerName>',
    '<X509SerialNumber>', serial, '</X509SerialNumber>',
    '</X509IssuerSerial>',
    '<X509Certificate>', certificado, '</X509Certificate>',
    '</X509Data></KeyInfo>'
  ].join('');
}

function construirVerificacion(config, requestId) {
  const solicitud = ['<des:VerificaSolicitudDescarga xmlns:des="', SAT_NS.DES, '">',
    '<des:solicitud RfcSolicitante="', escaparXml(config.rfc), '" IdSolicitud="', escaparXml(requestId), '"/>',
    '</des:VerificaSolicitudDescarga>'].join('');
  return ['<s:Envelope xmlns:s="', SAT_NS.SOAP, '"><s:Header/>', '<s:Body>', solicitud, '</s:Body></s:Envelope>'].join('');
}

function construirDescarga(config, packageId) {
  const solicitud = ['<des:PeticionDescargaMasivaTercerosEntrada xmlns:des="', SAT_NS.DES, '">',
    '<des:peticionDescarga RfcSolicitante="', escaparXml(config.rfc), '" IdPaquete="', escaparXml(packageId), '"/>',
    '</des:PeticionDescargaMasivaTercerosEntrada>'].join('');
  return ['<s:Envelope xmlns:s="', SAT_NS.SOAP, '"><s:Header/>', '<s:Body>', solicitud, '</s:Body></s:Envelope>'].join('');
}

function prepararCarpetaDrive(config, options, requestId) {
  if (!config.driveFolderId) {
    return null;
  }
  const carpetaRaiz = DriveApp.getFolderById(config.driveFolderId);
  const nombre = ['CFDI', options.tipoConsulta.toUpperCase(), options.startDate, options.endDate, requestId].join('_');
  return carpetaRaiz.createFolder(nombre);
}

function guardarZipEnDrive(folder, packageId, zipBytes) {
  const blob = Utilities.newBlob(zipBytes, 'application/zip', packageId + '.zip');
  folder.createFile(blob);
}

function parsearPaquete(paquete, options, estadoComprobante) {
  const resultados = [];
  const zipBlob = Utilities.newBlob(paquete.zipBytes, 'application/zip', paquete.id + '.zip');
  const blobs = Utilities.unzip(zipBlob);
  blobs.forEach(function (blob) {
    const nombre = blob.getName();
    if (!nombre || nombre.toLowerCase().indexOf('.xml') === -1) {
      return;
    }
    const xmlTexto = blob.getDataAsString('UTF-8');
    try {
      const registro = interpretarCfdi(xmlTexto, {
        archivoXml: nombre,
        paqueteId: paquete.id,
        origen: options.tipoConsulta,
        estadoComprobante: estadoComprobante,
        codigoDescarga: paquete.metadata.CodEstatus || null,
        mensajeDescarga: paquete.metadata.Mensaje || null,
      });
      resultados.push(registro);
    } catch (err) {
      console.error('No se pudo interpretar ' + nombre + ': ' + err.message);
    }
  });
  return resultados;
}

function interpretarCfdi(xmlTexto, metadata) {
  const documento = XmlService.parse(xmlTexto);
  const root = documento.getRootElement();
  const tipoRaw = obtenerAtributo(root, ['TipoDeComprobante', 'TipoDeComprob']);
  const tipo = normalizarTipo(tipoRaw || (root.getName().indexOf('Retenciones') !== -1 ? 'RET' : 'SIN-TIPO'));
  const fechaStr = obtenerAtributo(root, ['Fecha', 'FechaExp']);
  if (!fechaStr) {
    throw new Error('El CFDI no tiene atributo Fecha');
  }
  const fecha = new Date(fechaStr);
  const emisor = extraerEmisor(root);
  const receptor = extraerReceptor(root);
  const impuestos = extraerImpuestos(root);
  const totales = extraerTotales(root);
  const uuid = extraerUuid(root);
  return {
    uuid: uuid,
    fecha: fecha,
    tipo: tipo,
    origen: metadata.origen,
    estadoComprobante: metadata.estadoComprobante || null,
    emisorRfc: emisor.rfc,
    emisorNombre: emisor.nombre,
    receptorRfc: receptor.rfc,
    receptorNombre: receptor.nombre,
    moneda: obtenerAtributo(root, ['Moneda']) || null,
    subtotal: totales.subtotal,
    total: totales.total,
    formaPago: obtenerAtributo(root, ['FormaPago']) || null,
    metodoPago: obtenerAtributo(root, ['MetodoPago']) || null,
    usoCfdi: receptor.usoCfdi,
    serie: obtenerAtributo(root, ['Serie']) || null,
    folio: obtenerAtributo(root, ['Folio']) || null,
    impuestosTrasladados: impuestos.trasladados,
    impuestosRetenidos: impuestos.retenidos,
    paqueteId: metadata.paqueteId,
    archivoXml: metadata.archivoXml,
    codigoDescarga: metadata.codigoDescarga,
    mensajeDescarga: metadata.mensajeDescarga,
  };
}

function agruparPorHoja(registros) {
  const resultado = {};
  registros.forEach(function (registro) {
    const year = registro.fecha.getUTCFullYear();
    const month = registro.fecha.getUTCMonth() + 1;
    const nombre = [registro.tipo, '-', pad(year, 4), '-', pad(month, 2)].join(' ');
    if (!resultado[nombre]) {
      resultado[nombre] = [];
    }
    resultado[nombre].push(registro);
  });
  Object.keys(resultado).forEach(function (nombre) {
    resultado[nombre].sort(function (a, b) { return a.fecha - b.fecha; });
  });
  return resultado;
}

function escribirEnSheets(spreadsheetId, agrupados) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  Object.keys(agrupados).forEach(function (nombreHoja) {
    var hoja = spreadsheet.getSheetByName(nombreHoja);
    if (!hoja) {
      hoja = spreadsheet.insertSheet(nombreHoja);
    } else {
      hoja.clearContents();
    }
    const datos = agrupados[nombreHoja].map(function (registro) {
      return [
        registro.uuid,
        formatoFechaRegistro(registro.fecha),
        registro.tipo,
        registro.origen,
        registro.estadoComprobante,
        registro.emisorRfc,
        registro.emisorNombre,
        registro.receptorRfc,
        registro.receptorNombre,
        registro.moneda,
        registro.subtotal,
        registro.total,
        registro.formaPago,
        registro.metodoPago,
        registro.usoCfdi,
        registro.serie,
        registro.folio,
        registro.impuestosTrasladados,
        registro.impuestosRetenidos,
        registro.paqueteId,
        registro.archivoXml,
        registro.codigoDescarga,
        registro.mensajeDescarga,
      ];
    });
    const totalFilas = datos.length + 1;
    hoja.getRange(1, 1, 1, SAT_CFDI_HEADERS.length).setValues([SAT_CFDI_HEADERS]);
    if (datos.length) {
      hoja.getRange(2, 1, datos.length, SAT_CFDI_HEADERS.length).setValues(datos);
    }
    hoja.autoResizeColumns(1, SAT_CFDI_HEADERS.length);
    hoja.getRange(1, 1, totalFilas, SAT_CFDI_HEADERS.length).setWrap(false);
  });
}

function formatoFechaRegistro(fecha) {
  return Utilities.formatDate(fecha, 'GMT', "yyyy-MM-dd'T'HH:mm:ss");
}

function formatoFechaIso(fecha) {
  return Utilities.formatDate(fecha, 'GMT', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
}

function pad(number, length) {
  let str = String(number);
  while (str.length < length) {
    str = '0' + str;
  }
  return str;
}

function sha1DigestBase64(contenido) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, contenido, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(digest);
}

function firmarSha1(contenido, privateKeyPem) {
  const firma = Utilities.computeRsaSha1Signature(contenido, privateKeyPem);
  return Utilities.base64Encode(firma);
}

function limpiarCertificado(pem) {
  return pem.replace(/-----BEGIN CERTIFICATE-----/, '')
    .replace(/-----END CERTIFICATE-----/, '')
    .replace(/\s+/g, '');
}

function escaparXml(valor) {
  return String(valor)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buscarDescendiente(elemento, nombreLocal) {
  if (elemento.getName && elemento.getName() === nombreLocal) {
    return elemento;
  }
  const hijos = elemento.getChildren();
  for (var i = 0; i < hijos.length; i++) {
    const encontrado = buscarDescendiente(hijos[i], nombreLocal);
    if (encontrado) {
      return encontrado;
    }
  }
  return null;
}

function extraerAtributos(elemento) {
  const attrs = {};
  elemento.getAttributes().forEach(function (attr) {
    attrs[attr.getName()] = attr.getValue();
  });
  return attrs;
}

function obtenerAtributo(elemento, nombres) {
  if (!elemento) {
    return null;
  }
  for (var i = 0; i < nombres.length; i++) {
    const nombre = nombres[i];
    const attr = elemento.getAttribute(nombre);
    if (attr) {
      return attr.getValue();
    }
  }
  return null;
}

function extraerEmisor(root) {
  const emisor = root.getChildren().filter(function (child) {
    return child.getName() === 'Emisor';
  })[0];
  if (!emisor) {
    return { rfc: '', nombre: null };
  }
  const rfc = obtenerAtributo(emisor, ['Rfc', 'RfcE', 'RfcEmisor']) || '';
  const nombre = obtenerAtributo(emisor, ['Nombre', 'NomDenRazSocE']) || null;
  return { rfc: rfc, nombre: nombre };
}

function extraerReceptor(root) {
  const receptor = root.getChildren().filter(function (child) {
    return child.getName() === 'Receptor';
  })[0];
  if (!receptor) {
    return { rfc: null, nombre: null, usoCfdi: null };
  }
  const nacional = receptor.getChildren().filter(function (child) { return child.getName() === 'Nacional'; })[0];
  const extranjero = receptor.getChildren().filter(function (child) { return child.getName() === 'Extranjero'; })[0];
  if (nacional) {
    return {
      rfc: obtenerAtributo(nacional, ['RFCRecep', 'RfcR']) || obtenerAtributo(receptor, ['RfcR', 'RFCRecep']) || null,
      nombre: obtenerAtributo(nacional, ['Nombre', 'NomDenRazSocR']) || obtenerAtributo(receptor, ['NomDenRazSocR', 'Nombre']) || null,
      usoCfdi: obtenerAtributo(receptor, ['UsoCFDI']) || null,
    };
  }
  if (extranjero) {
    return {
      rfc: obtenerAtributo(extranjero, ['NumRegIdTrib']) || obtenerAtributo(receptor, ['NumRegIdTrib']) || null,
      nombre: obtenerAtributo(extranjero, ['Nombre', 'NomDenRazSocR']) || obtenerAtributo(receptor, ['NomDenRazSocR']) || null,
      usoCfdi: null,
    };
  }
  return {
    rfc: obtenerAtributo(receptor, ['Rfc', 'RfcR', 'RFCRecep']) || null,
    nombre: obtenerAtributo(receptor, ['Nombre', 'NomDenRazSocR']) || null,
    usoCfdi: obtenerAtributo(receptor, ['UsoCFDI']) || null,
  };
}

function extraerImpuestos(root) {
  const impuestos = root.getChildren().filter(function (child) { return child.getName() === 'Impuestos'; })[0];
  if (!impuestos) {
    return { trasladados: null, retenidos: null };
  }
  return {
    trasladados: obtenerAtributo(impuestos, ['TotalImpuestosTrasladados']) || null,
    retenidos: obtenerAtributo(impuestos, ['TotalImpuestosRetenidos']) || null,
  };
}

function extraerTotales(root) {
  const subtotal = obtenerAtributo(root, ['SubTotal', 'SubTot']) || null;
  var total = obtenerAtributo(root, ['Total', 'MontoTotOperacion']);
  if (!total) {
    const totales = root.getChildren().filter(function (child) { return child.getName() === 'Totales'; })[0];
    if (totales) {
      total = obtenerAtributo(totales, ['MontoTotOperacion', 'MontoTotGrav', 'MontoTotal']) || null;
    }
  }
  return { subtotal: subtotal, total: total };
}

function extraerUuid(root) {
  const complemento = root.getChildren().filter(function (child) { return child.getName() === 'Complemento'; })[0];
  if (!complemento) {
    return '';
  }
  for (var i = 0; i < complemento.getChildren().length; i++) {
    const hijo = complemento.getChildren()[i];
    if (hijo.getName() === 'TimbreFiscalDigital') {
      const uuid = obtenerAtributo(hijo, ['UUID']);
      return uuid || '';
    }
  }
  return '';
}

function normalizarTipo(tipo) {
  if (!tipo) {
    return 'SIN-TIPO';
  }
  var limpio = String(tipo).split('-')[0].trim().toUpperCase();
  return limpio || 'SIN-TIPO';
}

function normalizarSerial(serial) {
  if (!serial) {
    throw new Error('FIEL_CERT_SERIAL no puede estar vacío');
  }
  const limpio = serial.replace(/[^0-9A-Fa-f]/g, '');
  if (!limpio) {
    throw new Error('FIEL_CERT_SERIAL inválido');
  }
  if (/^[0-9]+$/.test(limpio)) {
    return limpio;
  }
  return BigInt('0x' + limpio).toString(10);
}
