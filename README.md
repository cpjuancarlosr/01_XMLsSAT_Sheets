# Automatización de descarga masiva de CFDI con Google Apps Script

Este repositorio entrega un script de **Google Apps Script** que permite:

1. Autenticarse en la Descarga Masiva de CFDI del SAT utilizando una FIEL.
2. Solicitar paquetes por rango de fechas, tipo de comprobante y estado (vigente, cancelado, todos).
3. Esperar la generación de los paquetes, descargarlos y (opcionalmente) guardarlos en Google Drive.
4. Descomprimir y parsear cada XML para escribirlo en pestañas de Google Sheets organizadas por tipo, año y mes (`I - 2024 - 05`, `E - 2023 - 12`, etc.).

El código reside en `apps_script/Code.gs` y puede pegarse directamente en un proyecto de Apps Script vinculado a la hoja de cálculo donde quieras recibir los datos.

## Preparación previa

### 1. Convierte tu FIEL a formatos compatibles

La API del SAT requiere la FIEL en PEM sin contraseña. Usa OpenSSL para preparar los archivos:

```bash
# Certificado (.cer) a PEM
openssl x509 -inform DER -in FIEL.cer -out FIEL_cert.pem

# Llave privada (.key) a PEM sin contraseña
openssl pkcs8 -inform DER -in FIEL.key -out FIEL_key.pem -passin pass:CONTRASENA

# Número de serie del certificado (en hexadecimal)
openssl x509 -inform DER -in FIEL.cer -serial -noout

# Nombre del emisor formateado RFC2253
openssl x509 -inform DER -in FIEL.cer -issuer -nameopt RFC2253 -noout
```

Guarda el contenido completo de `FIEL_cert.pem` y `FIEL_key.pem`. El número de serie puede pegarse tal cual (con o sin dos puntos); el script lo normaliza a decimal automáticamente.

### 2. Crea un proyecto de Apps Script

1. Abre la hoja de cálculo de destino en Google Sheets.
2. Ve a **Extensiones → Apps Script** y elimina cualquier código existente.
3. Copia y pega el contenido de `apps_script/Code.gs` en el editor y guarda los cambios.

### 3. Configura las propiedades del proyecto

En **Project Settings → Script properties** agrega las siguientes claves (todas como texto):

| Propiedad | Descripción |
|-----------|-------------|
| `FIEL_CERT_PEM` | Contenido completo del certificado en PEM (incluye encabezados `-----BEGIN CERTIFICATE-----`). |
| `FIEL_PRIVATE_KEY_PEM` | Contenido completo de la llave privada en PEM sin contraseña. |
| `FIEL_ISSUER_NAME` | Nombre del emisor obtenido con `openssl … -issuer -nameopt RFC2253`. |
| `FIEL_CERT_SERIAL` | Número de serie del certificado (hexadecimal o decimal). |
| `FIEL_RFC` | RFC del titular de la FIEL. |
| `GOOGLE_SHEETS_ID` | ID de la hoja donde se escribirán los datos (la misma en la que pegas el script). |
| `GOOGLE_DRIVE_FOLDER_ID` | (Opcional) ID de la carpeta raíz en Drive para guardar los ZIP descargados. |
| `SAT_ENVIRONMENT` | `PRODUCTION` o `TEST`. |
| `SAT_POLL_INTERVAL_SECONDS` | (Opcional) Segundos entre verificaciones al SAT. Predeterminado: 60. |
| `SAT_MAX_WAIT_MINUTES` | (Opcional) Tiempo máximo de espera antes de abortar. Predeterminado: 30 minutos. |
| `SAT_DEFAULT_OPTIONS` | (Opcional) JSON con los parámetros de consulta por defecto (ver ejemplo). |

### Ejemplo de `SAT_DEFAULT_OPTIONS`

```json
{
  "startDate": "2024-01-01",
  "endDate": "2024-01-31",
  "tipoConsulta": "emitidos",
  "estadoComprobante": "Vigente",
  "tipoSolicitud": "CFDI",
  "tipoComprobante": "I",
  "rfcEmisor": "AAA010101AAA",
  "rfcReceptores": ["BBB010101BBB"],
  "skipDriveUpload": false
}
```

Puedes sobrescribir cualquiera de estos valores cuando invoques el script manualmente.

## Uso del script

El archivo define dos funciones principales:

- `runDescargaMasivaCfdi()` ejecuta la descarga usando los valores definidos en `SAT_DEFAULT_OPTIONS`.
- `descargarCfdiMasivo(options)` permite pasar un objeto con parámetros específicos. Los campos soportados son:
  - `startDate` y `endDate` (formato `YYYY-MM-DD`).
  - `tipoConsulta`: `emitidos` o `recibidos`.
  - `estadoComprobante`: `Vigente`, `Cancelado` o `Todos`.
  - `tipoSolicitud`: `CFDI`, `Metadata`, etc. (valor del SAT).
  - `tipoComprobante`: opcional (`I`, `E`, `N`, `P`, etc.).
  - `rfcEmisor`, `rfcReceptor` (un solo RFC) y/o `rfcReceptores` (hasta 5 RFC), `rfcACuentaTerceros`, `complemento`.
  - `skipDriveUpload`: `true` para omitir el guardado de ZIP en Drive.

Ejemplo de ejecución manual desde el editor de Apps Script:

```javascript
function ejemplo() {
  descargarCfdiMasivo({
    startDate: '2024-04-01',
    endDate: '2024-04-30',
    tipoConsulta: 'recibidos',
    estadoComprobante: 'Vigente',
    rfcReceptores: ['AAA010101AAA', 'BBB010101BBB'],
    skipDriveUpload: true,
  });
}
```

La función devuelve un objeto con el `requestId`, la lista de `paquetes` descargados y las `hojasActualizadas`. Los detalles adicionales aparecen en el **Registro de ejecución** de Apps Script.

## Estructura de datos en Google Sheets

Cada pestaña creada sigue el formato `<Tipo> - <Año> - <Mes>` y contiene las columnas:

1. UUID
2. Fecha (UTC)
3. Tipo
4. Origen (`emitidos` o `recibidos`)
5. Estado del comprobante solicitado
6. RFC del emisor
7. Nombre del emisor
8. RFC del receptor
9. Nombre del receptor
10. Moneda
11. Subtotal
12. Total
13. Forma de pago
14. Método de pago
15. Uso CFDI
16. Serie
17. Folio
18. Total de impuestos trasladados
19. Total de impuestos retenidos
20. Id del paquete
21. Nombre del XML
22. Código de respuesta del SAT al descargar el paquete
23. Mensaje del SAT

Las columnas se autoajustan para facilitar la lectura.

## Consideraciones adicionales

- Asegúrate de que la FIEL esté vigente y que la cuenta tenga permisos para la Descarga Masiva.
- El SAT limita las solicitudes por día y por rango de fechas; revisa la documentación oficial para evitar bloqueos.
- Apps Script corre en la nube de Google: evita compartir la hoja con usuarios que no deban acceder a tu FIEL.
- Si necesitas otro flujo (por ejemplo, ejecutar desde un botón o un menú personalizado), puedes invocar `descargarCfdiMasivo` desde la UI de Google Sheets.

## Licencia

El código se entrega "tal cual" para uso educativo o personal. Ajusta la lógica según las políticas y obligaciones fiscales aplicables.
